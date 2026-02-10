import os
from datetime import timedelta, date
from pathlib import Path

from django import forms
from django.conf import settings
from django.contrib import admin, messages
from django.core.mail import EmailMultiAlternatives
from django.db.models import Q
from django.http import FileResponse
from django.urls import path, reverse
from django.utils import timezone
from django.utils.html import format_html
from django.template import Template, Context

from .models import SalesLead, SalesLeadImport, KpTemplate, LeadNote
from .services.import_sales_leads_xlsx import import_sales_leads_xlsx


# =========================
# ПОДСКАЗКИ ИСТОЧНИКОВ
# =========================
SOURCE_SUGGESTIONS = [
    "HH.ru",
    "Avito",
    "Сайт",
    "VK",
    "Telegram",
    "2ГИС",
    "Рекомендация",
    "Холодный обзвон",
    "Email-рассылка",
]


class SalesLeadAdminForm(forms.ModelForm):
    class Meta:
        model = SalesLead
        fields = "__all__"
        widgets = {
            "source": forms.TextInput(
                attrs={
                    "list": "source_suggestions",
                    "placeholder": "Выбери или напиши источник",
                    "style": "max-width: 320px;",
                }
            )
        }


# =========================
# ФИЛЬТРЫ
# =========================
class ReadyForKPFilter(admin.SimpleListFilter):
    title = "Готовы к КП"
    parameter_name = "ready_kp"

    def lookups(self, request, model_admin):
        return (("1", "Есть email + статус Новый"),)

    def queryset(self, request, queryset):
        if self.value() == "1":
            return queryset.filter(status="new").exclude(email__isnull=True).exclude(email="")
        return queryset


class EmailFilter(admin.SimpleListFilter):
    title = "Email"
    parameter_name = "email_state"

    def lookups(self, request, model_admin):
        return (
            ("with", "Показать с email"),
            ("without", "Показать без email"),
        )

    def queryset(self, request, queryset):
        if self.value() == "with":
            return queryset.exclude(email__isnull=True).exclude(email="")
        if self.value() == "without":
            return queryset.filter(Q(email__isnull=True) | Q(email=""))
        return queryset


class NoPhoneFilter(admin.SimpleListFilter):
    title = "Нет телефона"
    parameter_name = "no_phone"

    def lookups(self, request, model_admin):
        return (("1", "Без телефона"),)

    def queryset(self, request, queryset):
        if self.value() == "1":
            return queryset.filter(Q(phone__isnull=True) | Q(phone=""))
        return queryset


class AvitoTodayFilter(admin.SimpleListFilter):
    title = "Avito — сегодня"
    parameter_name = "avito_today"

    def lookups(self, request, model_admin):
        return (("today", "Лиды Avito за сегодня"),)

    def queryset(self, request, queryset):
        if self.value() == "today":
            return queryset.filter(source__iexact="Avito", created_at__date=date.today())
        return queryset


class KPSentNoReply3DaysFilter(admin.SimpleListFilter):
    title = "КП без ответа"
    parameter_name = "kp_no_reply"

    def lookups(self, request, model_admin):
        return (("3days", "КП → нет ответа 3 дня"),)

    def queryset(self, request, queryset):
        if self.value() == "3days":
            border = timezone.now() - timedelta(days=3)
            return queryset.filter(status="kp_sent", kp_sent_at__lte=border)
        return queryset


class LeadHasReminderTodayFilter(admin.SimpleListFilter):
    title = "Заметки/напоминания"
    parameter_name = "lead_reminders"

    def lookups(self, request, model_admin):
        return (
            ("today", "Есть напоминание на сегодня"),
            ("overdue", "Есть просроченное напоминание"),
        )

    def queryset(self, request, queryset):
        today = timezone.localdate()
        now = timezone.now()

        if self.value() == "today":
            return queryset.filter(
                notes__is_done=False,
                notes__remind_at__date=today,
            ).distinct()

        if self.value() == "overdue":
            return queryset.filter(
                notes__is_done=False,
                notes__remind_at__lt=now,
            ).distinct()

        return queryset


# =========================
# ШАБЛОНЫ КП
# =========================
@admin.register(KpTemplate)
class KpTemplateAdmin(admin.ModelAdmin):
    list_display = ("name", "is_active", "updated_at")
    list_editable = ("is_active",)
    ordering = ("-is_active", "-updated_at")


def render_template(text: str, lead: SalesLead, manager_name: str) -> str:
    ctx = Context({
        "vacancy": lead.vacancy or "",
        "city": lead.city or "",
        "company": lead.company_name or "",
        "manager": manager_name,
        "email": lead.email or "",
        "phone": lead.phone or "",
        "ad_url": getattr(lead, "ad_url", "") or "",
        "source": lead.source or "",
    })
    return Template(text).render(ctx)


# =========================
# ✅ ЗАМЕТКИ/КАЛЕНДАРЬ (INLINE)
# =========================
class LeadNoteInline(admin.TabularInline):
    model = LeadNote
    extra = 1
    fields = ("title", "text", "due_at", "remind_at", "is_done")
    show_change_link = True


@admin.register(LeadNote)
class LeadNoteAdmin(admin.ModelAdmin):
    list_display = ("lead", "title", "remind_at", "due_at", "is_done", "created_at")
    list_filter = ("is_done",)
    search_fields = ("title", "text", "lead__company_name", "lead__vacancy", "lead__email")
    readonly_fields = ("created_at",)

    actions = ("mark_done",)

    @admin.action(description="Отметить выполненным")
    def mark_done(self, request, queryset):
        updated = queryset.update(is_done=True)
        self.message_user(request, f"Готово. Выполнено: {updated}", level=messages.SUCCESS)


# =========================
# ИМПОРТ ЛИДОВ
# =========================
@admin.register(SalesLeadImport)
class SalesLeadImportAdmin(admin.ModelAdmin):
    list_display = ("id", "created_at", "excel_file", "template_link")
    readonly_fields = ("created_at",)
    fields = ("created_at", "excel_file")

    actions = ("run_import_now",)

    def template_link(self, obj):
        url = reverse("admin:download_avito_template")
        return format_html('<a href="{}">Скачать шаблон Avito</a>', url)

    template_link.short_description = "Шаблон"

    @admin.action(description="Загрузить лиды в «Лиды продаж»")
    def run_import_now(self, request, queryset):
        total_created = total_dup = total_bad = 0

        for obj in queryset:
            if not obj.excel_file:
                continue
            res = import_sales_leads_xlsx(obj.excel_file.path)
            total_created += res.get("created", 0)
            total_dup += res.get("skipped_dup", 0)
            total_bad += res.get("skipped_bad", 0)

        self.message_user(
            request,
            f"Готово. Создано: {total_created}, дубли: {total_dup}, битые строки: {total_bad}",
            level=messages.SUCCESS,
        )

    def get_urls(self):
        urls = super().get_urls()
        custom = [
            path(
                "download-avito-template/",
                self.admin_site.admin_view(self.download_avito_template),
                name="download_avito_template",
            ),
        ]
        return custom + urls

    def download_avito_template(self, request):
        template_path = Path(settings.BASE_DIR) / "app" / "static" / "templates" / "avito.xlsx"
        return FileResponse(open(template_path, "rb"), as_attachment=True, filename="avito_template.xlsx")


# =========================
# ЛИДЫ ПРОДАЖ
# =========================
@admin.register(SalesLead)
class SalesLeadAdmin(admin.ModelAdmin):
    form = SalesLeadAdminForm
    inlines = [LeadNoteInline]

    list_display = (
        "vacancy",
        "city",
        "source",
        "status",
        "manager",
        "next_reminder",
        "fill_contacts",
        "open_ad",
        "kp_sent_at",
        "created_at",
    )

    list_filter = (
        "source",
        "status",
        "city",
        "manager",
        ReadyForKPFilter,
        EmailFilter,
        NoPhoneFilter,
        AvitoTodayFilter,
        KPSentNoReply3DaysFilter,
        LeadHasReminderTodayFilter,
    )

    search_fields = ("vacancy", "city", "company_name", "email", "phone", "ad_url")
    readonly_fields = ("created_at", "kp_sent_at")

    actions = ("send_kp",)

    def fill_contacts(self, obj):
        url = reverse("admin:app_saleslead_change", args=[obj.id])
        return format_html('<a href="{}">✏️ Контакты</a>', url)

    fill_contacts.short_description = "Контакты"

    def open_ad(self, obj):
        return format_html('<a href="{}" target="_blank">Открыть</a>', obj.ad_url) if obj.ad_url else "-"

    open_ad.short_description = "Объявление"

    def next_reminder(self, obj):
        # ближайшее напоминание (не выполнено)
        note = obj.notes.filter(is_done=False, remind_at__isnull=False).order_by("remind_at").first()
        if not note:
            return "-"
        return timezone.localtime(note.remind_at).strftime("%d.%m.%Y %H:%M")

    next_reminder.short_description = "Ближайшее напоминание"

    @admin.action(description="Отправить КП и отметить «КП отправлено»")
    def send_kp(self, request, queryset):
        tpl = KpTemplate.objects.filter(is_active=True).order_by("-updated_at").first()

        sent = skipped_no_email = skipped_already = errors = 0

        for lead in queryset.select_related("manager"):
            if not lead.email:
                skipped_no_email += 1
                continue

            if lead.kp_sent_at:
                skipped_already += 1
                continue

            manager_name = lead.manager.get_full_name() if lead.manager else "Менеджер"

            try:
                subject = tpl.subject if tpl else "Kommercheskoe predlozhenie (KP)"
                html_body = render_template(tpl.body_html, lead, manager_name) if tpl else ""
                text_body = render_template(tpl.body_text, lead, manager_name) if tpl else ""

                if not text_body:
                    text_body = "Здравствуйте! Направляем коммерческое предложение. Подробности в письме."

                msg = EmailMultiAlternatives(
                    subject=subject,
                    body=text_body,
                    from_email=settings.EMAIL_HOST_USER,
                    to=[lead.email],
                )
                msg.encoding = "utf-8"

                if html_body:
                    msg.attach_alternative(html_body, "text/html")

                # Вложение Word — опционально (если поле есть и файл загружен)
                if getattr(tpl, "kp_docx", None) and tpl.kp_docx:
                    msg.attach_file(tpl.kp_docx.path)

                msg.send(fail_silently=False)

                lead.status = "kp_sent"
                lead.kp_sent_at = timezone.now()
                lead.save(update_fields=["status", "kp_sent_at"])
                sent += 1

            except Exception as e:
                errors += 1
                self.message_user(request, f"Ошибка отправки на {lead.email}: {e}", level=messages.ERROR)

        self.message_user(
            request,
            f"КП отправлено: {sent}, без email: {skipped_no_email}, уже отправляли: {skipped_already}, ошибок: {errors}",
            level=messages.SUCCESS if errors == 0 else messages.WARNING,
        )
