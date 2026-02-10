from decimal import Decimal

from django.conf import settings
from django.db import models
from django.core.exceptions import ValidationError
from django.utils import timezone


# ==========================
# СТАРЫЕ МОДЕЛИ (оставляем)
# ==========================

class Request(models.Model):
    received_at = models.DateTimeField(auto_now_add=True, verbose_name="Дата получения")
    source_email = models.EmailField(blank=True, default="", verbose_name="Почта отправителя")
    source_type = models.CharField(
        max_length=30,
        choices=[("merch", "Выкладка"), ("cleaning", "Клининг")],
        default="merch",
        verbose_name="Тип заявки",
    )
    file_hash = models.CharField(max_length=64, blank=True, default="", verbose_name="Хэш файла")
    excel_file = models.FileField(upload_to="requests/", verbose_name="Excel-файл заявки")

    class Meta:
        verbose_name = "Заявка"
        verbose_name_plural = "Заявки"


class Store(models.Model):
    city = models.CharField(max_length=100, verbose_name="Город")
    address = models.TextField(verbose_name="Адрес (без города)")
    address_raw = models.TextField(verbose_name="Адрес объекта (как в Excel)")

    status = models.CharField(
        max_length=10,
        choices=[("open", "Открыт"), ("closed", "Закрыт")],
        default="open",
        verbose_name="Статус",
    )

    assigned_employee = models.OneToOneField(
        "Employee",
        null=True,
        blank=True,
        on_delete=models.SET_NULL,
        verbose_name="Сотрудник",
        related_name="store",
    )

    current_hours_merch = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        default=Decimal("0.00"),
        verbose_name="Часы (текущие) выкладка",
    )
    current_hours_cleaning = models.DecimalField(
        max_digits=10,
        decimal_places=2,
        default=Decimal("0.00"),
        verbose_name="Часы (текущие) клининг",
    )

    class Meta:
        unique_together = ("city", "address")
        verbose_name = "Магазин"
        verbose_name_plural = "Магазины"

    def __str__(self):
        return f"{self.city}, {self.address}"


class Position(models.Model):
    code = models.CharField(
        max_length=50,
        unique=True,
        choices=[
            ("hall_worker", "Работник торгового зала"),
            ("cleaner", "Уборщица"),
        ],
        verbose_name="Должность",
    )

    class Meta:
        verbose_name = "Должность"
        verbose_name_plural = "Должности"

    def __str__(self):
        return self.get_code_display()


class Employee(models.Model):
    full_name = models.CharField(max_length=255, verbose_name="ФИО")

    full_name_norm = models.CharField(
        max_length=255,
        db_index=True,
        blank=True,
        default="",
        verbose_name="ФИО (нормализ.)",
    )

    email = models.EmailField(unique=True, verbose_name="Email")

    positions = models.ManyToManyField(
        Position,
        blank=True,
        verbose_name="Должности",
    )

    card_number = models.CharField(max_length=30, verbose_name="Номер карты")
    account_number = models.CharField(max_length=30, verbose_name="Номер счета")
    bik = models.CharField(max_length=15, verbose_name="БИК")
    bank_name = models.CharField(max_length=255, verbose_name="Банк")

    is_active = models.BooleanField(default=True, verbose_name="Активен")

    class Meta:
        verbose_name = "Сотрудник"
        verbose_name_plural = "Сотрудники"

    def save(self, *args, **kwargs):
        self.full_name_norm = (self.full_name or "").casefold()
        super().save(*args, **kwargs)

    def __str__(self):
        return self.full_name


class StoreShift(models.Model):
    store = models.ForeignKey("Store", on_delete=models.CASCADE, related_name="shifts", verbose_name="Магазин")
    date = models.DateField(db_index=True, verbose_name="Дата")
    service_type = models.CharField(
        max_length=30,
        choices=[("cleaning", "Клининг"), ("merch", "Выкладка")],
        verbose_name="Услуга",
    )
    employee = models.ForeignKey("Employee", null=True, blank=True, on_delete=models.SET_NULL, verbose_name="Сотрудник")
    hours = models.DecimalField(max_digits=6, decimal_places=2, default=Decimal("0.00"), verbose_name="Часы")
    comment = models.CharField(max_length=255, blank=True, default="", verbose_name="Комментарий")

    class Meta:
        verbose_name = "Смена"
        verbose_name_plural = "Смены"
        unique_together = ("store", "date", "service_type")
        indexes = [
            models.Index(fields=["date", "service_type"]),
            models.Index(fields=["employee", "date"]),
        ]

    def clean(self):
        super().clean()

        if self.employee and not self.employee.is_active:
            raise ValidationError({"employee": "Нельзя назначить неактивного сотрудника."})

        if self.employee:
            need = "cleaner" if self.service_type == "cleaning" else "hall_worker"
            if not self.employee.positions.filter(code=need).exists():
                raise ValidationError({"employee": f"Неверная должность для услуги {self.get_service_type_display()}."})

    def __str__(self):
        return f"{self.store} • {self.date} • {self.get_service_type_display()}"


class RequestLine(models.Model):
    request = models.ForeignKey(Request, on_delete=models.CASCADE, related_name="lines", verbose_name="Заявка")
    store = models.ForeignKey(Store, on_delete=models.CASCADE, verbose_name="Магазин")

    service_type = models.CharField(max_length=255, verbose_name="Вид оказываемых услуг")
    hours = models.DecimalField(max_digits=6, decimal_places=2, verbose_name="Часы")

    assigned_employee = models.ForeignKey(
        Employee, null=True, blank=True, on_delete=models.SET_NULL, verbose_name="Назначенный сотрудник"
    )

    row_hash = models.CharField(max_length=64, verbose_name="Хэш строки")

    def required_position(self) -> str:
        if self.request and self.request.source_type == "cleaning":
            return "cleaner"
        return "hall_worker"

    def clean(self):
        super().clean()

        emp = self.assigned_employee
        if not emp:
            return

        if not emp.is_active:
            raise ValidationError({"assigned_employee": "Нельзя назначить неактивного сотрудника."})

        need = self.required_position()
        if not emp.positions.filter(code=need).exists():
            position_labels = dict(Position._meta.get_field("code").choices)
            raise ValidationError({"assigned_employee": f"Неверная должность. Нужно: {position_labels[need]}."})

    class Meta:
        unique_together = ("request", "row_hash")
        verbose_name = "Строка заявки"
        verbose_name_plural = "Строки заявки"


class Reconciliation(models.Model):
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="Дата создания")

    file_left = models.FileField(upload_to="reconciliation/", verbose_name="Сверка от Заказчика")
    file_right = models.FileField(upload_to="reconciliation/", verbose_name="Выгрузка с базы")

    report_file = models.FileField(
        upload_to="reconciliation/reports/",
        null=True,
        blank=True,
        verbose_name="Отчёт (Excel)",
    )

    status = models.CharField(
        max_length=20,
        choices=[
            ("uploaded", "Загружено"),
            ("processing", "В обработке"),
            ("done", "Готово"),
            ("error", "Ошибка"),
        ],
        default="uploaded",
        verbose_name="Статус",
    )

    result = models.TextField(blank=True, default="", verbose_name="Результат сверки")

    class Meta:
        verbose_name = "Сверка"
        verbose_name_plural = "Сверка"

    def __str__(self):
        return f"Сверка #{self.pk} от {self.created_at:%d.%m.%Y %H:%M}"


# ==========================
# ПРОДАЖИ (CRM)
# ==========================

class SalesLead(models.Model):
    STATUS_CHOICES = [
        ("new", "Новый"),
        ("kp_sent", "КП отправлено"),
        ("reply", "Ответ получен"),
        ("negotiation", "Переговоры"),
        ("agreement", "Согласование условий"),
        ("deal", "Договор / запуск"),
        ("rejected", "Отказ / не актуально"),
    ]

    company_name = models.CharField(max_length=255, blank=True, default="", verbose_name="Компания")

    source = models.CharField(
        max_length=50,
        blank=True,
        default="",
        verbose_name="Источник",
        help_text="Например: HH.ru, Avito, VK, 2ГИС, Сайт, Рекомендация",
    )

    ad_url = models.CharField(
        max_length=700,
        blank=True,
        default="",
        verbose_name="Ссылка на объявление",
    )

    city = models.CharField(max_length=100, blank=True, default="", verbose_name="Город")

    email = models.EmailField(blank=True, default="", verbose_name="Email")
    phone = models.CharField(max_length=50, blank=True, default="", verbose_name="Телефон")

    vacancy = models.CharField(max_length=255, blank=True, default="", verbose_name="Вакансия")

    # не обязательное
    work_types = models.JSONField(blank=True, null=True, default=list, verbose_name="Тип работ")

    staff_count = models.PositiveIntegerField(null=True, blank=True, verbose_name="Кол-во персонала")

    comment = models.TextField(blank=True, default="", verbose_name="Комментарий")

    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default="new", verbose_name="Статус")

    manager = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        null=True,
        blank=True,
        on_delete=models.SET_NULL,
        related_name="sales_leads",
        verbose_name="Менеджер",
    )

    kp_sent_at = models.DateTimeField(null=True, blank=True, verbose_name="Дата отправки КП")

    created_at = models.DateTimeField(auto_now_add=True, verbose_name="Дата создания")

    class Meta:
        verbose_name = "Лид продаж"
        verbose_name_plural = "Лиды продаж"
        ordering = ["-created_at"]

    def __str__(self):
        return f"{self.company_name or self.vacancy or 'Лид'} ({self.city or '—'})"


class SalesLeadImport(models.Model):
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="Дата загрузки")
    excel_file = models.FileField(upload_to="sales_imports/", verbose_name="Excel/CSV файл лидов")

    class Meta:
        verbose_name = "Импорт лидов"
        verbose_name_plural = "Импорт лидов"

    def __str__(self):
        return f"Импорт лидов #{self.pk} от {self.created_at:%d.%m.%Y %H:%M}"


class SalesRoundRobinState(models.Model):
    """Состояние распределения по кругу."""
    last_manager = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        null=True,
        blank=True,
        on_delete=models.SET_NULL,
        verbose_name="Последний менеджер",
        related_name="round_robin_last_manager",
    )
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Распределение (состояние)"
        verbose_name_plural = "Распределение (состояние)"

    def __str__(self):
        return f"Последний менеджер: {self.last_manager_id or '-'}"


class KpTemplate(models.Model):
    name = models.CharField(max_length=120, default="Основной шаблон", verbose_name="Название")
    is_active = models.BooleanField(default=True, verbose_name="Активный шаблон")

    subject = models.CharField(
        max_length=200,
        default="Kommercheskoe predlozhenie (KP)",
        verbose_name="Тема письма",
    )

    body_text = models.TextField(
        blank=True,
        default="",
        verbose_name="Текст письма (TEXT)",
        help_text="Можно оставить пустым — тогда будет отправляться только HTML.",
    )

    body_html = models.TextField(
        blank=True,
        default="",
        verbose_name="Текст письма (HTML)",
        help_text="Переменные: {{vacancy}}, {{city}}, {{company}}, {{manager}}, {{email}}, {{phone}}, {{ad_url}}, {{source}}",
    )

    kp_docx = models.FileField(
        upload_to="kp_templates/",
        blank=True,
        null=True,
        verbose_name="КП (Word .docx)",
        help_text="Если загрузишь .docx — CRM может прикреплять его как вложение (опционально).",
    )

    short_text = models.TextField(
        blank=True,
        default="",
        verbose_name="Короткий текст письма (в тело)",
        help_text="Короткое сообщение в теле письма (если нужно).",
    )

    updated_at = models.DateTimeField(auto_now=True, verbose_name="Обновлено")

    class Meta:
        verbose_name = "Шаблон КП"
        verbose_name_plural = "Шаблоны КП"

    def __str__(self):
        return f"{self.name} ({'активный' if self.is_active else 'выкл'})"


# ✅ КАЛЕНДАРЬ / ЗАМЕТКИ / НАПОМИНАНИЯ (по-русски)
class LeadNote(models.Model):
    lead = models.ForeignKey(
        SalesLead,
        on_delete=models.CASCADE,
        related_name="notes",
        verbose_name="Лид",
    )

    title = models.CharField(
        "Заголовок",
        max_length=200,
        blank=True,
        default="",
    )

    text = models.TextField(
        "Текст заметки",
        blank=True,
        default="",
    )

    due_at = models.DateTimeField(
        "Дата выполнения",
        null=True,
        blank=True,
    )

    remind_at = models.DateTimeField(
        "Напомнить",
        null=True,
        blank=True,
    )

    is_done = models.BooleanField(
        "Выполнено",
        default=False,
    )

    reminded_at = models.DateTimeField(
        "Напоминание отправлено",
        null=True,
        blank=True,
    )

    author = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        verbose_name="Автор",
        null=True,
        blank=True,
        on_delete=models.SET_NULL,
    )

    created_at = models.DateTimeField(
        "Создано",
        auto_now_add=True,
    )

    class Meta:
        verbose_name = "Заметка"
        verbose_name_plural = "Заметки"
        ordering = ["-created_at"]

    def __str__(self):
        return self.title or "Заметка"
