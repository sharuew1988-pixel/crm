from django.core.management.base import BaseCommand
from urllib.parse import urlparse

from app.models import SalesLead


class Command(BaseCommand):
    help = "Fill vacancy from ad_url if empty"

    def handle(self, *args, **options):
        updated = 0
        for lead in SalesLead.objects.filter(vacancy=""):
            url = lead.ad_url or ""
            vac = ""
            try:
                parts = urlparse(url).path.strip("/").split("/")
                if parts:
                    vac = parts[-1].replace("-", " ").replace("_", " ").title()
            except Exception:
                pass

            if not vac:
                vac = "Линейный персонал"

            lead.vacancy = vac[:255]
            lead.company_name = (lead.company_name or vac)[:255]
            lead.save(update_fields=["vacancy", "company_name"])
            updated += 1

        self.stdout.write(self.style.SUCCESS(f"Updated: {updated}"))
