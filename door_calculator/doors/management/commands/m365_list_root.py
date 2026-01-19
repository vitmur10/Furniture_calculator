from django.core.management.base import BaseCommand
from django.conf import settings

from doors.services.m365_graph import (
    find_site_by_display_name,
    pick_drive,
    list_root_children,
)


class Command(BaseCommand):
    help = "Print root folders in configured M365 site(s) and drive"

    def handle(self, *args, **options):
        site_names = getattr(settings, "M365_SITE_DISPLAY_NAMES", None)
        if not site_names:
            site_names = [getattr(settings, "M365_SITE_DISPLAY_NAME")]

        drive_name = getattr(settings, "M365_DRIVE_NAME", None)
        if not drive_name:
            raise RuntimeError("M365_DRIVE_NAME is not set in settings.py")

        for site_name in site_names:
            self.stdout.write(self.style.NOTICE(f"\n=== SITE: {site_name} ==="))

            site = find_site_by_display_name(site_name)
            if not site:
                self.stdout.write(self.style.WARNING(f"Site not found: {site_name}"))
                continue

            drive = pick_drive(site["id"], drive_name)
            if not drive:
                self.stdout.write(self.style.WARNING(
                    f"Drive '{drive_name}' not found in site '{site_name}'"
                ))
                continue

            drive_id = drive["id"]
            items = list_root_children(drive_id)

            self.stdout.write(self.style.WARNING("ROOT FOLDERS IN DRIVE:"))
            found_any = False
            for it in items:
                if it.get("folder"):
                    found_any = True
                    self.stdout.write(f" - {it.get('name')}")

            if not found_any:
                self.stdout.write(self.style.WARNING(" (no folders found in root)"))

            # Підказка для settings.py
            self.stdout.write(self.style.NOTICE("\nSuggested settings.py snippet:"))
            self.stdout.write("M365_PROJECT_ROOT_FOLDERS = [")
            for it in items:
                if it.get("folder"):
                    self.stdout.write(f'    "{it.get("name")}",')
            self.stdout.write("]")
