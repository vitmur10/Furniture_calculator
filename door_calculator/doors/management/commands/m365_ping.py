from django.core.management.base import BaseCommand
from doors.services.m365_graph import graph_get


class Command(BaseCommand):
    help = "Ping Microsoft Graph with app-only token"

    def handle(self, *args, **options):
        data = graph_get("/sites?search=*")
        self.stdout.write(self.style.SUCCESS(f"OK. Sites found: {len(data.get('value', []))}"))
