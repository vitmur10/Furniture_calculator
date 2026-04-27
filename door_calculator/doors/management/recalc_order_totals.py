from decimal import Decimal

from django.core.management.base import BaseCommand

from doors.models import Order


class Command(BaseCommand):
    help = "Перераховує total_ks та total_cost для всіх замовлень з урахуванням коефіцієнтів"

    def handle(self, *args, **options):
        orders = Order.objects.prefetch_related("items__coefficients").all()
        total = orders.count()
        updated = 0

        for order in orders:
            total_ks_all = Decimal("0")
            total_cost_all = Decimal("0")

            for item in order.items.all():
                ks_base, coef = item.total_ks()
                ks_effective = Decimal(str(ks_base)) * Decimal(str(coef))
                total_ks_all += ks_effective
                total_cost_all += Decimal(str(item.total_cost()))

            order.total_ks = total_ks_all
            order.total_cost = total_cost_all
            order.save(update_fields=["total_ks", "total_cost"])
            updated += 1

            self.stdout.write(f"[{updated}/{total}] #{order.id} {order.order_name} — {total_ks_all:.2f} к/с, {total_cost_all:.2f} грн")

        self.stdout.write(self.style.SUCCESS(f"\nГотово. Оновлено {updated} замовлень."))
