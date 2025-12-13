from django.contrib import admin
from .models import Product, Addition, Coefficient, Rate, Order, OrderItem, AdditionItem, Worker, WorkLog, OrderProgress, Category, CompanyInfo
# Register your models here.

admin.site.register(Product)
admin.site.register(Addition)
admin.site.register(Coefficient)
admin.site.register(Rate)
admin.site.register(Order)
admin.site.register(OrderItem)
admin.site.register(AdditionItem)
admin.site.register(Worker)
admin.site.register(WorkLog)
admin.site.register(OrderProgress)
admin.site.register(Category)


@admin.register(CompanyInfo)
class CompanyInfoAdmin(admin.ModelAdmin):
    list_display = ("name", "phone", "email", "iban", "edrpou")
    fieldsets = (
        ("Основна інформація", {
            "fields": ("name", "logo")
        }),
        ("Контакти", {
            "fields": ("address", "phone", "email", "website")
        }),
        ("Банківські реквізити", {
            "fields": ("iban", "edrpou")
        }),
    )