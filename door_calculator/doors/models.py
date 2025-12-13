from django.db import models
from django.contrib.auth.models import User


class Category(models.Model):
    name = models.CharField(max_length=100, unique=True, verbose_name="Назва категорії")
    description = models.TextField(blank=True, null=True, verbose_name="Опис")

    class Meta:
        verbose_name = "Категорія"
        verbose_name_plural = "Категорії"
        ordering = ["name"]

    def __str__(self):
        return self.name


class Product(models.Model):
    category = models.ForeignKey(
        Category,
        on_delete=models.CASCADE,
        related_name="products",
        verbose_name="Категорія",
        null=True,
        blank=True
    )
    name = models.CharField(max_length=255, verbose_name="Назва виробу")
    base_ks = models.FloatField(verbose_name="Базові к/с")
    image = models.ImageField(upload_to="products/", blank=True, null=True, verbose_name="Зображення виробу")

    class Meta:
        verbose_name = "Виріб"
        verbose_name_plural = "Вироби"
        ordering = ["category", "name"]

    def __str__(self):
        return self.category.name if self.category else "No category"


class Addition(models.Model):
    name = models.CharField(max_length=255)
    ks_value = models.FloatField()

    # де доступне доповнення
    applies_globally = models.BooleanField(
        default=True,
        help_text="Якщо увімкнено — доступне для всіх виробів."
    )
    categories = models.ManyToManyField(
        Category, related_name="additions", blank=True,
        help_text="Доступне для виробів цих категорій."
    )
    products = models.ManyToManyField(
        Product, related_name="additions", blank=True,
        help_text="Доступне для конкретних виробів."
    )

    class Meta:
        verbose_name = "Доповнення"
        verbose_name_plural = "Доповнення"

    def __str__(self):
        return self.name


class Coefficient(models.Model):
    name = models.CharField(max_length=255)
    value = models.FloatField(default=1.0)

    # де доступний коефіцієнт
    applies_globally = models.BooleanField(
        default=True,
        help_text="Якщо увімкнено — доступний для всіх виробів, незалежно від зв’язків нижче."
    )
    categories = models.ManyToManyField(
        Category, related_name="coefficients", blank=True,
        help_text="Якщо вказано — коефіцієнт доступний для виробів цих категорій."
    )
    products = models.ManyToManyField(
        Product, related_name="coefficients", blank=True,
        help_text="Якщо вказано — коефіцієнт доступний для цих конкретних виробів."
    )

    class Meta:
        verbose_name = "Коефіцієнт"
        verbose_name_plural = "Коефіцієнти"

    def __str__(self):
        return f"{self.name} ×{self.value}"


class Rate(models.Model):
    price_per_ks = models.DecimalField(max_digits=10, decimal_places=2, default=10.00)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Вартість 1 к/с"
        verbose_name_plural = "Вартість 1 к/с"

    def __str__(self):
        return f"{self.price_per_ks} грн/к.с."


class Customer(models.Model):
    TYPE_CHOICES = [
        ("person", "Фізична особа"),
        ("company", "Юридична особа / ФОП"),
    ]

    type = models.CharField(
        "Тип замовника",
        max_length=20,
        choices=TYPE_CHOICES,
        default="person",
    )

    # основні поля
    name = models.CharField(
        "Ім'я / Назва замовника",
        max_length=255,
        help_text="Наприклад: Іван Петренко або ТОВ “БудМонтаж”",
    )
    contact_person = models.CharField(
        "Контактна особа",
        max_length=255,
        blank=True,
        null=True,
        help_text="Якщо це компанія — ПІБ контактної особи",
    )

    phone = models.CharField(
        "Телефон",
        max_length=50,
        blank=True,
        null=True,
    )
    email = models.EmailField(
        "Email",
        blank=True,
        null=True,
    )

    # реквізити / адреса (по бажанню)
    company_code = models.CharField(
        "ЄДРПОУ / ІПН",
        max_length=40,
        blank=True,
        null=True,
    )
    address = models.CharField(
        "Адреса",
        max_length=255,
        blank=True,
        null=True,
    )

    # необов'язкові поля
    telegram = models.CharField(
        "Telegram / нік",
        max_length=100,
        blank=True,
        null=True,
    )
    notes = models.TextField(
        "Нотатки по замовнику",
        blank=True,
        null=True,
    )

    created_at = models.DateTimeField("Дата створення", auto_now_add=True)

    class Meta:
        verbose_name = "Замовник"
        verbose_name_plural = "Замовники"
        ordering = ["-created_at"]

    def __str__(self):
        if self.type == "company" and self.contact_person:
            return f"{self.name} ({self.contact_person})"
        return self.name


class Order(models.Model):
    STATUS_CHOICES = [
        ("in_progress", "В роботі"),
        ("completed", "Завершено"),
        ("postponed", "Відкладено"),
        ("calculation", "Розрахунки")
    ]
    STATUS_CHOICES_FINANCE = [
        ("paid", "Сплачено"),
        ("awaiting_payment", "Очікує оплату"),
    ]
    customer = models.ForeignKey(
        "Customer",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="orders",
        verbose_name="Замовник",
    )
    order_number = models.CharField(max_length=50, unique=True)
    created_at = models.DateTimeField(auto_now_add=True)
    total_cost = models.DecimalField(max_digits=10, decimal_places=2, default=0)
    total_ks = models.DecimalField(max_digits=10, decimal_places=2, default=0)
    completion_percent = models.PositiveIntegerField(default=0)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default="in_progress")
    status_finance = models.CharField(max_length=20, choices=STATUS_CHOICES_FINANCE, default="postponed")
    sketch = models.ImageField(upload_to="sketches/", blank=True, null=True)

    def __str__(self):
        return f"Замовлення №{self.order_number}"


class OrderItem(models.Model):
    STATUS_CHOICES_ITEM = [
        ("pending", "Не розпочато"),
        ("in_progress", "В роботі"),
        ("done", "Готово"),
        ("canceled", "Скасовано"),
    ]

    order = models.ForeignKey(Order, on_delete=models.CASCADE, related_name="items")
    name = models.CharField(
        max_length=255,
        blank=True,
        help_text="Назва позиції (напр. 'Двостулкові двері')",
    )
    products = models.ManyToManyField(Product, blank=True)
    coefficients = models.ManyToManyField(Coefficient, blank=True)
    quantity = models.PositiveIntegerField(default=1)

    status = models.CharField(
        max_length=20,
        choices=STATUS_CHOICES_ITEM,
        default="pending",
        verbose_name="Статус позиції",
    )

    def total_ks(self):
        base_ks = sum(p.base_ks for p in self.products.all())
        add_ks = sum(a.total_ks() for a in self.addition_items.all())
        coef = 1
        for c in self.coefficients.all():
            coef *= c.value
        return (base_ks + add_ks) * self.quantity, coef

    def total_cost(self):
        base_ks, coef = self.total_ks()
        rate = Rate.objects.first()
        rate_val = float(rate.price_per_ks) if rate else 0
        return round(base_ks * coef * rate_val, 2)

    def __str__(self):
        return self.name or f"Позиція {self.id}"


class OrderImage(models.Model):
    order = models.ForeignKey(Order, on_delete=models.CASCADE, related_name="images")
    image = models.ImageField(upload_to="order_images/")
    uploaded_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"Фото для замовлення {self.order.order_number}"


class OrderImageMarker(models.Model):
    image = models.ForeignKey(
        OrderImage,
        on_delete=models.CASCADE,
        related_name="markers"
    )
    item = models.ForeignKey(
        OrderItem,
        on_delete=models.CASCADE,
        related_name="image_markers",
        null=True,
        blank=True,
    )
    # координати у відсотках від 0 до 100
    x = models.DecimalField(max_digits=6, decimal_places=2)
    y = models.DecimalField(max_digits=6, decimal_places=2)

    # 🎨 колір мітки (#RRGGBB)
    color = models.CharField(max_length=7, default="#FF0000")

    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.image.id} – {self.item or 'без позиції'} ({self.x}%, {self.y}%)"


class OrderFile(models.Model):
    order = models.ForeignKey(
        Order,
        on_delete=models.CASCADE,
        related_name="files",
        verbose_name="Замовлення",
    )
    file = models.FileField(upload_to="order_files/", verbose_name="Файл")
    description = models.CharField(
        "Опис / назва файлу", max_length=255, blank=True, null=True
    )
    uploaded_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Додатковий файл замовлення"
        verbose_name_plural = "Додаткові файли замовлення"

    def __str__(self):
        return self.description or f"Файл для замовлення {self.order.order_number}"


class OrderProgress(models.Model):
    order = models.ForeignKey(
        Order,
        on_delete=models.CASCADE,
        related_name="progress_logs"
    )
    date = models.DateField(auto_now_add=True)
    percent = models.PositiveIntegerField(default=0)
    comment = models.TextField(blank=True, null=True)

    # 🔹 позиції, які неможливо виконати
    problem_items = models.ManyToManyField(
        OrderItem,
        blank=True,
        related_name="problem_progresses",
        verbose_name="Позиції, які неможливо виконати",
    )

    class Meta:
        ordering = ["-date"]

    def __str__(self):
        return f"{self.order.order_number} — {self.percent}%"


class AdditionItem(models.Model):
    order_item = models.ForeignKey(OrderItem, on_delete=models.CASCADE, related_name="addition_items")
    addition = models.ForeignKey(Addition, on_delete=models.CASCADE)
    quantity = models.PositiveIntegerField(default=1)

    def total_ks(self):
        return self.addition.ks_value * self.quantity

    def __str__(self):
        return f"{self.addition.name} ×{self.quantity}"


class Worker(models.Model):
    name = models.CharField(max_length=100)
    position = models.CharField(max_length=100, blank=True, null=True)

    def __str__(self):
        return self.name


class WorkLog(models.Model):
    worker = models.ForeignKey(Worker, on_delete=models.CASCADE)
    order = models.ForeignKey("Order", on_delete=models.CASCADE, null=True, blank=True)
    date = models.DateField()
    hours = models.DecimalField(max_digits=5, decimal_places=2)
    comment = models.TextField(blank=True, null=True)

    def __str__(self):
        return f"{self.worker.name} — {self.date}"


class ItemProgress(models.Model):
    order_item = models.ForeignKey(OrderItem, on_delete=models.CASCADE, related_name="progress_history")
    date = models.DateField()
    percent_done = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        help_text="Відсоток виконання позиції на цю дату"
    )
    comment = models.TextField(blank=True, null=True)

    class Meta:
        verbose_name = "Прогрес позиції"
        verbose_name_plural = "Прогрес позицій"
        ordering = ["-date"]

    def __str__(self):
        return f"{self.order_item.name} — {self.percent_done}% ({self.date})"


class CompanyInfo(models.Model):
    name = models.CharField("Назва компанії", max_length=255)
    address = models.CharField("Адреса", max_length=255, blank=True, null=True)
    phone = models.CharField("Телефон", max_length=50, blank=True, null=True)
    email = models.EmailField("Email", blank=True, null=True)
    website = models.CharField("Сайт", max_length=255, blank=True, null=True)

    iban = models.CharField("IBAN", max_length=64, blank=True, null=True)
    edrpou = models.CharField("ЄДРПОУ", max_length=20, blank=True, null=True)

    logo = models.ImageField("Логотип", upload_to="company_logo/", blank=True, null=True)

    class Meta:
        verbose_name = "Реквізити компанії"
        verbose_name_plural = "Реквізити компанії"

    def __str__(self):
        return self.name
