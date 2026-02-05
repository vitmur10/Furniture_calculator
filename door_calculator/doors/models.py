from decimal import Decimal, ROUND_HALF_UP
from django.db import models


class OrderNameDirectory(models.Model):
    name = models.CharField(
        "Назва замовлення",
        max_length=255,
        unique=True,
        help_text="Шаблонна назва замовлення, напр. 'Двері в квартиру', 'Комплексні двері на об'єкт'"
    )
    description = models.TextField(
        "Опис / примітка",
        blank=True,
        null=True,
        help_text="За бажанням: де використовується, які особливості"
    )

    class Meta:
        verbose_name = "Назва замовлення (довідник)"
        verbose_name_plural = "Назви замовлень (довідник)"
        ordering = ["name"]

    def __str__(self):
        return self.name


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
        return self.name


class Addition(models.Model):
    name = models.CharField("Назва доповнення", max_length=255)
    ks_value = models.FloatField("Значення к/с")

    applies_globally = models.BooleanField(
        "Доступне для всіх",
        default=True,
        help_text="Якщо увімкнено — доступне для всіх виробів."
    )
    categories = models.ManyToManyField(
        Category,
        related_name="additions",
        blank=True,
        verbose_name="Категорії",
        help_text="Доступне для виробів цих категорій."
    )
    products = models.ManyToManyField(
        Product,
        related_name="additions",
        blank=True,
        verbose_name="Вироби",
        help_text="Доступне для конкретних виробів."
    )

    def __str__(self):
        return f"{self.name}={self.ks_value}"

    class Meta:
        verbose_name = "Доповнення"
        verbose_name_plural = "Доповнення"


class Coefficient(models.Model):
    name = models.CharField("Назва коефіцієнта", max_length=255)
    value = models.FloatField("Значення", default=1.0)

    applies_globally = models.BooleanField(
        "Доступний для всіх",
        default=True,
        help_text="Якщо увімкнено — доступний для всіх виробів."
    )
    categories = models.ManyToManyField(
        Category,
        related_name="coefficients",
        blank=True,
        verbose_name="Категорії"
    )
    products = models.ManyToManyField(
        Product,
        related_name="coefficients",
        blank=True,
        verbose_name="Вироби"
    )

    def __str__(self):
        return f"{self.name}={self.value}"

    class Meta:
        verbose_name = "Коефіцієнт"
        verbose_name_plural = "Коефіцієнти"


class Rate(models.Model):
    price_per_ks = models.DecimalField(
        "Вартість за 1 к/с",
        max_digits=10,
        decimal_places=2,
        default=10.00
    )
    updated_at = models.DateTimeField("Оновлено", auto_now=True)

    def __str__(self):
        return f"{self.price_per_ks}"
    class Meta:
        verbose_name = "Тариф"
        verbose_name_plural = "Тарифи"


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
        ("calculation", "Розрахунки"),
        ("in_progress", "В роботі"),
        ("completed", "Завершено"),
        ("postponed", "Відкладено"),
    ]
    STATUS_CHOICES_FINANCE = [
        ("paid", "Сплачено"),
        ("awaiting_payment", "Очікує оплату"),
        ("-----", "-----"),
    ]
    WORK_TYPE_CHOICES = [
        ("project", "Проєкт"),
        ("rework", "Переробка"),
    ]
    price_per_ks = models.DecimalField(
        "Ціна за 1 к/с (зафіксована)",
        max_digits=10,
        decimal_places=2,
        null=True,
        blank=True,
        help_text="Фіксується при створенні замовлення (або при першому розрахунку), щоб зміна Rate не впливала на старі замовлення.",
    )
    work_type = models.CharField(
        max_length=20,
        choices=WORK_TYPE_CHOICES,
        default="project",
        db_index=True,
        verbose_name="Тип (проєкт/переробка)",
    )
    markup_percent = models.DecimalField(
        "Націнка за замовлення (%)",
        max_digits=6,
        decimal_places=2,
        default=Decimal("0.00"),
    )
    customer = models.ForeignKey(
        "Customer",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="orders",
        verbose_name="Замовник",
    )
    order_name = models.CharField(
        "Назва замовлення",
        max_length=255,
        blank=True,
        null=True,
        help_text="Напр.: 'Двері на квартиру 12, під'їзд 3'"
    )

    order_number = models.CharField(max_length=50, unique=True)
    created_at = models.DateTimeField(auto_now_add=True)
    total_cost = models.DecimalField(max_digits=10, decimal_places=2, default=0)
    total_ks = models.DecimalField(max_digits=10, decimal_places=2, default=0)
    completion_percent = models.PositiveIntegerField(default=0)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default="in_progress")
    status_finance = models.CharField(max_length=20, choices=STATUS_CHOICES_FINANCE, default="postponed")
    sketch = models.ImageField(upload_to="sketches/", blank=True, null=True)
    source = models.CharField(
        "Джерело",
        max_length=20,
        default="local",
        help_text="local — локально, m365 — Microsoft 365"
    )
    remote_site_id = models.CharField(max_length=255, blank=True, null=True)
    remote_drive_id = models.CharField(max_length=255, blank=True, null=True)
    remote_folder_id = models.CharField(max_length=255, blank=True, null=True)
    remote_web_url = models.URLField(blank=True, null=True)

    def __str__(self):
        return f"Замовлення №{self.order_number}, назва замовлення{self.order_name}"

    class Meta:
        verbose_name = "Замовлення"
        verbose_name_plural = "Замовлення"
        ordering = ["-created_at"]


class OrderItemProduct(models.Model):
    order_item = models.ForeignKey("OrderItem", on_delete=models.CASCADE, related_name="product_items")
    product = models.ForeignKey("Product", on_delete=models.CASCADE)
    quantity = models.DecimalField(max_digits=10, decimal_places=2, default=1)

    class Meta:
        unique_together = ("order_item", "product")
        verbose_name = "Виріб у позиції"
        verbose_name_plural = "Вироби у позиціях"


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
    products_v2 = models.ManyToManyField(
        "Product",
        through="OrderItemProduct",
        blank=True,
        related_name="order_items_v2",
    )
    coefficients = models.ManyToManyField(Coefficient, blank=True)
    quantity = models.DecimalField(max_digits=10, decimal_places=2, default=1)

    status = models.CharField(
        max_length=20,
        choices=STATUS_CHOICES_ITEM,
        default="pending",
        verbose_name="Статус позиції",
    )
    markup_percent = models.DecimalField(
        "Індивідуальна націнка (%)",
        max_digits=6,
        decimal_places=2,
        null=True,
        blank=True,
        help_text="Якщо задано — перекриває націнку замовлення",
    )

    def effective_markup_percent(self) -> Decimal:
        """
        Повертає % націнки для позиції:
        - якщо задано у позиції → беремо його
        - інакше → беремо з замовлення
        """
        if self.markup_percent is not None:
            return Decimal(str(self.markup_percent))
        return Decimal(str(getattr(self.order, "markup_percent", 0) or 0))

    def base_cost(self):
        ks = self.ks_effective()  # або self.total_ks() якщо так задумано
        # беремо зафіксований тариф із замовлення, інакше fallback на Rate
        if self.order and self.order.price_per_ks is not None:
            price_per_ks = Decimal(str(self.order.price_per_ks))
        else:
            rate = Rate.objects.first()
            price_per_ks = Decimal(str(rate.price_per_ks)) if rate else Decimal("0")

        return ks * price_per_ks

    def total_ks(self):
        products_ks = Decimal("0")
        for op in self.product_items.select_related("product").all():
            base = Decimal(str(op.product.base_ks or 0))
            qty = Decimal(str(op.quantity or 1))
            products_ks += base * Decimal(qty)

        adds_ks = Decimal("0")
        for ai in self.addition_items.select_related("addition").all():
            adds_ks += Decimal(str(ai.total_ks() or 0))

        coef = Decimal("1.0")
        for c in self.coefficients.all():
            coef *= Decimal(str(c.value or 1))

        qty_item = Decimal(str(self.quantity or 1))
        ks_base = (products_ks + adds_ks) * qty_item

        return ks_base, coef

    def total_cost(self):
        ks_base, coef = self.total_ks()
        ks_effective = ks_base * coef

        rate = Decimal(str(self.order.price_per_ks or 0))
        base_price = ks_effective * rate

        markup = self.markup_percent
        if markup is None:
            markup = self.order.markup_percent or Decimal("0")

        return base_price * (Decimal("1") + (Decimal(str(markup)) / Decimal("100")))

    def __str__(self):
        return self.name or f"Позиція {self.id}"

    class Meta:
        verbose_name = "Позиція замовлення"
        verbose_name_plural = "Позиції замовлення"


class OrderImage(models.Model):
    """Фото замовлення з Microsoft 365 (Teams/SharePoint)"""

    order = models.ForeignKey(Order, on_delete=models.CASCADE, related_name="images")
    uploaded_at = models.DateTimeField(auto_now_add=True)

    # M365 remote reference (обов'язкові поля)
    remote_site_id = models.CharField(max_length=255)
    remote_drive_id = models.CharField(max_length=255)
    remote_item_id = models.CharField(max_length=255)
    remote_web_url = models.URLField(blank=True, null=True)
    remote_name = models.CharField(max_length=255, blank=True, null=True)
    remote_size = models.BigIntegerField(blank=True, null=True)

    class Meta:
        verbose_name = "Фото замовлення"
        verbose_name_plural = "Фото замовлення"
        constraints = [
            models.UniqueConstraint(
                fields=["remote_drive_id", "remote_item_id"],
                name="uniq_remote_image",
            )
        ]

    def get_image_url(self):
        """Повертає URL для відображення фото з M365"""
        from django.urls import reverse
        return reverse("m365_image_content", args=[self.id])

    def get_thumb_url(self):
        """Повертає URL для мініатюри з M365"""
        from django.urls import reverse
        return reverse("m365_image_thumb", args=[self.id])

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

    class Meta:
        verbose_name = "Мітка на фото"
        verbose_name_plural = "Мітки на фото"


class OrderFile(models.Model):
    SOURCE_CHOICES = [
        ("local", "Local upload"),
        ("m365", "Microsoft 365 (SharePoint/OneDrive)"),
    ]

    order = models.ForeignKey(Order, on_delete=models.CASCADE, related_name="files")
    file = models.FileField(upload_to="order_files/", blank=True, null=True)  # <- важливо

    description = models.CharField(max_length=255, blank=True, null=True)
    uploaded_at = models.DateTimeField(auto_now_add=True)

    # NEW: remote reference
    source = models.CharField(max_length=20, choices=SOURCE_CHOICES, default="local")
    remote_site_id = models.CharField(max_length=255, blank=True, null=True)
    remote_drive_id = models.CharField(max_length=255, blank=True, null=True)
    remote_item_id = models.CharField(max_length=255, blank=True, null=True)
    remote_web_url = models.URLField(blank=True, null=True)
    remote_name = models.CharField(max_length=255, blank=True, null=True)
    remote_size = models.BigIntegerField(blank=True, null=True)

    class Meta:
        verbose_name = "Файл замовлення"
        verbose_name_plural = "Файли замовлення"
        constraints = [
            models.UniqueConstraint(
                fields=["source", "remote_drive_id", "remote_item_id"],
                name="uniq_remote_file",
            )
        ]


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
        verbose_name = "Прогрес замовлення"
        verbose_name_plural = "Прогрес замовлень"
        ordering = ["-date"]

    def __str__(self):
        return f"{self.order.order_number} — {self.percent}%"


class AdditionItem(models.Model):
    order_item = models.ForeignKey(OrderItem, on_delete=models.CASCADE, related_name="addition_items")
    addition = models.ForeignKey(Addition, on_delete=models.CASCADE)
    quantity = models.DecimalField(max_digits=10, decimal_places=2, default=1)

    def total_ks(self):
        ks_value = Decimal(str(getattr(self.addition, "ks_value", 0) or 0))
        qty = Decimal(str(self.quantity or 0))
        return ks_value * qty

    def __str__(self):
        return f"{self.addition.name} ×{self.quantity}"

    class Meta:
        verbose_name = "Доповнення в позиції"
        verbose_name_plural = "Доповнення в позиції"


class Worker(models.Model):
    name = models.CharField(max_length=100)
    position = models.CharField(max_length=100, blank=True, null=True)

    def __str__(self):
        return self.name

    class Meta:
        verbose_name = "Працівник"
        verbose_name_plural = "Працівники"
        ordering = ["name"]


class WorkLog(models.Model):
    worker = models.ForeignKey(Worker, on_delete=models.CASCADE)
    order = models.ForeignKey("Order", on_delete=models.CASCADE, null=True, blank=True)
    date = models.DateField()
    hours = models.DecimalField(max_digits=5, decimal_places=2)
    comment = models.TextField(blank=True, null=True)

    def __str__(self):
        return f"{self.worker.name} — {self.date}"

    class Meta:
        verbose_name = "Журнал робіт"
        verbose_name_plural = "Журнали робіт"
        ordering = ["-date"]


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
