from django import forms
from django.utils.safestring import mark_safe
from .models import Product, Addition, Coefficient, OrderProgress, OrderItem


class ProductImageWidget(forms.RadioSelect):
    """Сучасний віджет для вибору виробів із картинками."""
    option_template_name = None

    def render(self, name, value, attrs=None, renderer=None):
        output = '<div class="modern-card-grid">'
        for product in self.choices.queryset:
            image_url = product.image.url if product.image else '/static/no_image.png'
            checked = 'checked' if str(product.id) == str(value) else ''
            active_class = 'active' if checked else ''
            output += f"""
            <label class="modern-card {active_class}">
                <input type="radio" name="{name}" value="{product.id}" {checked} hidden>
                <div class="modern-card-image">
                    <img src="{image_url}" alt="{product.name}">
                </div>
                <div class="modern-card-body">
                    <h6>{product.name}</h6>
                    <p>{product.complexity} к/с</p>
                </div>
            </label>
            """
        output += "</div>"
        return mark_safe(output)


class DoorCalculationForm(forms.Form):
    """Форма розрахунку вартості виробу."""
    product = forms.ModelChoiceField(
        queryset=Product.objects.all(),
        label="Оберіть виріб",
        widget=ProductImageWidget
    )
    additions = forms.ModelMultipleChoiceField(
        queryset=Addition.objects.all(),
        label="Доповнення",
        required=False,
        widget=forms.CheckboxSelectMultiple
    )
    coefficients = forms.ModelMultipleChoiceField(
        queryset=Coefficient.objects.all(),
        label="Коефіцієнти",
        required=False,
        widget=forms.CheckboxSelectMultiple
    )


class OrderItemMultipleChoiceField(forms.ModelMultipleChoiceField):
    """Красиві підписи: №замовлення — назва позиції."""

    def label_from_instance(self, obj):
        return f"№{obj.order.order_number} — {obj.name}"


class OrderProgressForm(forms.ModelForm):
    class Meta:
        model = OrderProgress
        fields = ["percent", "comment"]
        widgets = {
            "percent": forms.NumberInput(attrs={
                "class": "form-control",
                "min": 0,
                "max": 100,
                "step": 1,
            }),
            "comment": forms.Textarea(attrs={
                "class": "form-control",
                "rows": 2,
            }),
        }
        labels = {
            "percent": "Виконано, %",
            "comment": "Коментар (необов'язково)",
        }
