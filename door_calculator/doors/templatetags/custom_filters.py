from django import template
register = template.Library()

@register.filter
def get_item(dictionary, key):
    """Дозволяє отримати значення зі словника в шаблоні."""
    if dictionary and key in dictionary:
        return dictionary.get(key)
    return 0
