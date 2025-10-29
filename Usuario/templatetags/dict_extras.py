# Usuario/templatetags/dict_extras.py
from django import template

register = template.Library()

@register.filter
def dict_key(d, key):
    if d is None: return None
    return d.get(key)