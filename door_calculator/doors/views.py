import json
import os
from datetime import datetime, timedelta, date
from decimal import Decimal, ROUND_HALF_UP
from io import BytesIO
from django.conf import settings
from django.contrib import messages
from django.contrib.auth.decorators import login_required, user_passes_test
from django.db.models import Sum, Q
from django.http import HttpResponse, JsonResponse, HttpResponseBadRequest, FileResponse, StreamingHttpResponse, \
    Http404, QueryDict
from django.shortcuts import render, redirect, get_object_or_404
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.http import require_POST
from openpyxl import Workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from doors.services.m365_graph import get_app_token, list_children, search_in_folder, upload_bytes_to_folder
from .forms import OrderProgressForm
from .models import (
    Category, Product, Addition, Coefficient, Rate,
    Order, OrderItem, AdditionItem, WorkLog, Worker,
    OrderProgress, OrderImage, CompanyInfo, OrderFile, Customer, OrderImageMarker, OrderNameDirectory, OrderItemProduct
)
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.enums import TA_LEFT
import requests
from django.urls import reverse
from django.views.decorators.clickjacking import xframe_options_exempt
from django.utils.html import strip_tags
import logging
from math import ceil
from decimal import Decimal, ROUND_HALF_UP
from django.shortcuts import get_object_or_404, redirect, render
from django.template.loader import render_to_string
import re

logger = logging.getLogger(__name__)


def _debug_post(request, tag: str):
    keys = list(request.POST.keys())
    # покажемо вибірково найважливіше
    sample = {k: request.POST.getlist(k) for k in keys if k in (
        "save_markup", "order_markup",
        "bulk_coefficients", "bulk_scope", "bulk_mode",
        "bulk_coeff_ids", "selected_item_ids"
    ) or k.startswith("item_markup_")}
def _norm_m365_name(s: str) -> str:
    s = (s or "").lower()
    for ch in [" ", ".", ",", "-", "_", "’", "'", '"']:
        s = s.replace(ch, "")
    return s
def _is_m365_folder(it: dict) -> bool:
    return bool(it and it.get("folder"))
@login_required
def delete_order_file(request, file_id):
    of = get_object_or_404(OrderFile, id=file_id)
    order_id = of.order_id
    of.delete()
    return redirect("calculate_order", order_id=order_id)


def superuser_only(view_func):
    @login_required
    @user_passes_test(lambda u: u.is_superuser)
    def _wrapped(request, *args, **kwargs):
        return view_func(request, *args, **kwargs)

    return _wrapped


def _get_applicable_coefficients(products_qs):
    """
    Коefs: глобальні або прив'язані до категорій/конкретних продуктів.
    """
    prod_ids = list(products_qs.values_list("id", flat=True))
    cat_ids = list(products_qs.values_list("category_id", flat=True))
    return (Coefficient.objects.filter(
        Q(applies_globally=True) |
        Q(products__in=prod_ids) |
        Q(categories__in=cat_ids)
    )
            .distinct()
            .order_by("name"))


def _get_applicable_additions(products_qs):
    """
    Additions: глобальні або прив'язані до категорій/конкретних продуктів.
    """
    prod_ids = list(products_qs.values_list("id", flat=True))
    cat_ids = list(products_qs.values_list("category_id", flat=True))
    return (Addition.objects.filter(
        Q(applies_globally=True) |
        Q(products__in=prod_ids) |
        Q(categories__in=cat_ids)
    )
            .distinct()
            .order_by("name"))


def _recalc_order_totals(order: Order):
    total_ks_all = Decimal("0")
    total_cost_all = Decimal("0")
    for i in order.items.all():
        ks_base, coef = i.total_ks()
        # ВИПРАВЛЕННЯ: множимо на коефіцієнт (теплий профіль, великий розмір тощо)
        ks_effective = Decimal(str(ks_base)) * Decimal(str(coef))
        total_ks_all += ks_effective
        total_cost_all += Decimal(str(i.total_cost()))
    order.total_ks = total_ks_all
    order.total_cost = total_cost_all
    order.save(update_fields=["total_ks", "total_cost"])


def order_list(request):
    orders = Order.objects.all().order_by("-created_at")

    start_date = request.GET.get("start_date") or ""
    end_date = request.GET.get("end_date") or ""
    status = request.GET.get("status") or ""
    status_finance = request.GET.get("status_finance") or ""
    order_name = request.GET.get("order_name") or ""
    work_type = request.GET.get("work_type") or ""
    if start_date:
        orders = orders.filter(created_at__date__gte=start_date)
    if end_date:
        orders = orders.filter(created_at__date__lte=end_date)
    if status:
        orders = orders.filter(status=status)
    if status_finance:
        orders = orders.filter(status_finance=status_finance)
    if order_name:
        # фільтр по точній назві з дропдауну
        orders = orders.filter(order_name=order_name)
    if work_type:
        orders = orders.filter(work_type=work_type)
    # 🔹 Список доступних назв замовлень для фільтра (тільки не пусті)
    order_name_choices = (
        Order.objects
        .exclude(order_name__isnull=True)
        .exclude(order_name__exact="")
        .values_list("order_name", flat=True)
        .distinct()
        .order_by("order_name")
    )

    context = {
        "orders": orders,
        "status_choices": Order.STATUS_CHOICES,
        "status_finance_choices": Order.STATUS_CHOICES_FINANCE,
        "order_name_choices": order_name_choices,
        "work_type_choices": Order.WORK_TYPE_CHOICES,
    }
    return render(request, "doors/order_list.html", context)


def extract_kp_suffix_from_folder(folder: dict | None) -> str:
    name = ((folder or {}).get("name") or "").strip()

    # КП1 / КП 1 / КП-1
    m = re.search(r"кп[\s\-]*([0-9]+)", name, flags=re.IGNORECASE)
    if m:
        return f"КП{m.group(1)}"

    return "КП1"


def extract_d_suffix_from_folder(folder: dict | None) -> str:
    name = ((folder or {}).get("name") or "").strip()

    # Д1 / Д 1 / Д-1
    m = re.search(r"\bд[\s\-]*([0-9]+)\b", name, flags=re.IGNORECASE)
    if m:
        return f"Д{m.group(1)}"

    return "Д1"


def build_pdf_number(order: Order, mode: str, target_folder: dict | None = None) -> str:
    base = (order.order_number or "").strip()
    if not base:
        return ""

    # Для переробок не додаємо 2-гу частину
    if order.work_type == "rework":
        return base

    if mode == "precalc":
        return f"{base}{extract_kp_suffix_from_folder(target_folder)}"

    if mode == "final":
        if order.status == "in_progress":
            return f"{base}{extract_d_suffix_from_folder(target_folder)}"
        return base

    return base


@csrf_exempt
@require_POST
def update_status(request, order_id):
    order = get_object_or_404(Order, id=order_id)

    raw = (request.body or b"").decode("utf-8", errors="replace")
    data = json.loads(raw) if raw.strip() else {}

    # --- підтримка двох форматів ---
    # A) {"status": "...", "progress": 10}
    # B) {"kind": "...", "value": "..."}
    if "kind" in data and "value" in data:
        kind = data["kind"]
        value = data["value"]

        if kind == "status":
            data["status"] = value
        elif kind == "status_finance":
            data["status_finance"] = value
        elif kind == "progress":
            data["progress"] = value
        elif kind == "comment":
            data["comment"] = value

    before = {
        "status": order.status,
        "status_finance": getattr(order, "status_finance", None),
        "progress": order.completion_percent,
    }

    # звичайний статус
    if "status" in data:
        order.status = data["status"]

    # фінансовий статус
    if "status_finance" in data:
        order.status_finance = data["status_finance"]

    # прогрес
    progress = None
    if "progress" in data:
        try:
            progress = int(data["progress"])
        except (TypeError, ValueError):
            return JsonResponse({"success": False, "error": "progress must be int"}, status=400)
        order.completion_percent = progress

    order.save()
    order.refresh_from_db()

    if progress is not None:
        OrderProgress.objects.create(
            order=order,
            date=date.today(),
            percent=progress,
            comment=data.get("comment", "")
        )

    return JsonResponse({
        "success": True,
        "before": before,
        "after": {
            "status": order.status,
            "status_finance": getattr(order, "status_finance", None),
            "progress": order.completion_percent,
        },
        "received": data,
    })


@require_POST
def delete_order(request, order_id):
    order = get_object_or_404(Order, id=order_id)
    order.delete()

    # Якщо викликали через fetch (AJAX) — повернемо JSON
    if request.headers.get("x-requested-with") == "XMLHttpRequest":
        return JsonResponse({"success": True})

    # Якщо раптом звичайний POST — редірект на список
    return redirect("order_list")


def home(request):
    if request.method == "POST":
        order_number = request.POST.get("order_number")

        if not order_number:
            messages.error(request, "Введіть номер замовлення")
            return redirect("home")

        # Створюємо замовлення без фото (фото будуть з Teams)
        order = Order.objects.create(order_number=order_number)

        # Прибрано: локальне завантаження фото
        # main_sketch = request.FILES.get("sketch")
        # extra_images = request.FILES.getlist("images")
        # if main_sketch:
        #     order.sketch = main_sketch
        #     order.save()
        # if extra_images:
        #     for img in extra_images[:50]:
        #         OrderImage.objects.create(order=order, image=img)

        messages.info(request,
                      "Замовлення створено. Використайте команду 'python manage.py sync_m365_orders' "
                      "для синхронізації фото та файлів з Teams.")
        return redirect("calculate_order", order_id=order.id)

    return render(request, "doors/home.html")


ITEM_COLOR_PALETTE = [
    "#e6194b", "#3cb44b", "#ffe119", "#4363d8", "#f58231", "#911eb4",
    "#46f0f0", "#f032e6", "#bcf60c", "#fabebe", "#008080", "#e6beff",
    "#9a6324", "#fffac8", "#800000", "#aaffc3", "#808000", "#ffd8b1",
    "#000075", "#808080", "#000000", "#ffe4e1", "#ff1493", "#7fffd4",
    "#dc143c", "#00ced1", "#daa520", "#9932cc", "#00fa9a", "#f4a460",
    "#8b4513", "#2e8b57", "#ff4500", "#1e90ff", "#ff6347", "#32cd32",
    "#6495ed", "#ff00ff", "#b0c4de", "#8a2be2", "#7b68ee", "#20b2aa",
    "#ff69b4", "#cd5c5c", "#b22222", "#ff8c00", "#adff2f", "#40e0d0",
    "#ba55d3", "#5f9ea0", "#ffdab9", "#dda0dd", "#afeeee", "#deb887",
    "#ffb6c1", "#556b2f", "#4682b4", "#008b8b", "#7cfc00", "#fa8072",
    "#d2691e", "#00bfff", "#8fbc8f", "#da70d6", "#ffdead", "#bc8f8f",
    "#a52a2a", "#2f4f4f", "#8b008b", "#708090", "#c0c0c0", "#ffd700",
    "#00ff7f", "#7fffd4", "#ff7f50", "#dc143c", "#228b22", "#8a2be2",
    "#ff1493", "#00ff00", "#ff00ff", "#00ffff", "#ff4500", "#4169e1",
    "#ff8c00", "#90ee90", "#ff69b4", "#7b68ee", "#00fa9a", "#ffc0cb",
    "#ee82ee", "#7fff00", "#6a5acd", "#00bcd4", "#ff5722", "#4caf50"
]


def get_item_color(item_id: int | None) -> str:
    if not item_id:
        return "#ff4d4f"
    idx = (int(item_id) - 1) % len(ITEM_COLOR_PALETTE)
    return ITEM_COLOR_PALETTE[idx]


def calculate_order(request, order_id):
    order = get_object_or_404(Order, id=order_id)

    is_ajax = request.headers.get("x-requested-with") == "XMLHttpRequest"
    ajax_partial = False  # if True, return JSON with updated items table

    categories = Category.objects.prefetch_related("products").all()
    products = Product.objects.all().select_related("category")
    images = order.images.all()

    rate_obj = Rate.objects.first()
    current_rate = Decimal(str(rate_obj.price_per_ks)) if rate_obj else Decimal("0")

    if order.price_per_ks is None:
        order.price_per_ks = current_rate
        order.save(update_fields=["price_per_ks"])

    price_per_ks = Decimal(str(order.price_per_ks)) if order.price_per_ks is not None else current_rate

    # ---------------- helpers ----------------
    def _hsl_to_hex(h, s, l):
        import colorsys
        r, g, b = colorsys.hls_to_rgb(h / 360.0, l, s)
        return "#{:02x}{:02x}{:02x}".format(int(r * 255), int(g * 255), int(b * 255))

    def get_item_color(item_id: int) -> str:
        if not item_id:
            return "#ff4d4f"
        hue = (item_id * 137.508) % 360
        sat = 0.80
        light = 0.48 if (item_id % 2 == 0) else 0.62
        return _hsl_to_hex(hue, sat, light)

    def _to_decimal_or_none(val: str):
        val = (val or "").strip().replace(",", ".")
        if val == "":
            return None
        try:
            return Decimal(val)
        except:
            return None

    def _to_decimal_or_one(val: str):
        try:
            v = Decimal(str(val).replace(",", "."))
            return max(v, Decimal("0.01"))
        except:
            return Decimal("1")

    def _q2(x: Decimal) -> Decimal:
        return x.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    # ============================================================
    # POST: assign_customer
    # ============================================================
    if request.method == "POST" and "assign_customer" in request.POST:
        existing_id = request.POST.get("existing_customer") or ""
        customer = None

        if existing_id:
            customer = Customer.objects.filter(id=existing_id).first()
        else:
            cust_type = request.POST.get("customer_type") or "person"
            name = (request.POST.get("customer_name") or "").strip()
            phone = (request.POST.get("customer_phone") or "").strip()
            email = (request.POST.get("customer_email") or "").strip()
            address = (request.POST.get("customer_address") or "").strip()
            company_code = (request.POST.get("customer_company_code") or "").strip()

            if name:
                customer = Customer.objects.create(
                    type=cust_type,
                    name=name,
                    phone=phone or None,
                    email=email or None,
                    address=address or None,
                    company_code=company_code or None,
                )

        if customer:
            order.customer = customer
            order.save(update_fields=["customer"])

        if is_ajax:
            ajax_partial = True
        else:
            return redirect("calculate_order", order_id=order.id)

    # ============================================================
    # POST: bulk coefficients
    # ============================================================
    if request.method == "POST" and "bulk_coefficients" in request.POST:
        coeff_ids = request.POST.getlist("bulk_coeff_ids")
        scope = request.POST.get("bulk_scope", "all")
        mode = request.POST.get("bulk_mode", "add")
        selected_item_ids = request.POST.getlist("selected_item_ids")

        if coeff_ids:
            coefs = list(Coefficient.objects.filter(id__in=coeff_ids))

            target_qs = order.items.all()
            if scope == "selected":
                target_qs = target_qs.filter(id__in=selected_item_ids)

            if mode == "replace":
                for it in target_qs:
                    it.coefficients.set(coefs)
            else:
                for it in target_qs:
                    it.coefficients.add(*coefs)

            _recalc_order_totals(order)
            order.refresh_from_db()

        if is_ajax:
            ajax_partial = True
        else:
            return redirect("calculate_order", order_id=order.id)

    # ============================================================
    # POST: save_markup
    # ============================================================
    if request.method == "POST" and "save_markup" in request.POST:
        order_markup = _to_decimal_or_none(request.POST.get("order_markup")) or Decimal("0")
        order.markup_percent = order_markup
        order.save(update_fields=["markup_percent"])

        for it in order.items.all():
            key = f"item_markup_{it.id}"
            raw = request.POST.get(key)
            if raw:
                it.markup_percent = _to_decimal_or_none(raw)
                it.save(update_fields=["markup_percent"])

        _recalc_order_totals(order)
        order.refresh_from_db()

        if is_ajax:
            ajax_partial = True
        else:
            return redirect("calculate_order", order_id=order.id)

    # ============================================================
    # POST: attach/detach item (make item a "child" of another item)
    # ============================================================
    if request.method == "POST" and "attach_item" in request.POST:
        item_id = request.POST.get("attach_item_id")
        parent_id = request.POST.get("attach_parent_id") or None

        item = get_object_or_404(OrderItem, id=item_id, order=order)

        if parent_id:
            parent = get_object_or_404(OrderItem, id=parent_id, order=order)
            if parent.id == item.id:
                parent = None
        else:
            parent = None

        item.attached_to = parent
        item.save(update_fields=["attached_to"])

        _recalc_order_totals(order)
        order.refresh_from_db()

        if is_ajax:
            ajax_partial = True
        else:
            return redirect("calculate_order", order_id=order.id)

    # ============================================================
    # POST: copy item (duplicate existing position inside same order)
    # ============================================================
    if request.method == "POST" and "copy_item" in request.POST:
        src_id = request.POST.get("copy_item_id")
        src = get_object_or_404(OrderItem, id=src_id, order=order)

        new_item = OrderItem.objects.create(
            order=order,
            name=f"{src.name} (копія)",
            quantity=src.quantity,
            attached_to=src.attached_to,
        )

        if hasattr(src, "markup_percent"):
            new_item.markup_percent = src.markup_percent
            new_item.save(update_fields=["markup_percent"])

        new_item.coefficients.set(src.coefficients.all())

        for pi in src.product_items.all():
            OrderItemProduct.objects.create(
                order_item=new_item,
                product_id=pi.product_id,
                quantity=pi.quantity,
            )

        for ai in src.addition_items.all():
            AdditionItem.objects.create(
                order_item=new_item,
                addition_id=ai.addition_id,
                quantity=ai.quantity,
            )

        _recalc_order_totals(order)
        order.refresh_from_db()

        if is_ajax:
            ajax_partial = True
        else:
            return redirect("calculate_order", order_id=order.id)

    # ============================================================
    # POST: add item
    # ============================================================
    if request.method == "POST" and all(
        x not in request.POST
        for x in [
            "upload_images",
            "upload_files",
            "update_file",
            "delete_file",
            "assign_customer",
            "save_markup",
            "bulk_coefficients",
            "copy_item",
            "attach_item",
        ]
    ):
        order_name = (request.POST.get("order_name") or "").strip()
        order_name_template_id = request.POST.get("order_name_template") or ""
        fields_to_update = []

        if order_name_template_id:
            tpl = OrderNameDirectory.objects.filter(id=order_name_template_id).first()
            if tpl:
                order.order_name_template = tpl
                fields_to_update.append("order_name_template")
                if not order_name:
                    order_name = tpl.name

        if order_name:
            order.order_name = order_name
            fields_to_update.append("order_name")

        if fields_to_update:
            order.save(update_fields=fields_to_update)

        name = request.POST.get("name") or "Позиція"
        item_qty = _to_decimal_or_one(request.POST.get("item_qty", 1))

        calc_mode = (request.POST.get("calc_mode") or "products").strip().lower()

        facade_total_ks = _to_decimal_or_none(request.POST.get("facade_total_ks"))
        facade_total_cost = _to_decimal_or_none(request.POST.get("facade_total_cost"))

        selected_products = request.POST.getlist("products")
        selected_adds = request.POST.getlist("additions")
        selected_coefs = request.POST.getlist("coefficients")

        attach_parent_id = request.POST.get("attach_parent_id") or None
        attached_to = None
        if attach_parent_id:
            attached_to = OrderItem.objects.filter(id=attach_parent_id, order=order).first()

        item = OrderItem.objects.create(
            order=order,
            name=name,
            quantity=item_qty,
            attached_to=attached_to,
        )

        # ============================================================
        # FACADE MODE
        # ============================================================
        if calc_mode == "facade":
            ks_qty = facade_total_ks
            facade_json = request.POST.get("facade_data_json")

            if facade_json:
                try:
                    item.facade_data = json.loads(facade_json)
                except Exception:
                    item.facade_data = None
                item.save(update_fields=["facade_data"])

            if (ks_qty is None or ks_qty <= 0) and facade_total_cost is not None:
                if price_per_ks <= 0:
                    item.delete()
                    messages.error(request, "Не вдалося додати фасад: не заданий курс (ціна за кс).")
                    return redirect("calculate_order", order_id=order.id)
                ks_qty = (facade_total_cost / price_per_ks)

            if ks_qty is None or ks_qty <= 0:
                item.delete()
                messages.error(
                    request,
                    "Не вдалося додати фасад: значення фасаду (кс) порожнє або 0. Перерахуй фасад і спробуй ще раз."
                )
                return redirect("calculate_order", order_id=order.id)

            ks_qty = ks_qty.quantize(Decimal("0.001"), rounding=ROUND_HALF_UP)

            cat = Category.objects.filter(name__iexact="інше").first() or Category.objects.first()
            facade_product, _created = Product.objects.get_or_create(
                name="Фасад (розрахунок)",
                defaults={"base_ks": Decimal("1"), "category": cat} if cat else {"base_ks": Decimal("1")},
            )

            if not facade_product.base_ks or Decimal(str(facade_product.base_ks)) != Decimal("1"):
                facade_product.base_ks = Decimal("1")
                facade_product.save(update_fields=["base_ks"])

            OrderItemProduct.objects.create(order_item=item, product=facade_product, quantity=ks_qty)

            _recalc_order_totals(order)
            order.refresh_from_db()

            if is_ajax:
                ajax_partial = True
            else:
                return redirect("calculate_order", order_id=order.id)

        # ============================================================
        # PRODUCTS MODE
        # ============================================================
        if selected_products:
            for pid in selected_products:
                qty_field = f"prod_qty_{pid}"
                prod_qty = _to_decimal_or_one(request.POST.get(qty_field, 1))
                OrderItemProduct.objects.create(order_item=item, product_id=pid, quantity=prod_qty)

        if selected_coefs:
            item.coefficients.set(selected_coefs)

        for add_id in selected_adds:
            qty_field = f"add_qty_{add_id}"
            qty = _to_decimal_or_one(request.POST.get(qty_field, 1))
            AdditionItem.objects.create(order_item=item, addition_id=add_id, quantity=qty)

        _recalc_order_totals(order)
        order.refresh_from_db()

        if is_ajax:
            ajax_partial = True
        else:
            return redirect("calculate_order", order_id=order.id)

    # ============================================================
    # GET: prepare
    # ============================================================
    items = order.items.all().prefetch_related(
        "coefficients",
        "addition_items__addition",
        "product_items__product",
        "attached_items",
    ).select_related("attached_to")

    parents = [it for it in items if it.attached_to_id is None]
    children_map = {}
    for it in items:
        if it.attached_to_id:
            children_map.setdefault(it.attached_to_id, []).append(it)
    for pid, ch_list in children_map.items():
        ch_list.sort(key=lambda x: x.id)

    effective_ks = Decimal("0.00")
    total_sum = Decimal("0.00")
    formula_terms = []

    order_markup = Decimal(str(getattr(order, "markup_percent", 0) or 0))

    for it in items:
        it.color_hex = get_item_color(it.id)

        products_ks = Decimal("0")
        prod_terms = []

        for op in it.product_items.all():
            p = op.product
            base = Decimal(str(p.base_ks or 0))
            qty_p = Decimal(op.quantity or 1)
            products_ks += base * qty_p
            prod_terms.append(f"{base:.2f} × {qty_p}")

        adds_ks = Decimal("0")
        add_terms = []

        for ai in it.addition_items.all():
            qty_add = Decimal(getattr(ai, "quantity", 1) or 1)
            total_add = Decimal(str(ai.total_ks() or 0))
            adds_ks += total_add

            base_add = (total_add / qty_add) if qty_add > 0 else total_add
            base_add = _q2(base_add)
            add_terms.append(f"{base_add:.2f} × {qty_add}")

        qty = Decimal(it.quantity or 1)

        coef = Decimal("1.0")
        coef_terms = []
        coef_lines = []

        for c in it.coefficients.all():
            c_val = Decimal(str(c.value or 1))
            coef *= c_val
            coef_terms.append(f"{c_val:.2f}")
            coef_lines.append(f"• {c.name} ×{c_val:.2f}")

        products_formula = " + ".join(prod_terms) if prod_terms else "0.00"
        adds_formula = " + ".join(add_terms) if add_terms else "0.00"
        coef_part = f" × {' × '.join(coef_terms)}" if coef_terms else ""

        it.ks_formula = f"(({products_formula}) + ({adds_formula})) × {qty}{coef_part}"

        ks_base = (products_ks + adds_ks) * qty
        ks_effective = _q2(ks_base * coef)

        it.ks_products = _q2(products_ks)
        it.ks_adds = _q2(adds_ks)
        it.ks_qty = _q2(qty)
        it.ks_coef = _q2(coef)
        it.ks_effective = ks_effective

        effective_ks += ks_effective
        formula_terms.append(f"{ks_effective:.2f}")

        item_markup = it.markup_percent
        m = Decimal(str(item_markup)) if item_markup is not None else order_markup

        base_price = _q2(ks_effective * price_per_ks)
        final_price = _q2(base_price * (Decimal("1") + (m / Decimal("100"))))

        it.workshop_cost_value = base_price       # вартість роботи цеху (без торгової націнки)
        it.total_cost_value = final_price
        total_sum += final_price

        NL = "\n"
        prod_lines = [
            f"• {op.product.name}: {_q2(Decimal(str(op.product.base_ks)))} × {Decimal(op.quantity)} = "
            f"{_q2(Decimal(str(op.product.base_ks)) * Decimal(op.quantity))}"
            for op in it.product_items.all()
        ]
        products_breakdown = NL.join(prod_lines) if prod_lines else "—"

        add_lines = [
            f"• {ai.addition.name} ×{_q2(ai.quantity)}: {_q2(Decimal(str(ai.total_ks() or 0)))}"
            for ai in it.addition_items.all()
        ]
        addons_breakdown = NL.join(add_lines) if add_lines else "—"

        coefs_breakdown = NL.join(coef_lines) if coef_lines else "—"

        tail = f"Кількість: {it.ks_qty}"
        if coef_terms:
            tail += f"\nКоефіцієнт: {it.ks_coef}"

        it.ks_tooltip = (
            f"ПРОДУКТИ:\n{products_breakdown}\n\n"
            f"СУМА продуктів: {it.ks_products} к/с\n\n"
            f"ДОПОВНЕННЯ:\n{addons_breakdown}\n\n"
            f"СУМА доповнень: {it.ks_adds} к/с\n\n"
            f"КОЕФІЦІЄНТИ:\n{coefs_breakdown}\n\n"
            f"{tail}"
        )

    parents_with_children = []
    for p in parents:
        kids = children_map.get(p.id, [])
        parents_with_children.append((p, kids))

    effective_ks = _q2(effective_ks)
    total_sum = _q2(total_sum)
    formula_expression = " + ".join(formula_terms) if formula_terms else "0.00"

    # Підсумок роботи цеху (без торгової націнки) — сума base_price по всіх позиціях
    workshop_total = _q2(sum(
        (getattr(it, "workshop_cost_value", Decimal("0")) for it in items),
        Decimal("0")
    ))
    markup_total = _q2(total_sum - workshop_total)

    global_coeffs = Coefficient.objects.filter(applies_globally=True).order_by("name")
    category_coeffs = Coefficient.objects.filter(applies_globally=False).order_by("name")

    addons_global = Addition.objects.filter(applies_globally=True).order_by("name")

    addons_by_category = []
    for cat in categories:
        cat_adds = (
            Addition.objects
            .filter(applies_globally=False, categories=cat)
            .distinct()
            .order_by("name")
        )
        addons_by_category.append({"cat": cat, "addons": cat_adds})

    customers = Customer.objects.all().order_by("-created_at")

    markers_by_image = {}
    for img in images:
        markers_qs = OrderImageMarker.objects.filter(image=img).select_related("item").order_by("id")
        markers_by_image[img.id] = [
            {
                "x": float(m.x),
                "y": float(m.y),
                "item_name": m.item.name if m.item else "",
                "color": (m.color or (get_item_color(m.item_id) if m.item_id else "#ff4d4f")),
            }
            for m in markers_qs
        ]

    order_name_templates = OrderNameDirectory.objects.all().order_by("name")

    context = {
        "order": order,
        "categories": categories,
        "products": products,
        "global_coeffs": global_coeffs,
        "category_coeffs": category_coeffs,
        "addons": addons_global,
        "addons_global": addons_global,
        "addons_by_category": addons_by_category,
        "rate": price_per_ks,
        "items": items,
        "parents": parents,
        "children_map": children_map,
        "parents_with_children": parents_with_children,
        "total": total_sum,
        "workshop_total": workshop_total,
        "markup_total": markup_total,
        "customers": customers,
        "markers_by_image": markers_by_image,
        "effective_ks": effective_ks,
        "formula_expression": formula_expression,
        "order_name_templates": order_name_templates,
    }

    if ajax_partial:
        html = render_to_string("doors/partials/order_items.html", context, request=request)
        return JsonResponse({"ok": True, "items_html": html})

    return render(request, "doors/calculate_order.html", context)

def _draw_common_header(p, width, height, company, base_font):
    """
    Спільна шапка: логотип + реквізити.
    """
    y_top = height - 20 * mm

    # Логотип: спершу з CompanyInfo.logo, якщо нема — пробуємо старий статичний
    logo_drawn = False
    if company and company.logo:
        try:
            logo = ImageReader(company.logo.path)
            p.drawImage(
                logo,
                20 * mm,
                y_top - 20 * mm,
                width=40 * mm,
                height=20 * mm,
                preserveAspectRatio=True,
                mask="auto",
            )
            logo_drawn = True
        except Exception:
            pass

    if not logo_drawn:
        logo_path = os.path.join(settings.BASE_DIR, "doors", "static", "doors", "logo.png")
        if os.path.exists(logo_path):
            p.drawImage(
                logo_path,
                20 * mm,
                y_top - 20 * mm,
                width=40 * mm,
                height=20 * mm,
                preserveAspectRatio=True,
                mask="auto",
            )

    # Реквізити справа
    p.setFont(base_font, 9)
    x_info = width - 20 * mm
    text = p.beginText()
    text.setTextOrigin(x_info, y_top)
    text.setFont(base_font, 10)

    if company:
        text.textLine(company.name)
        text.setFont(base_font, 9)
        if company.address:
            text.textLine(company.address)
        if company.phone:
            text.textLine(f"Тел.: {company.phone}")
        if company.email:
            text.textLine(f"Email: {company.email}")
        if company.website:
            text.textLine(f"Сайт: {company.website}")
        if company.iban:
            text.textLine(f"IBAN: {company.iban}")
        if company.edrpou:
            text.textLine(f"ЄДРПОУ: {company.edrpou}")
    p.drawText(text)


def _draw_variant_1(p, width, height, base_font, order, final_total):
    """
    Варіант 1: простий блок з фінальною вартістю.
    """
    y_start = height - 45 * mm  # трохи нижче шапки

    # Заголовок
    p.setFont(base_font, 14)
    title = f"Фінальний документ замовлення №{order.order_number}"
    p.drawString(30 * mm, y_start, title)

    p.setFont(base_font, 11)
    p.drawString(30 * mm, y_start - 8 * mm, f"Дата: {order.created_at.strftime('%d.%m.%Y')}")

    # Блок з ціною
    y_box_top = y_start - 25 * mm
    box_x1 = 30 * mm
    box_x2 = width - 30 * mm
    box_y1 = y_box_top
    box_y2 = y_box_top - 30 * mm

    p.setLineWidth(1)
    p.rect(box_x1, box_y2, box_x2 - box_x1, box_y1 - box_y2)

    p.setFont(base_font, 11)
    p.drawCentredString(width / 2, box_y1 - 7 * mm, "Фінальна вартість замовлення")

    p.setFont(base_font, 20)
    p.drawCentredString(width / 2, box_y1 - 18 * mm, f"{final_total:.2f} грн")

    # Маленька нотатка
    p.setFont(base_font, 8)
    p.drawString(
        30 * mm,
        box_y2 - 10 * mm,
        "Сума вказана з урахуванням узгодженої націнки та додаткових витрат (доставка, пакування тощо).",
    )


def _draw_variant_2(p, width, height, base_font, order, final_total):
    """
    Варіант 2: більш «комерційна пропозиція» з плашкою.
    """
    y_start = height - 45 * mm

    # Заголовок
    p.setFont(base_font, 16)
    title = f"Комерційна пропозиція №{order.order_number}"
    p.drawCentredString(width / 2, y_start, title)

    p.setFont(base_font, 11)
    p.drawCentredString(
        width / 2,
        y_start - 7 * mm,
        f"Дата: {order.created_at.strftime('%d.%m.%Y')}",
    )

    # Плашка з фінальною вартістю
    y_box_top = y_start - 25 * mm
    box_x1 = 35 * mm
    box_x2 = width - 35 * mm
    box_y1 = y_box_top
    box_y2 = y_box_top - 35 * mm

    p.setLineWidth(1.2)
    p.roundRect(box_x1, box_y2, box_x2 - box_x1, box_y1 - box_y2, 4 * mm)

    p.setFont(base_font, 10)
    p.drawCentredString(width / 2, box_y1 - 8 * mm, "Підсумкова вартість пропозиції")

    p.setFont(base_font, 22)
    p.drawCentredString(width / 2, box_y1 - 20 * mm, f"{final_total:.2f} грн")

    # Нотатка знизу
    p.setFont(base_font, 8)
    p.drawString(
        30 * mm,
        box_y2 - 10 * mm,
        "Дана комерційна пропозиція носить інформаційний характер. "
        "Умови оплати та поставки уточнюються з менеджером.",
    )


def _q2(x: Decimal) -> Decimal:
    return x.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def build_item_formula_parts(it):
    """
    Деталізація для internal PDF:
      products_sum, products_terms
      adds_sum, adds_terms
      qty
      coef
      ks_effective
      ks_formula
    """

    def _q2(x: Decimal) -> Decimal:
        return Decimal(x).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    products_sum = Decimal("0")
    prod_terms = []

    for op in it.product_items.select_related("product").all():
        p = op.product
        base = Decimal(str(p.base_ks or 0))
        qty_p = Decimal(str(op.quantity or 1))
        products_sum += base * qty_p
        prod_terms.append(f"{base:.2f}×{qty_p}")

    adds_sum = Decimal("0")
    add_terms = []

    for ai in it.addition_items.select_related("addition").all():
        qty_add = Decimal(str(getattr(ai, "quantity", 1) or 1))
        total_add = Decimal(str(ai.total_ks() or 0))
        adds_sum += total_add

        base_add = (total_add / qty_add) if qty_add > 0 else total_add
        base_add = _q2(base_add)
        add_terms.append(f"{base_add:.2f}×{qty_add}")

    qty = Decimal(str(it.quantity or 1))

    # ✅ правильний коеф: множення, базове значення 1.0
    coef = Decimal("1.0")
    coef_terms = []
    for c in it.coefficients.all():
        c_val = Decimal(str(c.value or 1))
        coef *= c_val
        coef_terms.append(f"{c_val:.2f}")

    products_formula = " + ".join(prod_terms) if prod_terms else "0.00"
    adds_formula = " + ".join(add_terms) if add_terms else "0.00"

    ks_base = (products_sum + adds_sum) * qty
    ks_effective = _q2(ks_base * coef)

    # ✅ Формула: коеф показуємо тільки якщо він реально є (тобто були вибрані коефи)
    coef_part = f" × {' × '.join(coef_terms)}" if coef_terms else ""
    ks_formula = f"(({products_formula}) + ({adds_formula})) × {qty}{coef_part}"

    return {
        "products_sum": _q2(products_sum),
        "products_terms": products_formula,
        "adds_sum": _q2(adds_sum),
        "adds_terms": adds_formula,
        "qty": qty,
        "coef": _q2(coef),
        "ks_effective": ks_effective,
        "ks_formula": ks_formula,
    }


def generate_pdf(request, order_id):
    """
    Генерація PDF по замовленню.

    Режими:
      - детальний (за замовчуванням) — таблиця з позиціями + окрема таблиця додаткових послуг
      - спрощений (?simple=1) — без основної таблиці, лише таблиця додаткових послуг (якщо є) + підсумки
      - внутрішній (?internal=1) — внутрішній розрахунок (копія таблиці + формула)

    Параметри GET:
      ?markup=10     — націнка, % (override для PDF; якщо не задано — береться з order/item)
      ?delivery=300  — доставка, грн
      ?packing=200   — пакування, грн
      ?simple=1      — спрощений варіант
      ?internal=1    — внутрішній PDF (ігнорує download)
      ?download=1    — скачати файл (для internal ігнорується)
    """
    order = Order.objects.select_related("customer").get(id=order_id)
    items_qs = OrderItem.objects.filter(order=order).prefetch_related(
        "coefficients",
        "addition_items__addition",
        "product_items__product",
    )
    items = list(items_qs)
    company = CompanyInfo.objects.first()

    # ---------- helpers ----------
    def to_decimal(v, default="0"):
        """
        Стабільний парсер числа:
        - підтримка коми
        - прибирає пробіли
        - прибирає символ % (щоб ?markup=10% не ламався)
        """
        if v in (None, ""):
            return Decimal(default)
        try:
            s = str(v).replace(",", ".").strip()
            s = s.replace("%", "").strip()
            return Decimal(s) if s else Decimal(default)
        except Exception:
            return Decimal(default)

    def _q2(x):
        return to_decimal(x, "0").quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    def safe_text(x: str) -> str:
        return strip_tags(str(x or "")).replace("\n", " ").strip()

    def fmt_qty(q: Decimal) -> str:
        """Гарне відображення кількості: 2.00 -> 2, 1.50 -> 1.5"""
        q = to_decimal(q, "0")
        if q == q.to_integral_value():
            return str(int(q))
        s = format(q.normalize(), "f")
        return s.rstrip("0").rstrip(".") if "." in s else s

    def extract_ks_from_obj(obj):
        """
        Повертає КС з об'єкта, якщо є (Order/OrderItem), або 0.
        """
        for attr in (
            "total_ks", "total_ks_cached", "ks_total", "total_ks_value",
            "ks_effective", "effective_ks", "ks_value", "ks"
        ):
            if hasattr(obj, attr):
                val = getattr(obj, attr)
                if callable(val):
                    val = val()
                return to_decimal(val, "0")
        return Decimal("0")

    def extract_markup_from_obj(obj):
        """
        Повертає націнку (%) з об'єкта, якщо поле названо по-різному.
        Важливо: ми відрізняємо "поля немає" від "значення 0".
        """
        candidates = ("markup_percent", "markup", "markup_percentage", "markup_value")
        for attr in candidates:
            if hasattr(obj, attr):
                val = getattr(obj, attr)
                if callable(val):
                    val = val()
                return True, to_decimal(val, "0")
        return False, Decimal("0")

    def calc_production_days_from_ks(total_ks: Decimal):
        """
        0.75 кс/год * 2 працівники * 8 год/день => 12 кс/день
        +30% запас. Округлення до цілого дня.
        """
        if total_ks is None or total_ks <= 0:
            return None

        hours_total = total_ks / Decimal("0.75")
        hours_per_day_all_workers = Decimal("2") * Decimal("8")
        days_raw = hours_total / hours_per_day_all_workers
        days_with_margin = days_raw * Decimal("1.3")

        days_int = int(days_with_margin.to_integral_value(rounding=ROUND_HALF_UP))
        return max(days_int, 1)

    # ---------- GET params ----------
    markup_percent_get = request.GET.get("markup")
    markup_override = None
    if markup_percent_get not in (None, ""):
        markup_override = to_decimal(markup_percent_get, "0")

    delivery = to_decimal(request.GET.get("delivery"), "0")
    packing = to_decimal(request.GET.get("packing"), "0")
    simple_mode = request.GET.get("simple") == "1"
    internal_mode = request.GET.get("internal") == "1"

    # ---------- базові підрахунки за один прохід ----------
    constructions_total = Decimal("0")
    positions_count = len(items)

    base_without_markup = Decimal("0")
    item_costs = []  # (it, raw_dec)

    total_ks_sum = Decimal("0")

    for it in items:
        qty = to_decimal(getattr(it, "quantity", 1) or 1, "1")
        constructions_total += qty

        raw = it.total_cost() if callable(getattr(it, "total_cost", None)) else getattr(it, "total_cost", 0)
        raw_dec = to_decimal(raw, "0")
        item_costs.append((it, raw_dec))
        base_without_markup += raw_dec

        total_ks_sum += extract_ks_from_obj(it)

    # Націнка для детального/спрощеного:
    # якщо ?markup= передали — застосувати її, інакше множник 1.0
    if markup_override is not None and markup_override > 0:
        markup_factor = (Decimal("100") + markup_override) / Decimal("100")
    else:
        markup_factor = Decimal("1.0")

    base_with_markup = (base_without_markup * markup_factor).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    final_total = (base_with_markup + delivery + packing).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    # ---------- розрахунок орієнтовного терміну виготовлення ----------
    total_ks = extract_ks_from_obj(order)

    if total_ks <= 0:
        total_ks = total_ks_sum

    # fallback (як у internal)
    if total_ks <= 0:
        ks_eff_sum = Decimal("0")
        for it in items:
            try:
                parts = build_item_formula_parts(it)
                ks_eff_sum += to_decimal(parts.get("ks_effective", 0), "0")
            except Exception:
                pass
        total_ks = ks_eff_sum

    production_days = calc_production_days_from_ks(total_ks) or 1

    # ---------- старт PDF ----------
    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    # ---------- шрифт ----------
    font_path = os.path.join(settings.BASE_DIR, "doors", "static", "fonts", "DejaVuSerif.ttf")
    if os.path.exists(font_path):
        pdfmetrics.registerFont(TTFont("DejaVuSerif", font_path))
        base_font = "DejaVuSerif"
    else:
        base_font = "Helvetica"

    # ---------- стилі ----------
    styles = getSampleStyleSheet()

    cell_style = ParagraphStyle(
        name="CellStyle",
        parent=styles["Normal"],
        fontName=base_font,
        fontSize=9,
        leading=11,
        alignment=TA_LEFT,
        wordWrap="CJK",
    )

    cell_center_style = ParagraphStyle(
        name="CellCenterStyle",
        parent=cell_style,
        alignment=1,
    )

    cell_right_style = ParagraphStyle(
        name="CellRightStyle",
        parent=cell_style,
        alignment=2,
    )

    formula_style = ParagraphStyle(
        name="FormulaStyle",
        parent=styles["Normal"],
        fontName=base_font,
        fontSize=8,
        leading=10,
        alignment=TA_LEFT,
        wordWrap="CJK",
    )

    # ---------- шапка ----------
    if company and getattr(company, "logo", None):
        try:
            logo = ImageReader(company.logo.path)
            p.drawImage(logo, 40, height - 140, width=160, preserveAspectRatio=True, mask="auto")
        except Exception:
            pass

    x_right = width - 40
    p.setFont(base_font, 12)
    if company:
        p.drawRightString(x_right, height - 60, safe_text(getattr(company, "name", "")))
        p.setFont(base_font, 10)
        if getattr(company, "address", None):
            p.drawRightString(x_right, height - 80, safe_text(company.address))
        if getattr(company, "phone", None):
            p.drawRightString(x_right, height - 100, f"Тел.: {safe_text(company.phone)}")
        if getattr(company, "email", None):
            p.drawRightString(x_right, height - 120, f"Email: {safe_text(company.email)}")
        if getattr(company, "edrpou", None):
            p.drawRightString(x_right, height - 140, f"ЄДРПОУ: {safe_text(company.edrpou)}")
        if getattr(company, "iban", None):
            p.drawRightString(x_right, height - 160, f"IBAN: {safe_text(company.iban)}")

    # ---------- заголовок ----------
    title_y = height - 170

    p.setFont(base_font, 15)
    if internal_mode:
        title = "Внутрішній розрахунок"
    else:
        title = "Комерційна пропозиція" if simple_mode else "Фінальний документ замовлення"
    p.drawString(40, title_y, title)

    p.setFont(base_font, 11)
    pdf_order_number = (request.GET.get("pdf_number") or "").strip()
    order_number = pdf_order_number or getattr(order, "order_number", str(order.id))
    created_at = getattr(order, "created_at", None)
    p.drawString(40, title_y - 20, f"Замовлення №: {order_number}")
    p.drawString(40, title_y - 38, f"Дата: {created_at.strftime('%d.%m.%Y') if created_at else ''}")

    # Замовник
    customer_name = ""
    if hasattr(order, "customer") and order.customer:
        customer_name = safe_text(getattr(order.customer, "name", "")) or safe_text(str(order.customer))
        p.drawString(40, title_y - 56, f"Замовник: {customer_name}")
    else:
        p.drawString(40, title_y - 56, "Замовник: ____________________")

    # Кількість конструкцій + кількість позицій
    p.drawString(40, title_y - 74, f"Кількість конструкцій у замовленні: {fmt_qty(constructions_total)}")
    p.setFont(base_font, 10)
    p.drawString(40, title_y - 90, f"Кількість позицій у замовленні: {positions_count}")

    current_y = title_y - 115

    # =====================================================================
    # ======================== INTERNAL MODE ===============================
    # =====================================================================
    if internal_mode:
        internal_items = items

        rate_obj = Rate.objects.first()
        current_rate = Decimal(str(rate_obj.price_per_ks)) if rate_obj else Decimal("0")

        if getattr(order, "price_per_ks", None) is None:
            order.price_per_ks = current_rate
            order.save(update_fields=["price_per_ks"])

        rate = Decimal(str(getattr(order, "price_per_ks", None) or current_rate))

        data = [[
            Paragraph("№", cell_center_style),
            Paragraph("Позиція", cell_center_style),
            Paragraph("Qty", cell_center_style),
            Paragraph("Формула", cell_center_style),
            Paragraph("К/С", cell_center_style),
            Paragraph("ТН%", cell_center_style),
            Paragraph("Без ТН", cell_center_style),
            Paragraph("Ціна з ТН", cell_center_style),
            Paragraph("ТН", cell_center_style),
        ]]

        effective_ks_sum = Decimal("0")
        total_sum = Decimal("0")
        total_qty_sum = Decimal("0")
        total_base_sum = Decimal("0")
        total_markup_sum = Decimal("0")

        idx = 1
        for it in internal_items:
            name = safe_text(getattr(it, "name", ""))
            qty_item = to_decimal(getattr(it, "quantity", 1) or 1, "1")

            parts = build_item_formula_parts(it)
            formula = safe_text(parts.get("ks_formula", ""))
            formula = (
                formula
                .replace(" x ", " ×\u200b")
                .replace(" + ", " +\u200b")
                .replace(" - ", " -\u200b")
                .replace(" / ", " /\u200b")
            )

            ks_base, coef = it.total_ks()
            ks_eff = _q2(Decimal(str(ks_base)) * Decimal(str(coef)))

            m = _q2(Decimal(str(it.effective_markup_percent() or 0)))
            final_price = _q2(Decimal(str(it.total_cost() or 0)))

            denom = Decimal("1") + (m / Decimal("100"))
            base_price = _q2(final_price / denom) if denom != 0 else Decimal("0")

            effective_ks_sum += ks_eff
            total_sum += final_price
            total_qty_sum += qty_item
            total_base_sum += base_price
            total_markup_sum += (final_price - base_price)

            mark_up = final_price - base_price

            data.append([
                Paragraph(str(idx), cell_center_style),
                Paragraph(name, cell_style),
                Paragraph(fmt_qty(qty_item), cell_center_style),
                Paragraph(formula, formula_style),
                Paragraph(f"{ks_eff:.2f}", cell_center_style),
                Paragraph(f"{m:.2f}", cell_right_style),
                Paragraph(f"{base_price:.2f}", cell_right_style),
                Paragraph(f"{final_price:.2f}", cell_right_style),
                Paragraph(f"{mark_up:.2f}", cell_right_style),
            ])
            idx += 1

        # Рядок "Разом" тільки для внутрішньої КП
        data.append([
            Paragraph("", cell_center_style),
            Paragraph("Разом", cell_style),
            Paragraph(fmt_qty(total_qty_sum), cell_center_style),
            Paragraph("", formula_style),
            Paragraph(f"{_q2(effective_ks_sum):.2f}", cell_center_style),
            Paragraph("", cell_right_style),
            Paragraph(f"{_q2(total_base_sum):.2f}", cell_right_style),
            Paragraph(f"{_q2(total_sum):.2f}", cell_right_style),
            Paragraph(f"{_q2(total_markup_sum):.2f}", cell_right_style),
        ])

        col_widths = [20, 80, 35, 140, 35, 30, 70, 70, 70]
        tbl = Table(data, colWidths=col_widths)
        tbl.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
            ("FONTNAME", (0, 0), (-1, -1), base_font),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0d6efd")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEADING", (0, 0), (-1, -1), 10),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#f2f2f2")),
        ]))

        _, h = tbl.wrap(0, 0)
        y = current_y - h
        if y < 140:
            p.showPage()
            current_y = height - 80
            y = current_y - h

        tbl.drawOn(p, 40, y)
        current_y = y - 20

        effective_ks_sum_q = _q2(effective_ks_sum)
        total_sum_q = _q2(total_sum)

        production_days_internal = calc_production_days_from_ks(effective_ks_sum) or 1

        if current_y < 160:
            p.showPage()
            current_y = height - 80

        p.setFont(base_font, 12)
        p.drawString(40, current_y, "Формула розрахунку")
        current_y -= 14

        p.setFont(base_font, 10)
        p.drawString(40, current_y, f"Σ к/с: {effective_ks_sum_q:.2f} к/с")
        current_y -= 16

        p.setFont(base_font, 11)
        p.drawString(40, current_y, f"(Σ позицій) × {_q2(rate):.2f} грн")
        current_y -= 16

        p.setFont(base_font, 12)
        p.drawString(40, current_y, f"= {total_sum_q:.2f} грн")
        current_y -= 18

        extras_total = _q2(delivery + packing)
        if extras_total > 0:
            p.setFont(base_font, 10)
            if delivery > 0:
                p.drawString(40, current_y, f"+ Доставка: {_q2(delivery):.2f} грн")
                current_y -= 14
            if packing > 0:
                p.drawString(40, current_y, f"+ Пакування: {_q2(packing):.2f} грн")
                current_y -= 14

            p.setFont(base_font, 12)
            p.drawString(40, current_y, f"Разом: {_q2(total_sum_q + extras_total):.2f} грн")
            current_y -= 18

        text = p.beginText()
        text.setTextOrigin(40, current_y)
        text.setFont(base_font, 9)
        text.setLeading(12)

        for line in [
            f"Орієнтовний термін виготовлення {production_days_internal} робочих днів.",
            "Дата початку робіт призначається за наявності матеріалу та проєкту",
            "на виготовлення замовлення і залежить від завантаження виробництва",
            "Якщо в процесі перевірки креслення виявиться, що не повністю",
            "розкритий обсяг робіт, невраховані роботи додатково збільшать",
            "вартість проєкту.",
        ]:
            text.textLine(line)
        p.drawText(text)

        p.showPage()
        p.save()
        buffer.seek(0)

        filename = f"Внутрішній_розрахунок_{order_number}.pdf"
        resp = HttpResponse(buffer, content_type="application/pdf")
        resp["Content-Disposition"] = f'inline; filename="{filename}"'
        return resp

    if not simple_mode and item_costs:
        main_data = [[
            Paragraph("№", cell_center_style),
            Paragraph("Позиція", cell_center_style),
            Paragraph("Кількість", cell_center_style),
            Paragraph("Вартість за одиницю, грн", cell_center_style),
            Paragraph("Сума, грн", cell_center_style),
        ]]

        for idx, (it, raw_dec) in enumerate(item_costs, start=1):
            total_with_markup = (raw_dec * markup_factor).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            qty = to_decimal(getattr(it, "quantity", 1) or 1, "1")
            unit_cost = (
                (total_with_markup / qty).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                if qty > 0
                else Decimal("0.00")
            )

            main_data.append([
                Paragraph(str(idx), cell_center_style),
                Paragraph(safe_text(getattr(it, "name", "")), cell_style),
                Paragraph(fmt_qty(qty), cell_center_style),
                Paragraph(f"{unit_cost:.2f}", cell_right_style),
                Paragraph(f"{total_with_markup:.2f}", cell_right_style),
            ])

        main_table = Table(main_data, colWidths=[30, 250, 60, 110, 90])
        main_table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
            ("FONTNAME", (0, 0), (-1, -1), base_font),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("BACKGROUND", (0, 0), (-1, 0), colors.white),
            ("TEXTCOLOR", (0, 0), (-1, -1), colors.black),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (1, 1), (1, -1), 6),
            ("RIGHTPADDING", (1, 1), (1, -1), 6),
        ]))

        _, main_h = main_table.wrap(0, 0)
        main_y = current_y - main_h
        if main_y < 60:
            main_y = 60
        main_table.drawOn(p, 40, main_y)
        current_y = main_y - 30

    extras_rows = []
    if delivery > 0:
        extras_rows.append(("Доставка", delivery))
    if packing > 0:
        extras_rows.append(("Пакування", packing))

    if extras_rows:
        extras_data = [[
            Paragraph("№", cell_center_style),
            Paragraph("Додаткові послуги", cell_center_style),
            Paragraph("Кількість", cell_center_style),
            Paragraph("Вартість за одиницю, грн", cell_center_style),
            Paragraph("Сума, грн", cell_center_style),
        ]]
        for idx, (name, value) in enumerate(extras_rows, start=1):
            val_str = f"{_q2(value):.2f}"
            extras_data.append([
                Paragraph(str(idx), cell_center_style),
                Paragraph(name, cell_style),
                Paragraph("1", cell_center_style),
                Paragraph(val_str, cell_right_style),
                Paragraph(val_str, cell_right_style),
            ])

        extras_table = Table(extras_data, colWidths=[30, 230, 70, 130, 80])
        extras_table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
            ("FONTNAME", (0, 0), (-1, -1), base_font),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
            ("BACKGROUND", (0, 0), (-1, 0), colors.white),
            ("TEXTCOLOR", (0, 0), (-1, -1), colors.black),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))

        _, extras_h = extras_table.wrap(0, 0)
        extras_y = current_y - extras_h
        if extras_y < 60:
            extras_y = 60
        extras_table.drawOn(p, 40, extras_y)
        current_y = extras_y - 30

    y_final = current_y - 10

    p.setFont(base_font, 14)
    p.drawString(40, y_final, f"Фінальна сума до оплати: {final_total:.2f} грн")

    disclaimer_y = y_final - 35
    text = p.beginText()
    text.setTextOrigin(40, disclaimer_y)
    text.setFont(base_font, 9)
    text.setLeading(12)

    for line in [
        f"Орієнтовний термін виготовлення {production_days} робочих днів.",
        "Дата початку робіт призначається за наявності матеріалу та проєкту",
        "на виготовлення замовлення і залежить від завантаження виробництва",
        "Якщо в процесі перевірки креслення виявиться, що не повністю",
        "розкритий обсяг робіт, невраховані роботи додатково збільшать",
        "вартість проєкту.",
    ]:
        text.textLine(line)
    p.drawText(text)

    p.showPage()
    p.save()
    buffer.seek(0)

    if customer_name:
        safe_name = customer_name.strip().replace(" ", "_")
        filename = f"order_{order_number}_{safe_name}.pdf"
    else:
        filename = f"order_{order_number}.pdf"

    if request.GET.get("download") == "1":
        return FileResponse(buffer, as_attachment=True, filename=filename, content_type="application/pdf")

    response = HttpResponse(buffer, content_type="application/pdf")
    response["Content-Disposition"] = f'inline; filename="{filename}"'
    return response

def worklog_list(request):
    # Отримуємо всі записи
    logs = WorkLog.objects.select_related("worker", "order").order_by("-date")
    workers = Worker.objects.all()
    orders = Order.objects.all()

    # Отримуємо параметри фільтрів
    worker_id = request.GET.get("worker")
    order_id = request.GET.get("order")
    start_date = request.GET.get("start_date")
    end_date = request.GET.get("end_date")

    # Фільтрація
    if worker_id:
        logs = logs.filter(worker_id=worker_id)
    if order_id:
        logs = logs.filter(order_id=order_id)
    if start_date:
        logs = logs.filter(date__gte=start_date)
    if end_date:
        logs = logs.filter(date__lte=end_date)

    # Підсумки годин по кожному працівнику
    totals = (
        logs.values("worker__name", "worker__position")
        .annotate(
            total_hours=Sum("hours"),
            work_hours=Sum("work_hours"),
        )
        .order_by("worker__name")
    )

    return render(request, "doors/worklog_list.html", {
        "logs": logs,
        "totals": totals,
        "workers": workers,
        "orders": orders,
        "selected_worker": worker_id,
        "selected_order": order_id,
        "start_date": start_date,
        "end_date": end_date,
    })


@require_POST
def worklog_delete(request, pk):
    log = get_object_or_404(WorkLog, pk=pk)

    worker_name = log.worker.name if log.worker else "—"
    log_date = log.date.strftime("%d.%m.%Y") if log.date else "—"

    log.delete()

    messages.success(
        request,
        f"Запис робочого часу працівника {worker_name} за {log_date} видалено."
    )

    redirect_url = request.META.get("HTTP_REFERER")
    return redirect(redirect_url if redirect_url else "worklog_list")


def report_view(request):
    start_date_raw = request.GET.get("start_date")
    end_date_raw = request.GET.get("end_date")
    export = request.GET.get("export")

    # Парсимо дати
    start_date = None
    end_date = None
    try:
        if start_date_raw:
            start_date = datetime.strptime(start_date_raw, "%Y-%m-%d").date()
        if end_date_raw:
            end_date = datetime.strptime(end_date_raw, "%Y-%m-%d").date()
    except ValueError:
        start_date = None
        end_date = None

    orders = Order.objects.all().order_by("-created_at")

    # Фільтрація по даті створення замовлення
    if start_date:
        orders = orders.filter(created_at__date__gte=start_date)
    if end_date:
        orders = orders.filter(created_at__date__lte=end_date)

    # Розрахунок % виконання на дату (end_date або зараз) по OrderProgress
    calc_date = end_date or date.today()
    for order in orders:
        op = (
            OrderProgress.objects
            .filter(order=order, date__lte=calc_date)
            .order_by("-date")
            .first()
        )
        if op:
            order.calculated_progress = float(op.percent)
        else:
            order.calculated_progress = float(order.completion_percent or 0)

    # Поділ на активні/відкладені
    active_orders = [o for o in orders if o.status != "postponed"]
    postponed_orders = [o for o in orders if o.status == "postponed"]

    # Загальна вартість по активним
    total_value = (
            Order.objects
            .filter(id__in=[o.id for o in active_orders])
            .aggregate(Sum("total_cost"))["total_cost__sum"]
            or Decimal("0")
    )

    # Середній % виконання по активним
    if active_orders:
        avg_progress = sum(o.calculated_progress for o in active_orders) / len(active_orders)
    else:
        avg_progress = 0

    # Експорт в Excel
    if export == "excel":
        wb = Workbook()
        ws = wb.active
        ws.title = "Звіт по замовленнях"
        ws.append(["№", "Номер", "Статус", "Вартість (грн)", "К/С", "Виконано (%)", "Дата"])

        for idx, order in enumerate(orders, start=1):
            ws.append([
                idx,
                order.order_number,
                order.get_status_display(),
                float(order.total_cost or 0),
                float(order.total_ks or 0),
                round(order.calculated_progress, 1),
                order.created_at.strftime("%d.%m.%Y"),
            ])

        ws.append([])
        ws.append([
            "",
            "Разом:",
            "",
            float(total_value),
            "",
            f"{avg_progress:.1f}%",
            "",
        ])

        response = HttpResponse(
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        response["Content-Disposition"] = f'attachment; filename="report_{datetime.now().strftime("%Y%m%d")}.xlsx"'
        wb.save(response)
        return response

    return render(request, "doors/report.html", {
        "orders": orders,
        "start_date": start_date,
        "end_date": end_date,
        "total_value": total_value,
        "avg_progress": avg_progress,
        "active_orders": active_orders,
        "postponed_orders": postponed_orders,
    })


def report_period_view(request):
    start_date_raw = request.GET.get("start_date")
    end_date_raw = request.GET.get("end_date")

    # Якщо дати не вибрані — останні 7 днів
    if not start_date_raw or not end_date_raw:
        end_date = datetime.today().date()
        start_date = end_date - timedelta(days=7)
    else:
        try:
            start_date = datetime.strptime(start_date_raw, "%Y-%m-%d").date()
            end_date = datetime.strptime(end_date_raw, "%Y-%m-%d").date()
        except ValueError:
            return HttpResponseBadRequest("Некоректні дати")

    if start_date > end_date:
        start_date, end_date = end_date, start_date

    # Норма виробітку (к/с за год). Можеш винести в settings або модель.
    HV = Decimal("0.75")

    # Години працівників
    logs = (
        WorkLog.objects
        .select_related("worker", "order")
        .filter(date__range=[start_date, end_date])
    )

    # Прогрес за період (щоб визначити перелік замовлень)
    progress_in_period_qs = (
        OrderProgress.objects
        .select_related("order")
        .filter(date__range=[start_date, end_date])
    )

    # ----- СПИСОК ЗАМОВЛЕНЬ (БЕЗ ДУБЛІВ) -----
    order_ids = progress_in_period_qs.values_list("order_id", flat=True).distinct()
    orders_qs = Order.objects.filter(id__in=order_ids).order_by("order_number")

    orders = [{"number": o.order_number, "name": o.order_name, "total_ks": o.total_ks} for o in orders_qs]

    # ----- СПИСОК ПРАЦІВНИКІВ (ID + name) -----
    workers = list(
        logs.order_by("worker__name")
        .values("worker_id", "worker__name")
        .distinct()
    )
    workers = [{"id": w["worker_id"], "name": w["worker__name"]} for w in workers]

    # ===== МАПА ГОДИН (date, worker_id) -> total hours =====
    hours_map = {
        (x["date"], x["worker_id"]): Decimal(str(x["total"] or 0))
        for x in logs.values("date", "worker_id").annotate(total=Sum("hours"))
    }

    # ===== Прогреси для замовлень до end_date (щоб дістати % на start/end і по днях) =====
    order_numbers = [o["number"] for o in orders]
    progress_all_qs = (
        OrderProgress.objects
        .filter(order__order_number__in=order_numbers, date__lte=end_date)
        .order_by("order__order_number", "date")
        .values("order__order_number", "date", "percent")
    )

    progress_by_order = {}
    for p in progress_all_qs:
        num = p["order__order_number"]
        progress_by_order.setdefault(num, []).append(
            (p["date"], Decimal(str(p["percent"])))
        )

    # helper: останній % <= target_date
    def get_last_percent(num, target_date: date) -> Decimal:
        arr = progress_by_order.get(num, [])
        last = Decimal("0")
        for d, perc in arr:
            if d <= target_date:
                last = perc
            else:
                break
        return last

    # ===== ТАБЛИЦЯ ПО ДНЯХ =====
    table = []
    totals_workers = {w["id"]: Decimal("0") for w in workers}
    total_all = Decimal("0")  # Tфакт (години за період)

    current_date = start_date
    while current_date <= end_date:
        row = {"date": current_date, "total": Decimal("0")}

        # % виконання по кожному замовленню станом на current_date
        for o in orders:
            num = o["number"]
            row[num] = float(get_last_percent(num, current_date))  # для красивого виводу

        # години по працівниках
        for w in workers:
            wid = w["id"]
            h = hours_map.get((current_date, wid), Decimal("0"))
            row[wid] = h
            totals_workers[wid] += h
            row["total"] += h

        total_all += row["total"]
        table.append(row)
        current_date += timedelta(days=1)

    # ===== ПІДСУМКОВИЙ % ПО ЗАМОВЛЕННЯХ (на end_date) =====
    totals_orders = {}
    for o in orders:
        num = o["number"]
        totals_orders[num] = float(get_last_percent(num, end_date))

    # ===== РОЗРАХУНОК ПО МЕТОДИЧЦІ ЗАМОВНИКА =====
    # Vза_період[i] = ( %кін - %до ) * Vz / 100
    # Vz беремо як total_ks (к/с) у Order
    orders_map = {o.order_number: o for o in orders_qs}

    start_percents = {num: get_last_percent(num, start_date) for num in order_numbers}
    end_percents = {num: get_last_percent(num, end_date) for num in order_numbers}

    vza_period_by_order = {}
    vzag = Decimal("0")

    for num in order_numbers:
        order_obj = orders_map.get(num)
        vz = Decimal(str(getattr(order_obj, "total_ks", 0) or 0))
        d_percent = end_percents[num] - start_percents[num]
        vza = (d_percent * vz) / Decimal("100")

        # якщо не хочемо від’ємні значення (регрес) — обрізаємо
        if vza < 0:
            vza = Decimal("0")

        vza_period_by_order[num] = vza
        vzag += vza

    t_norma = (vzag / HV) if HV > 0 else Decimal("0")  # години по нормі
    t_fact = total_all  # фактичні години
    eff_percent = (t_norma / t_fact * Decimal("100")) if t_fact > 0 else Decimal("0")

    work_days = (end_date - start_date).days + 1

    context = {
        "start_date": start_date,
        "end_date": end_date,
        "orders": orders,
        "workers": workers,
        "table": table,
        "totals_orders": totals_orders,
        "totals_workers": totals_workers,
        "total_all": total_all,  # Tфакт
        "work_days": work_days,

        # нові показники по методичці
        "hv": HV,
        "vzag": vzag,
        "t_norma": t_norma,
        "t_fact": t_fact,
        "eff_percent": eff_percent,
        "vza_period_by_order": vza_period_by_order,
    }

    # ===== Якщо просили PDF =====
    if request.GET.get("export") == "pdf" and table:
        # тут залишаєш свій існуючий код PDF без змін
        # (якщо треба — потім доповнимо PDF цими метриками)
        pass

    return render(request, "doors/report_period.html", context)


def worklog_add(request):
    if request.method == "POST":
        worker_id = request.POST.get("worker")
        hours = request.POST.get("hours")
        work_hours = request.POST.get("work_hours")  # НОВЕ
        comment = request.POST.get("comment", "")
        date_str = request.POST.get("date")

        # Перевірка дати
        if not date_str:
            date = datetime.today().date()
        else:
            try:
                date = datetime.strptime(date_str, "%Y-%m-%d").date()
            except ValueError:
                return HttpResponseBadRequest("Некоректна дата")

        # Перевірка чи вибрано працівника
        if not worker_id:
            return HttpResponseBadRequest("Оберіть працівника")

        # Створення запису
        WorkLog.objects.create(
            worker_id=worker_id,
            date=date,
            hours=hours,
            work_hours=work_hours or None,  # НОВЕ
            comment=comment,
        )

        return redirect("worklog_list")

    # Якщо GET — показуємо форму
    workers = Worker.objects.all()
    today = datetime.today()
    return render(request, "doors/worklog_add.html", {
        "workers": workers,
        "today": today,
    })


def order_history(request, order_id):
    order = get_object_or_404(Order, id=order_id)
    history = order.progress_logs.all()
    return render(request, "doors/partials/order_history.html", {"history": history, "order": order})


@csrf_exempt
def update_completion(request, order_id):
    """Оновлює відсоток виконання замовлення прямо зі списку."""
    if request.method == "POST":
        try:
            data = json.loads(request.body)
            percent = int(data.get("completion_percent", 0))
            order = Order.objects.get(id=order_id)
            order.completion_percent = max(0, min(100, percent))  # обмежуємо 0–100

            # якщо 100% — автоматично завершено
            if order.completion_percent == 100:
                order.status = "completed"

            order.save()
            return JsonResponse({"success": True, "percent": order.completion_percent})
        except Exception as e:
            return JsonResponse({"success": False, "error": str(e)}, status=400)

    return JsonResponse({"success": False, "error": "Invalid request"}, status=405)


def add_item_progress(request):
    """
    Додаємо прогрес по замовленню +
    позначаємо позиції, які неможливо виконати:
      - обрані позиції отримують status="impossible"
      - замовлення переходить у статус "postponed", якщо є хоча б одна така позиція
    Працюємо тільки із замовленнями в статусі "В роботі".
    """

    WORK_STATUS = "in_progress"

    order_id = request.GET.get("order") or request.POST.get("order_id")
    selected_order = (
        Order.objects.filter(id=order_id, status=WORK_STATUS).first()
        if order_id
        else None
    )

    if request.method == "POST":
        if not selected_order:
            messages.error(request, "Спочатку оберіть замовлення в статусі 'В роботі'.")
            return redirect("item_progress_add")

        form = OrderProgressForm(request.POST)
        if form.is_valid():
            progress = form.save(commit=False)
            progress.order = selected_order
            progress.save()

            problem_ids = request.POST.getlist("problem_items")
            has_problems = False

            if problem_ids:
                items_qs = OrderItem.objects.filter(
                    id__in=problem_ids,
                    order=selected_order,
                )

                if hasattr(progress, "problem_items"):
                    progress.problem_items.set(items_qs)

                items_qs.update(status="impossible")
                has_problems = items_qs.exists()

            selected_order.completion_percent = progress.percent
            fields_to_update = ["completion_percent"]

            if has_problems:
                selected_order.status = "postponed"
                fields_to_update.append("status")
            else:
                if progress.percent >= 100:
                    selected_order.status = "completed"
                    fields_to_update.append("status")

            selected_order.save(update_fields=fields_to_update)

            messages.success(
                request,
                f"Прогрес по замовленню №{selected_order.order_number} "
                f"оновлено до {progress.percent}%."
                + (" Замовлення відкладено через проблемні позиції." if has_problems else ""),
            )
            return redirect(f"{request.path}?order={selected_order.id}")
    else:
        form = OrderProgressForm()

    order_items_for_selection = (
        selected_order.items.all() if selected_order else OrderItem.objects.none()
    )

    latest_progress = (
        OrderProgress.objects
        .select_related("order")
        .prefetch_related("problem_items")
        .order_by("-date", "-id")[:10]
    )

    all_orders = (
        Order.objects
        .filter(status=WORK_STATUS)
        .order_by("-created_at")
    )

    return render(
        request,
        "doors/item_progress_add.html",
        {
            "form": form,
            "selected_order": selected_order,
            "order_items_for_selection": order_items_for_selection,
            "latest_progress": latest_progress,
            "all_orders": all_orders,
        },
    )


@require_POST
def delete_item_progress(request, pk):
    """
    Видаляє запис прогресу та перераховує стан замовлення.
    """
    progress = get_object_or_404(
        OrderProgress.objects.select_related("order").prefetch_related("problem_items"),
        pk=pk
    )
    order = progress.order

    # id позицій, які були прив'язані до цього запису
    deleted_problem_item_ids = list(progress.problem_items.values_list("id", flat=True))

    progress.delete()

    # Які позиції все ще позначені як проблемні в інших записах цього замовлення
    remaining_problem_item_ids = set(
        OrderProgress.objects.filter(order=order, problem_items__isnull=False)
        .values_list("problem_items__id", flat=True)
    )

    # Ті позиції, які були проблемними тільки в видаленому записі — повертаємо назад
    ids_to_restore = [
        item_id for item_id in deleted_problem_item_ids
        if item_id not in remaining_problem_item_ids
    ]

    if ids_to_restore:
        # Якщо у вас інший "звичайний" статус позиції — заміни "new" на свій
        OrderItem.objects.filter(id__in=ids_to_restore, order=order).update(status="new")

    # Беремо останній актуальний запис прогресу
    last_progress = (
        OrderProgress.objects.filter(order=order)
        .order_by("-date", "-id")
        .first()
    )

    has_problems = bool(remaining_problem_item_ids)

    if last_progress:
        order.completion_percent = last_progress.percent

        if has_problems:
            order.status = "postponed"
        elif last_progress.percent >= 100:
            order.status = "completed"
        else:
            order.status = "in_progress"
    else:
        order.completion_percent = 0
        order.status = "postponed" if has_problems else "in_progress"

    order.save(update_fields=["completion_percent", "status"])

    messages.success(request, f"Запис прогресу по замовленню №{order.order_number} видалено.")
    return redirect(f"{request.META.get('HTTP_REFERER', '') or '/'}")


def options_for_products(request):
    """
    GET /options-for-products/?ids=1&ids=3&ids=5
    Повертає лише ті доповнення/коефіцієнти, які підходять під вибрані продукти.
    """
    ids = request.GET.getlist("ids")
    products_qs = Product.objects.filter(id__in=ids)

    coeffs = _get_applicable_coefficients(products_qs)
    adds = _get_applicable_additions(products_qs)

    return JsonResponse({
        "coefficients": [
            {"id": c.id, "name": c.name, "value": c.value}
            for c in coeffs
        ],
        "additions": [
            {"id": a.id, "name": a.name, "ks_value": a.ks_value}
            for a in adds
        ],
    })


def order_item_edit(request, item_id):
    item = get_object_or_404(OrderItem, id=item_id)
    order = item.order

    all_products = Product.objects.select_related("category").all()
    all_additions = Addition.objects.all()
    all_coeffs = Coefficient.objects.all()

    def _to_decimal_or_one(val):
        try:
            v = Decimal(str(val).replace(",", "."))
            return max(v, Decimal("0.01"))
        except Exception:
            return Decimal("1")

    def _to_decimal_or_none(val):
        val = (val or "").strip().replace(",", ".")
        if val == "":
            return None
        try:
            return Decimal(val)
        except Exception:
            return None

    def _is_missing(val):
        return val is None or str(val).strip() == ""

    if request.method == "POST":
        # базові поля
        item.name = request.POST.get("name") or item.name
        qty_raw = request.POST.get("quantity", None)
        if _is_missing(qty_raw):
            # якщо не прийшло/пусто — не чіпаємо
            pass
        else:
            item.quantity = _to_decimal_or_one(qty_raw)

            # =========================
            # ФАСАД: редагування параметрів фасаду
            # =========================
            is_facade_item = item.product_items.filter(product__name="Фасад (розрахунок)").exists()

            if is_facade_item:
                facade_total_ks = _to_decimal_or_none(request.POST.get("facade_total_ks"))
                facade_json = request.POST.get("facade_data_json")

                if facade_json:
                    try:
                        item.facade_data = json.loads(facade_json)
                    except Exception:
                        item.facade_data = None

                if facade_total_ks is not None and facade_total_ks > 0:
                    carrier = item.product_items.filter(product__name="Фасад (розрахунок)").first()
                    if carrier:
                        carrier.quantity = facade_total_ks
                        carrier.save(update_fields=["quantity"])

                item.save()
                _recalc_order_totals(order)

                messages.success(request, "Фасадну позицію успішно оновлено ✅")
                return redirect("calculate_order", order_id=order.id)
        # =========================
        # ВИРОБИ з кількістю (OrderItemProduct)
        # =========================
        selected_products = set(request.POST.getlist("products"))  # рядки id

        existing_prod_map = {str(pi.product_id): pi for pi in item.product_items.all()}

        to_update = []
        to_create = []
        to_delete_ids = []

        for p in all_products:
            pid = str(p.id)
            qty_field = f"prod_qty_{p.id}"

            if pid in selected_products:
                qty_raw = request.POST.get(qty_field, None)

                # FIX: якщо qty не прийшов/пустий — залишаємо старе значення
                if _is_missing(qty_raw) and pid in existing_prod_map:
                    qty = existing_prod_map[pid].quantity
                else:
                    qty = _to_decimal_or_one(qty_raw or "1")

                if pid in existing_prod_map:
                    pi = existing_prod_map[pid]
                    pi.quantity = qty
                    to_update.append(pi)
                else:
                    to_create.append(
                        OrderItemProduct(order_item=item, product=p, quantity=qty)
                    )
            else:
                if pid in existing_prod_map:
                    to_delete_ids.append(existing_prod_map[pid].id)

        if to_delete_ids:
            OrderItemProduct.objects.filter(id__in=to_delete_ids).delete()
        if to_update:
            OrderItemProduct.objects.bulk_update(to_update, ["quantity"])
        if to_create:
            OrderItemProduct.objects.bulk_create(to_create)

        # =========================
        # КОЕФІЦІЄНТИ
        # =========================
        selected_coeffs = request.POST.getlist("coefficients")
        item.coefficients.set(selected_coeffs or [])

        # =========================
        # ДОПОВНЕННЯ з кількістю (AdditionItem)
        # =========================
        selected_adds = set(request.POST.getlist("additions"))
        existing_add_map = {str(ai.addition_id): ai for ai in item.addition_items.all()}

        to_update_add = []
        to_create_add = []
        to_delete_add_ids = []

        for add in all_additions:
            aid = str(add.id)
            qty_field = f"add_qty_{add.id}"

            if aid in selected_adds:
                qty_raw = request.POST.get(qty_field, None)

                # FIX: якщо qty не прийшов/пустий — залишаємо старе значення
                if _is_missing(qty_raw) and aid in existing_add_map:
                    qty = existing_add_map[aid].quantity
                else:
                    qty = _to_decimal_or_one(qty_raw or "1")

                if aid in existing_add_map:
                    ai = existing_add_map[aid]
                    ai.quantity = qty
                    to_update_add.append(ai)
                else:
                    to_create_add.append(
                        AdditionItem(order_item=item, addition=add, quantity=qty)
                    )
            else:
                if aid in existing_add_map:
                    to_delete_add_ids.append(existing_add_map[aid].id)

        if to_delete_add_ids:
            AdditionItem.objects.filter(id__in=to_delete_add_ids).delete()
        if to_update_add:
            AdditionItem.objects.bulk_update(to_update_add, ["quantity"])
        if to_create_add:
            AdditionItem.objects.bulk_create(to_create_add)

        item.save()
        _recalc_order_totals(order)

        messages.success(request, "Позицію успішно оновлено ✅")
        return redirect("calculate_order", order_id=order.id)

    # =========================
    # GET: підготовка стану для форми
    # =========================
    selected_products_ids = set(item.product_items.values_list("product_id", flat=True))
    product_qty = {pi.product_id: pi.quantity for pi in item.product_items.all()}

    selected_coeffs_ids = set(item.coefficients.values_list("id", flat=True))
    addition_qty = {ai.addition_id: ai.quantity for ai in item.addition_items.all()}
    is_facade_item = item.product_items.filter(product__name="Фасад (розрахунок)").exists()
    facade = item.facade_data or {}
    facade_total_ks = None

    carrier = item.product_items.filter(product__name="Фасад (розрахунок)").first()
    if carrier:
        facade_total_ks = carrier.quantity
    return render(
        request,
        "doors/order_item_edit.html",
        {
            "order": order,
            "item": item,
            "products": all_products,
            "coefficients": all_coeffs,
            "addons": all_additions,
            "selected_products_ids": selected_products_ids,
            "product_qty": product_qty,
            "selected_coeffs_ids": selected_coeffs_ids,
            "addition_qty": addition_qty,
            "is_facade_item": is_facade_item,
            "facade": facade,
            "facade_total_ks": facade_total_ks,
            "facade_data_json": json.dumps(facade, ensure_ascii=False) if facade else "",
        },
    )


def order_item_delete(request, item_id):
    item = get_object_or_404(OrderItem, id=item_id)
    order_id = item.order_id
    item.delete()
    # перерахунок після видалення
    order = get_object_or_404(Order, id=order_id)
    _recalc_order_totals(order)
    messages.info(request, "Позицію видалено.")
    return redirect("calculate_order", order_id=order_id)


@login_required
def annotate_order_image(request, image_id: int):
    """
    Сторінка розмітки конкретного фото замовлення:
    - вибір позиції (OrderItem)
    - клік по фото -> додається мітка
    - вибір кольору для міток
    - видалення окремих міток
    - «скинути мітки» тільки для обраної позиції (на фронті)
    Збереження: усі мітки цього зображення з фронту перезаписуються.
    Працює як з локальними, так і з M365 фото.
    """
    image = get_object_or_404(OrderImage, id=image_id)
    order = image.order
    items = order.items.all().order_by("id")

    # ✅ Формуємо URL на картинку: local або M365
    if getattr(image, "image", None):
        # ImageField існує, але може бути пустим
        if image.image:
            image_url = image.image.url
        else:
            image_url = reverse("m365_image_content", args=[image.id])
    else:
        # якщо раптом поля image нема (інша модель/міграції)
        image_url = reverse("m365_image_content", args=[image.id])

    # -------------------------
    # POST — зберігаємо всі мітки з JSON
    # -------------------------
    if request.method == "POST":
        markers_json = request.POST.get("markers_json") or "[]"
        print("markers_json:", markers_json)
        try:
            data = json.loads(markers_json)
            if not isinstance(data, list):
                data = []
        except json.JSONDecodeError:
            data = []

        # повністю чистимо мітки для цього зображення
        OrderImageMarker.objects.filter(image=image).delete()

        bulk = []
        for m in data:
            if not isinstance(m, dict):
                continue

            # координати приходять як % (0..100) — збережемо так само
            try:
                x = Decimal(str(m.get("x", 0)))
                y = Decimal(str(m.get("y", 0)))
            except Exception:
                continue

            # невеликий clamp, щоб не зберігати сміття
            if x < 0:
                x = Decimal("0")
            if y < 0:
                y = Decimal("0")
            if x > 100:
                x = Decimal("100")
            if y > 100:
                y = Decimal("100")

            item_id = m.get("item_id") or None
            color = m.get("color") or "#FF0000"

            bulk.append(
                OrderImageMarker(
                    image=image,
                    item_id=item_id,
                    x=x,
                    y=y,
                    color=color,
                )
            )

        if bulk:
            OrderImageMarker.objects.bulk_create(bulk)

        return redirect("calculate_order", order_id=order.id)

    # -------------------------
    # GET — збираємо всі існуючі мітки цього зображення
    # -------------------------
    markers_qs = (
        OrderImageMarker.objects.filter(image=image)
        .select_related("item")
        .order_by("id")
    )

    markers = [
        {
            "id": m.id,
            "x": float(m.x),
            "y": float(m.y),
            "item_id": m.item_id,
            "item_name": m.item.name if m.item else "Без позиції",
            "color": m.color or "#FF0000",
        }
        for m in markers_qs
    ]

    context = {
        "image": image,
        "order": order,
        "items": items,
        "markers": markers,
        "image_url": image_url,  # ✅ головне: у шаблоні використовуй тільки це
    }

    return render(request, "doors/annotate_order_image.html", context)


@csrf_exempt
def add_order_name(request):
    if request.method != "POST":
        return JsonResponse({"error": "Invalid request"}, status=400)

    data = json.loads(request.body)
    name = data.get("name", "").strip()

    if not name:
        return JsonResponse({"error": "Empty name"}, status=400)

    tpl, created = OrderNameDirectory.objects.get_or_create(name=name)

    return JsonResponse({"id": tpl.id, "name": tpl.name})


"""Teams"""


def _stream_graph_content(graph_url: str, access_token: str):
    """
    Стрімить контент з Microsoft Graph (щоб не тримати файл у памʼяті).
    """
    r = requests.get(
        graph_url,
        headers={"Authorization": f"Bearer {access_token}"},
        stream=True,
        timeout=60,
        allow_redirects=True,
    )

    # Важливо: Graph інколи відповідає 302 на pre-auth download URL — allow_redirects=True це покриває
    if not r.ok:
        return HttpResponse(r.text, status=r.status_code, content_type="text/plain")

    content_type = r.headers.get("Content-Type", "application/octet-stream")
    resp = StreamingHttpResponse(r.iter_content(chunk_size=1024 * 64), content_type=content_type)

    # корисно: кеш відключити, щоб завжди брало актуальне
    resp["Cache-Control"] = "no-store"
    return resp


def order_file_download(request, file_id: int):
    of = get_object_or_404(OrderFile, id=file_id)

    # 1) Якщо локальний файл
    if of.file:
        resp = StreamingHttpResponse(of.file.open("rb"), content_type="application/octet-stream")
        filename = of.file.name.split("/")[-1]
        resp["Content-Disposition"] = f'attachment; filename="{filename}"'
        return resp

    # 2) Якщо remote файл (M365)
    if of.source != "m365" or not of.remote_drive_id or not of.remote_item_id:
        raise Http404("File not available")

    token = get_app_token()
    url = f"https://graph.microsoft.com/v1.0/drives/{of.remote_drive_id}/items/{of.remote_item_id}/content"

    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, stream=True, timeout=180)
    if r.status_code >= 400:
        raise Http404(f"Graph error: {r.status_code}")

    filename = of.remote_name or "file"
    resp = StreamingHttpResponse(
        r.iter_content(chunk_size=1024 * 256),
        content_type=r.headers.get("Content-Type", "application/octet-stream"),
    )
    resp["Content-Disposition"] = f'attachment; filename="{filename}"'
    return resp


def _is_image_name(name: str | None) -> bool:
    if not name:
        return False
    ext = os.path.splitext(name.lower())[1]
    return ext in IMAGE_EXTS


def _stream_graph(url: str, token: str) -> StreamingHttpResponse:
    r = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        stream=True,
        timeout=60,
    )
    if r.status_code == 404:
        raise Http404("Remote file not found")
    if not r.ok:
        raise Http404(f"Graph error: {r.status_code}")

    resp = StreamingHttpResponse(
        streaming_content=r.iter_content(chunk_size=1024 * 256),
        status=200,
    )
    ct = r.headers.get("Content-Type")
    if ct:
        resp["Content-Type"] = ct
    return resp


@login_required
def m365_file_content(request, file_id: int):
    of = get_object_or_404(OrderFile, id=file_id)

    if of.source != "m365" or not of.remote_drive_id or not of.remote_item_id:
        raise Http404("Not an M365 file")

    token = get_app_token()

    # content stream
    url = f"https://graph.microsoft.com/v1.0/drives/{of.remote_drive_id}/items/{of.remote_item_id}/content"
    resp = _stream_graph(url, token)

    # filename (щоб нормально завантажувалось)
    filename = of.remote_name or of.description or "file"
    resp["Content-Disposition"] = f'inline; filename="{filename}"'
    return resp


def m365_download_bytes(*, drive_id: str, item_id: str):
    """
    Завантажує байти файлу з Microsoft Graph:
    GET /drives/{drive_id}/items/{item_id}/content
    Повертає (bytes, content_type)
    """
    token = get_app_token()  # у тебе вже є
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"

    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60, allow_redirects=True)
    if not r.ok:
        raise RuntimeError(f"Graph download failed HTTP {r.status_code}: {r.text}")

    content_type = r.headers.get("Content-Type")
    return r.content, content_type


@login_required
def m365_file_thumb(request, file_id: int):
    of = get_object_or_404(OrderFile, id=file_id)

    if of.source != "m365" or not of.remote_drive_id or not of.remote_item_id:
        raise Http404("Not an M365 file")

    # thumb має сенс тільки для картинок
    name = of.remote_name or of.description
    if not _is_image_name(name):
        raise Http404("Not an image")

    token = get_app_token()

    # Graph thumbnails: /thumbnails/0/medium/content
    url = (
        f"https://graph.microsoft.com/v1.0/drives/{of.remote_drive_id}"
        f"/items/{of.remote_item_id}/thumbnails/0/medium/content"
    )
    return _stream_graph(url, token)


@login_required
def m365_image_content(request, image_id: int):
    image = get_object_or_404(OrderImage, id=image_id)

    if not image.remote_drive_id or not image.remote_item_id:
        raise Http404("Not an M365 image")

    token = get_app_token()

    # /content — віддає байти файла
    graph_url = f"https://graph.microsoft.com/v1.0/drives/{image.remote_drive_id}/items/{image.remote_item_id}/content"
    return _stream_graph_content(graph_url, token)


@login_required
def m365_image_thumb(request, image_id: int):
    image = get_object_or_404(OrderImage, id=image_id)

    if not image.remote_drive_id or not image.remote_item_id:
        raise Http404("Not an M365 image")

    token = get_app_token()

    # thumbnails — дає меншу картинку, якщо є
    graph_url = (
        f"https://graph.microsoft.com/v1.0/drives/{image.remote_drive_id}"
        f"/items/{image.remote_item_id}/thumbnails/0/medium/content"
    )
    return _stream_graph_content(graph_url, token)


@login_required
@xframe_options_exempt
def order_file_inline(request, file_id):
    f = get_object_or_404(OrderFile, id=file_id)

    if f.source == "m365":
        content_bytes, content_type = m365_download_bytes(
            drive_id=f.remote_drive_id,
            item_id=f.remote_item_id,
        )
        resp = HttpResponse(content_bytes, content_type=content_type or "application/pdf")
        resp["Content-Disposition"] = f'inline; filename="{(f.remote_name or "file.pdf")}"'
        return resp

    if f.file:
        resp = HttpResponse(f.file.read(), content_type="application/pdf")
        resp["Content-Disposition"] = f'inline; filename="{f.file.name.split("/")[-1]}"'
        return resp

    return HttpResponse(status=404)


def _lower(s: str) -> str:
    return (s or "").lower()


def _is_folder(it: dict) -> bool:
    return "folder" in (it or {})


def _find_child_folder_by_contains(drive_id: str, parent_id: str, needle: str):
    n = _lower(needle)
    for x in list_children(drive_id, parent_id):
        if _is_folder(x) and n in _lower(x.get("name", "")):
            return x
    return None


def _unique_by_id(items: list[dict]) -> list[dict]:
    d = {}
    for it in items:
        if it and it.get("id"):
            d[it["id"]] = it
    return list(d.values())


def resolve_target_folders_for_normal_project(order: Order, mode: str) -> list[dict]:
    """
    mode:
      - "precalc": Проект -> 2-Комерційна пропозиція -> (усі папки з 'КП') -> 1 Розрахунок матеріалів -> (усі 'Для КС')
      - "final"  : Проект -> 4-Проектування -> (усі папки з 'Проект') -> (усі 'Для КС')
    Повертає список leaf-папок 'Для КС' (може бути багато).
    """
    drive_id = order.remote_drive_id
    root_id = order.remote_folder_id

    if not drive_id or not root_id:
        return []

    if mode == "precalc":
        # 2-Комерційна пропозиція
        f_cp = _find_child_folder_by_contains(drive_id, root_id, "2-Комерційна пропозиція")
        if not f_cp:
            return []

        # усі КП* всередині
        cps = [x for x in list_children(drive_id, f_cp["id"]) if _is_folder(x) and "кп" in _lower(x.get("name", ""))]

        result = []
        for cp in cps:
            f_calc = _find_child_folder_by_contains(drive_id, cp["id"], "1 Розрахунок матеріалів")
            if not f_calc:
                continue
            # знайти ВСІ "Для КС" всередині f_calc через list_children (без пошуку)
            children = list_children(drive_id, f_calc["id"]) or []
            result.extend([x for x in children if _is_folder(x) and "для кс" in _lower(x.get("name", ""))])
        return _unique_by_id(result)

    if mode == "final":
        # 4-Проектування
        f_proj = _find_child_folder_by_contains(drive_id, root_id, "4-Проектування")
        if not f_proj:
            return []

        # В роботу
        f_in_work = _find_child_folder_by_contains(drive_id, f_proj["id"], "В роботу")
        if not f_in_work:
            return []

        # Всі підпапки в "2 В роботу" (Проект 1, Проект 2, ...)
        project_folders = [x for x in list_children(drive_id, f_in_work["id"]) if _is_folder(x)]

        result = []
        for pf in project_folders:
            # беремо ВСІ "Для КС" всередині кожного проекту через list_children
            children = list_children(drive_id, pf["id"]) or []
            result.extend([
                x for x in children
                if _is_folder(x) and "для кс" in _lower(x.get("name", ""))
            ])

        return _unique_by_id(result)

    return []
def find_folder_contains_all(children, *needles: str):
    nn = [_lower(x) for x in needles if x]
    for it in children or []:
        if not _is_folder(it):
            continue
        name = _lower(it.get("name", ""))
        if all(n in name for n in nn):
            return it
    return None
def _pick_child_folder_contains(drive_id: str, parent_id: str, needle: str):
    needle_n = _norm_m365_name(needle)
    for it in list_children(drive_id, parent_id) or []:
        if _is_m365_folder(it) and needle_n in _norm_m365_name(it.get("name", "")):
            return it
    return None

def resolve_rework_destination_folder(drive_id: str, project_folder_id: str, is_final: bool = False) -> dict:
    """
    Для переробок:
      - precalc -> 2 КП попереднє
      - final   -> 4 КП в роботу

    Без варіантів і без вибору: шукаємо один конкретний folder.
    """
    folder_name = "4 КП в роботу" if is_final else "2 КП попереднє"

    children = list_children(drive_id, project_folder_id) or []

    for it in children:
        if not _is_m365_folder(it):
            continue

        name = (it.get("name") or "").strip()
        if name == folder_name:
            return it

    child_names = [x.get("name", "") for x in children if _is_m365_folder(x)]

    raise RuntimeError(
        f"Не знайдено цільову папку для переробки. "
        f"Очікується: '{folder_name}'. "
        f"Доступні папки: {child_names}"
    )

@require_POST
def sync_internal_pdf(request, order_id):
    """
    POST JSON:
      {
        "mode": "precalc" | "final",
        "markup": 10,
        "delivery": 300,
        "packing": 200
      }

    Синхронізує PDF в Teams/SharePoint у відповідні папки.

    Логіка папок:
      - work_type="project": шукає leaf-папки "Для КС" (може бути кілька)
      - work_type="rework": знаходить 1 destination папку в корені проєкту
    """
    order = Order.objects.filter(id=order_id).first()
    if not order:
        return JsonResponse({"ok": False, "error": "Order not found"}, status=404)

    if order.source != "m365" or not order.remote_drive_id or not order.remote_folder_id:
        return JsonResponse({"ok": False, "error": "Order is not linked to M365 project folder"}, status=400)

    try:
        payload = json.loads(request.body.decode("utf-8") or "{}")
    except Exception:
        payload = {}

    mode = (payload.get("mode") or "").strip().lower()
    if mode not in ("precalc", "final"):
        return JsonResponse({"ok": False, "error": "Invalid mode. Use 'precalc' or 'final'."}, status=400)

    def _num(v, default=0):
        try:
            if v is None or v == "":
                return default
            return float(v)
        except Exception:
            return default

    markup = _num(payload.get("markup", 0), 0)
    delivery = _num(payload.get("delivery", 0), 0)
    packing = _num(payload.get("packing", 0), 0)

    if order.work_type == "rework":
        is_final = (mode == "final")
        try:
            dest = resolve_rework_destination_folder(
                drive_id=order.remote_drive_id,
                project_folder_id=order.remote_folder_id,
                is_final=is_final,
            )
        except RuntimeError as e:
            return JsonResponse({"ok": False, "error": str(e)}, status=404)

        target_folders = [dest]
    else:
        target_folders = resolve_target_folders_for_normal_project(order, mode)
        chosen_id = (payload.get("target_folder_id") or "").strip()

        if order.work_type != "rework" and len(target_folders) > 1 and not chosen_id:
            return JsonResponse({
                "ok": False,
                "error": "Multiple target folders found. Please choose one.",
                "candidates": [
                    {"id": x.get("id"), "name": x.get("name"), "webUrl": x.get("webUrl")}
                    for x in target_folders
                    if x and x.get("id")
                ],
            }, status=409)

        if order.work_type != "rework" and chosen_id:
            by_id = {x["id"]: x for x in target_folders if x and x.get("id")}
            if chosen_id not in by_id:
                return JsonResponse({"ok": False, "error": "target_folder_id is not among candidates"}, status=400)
            target_folders = [by_id[chosen_id]]

    if not target_folders:
        return JsonResponse({
            "ok": False,
            "error": "Не знайдено цільових папок для синхронізації."
        }, status=404)

    valid_target_folders = [folder for folder in target_folders if folder and folder.get("id")]
    if not valid_target_folders:
        return JsonResponse({
            "ok": False,
            "error": "Цільові папки знайдені, але не містять коректного folder_id."
        }, status=400)

    def _render_pdf_bytes(get_params: dict) -> bytes:
        old_get = request.GET
        q = QueryDict(mutable=True)
        for k, v in get_params.items():
            q[k] = str(v)
        request.GET = q
        try:
            resp = generate_pdf(request, order_id)
            content = getattr(resp, "content", None)
            if content is None:
                content = b"".join(resp.streaming_content)
            return content or b""
        finally:
            request.GET = old_get

    base_params = {
        "markup": markup,
        "delivery": delivery,
        "packing": packing,
    }

    mode_label = "Попередній" if mode == "precalc" else "Фінальний"

    uploaded_folders = 0
    uploaded_files = []

    for folder in valid_target_folders:
        folder_id = folder.get("id")
        pdf_number = build_pdf_number(order, mode, folder)

        if order.work_type == "rework":
            pdfs = [
                (
                    "detailed",
                    {**base_params, "pdf_number": pdf_number},
                    f"{mode_label}_Детальний_{pdf_number}.pdf"
                ),
                (
                    "offer",
                    {**base_params, "simple": 1, "pdf_number": pdf_number},
                    f"{mode_label}_Комерційна_пропозиція_{pdf_number}.pdf"
                ),
                (
                    "internal",
                    {**base_params, "internal": 1, "pdf_number": pdf_number},
                    f"{mode_label}_Внутрішній_розрахунок_{pdf_number}.pdf"
                ),
            ]
        else:
            pdfs = [
                (
                    "internal",
                    {**base_params, "internal": 1, "pdf_number": pdf_number},
                    f"{mode_label}_Внутрішній_розрахунок_{pdf_number}.pdf"
                ),
            ]

        uploaded_any_in_folder = False

        for label, params, filename in pdfs:
            pdf_bytes = _render_pdf_bytes(params)
            if not pdf_bytes:
                return JsonResponse({"ok": False, "error": f"Failed to render PDF: {label}"}, status=500)

            upload_bytes_to_folder(
                drive_id=order.remote_drive_id,
                folder_id=folder_id,
                filename=filename,
                content=pdf_bytes,
                content_type="application/pdf",
            )
            uploaded_files.append(filename)
            uploaded_any_in_folder = True

        if uploaded_any_in_folder:
            uploaded_folders += 1

    if uploaded_folders == 0:
        return JsonResponse({
            "ok": False,
            "error": "PDF не було завантажено в жодну папку."
        }, status=400)

    return JsonResponse({
        "ok": True,
        "mode": mode,
        "work_type": order.work_type,
        "uploaded_to": uploaded_folders,
        "files": uploaded_files,
    })


