from django.urls import path
from . import views
from django.conf import settings
from django.conf.urls.static import static

urlpatterns = [
    path("", views.order_list, name="home"),
    path("order/<int:order_id>/", views.calculate_order, name="calculate_order"),
    path("generate-pdf/<int:order_id>/", views.generate_pdf, name="generate_pdf"),
    path("update-status/<int:order_id>/", views.update_status, name="update_status"),
    path("worklog/", views.worklog_list, name="worklog_list"),
    path("report/", views.report_view, name="report_view"),
    path("report/period/", views.report_period_view, name="report_period"),
    path("worklog/", views.worklog_list, name="worklog_list"),
    path("worklog/add/", views.worklog_add, name="worklog_add"),
    path("update-completion/<int:order_id>/", views.update_completion, name="update_completion"),
    path("progress/add/", views.add_item_progress, name="item_progress_add"),
    path("options-for-products/", views.options_for_products, name="options_for_products"),
    path("order/item/<int:item_id>/edit/", views.order_item_edit, name="order_item_edit"),
    path("order/item/<int:item_id>/delete/", views.order_item_delete, name="order_item_delete"),
    path("order/<int:order_id>/delete/", views.delete_order, name="order_delete"),
    path("order-image/<int:image_id>/annotate/",views.annotate_order_image,name="annotate_order_image"),
    path("order-name/add/", views.add_order_name, name="add_order_name"),
    path("order-file/<int:file_id>/download/", views.order_file_download, name="order_file_download"),
    path("order-file/<int:file_id>/delete/", views.delete_order_file, name="delete_order_file"),
    path("m365/file/<int:file_id>/content/", views.m365_file_content, name="m365_file_content"),
    path("m365/file/<int:file_id>/thumb/", views.m365_file_thumb, name="m365_file_thumb"),
        # Додати для фото:
    path("m365/image/<int:image_id>/content/", views.m365_image_content, name="m365_image_content"),
    path("m365/image/<int:image_id>/thumb/", views.m365_image_thumb, name="m365_image_thumb"),
    path("m365/file/<int:file_id>/inline/", views.order_file_inline, name="order_file_inline"),

]

