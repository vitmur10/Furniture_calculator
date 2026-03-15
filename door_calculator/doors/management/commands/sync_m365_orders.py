import os
from typing import Iterable, List, Dict

from django.conf import settings
from django.core.management.base import BaseCommand, CommandError
from django.db import transaction
from django.utils import timezone
from django.utils.dateparse import parse_datetime

from doors.models import Order, OrderFile, OrderImage
from doors.services.m365_graph import (
    find_site_by_display_name,
    pick_drive,
    list_root_children,
    list_children,
    search_in_folder,
)

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp", ".tiff"}


def _norm(s: str) -> str:
    s = (s or "").lower()
    for ch in [" ", ".", ",", "-", "_", "’", "'", '"']:
        s = s.replace(ch, "")
    return s


def _is_folder(it: dict) -> bool:
    return bool(it and it.get("folder"))


def _is_image(name: str) -> bool:
    if not name:
        return False
    ext = os.path.splitext(name.lower())[1]
    return ext in IMAGE_EXTS


def _is_our_pdf(name: str) -> bool:
    if not name:
        return False

    name = name.lower()

    if not name.endswith(".pdf"):
        return False

    return (
        "попередній_" in name
        or "фінальний_" in name
    )


def _parse_m365_created_dt(item: dict):
    raw = (item or {}).get("createdDateTime")
    if not raw:
        return timezone.now()

    dt = parse_datetime(raw)
    if dt is None:
        return timezone.now()

    if timezone.is_naive(dt):
        dt = timezone.make_aware(dt, timezone.get_current_timezone())

    return timezone.localtime(dt)


def get_type_letter(site_name: str) -> str:
    return "П" if "перероб" in (site_name or "").lower() else "О"


def make_order_number_base(folder: dict, site_name: str) -> str:
    created_dt = _parse_m365_created_dt(folder)
    date_part = created_dt.strftime("%d.%m.%y")
    type_letter = get_type_letter(site_name)
    return f"{date_part}{type_letter}"


def make_unique_order_number(folder: dict, site_name: str) -> str:
    base = make_order_number_base(folder, site_name)

    existing_numbers = set(
        Order.objects.filter(order_number__startswith=base)
        .values_list("order_number", flat=True)
    )

    index = 1
    while f"{base}{index}" in existing_numbers:
        index += 1

    return f"{base}{index}"


def _safe_search_in_folder(drive_id: str, parent_id: str, needle: str) -> List[dict]:
    try:
        return search_in_folder(drive_id, parent_id, needle) or []
    except Exception:
        try:
            return list_children(drive_id, parent_id) or []
        except Exception:
            return []


def pick_child_folder_contains(drive_id: str, parent_id: str, needle: str):
    needle_n = _norm(needle)
    for it in list_children(drive_id, parent_id) or []:
        if _is_folder(it) and needle_n in _norm(it.get("name", "")):
            return it
    return None


def pick_child_folders_contains(drive_id: str, parent_id: str, needle: str) -> List[dict]:
    needle_n = _norm(needle)
    out = []
    for it in list_children(drive_id, parent_id) or []:
        if _is_folder(it) and needle_n in _norm(it.get("name", "")):
            out.append(it)
    return out


def pick_search_folder_contains(drive_id: str, parent_id: str, needle: str):
    needle_n = _norm(needle)
    for it in _safe_search_in_folder(drive_id, parent_id, needle):
        if _is_folder(it) and needle_n in _norm(it.get("name", "")):
            return it
    return None


def pick_search_folders_contains(drive_id: str, parent_id: str, needle: str) -> List[dict]:
    needle_n = _norm(needle)
    out = []
    for it in _safe_search_in_folder(drive_id, parent_id, needle):
        if _is_folder(it) and needle_n in _norm(it.get("name", "")):
            out.append(it)
    return out


def resolve_leaf_folders_by_chain(drive_id: str, start_folder_id: str, chain: List[Dict]) -> List[dict]:
    current_ids = [start_folder_id]

    for step in chain:
        step_type = step["type"]
        value = step["value"]

        next_items: List[dict] = []

        for cid in current_ids:

            if step_type == "child_contains":
                found = pick_child_folder_contains(drive_id, cid, value)
                if found:
                    next_items.append(found)

            elif step_type == "child_all_contains":
                next_items.extend(pick_child_folders_contains(drive_id, cid, value))

            elif step_type == "search_contains":
                found = pick_search_folder_contains(drive_id, cid, value)
                if found:
                    next_items.append(found)

            elif step_type == "search_all_contains":
                next_items.extend(pick_search_folders_contains(drive_id, cid, value))

        if not next_items:
            return []

        current_ids = [it["id"] for it in next_items]

    return next_items


def iter_files_direct(drive_id: str, folder_id: str) -> Iterable[dict]:
    for it in list_children(drive_id, folder_id) or []:
        if "file" in it:
            yield it


class Command(BaseCommand):
    help = "Sync Orders from Microsoft 365 (Teams/SharePoint)"

    def handle(self, *args, **options):

        site_names = getattr(settings, "M365_SITE_DISPLAY_NAMES", None)
        if not site_names:
            site_names = [getattr(settings, "M365_SITE_DISPLAY_NAME", None)]

        drive_name = getattr(settings, "M365_DRIVE_NAME", None)
        chains = getattr(settings, "M365_SYNC_CHAINS", None)

        if not site_names:
            raise CommandError("M365 site not configured")

        if not drive_name:
            raise CommandError("M365_DRIVE_NAME missing")

        if not chains:
            raise CommandError("M365_SYNC_CHAINS missing")

        created_orders = 0
        deleted_orders = 0
        created_files = 0
        created_images = 0

        for site_name in site_names:

            site = find_site_by_display_name(site_name)
            if not site:
                continue

            drive = pick_drive(site["id"], drive_name)
            if not drive:
                continue

            drive_id = drive["id"]

            root_items = list_root_children(drive_id) or []
            project_folders = [it for it in root_items if _is_folder(it)]

            current_root_ids = {it["id"] for it in project_folders}

            for folder in project_folders:

                folder_id = folder["id"]
                folder_name = folder.get("name", "")
                folder_url = folder.get("webUrl", "")

                chain_leafs = {}

                for chain_name, chain in chains.items():
                    leafs = resolve_leaf_folders_by_chain(drive_id, folder_id, chain)
                    if leafs:
                        chain_leafs[chain_name] = leafs

                if not chain_leafs:
                    continue

                with transaction.atomic():

                    order = Order.objects.filter(
                        source="m365",
                        remote_folder_id=folder_id,
                    ).first()

                    if not order:
                        order = Order.objects.create(
                            order_name=folder_name,
                            order_number=make_unique_order_number(folder, site_name),
                            source="m365",
                            remote_site_id=site["id"],
                            remote_drive_id=drive_id,
                            remote_folder_id=folder_id,
                            remote_web_url=folder_url,
                        )

                        created_orders += 1

                    for leafs in chain_leafs.values():

                        for leaf in leafs:

                            for it in iter_files_direct(drive_id, leaf["id"]):

                                file_id = it["id"]
                                name = it.get("name", "")
                                web = it.get("webUrl", "")

                                if _is_image(name):

                                    if not OrderImage.objects.filter(
                                        remote_item_id=file_id
                                    ).exists():

                                        OrderImage.objects.create(
                                            order=order,
                                            remote_site_id=site["id"],
                                            remote_drive_id=drive_id,
                                            remote_item_id=file_id,
                                            remote_web_url=web,
                                            remote_name=name,
                                        )

                                        created_images += 1

                                else:

                                    if _is_our_pdf(name):
                                        continue

                                    if not OrderFile.objects.filter(
                                        remote_item_id=file_id
                                    ).exists():

                                        OrderFile.objects.create(
                                            order=order,
                                            source="m365",
                                            remote_site_id=site["id"],
                                            remote_drive_id=drive_id,
                                            remote_item_id=file_id,
                                            remote_web_url=web,
                                            remote_name=name,
                                        )

                                        created_files += 1

            stale_orders = Order.objects.filter(
                source="m365",
                remote_site_id=site["id"],
                remote_drive_id=drive_id,
            ).exclude(
                remote_folder_id__in=current_root_ids
            )

            for order in stale_orders:
                order.delete()
                deleted_orders += 1

        self.stdout.write(self.style.SUCCESS(
            f"Sync finished. "
            f"created_orders={created_orders}, "
            f"created_files={created_files}, "
            f"created_images={created_images}, "
            f"deleted_orders={deleted_orders}"
        ))