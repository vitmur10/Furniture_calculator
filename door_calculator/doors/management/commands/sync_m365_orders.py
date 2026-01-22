import hashlib
import os
import uuid
from typing import Iterable, List, Dict

from django.conf import settings
from django.core.management.base import BaseCommand, CommandError
from django.db import IntegrityError, transaction

from doors.models import Order, OrderFile, OrderImage
from doors.services.m365_graph import (
    find_site_by_display_name,
    pick_drive,
    list_root_children,
    list_children,
    search_in_folder,
)

IGNORED_CALC_FILE_MARKERS = [
    "попередній_",
    "фінальний_",
    "комерційна_пропозиція",
    "внутрішній_розрахунок",
    "_розрахунок_",
    "_кп_",
]

def is_own_generated_calc_file(filename: str) -> bool:
    name = (filename or "").lower()
    return any(marker in name for marker in IGNORED_CALC_FILE_MARKERS)

# ----------------------------
# Helpers
# ----------------------------

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


def make_order_number(folder_id: str) -> str:
    h = hashlib.sha1(folder_id.encode("utf-8")).hexdigest()[:10].upper()
    return h


# -------- folder pickers (single / all) --------

def pick_child_folder_contains(drive_id: str, parent_id: str, needle: str):
    """FIRST direct child folder whose name contains needle."""
    needle_n = _norm(needle)
    for it in list_children(drive_id, parent_id) or []:
        if _is_folder(it) and needle_n in _norm(it.get("name", "")):
            return it
    return None


def pick_child_folders_contains(drive_id: str, parent_id: str, needle: str) -> List[dict]:
    """ALL direct child folders whose name contains needle."""
    needle_n = _norm(needle)
    out = []
    for it in list_children(drive_id, parent_id) or []:
        if _is_folder(it) and needle_n in _norm(it.get("name", "")):
            out.append(it)
    return out


def pick_search_folder_contains(drive_id: str, parent_id: str, needle: str):
    """FIRST folder from Graph search results in folder parent_id whose name contains needle."""
    needle_n = _norm(needle)
    for it in search_in_folder(drive_id, parent_id, needle) or []:
        if _is_folder(it) and needle_n in _norm(it.get("name", "")):
            return it
    return None


def pick_search_folders_contains(drive_id: str, parent_id: str, needle: str) -> List[dict]:
    """ALL folders from Graph search results in folder parent_id whose name contains needle."""
    needle_n = _norm(needle)
    out = []
    for it in search_in_folder(drive_id, parent_id, needle) or []:
        if _is_folder(it) and needle_n in _norm(it.get("name", "")):
            out.append(it)
    return out


# -------- chain resolver (supports branching) --------

def resolve_leaf_folders_by_chain(drive_id: str, start_folder_id: str, chain: List[Dict]) -> List[dict]:
    """
    Returns LIST of leaf folder items.
    Supports branching steps that return multiple folders.

    Step types:
      - child_contains:       pick FIRST direct child containing value (per current folder)
      - child_all_contains:   pick ALL direct children containing value (per current folder)
      - search_contains:      pick FIRST search match containing value (per current folder)
      - search_all_contains:  pick ALL search matches containing value (per current folder)

    Algorithm:
      current_ids = [start_folder_id]
      for step in chain:
          next_folders = []
          for cid in current_ids:
              ... append found folder(s) ...
          current_ids = ids(next_folders)
      return last found folders (with id+name+webUrl if available)
    """
    current_ids = [start_folder_id]
    current_items = [{"id": start_folder_id}]  # best-effort items

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

            else:
                raise ValueError(f"Unknown step type: {step_type}")

        if not next_items:
            return []

        # de-dup by id (Graph search can return duplicates)
        seen = set()
        deduped = []
        for it in next_items:
            fid = it.get("id")
            if not fid or fid in seen:
                continue
            seen.add(fid)
            deduped.append(it)

        current_items = deduped
        current_ids = [it["id"] for it in deduped]

    return current_items


# -------- file iterators --------

def iter_files_direct(drive_id: str, folder_id: str) -> Iterable[dict]:
    for it in list_children(drive_id, folder_id) or []:
        if "file" in it:
            yield it


def iter_files_recursive(drive_id: str, folder_id: str) -> Iterable[dict]:
    stack = [folder_id]
    while stack:
        cur = stack.pop()
        for it in list_children(drive_id, cur) or []:
            if it.get("folder"):
                stack.append(it["id"])
                continue
            if "file" in it:
                yield it


def detect_work_type(site: dict, site_name_fallback: str = "") -> str:
    """
    Визначає, це "project" чи "rework" по назві сайту.
    Правило: якщо містить "Переробка" (case-sensitive як у тебе) -> rework.
    """
    site_display = (site.get("displayName", "") or site.get("name", "") or site_name_fallback or "").strip()

    # можна зробити більш стійко до регістру:
    # if "переробка" in site_display.lower():
    if "Переробка" in site_display:
        return "rework"
    return "project"


# ----------------------------
# Command
# ----------------------------

class Command(BaseCommand):
    help = (
        "Sync Orders from M365 (SharePoint) using folder chains.\n"
        "Projects are auto-detected as root folders.\n"
        "IMPORTANT: supports multiple 'КП*' and multiple 'Для КС*' folders (imports from ALL found leaf folders)."
    )

    def add_arguments(self, parser):
        parser.add_argument(
            "--limit",
            type=int,
            default=0,
            help="Limit number of root folders (projects) to process per site (0 = no limit).",
        )

    def handle(self, *args, **options):
        limit = int(options.get("limit") or 0)

        site_names = getattr(settings, "M365_SITE_DISPLAY_NAMES", None)
        if not site_names:
            site_names = [getattr(settings, "M365_SITE_DISPLAY_NAME", None)]
        site_names = [s for s in (site_names or []) if s]
        if not site_names:
            raise CommandError("No M365 site names configured (M365_SITE_DISPLAY_NAME(S))")

        drive_name = getattr(settings, "M365_DRIVE_NAME", None)
        if not drive_name:
            raise CommandError("M365_DRIVE_NAME is not set in settings.py")

        chains = getattr(settings, "M365_SYNC_CHAINS", None)
        if not chains or not isinstance(chains, dict):
            raise CommandError("M365_SYNC_CHAINS is empty or invalid in settings.py")

        level3_mode = getattr(settings, "M365_LEVEL3_MODE", "direct")
        if level3_mode not in ("direct", "recursive"):
            level3_mode = "direct"

        created_orders = 0
        updated_orders = 0
        created_files = 0
        created_images = 0
        skipped_root_folders_without_structure = 0

        for site_name in site_names:
            self.stdout.write(self.style.NOTICE(f"\n=== SITE: {site_name} ==="))

            site = find_site_by_display_name(site_name)
            if not site:
                self.stdout.write(self.style.WARNING(f"Site not found: {site_name}"))
                continue

            # ✅ тут визначаємо тип роботи
            work_type = detect_work_type(site, site_name_fallback=site_name)

            drive = pick_drive(site["id"], drive_name)
            if not drive:
                self.stdout.write(self.style.WARNING(f"Drive '{drive_name}' not found in site '{site_name}'"))
                continue

            drive_id = drive["id"]

            root_items = list_root_children(drive_id) or []
            project_candidates = [it for it in root_items if _is_folder(it)]
            if limit > 0:
                project_candidates = project_candidates[:limit]

            self.stdout.write(self.style.NOTICE(f"Root folder candidates: {len(project_candidates)}"))
            self.stdout.write(self.style.NOTICE(f"Detected work_type: {work_type}"))

            for folder in project_candidates:
                folder_id = folder["id"]
                folder_name = folder.get("name", "") or ""
                folder_url = folder.get("webUrl", "") or ""

                # Resolve leaf folders for each chain (can be MANY leaf folders now)
                chain_to_leafs: Dict[str, List[dict]] = {}
                for chain_name, chain in chains.items():
                    leafs = resolve_leaf_folders_by_chain(drive_id, folder_id, chain)
                    if leafs:
                        chain_to_leafs[chain_name] = leafs

                if not chain_to_leafs:
                    skipped_root_folders_without_structure += 1
                    continue

                # Create/update Order
                with transaction.atomic():
                    order = Order.objects.filter(source="m365", remote_folder_id=folder_id).first()

                    if not order:
                        try:
                            order = Order.objects.create(
                                order_name=folder_name,
                                order_number=make_order_number(folder_id),
                                source="m365",
                                work_type=work_type,  # ✅ NEW
                                remote_site_id=site["id"],
                                remote_drive_id=drive_id,
                                remote_folder_id=folder_id,
                                remote_web_url=folder_url,
                            )
                            created_orders += 1
                        except IntegrityError:
                            order = Order.objects.create(
                                order_name=folder_name,
                                order_number=f"M365-{uuid.uuid4().hex[:10].upper()}",
                                source="m365",
                                work_type=work_type,  # ✅ NEW
                                remote_site_id=site["id"],
                                remote_drive_id=drive_id,
                                remote_folder_id=folder_id,
                                remote_web_url=folder_url,
                            )
                            created_orders += 1
                    else:
                        changed = False
                        update_fields = []

                        if order.order_name != folder_name:
                            order.order_name = folder_name
                            changed = True
                            update_fields.append("order_name")

                        if order.remote_web_url != folder_url:
                            order.remote_web_url = folder_url
                            changed = True
                            update_fields.append("remote_web_url")

                        # ✅ NEW: оновлюємо тип
                        if getattr(order, "work_type", None) != work_type:
                            order.work_type = work_type
                            changed = True
                            update_fields.append("work_type")

                        if changed:
                            order.save(update_fields=update_fields)
                            updated_orders += 1

                    # Import from ALL leaf folders
                    for chain_name, leafs in chain_to_leafs.items():
                        self.stdout.write(self.style.NOTICE(
                            f"[{order.order_number}] {chain_name}: found {len(leafs)} leaf folder(s)"
                        ))

                        for leaf in leafs:
                            leaf_id = leaf["id"]
                            leaf_name = leaf.get("name") or leaf_id

                            # Choose file iteration mode
                            items_iter = (
                                iter_files_recursive(drive_id, leaf_id)
                                if level3_mode == "recursive"
                                else iter_files_direct(drive_id, leaf_id)
                            )

                            # Optional: show where we read from
                            self.stdout.write(self.style.NOTICE(
                                f"[{order.order_number}] {chain_name} -> {leaf_name}"
                            ))

                            for it in items_iter:
                                remote_item_id = it["id"]
                                file_name = it.get("name", "") or ""
                                web_url = it.get("webUrl", "") or ""
                                size = it.get("size") or 0

                                # ❌ пропускаємо файли, які ми самі заливаємо
                                if is_own_generated_calc_file(file_name):
                                    continue

                                web_url = it.get("webUrl", "") or ""
                                size = it.get("size") or 0
                                if _is_image(file_name):
                                    existing = OrderImage.objects.filter(
                                        remote_drive_id=drive_id,
                                        remote_item_id=remote_item_id,
                                    ).first()
                                    if existing:
                                        if existing.order_id != order.id:
                                            existing.order = order
                                            existing.save(update_fields=["order"])
                                        continue

                                    OrderImage.objects.create(
                                        order=order,
                                        remote_site_id=site["id"],
                                        remote_drive_id=drive_id,
                                        remote_item_id=remote_item_id,
                                        remote_web_url=web_url,
                                        remote_name=file_name,
                                        remote_size=size,
                                    )
                                    created_images += 1
                                else:
                                    existing_f = OrderFile.objects.filter(
                                        remote_drive_id=drive_id,
                                        remote_item_id=remote_item_id,
                                    ).first()
                                    if existing_f:
                                        if existing_f.order_id != order.id:
                                            existing_f.order = order
                                            existing_f.save(update_fields=["order"])
                                        continue

                                    OrderFile.objects.create(
                                        order=order,
                                        source="m365",
                                        remote_site_id=site["id"],
                                        remote_drive_id=drive_id,
                                        remote_item_id=remote_item_id,
                                        remote_web_url=web_url,
                                        remote_name=file_name,
                                        remote_size=size,
                                        description=file_name,
                                    )
                                    created_files += 1

        self.stdout.write(self.style.SUCCESS(
            "Done. "
            f"created_orders={created_orders}, updated_orders={updated_orders}, "
            f"created_images={created_images}, created_files={created_files}, "
            f"skipped_root_folders_without_structure={skipped_root_folders_without_structure}"
        ))
