# -*- coding: utf-8 -*-
"""
Script:   AUK Window Daylighting Compliance Checker
Desc:     Checks each room against the 1/20th window area requirement per Building Regs Part F/L.
          Includes both OST_Windows and glazed curtain wall panels in clear opening calculation.
Author:   Aukett Swanke Digital
Usage:    Open a project with placed Rooms and Windows/Curtain Walls (host or linked). Run with no selection.
Result:   Summary table per room + per-window/panel detail listing with mark values.
"""
from __future__ import division
from pyrevit import revit, DB
from pyrevit import script
from pyrevit import forms

logger = script.get_logger()
output = script.get_output()
doc    = revit.doc

# ── Constants ─────────────────────────────────────────────────────────────────
RATIO_REQUIREMENT = 1.0 / 20.0
PASS_LABEL        = u"PASS"
FAIL_LABEL        = u"FAIL"
WARN_LABEL        = u"NO OPENINGS"
SQ_FT_TO_SQ_M     = 0.092903

WIDTH_PARAM_NAMES  = [
    "Width", "Rough Width", "Frame Width", "Overall Width",
    "Window Width", "Clear Width", "Opening Width",
]
HEIGHT_PARAM_NAMES = [
    "Height", "Rough Height", "Frame Height", "Overall Height",
    "Window Height", "Clear Height", "Opening Height",
]

_BUILTIN_WIDTH  = [
    DB.BuiltInParameter.FAMILY_WIDTH_PARAM,
    DB.BuiltInParameter.GENERIC_WIDTH,
    DB.BuiltInParameter.DOOR_WIDTH,
]
_BUILTIN_HEIGHT = [
    DB.BuiltInParameter.FAMILY_HEIGHT_PARAM,
    DB.BuiltInParameter.GENERIC_HEIGHT,
    DB.BuiltInParameter.DOOR_HEIGHT,
]

_S4_OFFSETS = [
    ( 0.984,  0.0), (-0.984,  0.0),
    ( 0.0,  0.984), ( 0.0,  -0.984),
    ( 1.969,  0.0), (-1.969,  0.0),
    ( 0.0,  1.969), ( 0.0,  -1.969),
    ( 0.984,  0.984), (-0.984,  0.984),
    ( 0.984, -0.984), (-0.984, -0.984),
]

# Keywords indicating a glazed curtain panel (case-insensitive check)
_GLAZED_KEYWORDS = [
    "glaz", "glass", "transparent", "vision", "clear",
    "curtain panel", "system panel",
]

# Keywords indicating an opaque/non-glazed panel (exclusion list)
_OPAQUE_KEYWORDS = [
    "opaque", "solid", "spandrel", "insulated", "metal",
    "stone", "timber", "wood", "blank", "infill",
]


# ── Helpers ───────────────────────────────────────────────────────────────────
def to_sqm(v):
    return float(v) * SQ_FT_TO_SQ_M


def to_mm(ft):
    """Convert internal feet to mm for display."""
    return float(ft) * 304.8


def eid_int(eid):
    try:
        return int(eid.Value)
    except AttributeError:
        return int(eid.IntegerValue)


def safe_name(element):
    """Safely get element Name across IronPython contexts."""
    try:
        return element.Name
    except Exception:
        pass
    try:
        return DB.Element.Name.__get__(element)
    except Exception:
        pass
    return u""


def _read_named_param(element, names):
    for name in names:
        try:
            p = element.LookupParameter(name)
            if p and p.StorageType == DB.StorageType.Double and p.HasValue:
                v = float(p.AsDouble())
                if v > 0.0:
                    return v
        except Exception:
            pass
    return 0.0


def _read_builtin_param(element, bips):
    for bip in bips:
        try:
            p = element.get_Parameter(bip)
            if p and p.HasValue:
                v = float(p.AsDouble())
                if v > 0.0:
                    return v
        except Exception:
            pass
    return 0.0


def get_window_symbol(instance, linked_doc):
    try:
        type_id = instance.GetTypeId()
        if type_id and eid_int(type_id) > 0:
            owner = linked_doc if linked_doc is not None else doc
            return owner.GetElement(type_id)
    except Exception:
        pass
    try:
        return instance.Symbol
    except Exception:
        pass
    return None


def get_window_mark(instance, linked_doc):
    """
    Return the Mark value of a window/panel instance.
    Tries instance ALL_MODEL_MARK built-in first, then 'Mark' named param,
    then falls back to the type name.
    """
    # Instance mark (built-in)
    try:
        p = instance.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
        if p and p.HasValue:
            val = p.AsString()
            if val and val.strip():
                return val.strip()
    except Exception:
        pass
    # Named instance param
    try:
        p = instance.LookupParameter("Mark")
        if p and p.HasValue:
            val = p.AsString()
            if val and val.strip():
                return val.strip()
    except Exception:
        pass
    # Fall back to type name
    symbol = get_window_symbol(instance, linked_doc)
    if symbol:
        try:
            return symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or u"-"
        except Exception:
            pass
    return u"-"


def get_window_type_name(instance, linked_doc):
    """Return family : type name string."""
    symbol = get_window_symbol(instance, linked_doc)
    if symbol is None:
        return u"-"
    try:
        type_name   = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or u""
        family_name = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString() or u""
        if family_name and type_name:
            return u"{} : {}".format(family_name, type_name)
        return type_name or family_name or u"-"
    except Exception:
        return u"-"


def window_clear_opening_sqm(instance, linked_doc=None):
    """Returns (area_sqm, w_ft, h_ft, source_label)."""
    symbol = get_window_symbol(instance, linked_doc)

    w = _read_named_param(instance, WIDTH_PARAM_NAMES)
    if w <= 0.0 and symbol:
        w = _read_named_param(symbol, WIDTH_PARAM_NAMES)

    h = _read_named_param(instance, HEIGHT_PARAM_NAMES)
    if h <= 0.0 and symbol:
        h = _read_named_param(symbol, HEIGHT_PARAM_NAMES)

    if w > 0.0 and h > 0.0:
        return to_sqm(w * h), w, h, u"param"

    if w <= 0.0:
        w = _read_builtin_param(instance, _BUILTIN_WIDTH)
    if w <= 0.0 and symbol:
        w = _read_builtin_param(symbol, _BUILTIN_WIDTH)
    if h <= 0.0:
        h = _read_builtin_param(instance, _BUILTIN_HEIGHT)
    if h <= 0.0 and symbol:
        h = _read_builtin_param(symbol, _BUILTIN_HEIGHT)

    if w > 0.0 and h > 0.0:
        return to_sqm(w * h), w, h, u"builtin"

    # Bounding-box XY fallback
    try:
        bb = instance.get_BoundingBox(None)
        if bb:
            dx = abs(float(bb.Max.X - bb.Min.X))
            dy = abs(float(bb.Max.Y - bb.Min.Y))
            dz = abs(float(bb.Max.Z - bb.Min.Z))
            xy_max = max(dx, dy)
            return to_sqm(xy_max * dz), xy_max, dz, u"bbox"
    except Exception:
        pass

    return 0.0, 0.0, 0.0, u"none"


# ── Curtain Panel Helpers ─────────────────────────────────────────────────────
def _get_panel_full_name(panel, linked_doc):
    """Return lowercase combined family:type name for keyword matching."""
    symbol = get_window_symbol(panel, linked_doc)
    parts = []
    if symbol:
        try:
            fn = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString()
            if fn:
                parts.append(fn)
        except Exception:
            pass
        try:
            tn = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            if tn:
                parts.append(tn)
        except Exception:
            pass
    if not parts:
        n = safe_name(panel)
        if n:
            parts.append(n)
    return u" ".join(parts).lower()


def _check_material_transparency(panel, linked_doc):
    """
    Check if the panel's material has transparency > 0 (indicates glass).
    Returns True if glazed, False if opaque, None if indeterminate.
    """
    owner = linked_doc if linked_doc is not None else doc
    try:
        mat_ids = panel.GetMaterialIds(False)
        for mat_id in mat_ids:
            mat = owner.GetElement(mat_id)
            if mat is None:
                continue
            try:
                transparency = mat.Transparency
                if transparency > 0:
                    return True
            except Exception:
                pass
            # Check material name for glass keywords
            mat_name = safe_name(mat).lower()
            if any(kw in mat_name for kw in ["glass", "glaz", "transparent", "clear"]):
                return True
    except Exception:
        pass
    return None


def is_glazed_panel(panel, linked_doc):
    """
    Determine if a curtain panel is glazed (contributes to clear opening).
    Strategy:
      1. Exclude if family/type name contains opaque keywords
      2. Include if family/type name contains glazed keywords
      3. Check material transparency as tiebreaker
      4. Default System Panel (empty wall / no family) is treated as glazed
    """
    full_name = _get_panel_full_name(panel, linked_doc)

    # Exclude known opaque types
    for kw in _OPAQUE_KEYWORDS:
        if kw in full_name:
            return False

    # Include known glazed types
    for kw in _GLAZED_KEYWORDS:
        if kw in full_name:
            return True

    # Check if it's the default "System Panel" (basic curtain wall glazing)
    if "system panel" in full_name or full_name == u"":
        return True

    # Material transparency check
    mat_result = _check_material_transparency(panel, linked_doc)
    if mat_result is not None:
        return mat_result

    # If no name match and no material info, conservatively exclude
    return False


def curtain_panel_area_sqm(panel, linked_doc=None):
    """
    Returns (area_sqm, w_ft, h_ft, source_label) for a curtain panel.
    Resolution order:
      1. CURTAIN_WALL_PANELS_AREA built-in parameter (most accurate)
      2. HOST_AREA_COMPUTED built-in
      3. Width × Height named/built-in params
      4. Bounding-box fallback
    """
    # Strategy 1: Panel area parameter (sq ft internal)
    for bip in [DB.BuiltInParameter.HOST_AREA_COMPUTED]:
        try:
            p = panel.get_Parameter(bip)
            if p and p.HasValue:
                area_sqft = float(p.AsDouble())
                if area_sqft > 0.0:
                    return to_sqm(area_sqft), 0.0, 0.0, u"area_param"
        except Exception:
            pass

    # Strategy 1b: Named "Area" parameter
    try:
        p = panel.LookupParameter("Area")
        if p and p.StorageType == DB.StorageType.Double and p.HasValue:
            area_sqft = float(p.AsDouble())
            if area_sqft > 0.0:
                return to_sqm(area_sqft), 0.0, 0.0, u"area_param"
    except Exception:
        pass

    # Strategy 2: Width × Height (reuse window logic)
    symbol = get_window_symbol(panel, linked_doc)

    w = _read_named_param(panel, WIDTH_PARAM_NAMES)
    if w <= 0.0 and symbol:
        w = _read_named_param(symbol, WIDTH_PARAM_NAMES)

    h = _read_named_param(panel, HEIGHT_PARAM_NAMES)
    if h <= 0.0 and symbol:
        h = _read_named_param(symbol, HEIGHT_PARAM_NAMES)

    if w > 0.0 and h > 0.0:
        return to_sqm(w * h), w, h, u"param"

    # Built-in width/height
    if w <= 0.0:
        w = _read_builtin_param(panel, _BUILTIN_WIDTH)
    if w <= 0.0 and symbol:
        w = _read_builtin_param(symbol, _BUILTIN_WIDTH)
    if h <= 0.0:
        h = _read_builtin_param(panel, _BUILTIN_HEIGHT)
    if h <= 0.0 and symbol:
        h = _read_builtin_param(symbol, _BUILTIN_HEIGHT)

    if w > 0.0 and h > 0.0:
        return to_sqm(w * h), w, h, u"builtin"

    # Strategy 3: Bounding-box — use the two largest dimensions (skip wall depth)
    try:
        bb = panel.get_BoundingBox(None)
        if bb:
            dx = abs(float(bb.Max.X - bb.Min.X))
            dy = abs(float(bb.Max.Y - bb.Min.Y))
            dz = abs(float(bb.Max.Z - bb.Min.Z))
            dims = sorted([dx, dy, dz], reverse=True)
            # Two largest = panel face dimensions, smallest = thickness
            return to_sqm(dims[0] * dims[1]), dims[0], dims[1], u"bbox"
    except Exception:
        pass

    return 0.0, 0.0, 0.0, u"none"


def get_panel_centroid(panel, transform):
    """
    Get the face centroid of a curtain panel in host coordinates.
    Uses the panel's GetTransform() origin if available (more accurate than bbox centre
    for panels that sit flush with the curtain wall face).
    Falls back to bounding-box midpoint.
    """
    mid = None

    # Try panel transform origin (centre of the panel face)
    try:
        panel_transform = panel.GetTransform()
        if panel_transform is not None:
            mid = panel_transform.Origin
    except Exception:
        pass

    # Fallback: bounding-box centre
    if mid is None:
        try:
            bb = panel.get_BoundingBox(None)
            if bb:
                mid = DB.XYZ(
                    (bb.Min.X + bb.Max.X) * 0.5,
                    (bb.Min.Y + bb.Max.Y) * 0.5,
                    (bb.Min.Z + bb.Max.Z) * 0.5,
                )
        except Exception:
            pass

    # Apply link transform
    if mid is not None and transform is not None:
        mid = transform.OfPoint(mid)

    return mid


# ── Shared geometry helpers ───────────────────────────────────────────────────
def midpoint_in_host_coords(element, transform):
    try:
        bb = element.get_BoundingBox(None)
        if bb:
            mid = DB.XYZ(
                (bb.Min.X + bb.Max.X) * 0.5,
                (bb.Min.Y + bb.Max.Y) * 0.5,
                (bb.Min.Z + bb.Max.Z) * 0.5
            )
            if transform is not None:
                mid = transform.OfPoint(mid)
            return mid
    except Exception:
        pass
    return None


def get_room_area_sqm(room):
    try:
        p = room.get_Parameter(DB.BuiltInParameter.ROOM_AREA)
        if p and p.HasValue:
            return to_sqm(p.AsDouble())
    except Exception:
        pass
    return 0.0


def get_room_label(room):
    try:
        name   = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME).AsString() or u"Unnamed"
        number = room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER).AsString() or u"?"
        return u"{} - {}".format(number, name)
    except Exception:
        return u"Unknown Room"


def get_room_base_z(room):
    try:
        level = doc.GetElement(room.LevelId)
        if level:
            base = float(level.Elevation)
            p = room.get_Parameter(DB.BuiltInParameter.ROOM_LOWER_OFFSET)
            if p and p.HasValue:
                base += float(p.AsDouble())
            return base
    except Exception:
        pass
    return 0.0


def get_room_upper_z(room):
    try:
        upper_id    = room.get_Parameter(DB.BuiltInParameter.ROOM_UPPER_LEVEL).AsElementId()
        upper_level = doc.GetElement(upper_id)
        if upper_level:
            upper = float(upper_level.Elevation)
            p = room.get_Parameter(DB.BuiltInParameter.ROOM_UPPER_OFFSET)
            if p and p.HasValue:
                upper += float(p.AsDouble())
            return upper
    except Exception:
        pass
    return get_room_base_z(room) + float(2500.0 / 304.8)


# ── 2D polygon ────────────────────────────────────────────────────────────────
def point_in_polygon_2d(px, py, polygon_pts):
    n = len(polygon_pts)
    inside = False
    j = n - 1
    for i in range(n):
        xi, yi = polygon_pts[i]
        xj, yj = polygon_pts[j]
        if ((yi > py) != (yj > py)) and (px < (xj - xi) * (py - yi) / (yj - yi + 1e-12) + xi):
            inside = not inside
        j = i
    return inside


def get_room_boundary_polygon(room):
    opts = DB.SpatialElementBoundaryOptions()
    opts.SpatialElementBoundaryLocation = DB.SpatialElementBoundaryLocation.Center
    try:
        loops = room.GetBoundarySegments(opts)
        if not loops:
            return []
        return [(float(seg.GetCurve().GetEndPoint(0).X),
                 float(seg.GetCurve().GetEndPoint(0).Y)) for seg in loops[0]]
    except Exception:
        return []


def poly_test_with_offsets(mid, poly):
    if not poly or mid is None:
        return False
    if point_in_polygon_2d(float(mid.X), float(mid.Y), poly):
        return True
    for dx, dy in _S4_OFFSETS:
        if point_in_polygon_2d(float(mid.X) + dx, float(mid.Y) + dy, poly):
            return True
    return False


# ── Collection ────────────────────────────────────────────────────────────────
def collect_rooms():
    return [r for r in
            DB.FilteredElementCollector(doc)
              .OfCategory(DB.BuiltInCategory.OST_Rooms)
              .WhereElementIsNotElementType()
              .ToElements()
            if r.Area > 0]


def collect_all_windows():
    results    = []
    host_wins  = list(
        DB.FilteredElementCollector(doc)
          .OfCategory(DB.BuiltInCategory.OST_Windows)
          .WhereElementIsNotElementType()
          .ToElements()
    )
    for w in host_wins:
        results.append((w, None, None))
    host_count = len(host_wins)

    linked_count = 0
    try:
        for link_inst in DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance).ToElements():
            try:
                link_doc = link_inst.GetLinkDocument()
                if link_doc is None:
                    continue
                xform = link_inst.GetTotalTransform()
                wins  = list(
                    DB.FilteredElementCollector(link_doc)
                      .OfCategory(DB.BuiltInCategory.OST_Windows)
                      .WhereElementIsNotElementType()
                      .ToElements()
                )
                for w in wins:
                    results.append((w, xform, link_doc))
                linked_count += len(wins)
            except Exception as e:
                logger.warning("Skipped link {}: {}".format(eid_int(link_inst.Id), str(e)))
    except Exception as e:
        logger.warning("Could not query linked models: {}".format(str(e)))

    logger.info("Windows: {} host + {} linked = {} total.".format(
        host_count, linked_count, len(results)
    ))
    return results, host_count, linked_count


def collect_all_curtain_panels():
    """
    Collect glazed curtain panels from host + linked models.
    Returns (panel_tuples, host_count, linked_count, skipped_opaque).
    Each tuple: (panel, transform_or_None, linked_doc_or_None).
    """
    results = []
    skipped = 0

    # Host model panels
    host_panels = list(
        DB.FilteredElementCollector(doc)
          .OfCategory(DB.BuiltInCategory.OST_CurtainWallPanels)
          .WhereElementIsNotElementType()
          .ToElements()
    )
    host_count = 0
    for p in host_panels:
        if is_glazed_panel(p, None):
            results.append((p, None, None))
            host_count += 1
        else:
            skipped += 1

    # Linked model panels
    linked_count = 0
    try:
        for link_inst in DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance).ToElements():
            try:
                link_doc = link_inst.GetLinkDocument()
                if link_doc is None:
                    continue
                xform = link_inst.GetTotalTransform()
                panels = list(
                    DB.FilteredElementCollector(link_doc)
                      .OfCategory(DB.BuiltInCategory.OST_CurtainWallPanels)
                      .WhereElementIsNotElementType()
                      .ToElements()
                )
                for p in panels:
                    if is_glazed_panel(p, link_doc):
                        results.append((p, xform, link_doc))
                        linked_count += 1
                    else:
                        skipped += 1
            except Exception as e:
                logger.warning("Skipped link panels {}: {}".format(eid_int(link_inst.Id), str(e)))
    except Exception as e:
        logger.warning("Could not query linked models for panels: {}".format(str(e)))

    logger.info("Curtain panels (glazed): {} host + {} linked = {} total ({} opaque skipped).".format(
        host_count, linked_count, len(results), skipped
    ))
    return results, host_count, linked_count, skipped


# ── Assignment ────────────────────────────────────────────────────────────────
def _assign_element_to_room(mid, rooms, room_z_centre, polygon_cache, element_id):
    """
    Run the four-strategy cascade to find the matching room.
    Returns (room_id, strategy_label) or (None, None).
    """
    if mid is None:
        return None, None

    for room in rooms:
        rid  = eid_int(room.Id)
        poly = polygon_cache.get(rid, [])

        # S1: IsPointInRoom at transformed midpoint ± dz
        for dz in [0.0, 0.5, -0.5]:
            try:
                if room.IsPointInRoom(DB.XYZ(mid.X, mid.Y, mid.Z + dz)):
                    return rid, u"S1"
            except Exception:
                pass

        # S2: IsPointInRoom with Z clamped to room centre
        try:
            if room.IsPointInRoom(DB.XYZ(mid.X, mid.Y, room_z_centre[rid])):
                return rid, u"S2"
        except Exception:
            pass

        # S3: 2D point-in-polygon
        if poly and point_in_polygon_2d(float(mid.X), float(mid.Y), poly):
            return rid, u"S3"

        # S4: 2D polygon with XY offsets
        if poly_test_with_offsets(mid, poly):
            return rid, u"S4"

    return None, None


def map_elements_to_rooms(rooms, window_tuples, panel_tuples):
    """
    Assign windows and curtain panels to rooms.
    Returns room_id_map: { rid: { 'windows': [(w, lnk)], 'panels': [(p, lnk)] } }
    """
    room_id_map   = {eid_int(r.Id): {"windows": [], "panels": []} for r in rooms}
    polygon_cache = {}
    room_z_centre = {}
    assigned_ids  = set()
    win_assigned  = 0
    win_unassigned_ids = []
    pnl_assigned  = 0
    pnl_unassigned_ids = []

    for room in rooms:
        rid = eid_int(room.Id)
        base  = get_room_base_z(room)
        upper = get_room_upper_z(room)
        room_z_centre[rid] = (base + upper) * 0.5
        polygon_cache[rid] = get_room_boundary_polygon(room)

    # Assign windows
    for (win, transform, linked_doc) in window_tuples:
        win_key = (eid_int(win.Id), id(transform))
        if win_key in assigned_ids:
            continue

        mid = midpoint_in_host_coords(win, transform)
        rid, strat = _assign_element_to_room(mid, rooms, room_z_centre, polygon_cache, eid_int(win.Id))

        if rid is not None:
            room_id_map[rid]["windows"].append((win, linked_doc))
            assigned_ids.add(win_key)
            win_assigned += 1
            logger.debug("Win {} -> room {} {}.".format(eid_int(win.Id), rid, strat))
        else:
            win_unassigned_ids.append(eid_int(win.Id))

    # Assign curtain panels
    for (panel, transform, linked_doc) in panel_tuples:
        pnl_key = (eid_int(panel.Id), id(transform))
        if pnl_key in assigned_ids:
            continue

        # Use panel-specific centroid (GetTransform origin preferred over bbox centre)
        mid = get_panel_centroid(panel, transform)
        rid, strat = _assign_element_to_room(mid, rooms, room_z_centre, polygon_cache, eid_int(panel.Id))

        if rid is not None:
            room_id_map[rid]["panels"].append((panel, linked_doc))
            assigned_ids.add(pnl_key)
            pnl_assigned += 1
            logger.debug("Panel {} -> room {} {}.".format(eid_int(panel.Id), rid, strat))
        else:
            pnl_unassigned_ids.append(eid_int(panel.Id))

    return (room_id_map,
            win_assigned, len(win_unassigned_ids), win_unassigned_ids,
            pnl_assigned, len(pnl_unassigned_ids), pnl_unassigned_ids)


# ── Validation ────────────────────────────────────────────────────────────────
def validate():
    rooms = collect_rooms()
    if not rooms:
        forms.alert(
            "No placed rooms found.\nEnsure rooms are placed and have area > 0.",
            exitscript=True
        )
    return rooms


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    rooms = validate()

    window_tuples, win_host, win_linked = collect_all_windows()
    panel_tuples, pnl_host, pnl_linked, pnl_opaque_skipped = collect_all_curtain_panels()

    total_wins   = win_host + win_linked
    total_panels = pnl_host + pnl_linked

    if not window_tuples and not panel_tuples:
        forms.alert("No window instances or glazed curtain panels found in this model or linked models.", ok=True)

    (room_map,
     win_assigned, win_unassigned_count, win_unassigned_ids,
     pnl_assigned, pnl_unassigned_count, pnl_unassigned_ids) = map_elements_to_rooms(
        rooms, window_tuples, panel_tuples
    )

    # ── Build per-room results ────────────────────────────────────────────────
    summary_rows  = []
    detail_blocks = []
    pass_count    = 0
    fail_count    = 0
    warn_count    = 0
    err_count     = 0
    bbox_warned   = False

    for room in sorted(rooms, key=lambda r: get_room_label(r)):
        try:
            rid       = eid_int(room.Id)
            label     = get_room_label(room)
            room_area = get_room_area_sqm(room)
            req_area  = room_area * RATIO_REQUIREMENT

            # Clickable room link (rooms are always in host doc)
            room_link = output.linkify(room.Id, title=label)

            room_data   = room_map.get(rid, {"windows": [], "panels": []})
            win_tuples  = room_data["windows"]
            pnl_tuples  = room_data["panels"]
            opening_count = len(win_tuples) + len(pnl_tuples)

            total_open = 0.0
            used_bbox  = False
            detail_rows = []

            # Process windows
            for (w, lnk_doc) in win_tuples:
                area, w_ft, h_ft, source = window_clear_opening_sqm(w, lnk_doc)
                total_open += area
                if source == u"bbox":
                    used_bbox   = True
                    bbox_warned = True

                mark      = get_window_mark(w, lnk_doc)
                type_name = get_window_type_name(w, lnk_doc)

                # Linkify host-doc windows; linked elements show plain ID
                if lnk_doc is None:
                    id_cell = output.linkify(w.Id, title=mark)
                else:
                    id_cell = u"{} (linked)".format(mark)

                if w_ft > 0.0 and h_ft > 0.0:
                    dim_str = u"{:.0f} \u00d7 {:.0f} mm".format(to_mm(w_ft), to_mm(h_ft))
                else:
                    dim_str = u"-"

                detail_rows.append([
                    id_cell,
                    type_name,
                    u"Window",
                    dim_str,
                    u"{:.3f} m\u00b2".format(area),
                    u"[{}]".format(source),
                ])

            # Process curtain panels
            for (p, lnk_doc) in pnl_tuples:
                area, w_ft, h_ft, source = curtain_panel_area_sqm(p, lnk_doc)
                total_open += area
                if source == u"bbox":
                    used_bbox   = True
                    bbox_warned = True

                mark      = get_window_mark(p, lnk_doc)
                type_name = get_window_type_name(p, lnk_doc)

                # Linkify host-doc panels; linked elements show plain ID
                if lnk_doc is None:
                    id_cell = output.linkify(p.Id, title=mark)
                else:
                    id_cell = u"{} (linked)".format(mark)

                if w_ft > 0.0 and h_ft > 0.0:
                    dim_str = u"{:.0f} \u00d7 {:.0f} mm".format(to_mm(w_ft), to_mm(h_ft))
                else:
                    dim_str = u"-"

                detail_rows.append([
                    id_cell,
                    type_name,
                    u"Curtain Panel",
                    dim_str,
                    u"{:.3f} m\u00b2".format(area),
                    u"[{}]".format(source),
                ])

            total_open = float(total_open)

            # Compose count string
            win_c = len(win_tuples)
            pnl_c = len(pnl_tuples)
            count_parts = []
            if win_c > 0:
                count_parts.append(u"{} win".format(win_c))
            if pnl_c > 0:
                count_parts.append(u"{} pnl".format(pnl_c))
            count_str = u" + ".join(count_parts) if count_parts else u"0"

            if room_area <= 0.0:
                status    = u"NO AREA"
                ratio_str = u"-"
                warn_count += 1
            elif opening_count == 0:
                status    = WARN_LABEL
                ratio_str = u"0.000 / {:.3f}".format(RATIO_REQUIREMENT)
                warn_count += 1
            else:
                ratio = total_open / room_area
                if total_open >= req_area:
                    status = PASS_LABEL
                    pass_count += 1
                else:
                    status = FAIL_LABEL
                    fail_count += 1
                ratio_str = u"{:.3f} / {:.3f}".format(ratio, RATIO_REQUIREMENT)
                if used_bbox:
                    status += u" \u26a0"

            summary_rows.append([
                room_link,
                u"{:.2f} m\u00b2".format(room_area),
                count_str,
                u"{:.3f} m\u00b2".format(total_open),
                u"{:.3f} m\u00b2".format(req_area),
                ratio_str,
                status,
            ])

            detail_blocks.append((room_link, label, status, detail_rows))

        except Exception as e:
            err_count += 1
            logger.warning("Error processing room {}: {}".format(room.Id, str(e)))

    # ── Output ────────────────────────────────────────────────────────────────
    output.print_md(u"# AUK Window Daylighting Compliance \u2014 1/20th Rule")
    output.print_md(
        u"Requirement: Total clear opening \u2265 **1/20th** of room floor area.\n"
        u" Windows: {} ({} host / {} linked)"
        u" | Curtain panels (glazed): {} ({} host / {} linked, {} opaque skipped)"
        u"\n Assigned: {} win + {} pnl"
        u" | Unassigned: {} win + {} pnl.".format(
            total_wins, win_host, win_linked,
            total_panels, pnl_host, pnl_linked, pnl_opaque_skipped,
            win_assigned, pnl_assigned,
            win_unassigned_count, pnl_unassigned_count,
        )
    )

    # Summary table
    output.print_md(u"## Summary")
    if summary_rows:
        output.print_table(
            summary_rows,
            columns=[
                u"Room", u"Room Area", u"Openings",
                u"Clear Opening", u"Min Required",
                u"Ratio (actual / min)", u"Status",
            ]
        )
    else:
        output.print_md(u"> No rooms could be processed.")

    summary_line = u"**{} PASS \u2022 {} FAIL \u2022 {} NO OPENINGS**".format(
        pass_count, fail_count, warn_count
    )
    if err_count:
        summary_line += u" \u2022 {} ERROR".format(err_count)
    output.print_md(summary_line)

    # Per-room detail
    output.print_md(u"---")
    output.print_md(u"## Opening Detail by Room")

    for (room_link, room_label, status, detail_rows) in detail_blocks:
        # Room header with clickable link
        output.print_md(u"### {} \u2014 {}".format(room_label, status))
        output.print_md(u"\u2192 Select room: {}".format(room_link))
        if detail_rows:
            output.print_table(
                detail_rows,
                columns=[
                    u"Mark",
                    u"Family : Type",
                    u"Element",
                    u"W \u00d7 H",
                    u"Clear Opening",
                    u"Source",
                ]
            )
        else:
            output.print_md(u"> No openings assigned to this room.")

    # Footnotes
    if bbox_warned:
        output.print_md(
            u"\n> \u26a0 **Source [bbox]** means no Width/Height/Area parameters were found; "
            u"bounding-box estimation used. Add Width/Height to the family for accuracy."
        )

    if win_unassigned_count > 0:
        # Linkify unassigned host-doc window IDs
        id_links = []
        for uid in win_unassigned_ids:
            try:
                id_links.append(output.linkify(DB.ElementId(uid)))
            except Exception:
                id_links.append(str(uid))
        output.print_md(
            u"\n> **{} window(s) unassigned** (IDs: {}) \u2014 no matching room boundary found.".format(
                win_unassigned_count,
                u", ".join(id_links)
            )
        )

    if pnl_unassigned_count > 0:
        # Linkify unassigned host-doc panel IDs
        id_links = []
        for uid in pnl_unassigned_ids:
            try:
                id_links.append(output.linkify(DB.ElementId(uid)))
            except Exception:
                id_links.append(str(uid))
        output.print_md(
            u"\n> **{} curtain panel(s) unassigned** (IDs: {}) \u2014 no matching room boundary found.".format(
                pnl_unassigned_count,
                u", ".join(id_links)
            )
        )


main()
