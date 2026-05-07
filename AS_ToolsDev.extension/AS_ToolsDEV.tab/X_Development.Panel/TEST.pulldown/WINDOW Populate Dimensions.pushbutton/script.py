# -*- coding: utf-8 -*-
"""
Script:   WINDOW Populate Dimensions
Desc:     Reads Width and Height from all windows in the active model and linked
          models, then writes the resolved values into user-selected instance
          parameters on host windows. Linked windows are reported read-only.
Author:   Aukett Swanke Digital
Usage:    Run from the Building Elements panel. No prior selection required.
Result:   Chosen parameters populated with resolved Width/Height values on host
          windows. Output table shows per-window status and source tag.
"""
from __future__ import division

from pyrevit import revit, DB, script, forms
from Snippets._context_manager import ef_Transaction

logger = script.get_logger()
output = script.get_output()
doc    = revit.doc

# ── Constants ─────────────────────────────────────────────────────────────────
FT_TO_MM = 304.8

WIDTH_PARAM_NAMES  = [
    u"Width", u"Rough Width", u"Frame Width", u"Overall Width",
    u"Window Width", u"Clear Width", u"Opening Width", u"RO Width",
]
HEIGHT_PARAM_NAMES = [
    u"Height", u"Rough Height", u"Frame Height", u"Overall Height",
    u"Window Height", u"Clear Height", u"Opening Height", u"RO Height",
]

# Only use window-specific built-ins. GENERIC_* and DOOR_* are excluded —
# they are not axis-safe for windows and can return the wrong dimension.
_BUILTIN_WIDTH  = [DB.BuiltInParameter.FAMILY_WIDTH_PARAM]
_BUILTIN_HEIGHT = [DB.BuiltInParameter.FAMILY_HEIGHT_PARAM]


# ── Helpers ───────────────────────────────────────────────────────────────────
def to_mm(ft):
    return float(ft) * FT_TO_MM


def eid_int(eid):
    try:
        return int(eid.Value)
    except AttributeError:
        return int(eid.IntegerValue)


def get_window_symbol(instance, linked_doc):
    """Return the ElementType for a window instance (looked up in owner doc)."""
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


def _read_named_param(element, names):
    """Return first positive Double value from a list of named parameters."""
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
    """Return first positive Double value from a list of BuiltInParameters."""
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


def resolve_width_height(instance, linked_doc=None):
    """
    Return (w_ft, h_ft, source_label) for a window instance.
    Resolution order:
      1. Named params on instance
      2. Named params on type
      3. Built-in params on instance then type
      4. Bounding-box fallback (max(dx,dy) x dz)
    source_label: 'param' | 'builtin' | 'bbox' | 'none'
    """
    symbol = get_window_symbol(instance, linked_doc)

    w = _read_named_param(instance, WIDTH_PARAM_NAMES)
    if w <= 0.0 and symbol:
        w = _read_named_param(symbol, WIDTH_PARAM_NAMES)

    h = _read_named_param(instance, HEIGHT_PARAM_NAMES)
    if h <= 0.0 and symbol:
        h = _read_named_param(symbol, HEIGHT_PARAM_NAMES)

    if w > 0.0 and h > 0.0:
        return w, h, u"param"

    if w <= 0.0:
        w = _read_builtin_param(instance, _BUILTIN_WIDTH)
    if w <= 0.0 and symbol:
        w = _read_builtin_param(symbol, _BUILTIN_WIDTH)

    if h <= 0.0:
        h = _read_builtin_param(instance, _BUILTIN_HEIGHT)
    if h <= 0.0 and symbol:
        h = _read_builtin_param(symbol, _BUILTIN_HEIGHT)

    if w > 0.0 and h > 0.0:
        return w, h, u"builtin"

    # Bounding-box fallback.
    # Height is always the world-Z extent (unambiguously vertical).
    # Width is the extent along the instance's local Y axis (along-wall direction),
    # which is correct for any wall orientation — not just axis-aligned walls.
    try:
        bb = instance.get_BoundingBox(None)
        if bb:
            h_bbox = abs(float(bb.Max.Z - bb.Min.Z))
            try:
                basis_y = instance.GetTransform().BasisY
                pts = [
                    DB.XYZ(bb.Min.X, bb.Min.Y, 0.0),
                    DB.XYZ(bb.Max.X, bb.Min.Y, 0.0),
                    DB.XYZ(bb.Min.X, bb.Max.Y, 0.0),
                    DB.XYZ(bb.Max.X, bb.Max.Y, 0.0),
                ]
                proj = [pt.DotProduct(basis_y) for pt in pts]
                w_bbox = max(proj) - min(proj)
            except Exception:
                # Axis-aligned fallback if transform is unavailable
                dx = abs(float(bb.Max.X - bb.Min.X))
                dy = abs(float(bb.Max.Y - bb.Min.Y))
                w_bbox = max(dx, dy)
            if w_bbox > 0.0 and h_bbox > 0.0:
                return w_bbox, h_bbox, u"bbox"
    except Exception:
        pass

    return 0.0, 0.0, u"none"


def get_type_label(instance, linked_doc):
    """Return 'FamilyName : TypeName' string for a window instance."""
    symbol = get_window_symbol(instance, linked_doc)
    if symbol is None:
        return u"-"
    try:
        fam = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString() or u""
        typ = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or u""
        if fam and typ:
            return u"{} : {}".format(fam, typ)
        return typ or fam or u"-"
    except Exception:
        return u"-"


# ── Collection ────────────────────────────────────────────────────────────────
def collect_host_windows():
    return list(
        DB.FilteredElementCollector(doc)
          .OfCategory(DB.BuiltInCategory.OST_Windows)
          .WhereElementIsNotElementType()
          .ToElements()
    )


def collect_linked_windows():
    """Return list of (window_instance, total_transform, link_doc)."""
    results = []
    try:
        for link_inst in DB.FilteredElementCollector(doc)\
                .OfClass(DB.RevitLinkInstance).ToElements():
            try:
                link_doc = link_inst.GetLinkDocument()
                if link_doc is None:
                    continue
                xform = link_inst.GetTotalTransform()
                for w in DB.FilteredElementCollector(link_doc)\
                        .OfCategory(DB.BuiltInCategory.OST_Windows)\
                        .WhereElementIsNotElementType().ToElements():
                    results.append((w, xform, link_doc))
            except Exception as e:
                logger.warning(u"Skipped link {}: {}".format(eid_int(link_inst.Id), str(e)))
    except Exception as e:
        logger.warning(u"Could not query linked models: {}".format(str(e)))
    return results


# ── Parameter discovery ───────────────────────────────────────────────────────
def discover_writable_params(windows, sample=50):
    """
    Return sorted list of writable Double/String parameter names found on
    sampled host window instances AND their types.
    User-defined project parameters are often bound at the type level, so
    both scopes must be checked.
    """
    names = set()
    for w in windows[:sample]:
        # Instance-level parameters
        try:
            for p in w.Parameters:
                if (not p.IsReadOnly
                        and p.Definition is not None
                        and p.StorageType in (DB.StorageType.Double, DB.StorageType.String)):
                    names.add(p.Definition.Name)
        except Exception:
            pass
        # Type-level parameters
        sym = get_window_symbol(w, None)
        if sym:
            try:
                for p in sym.Parameters:
                    if (not p.IsReadOnly
                            and p.Definition is not None
                            and p.StorageType in (DB.StorageType.Double, DB.StorageType.String)):
                        names.add(p.Definition.Name)
            except Exception:
                pass
    return sorted(names)


# ── Write helpers ─────────────────────────────────────────────────────────────
def _is_length_param(p):
    """
    Return True when the parameter stores length data (feet internally).
    Handles both pre-2022 ParameterType API and 2022+ ForgeTypeId API.
    """
    try:
        # Revit 2022+ uses ForgeTypeId / SpecTypeId
        spec = p.Definition.GetDataType()
        try:
            return spec == DB.SpecTypeId.Length
        except AttributeError:
            return u"length" in str(spec).lower()
    except AttributeError:
        pass
    try:
        # Pre-2022
        return p.Definition.ParameterType == DB.ParameterType.Length
    except Exception:
        pass
    # Unknown Double parameter — default to treating as Length (safe for dimensions)
    return True


def _write_to_param_object(p, value_ft):
    """Write value_ft to a resolved parameter object. Returns True on success."""
    try:
        if p.StorageType == DB.StorageType.Double:
            p.Set(value_ft if _is_length_param(p) else to_mm(value_ft))
            return True
        if p.StorageType == DB.StorageType.String:
            p.Set(str(int(round(to_mm(value_ft)))))
            return True
    except Exception:
        pass
    return False


def write_dimension_to_param(elem, param_name, value_ft):
    """
    Write value_ft (Revit internal feet) to param_name on elem.
    Tries the instance parameter first; falls back to the type parameter so
    that project parameters bound at the type level are also handled.
      - Length param  → write feet (Revit handles display unit conversion)
      - Number param  → write mm value
      - String param  → write rounded mm as text string
    Returns True on success.
    """
    # Try instance parameter first
    p = elem.LookupParameter(param_name)
    if p and not p.IsReadOnly:
        if _write_to_param_object(p, value_ft):
            return True

    # Fall back to type parameter
    sym = get_window_symbol(elem, None)
    if sym:
        p = sym.LookupParameter(param_name)
        if p and not p.IsReadOnly:
            if _write_to_param_object(p, value_ft):
                return True

    logger.warning(u"Write failed on element {} param '{}': not found or read-only.".format(
        eid_int(elem.Id), param_name))
    return False


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    # ── 1. Collect host windows ───────────────────────────────────────────────
    host_wins = collect_host_windows()
    if not host_wins:
        forms.alert(
            u"No window instances found in the active model.",
            title=u"Window Dimension Populator",
            exitscript=True
        )
        return

    # ── 2. Discover writable parameters on host windows ───────────────────────
    param_names = discover_writable_params(host_wins)
    if not param_names:
        forms.alert(
            u"No writable Double or String parameters found on host window instances or types.\n\n"
            u"Add the required shared parameters to the Window category before running.",
            title=u"No Writable Parameters Found",
            exitscript=True
        )
        return

    # ── 3. User selects target parameters ────────────────────────────────────
    width_param = forms.SelectFromList.show(
        param_names,
        title=u"Select parameter \u2192 WINDOW WIDTH",
        multiselect=False,
        button_name=u"Select"
    )
    if not width_param:
        return

    height_param = forms.SelectFromList.show(
        param_names,
        title=u"Select parameter \u2192 WINDOW HEIGHT",
        multiselect=False,
        button_name=u"Select"
    )
    if not height_param:
        return

    # ── 4. Collect linked windows (resolved before transaction) ───────────────
    linked_wins = collect_linked_windows()

    # ── 5. Resolve and write host windows inside a single transaction ─────────
    host_rows  = []
    updated    = 0
    skipped    = 0
    bbox_count = 0

    with ef_Transaction(doc, u"Populate Window Width & Height"):
        for win in host_wins:
            w_ft, h_ft, source = resolve_width_height(win, None)

            ok_w = write_dimension_to_param(win, width_param,  w_ft) if w_ft > 0.0 else False
            ok_h = write_dimension_to_param(win, height_param, h_ft) if h_ft > 0.0 else False
            status = u"OK" if (ok_w or ok_h) else u"SKIP"

            if ok_w or ok_h:
                updated += 1
            else:
                skipped += 1
            if source == u"bbox":
                bbox_count += 1

            host_rows.append([
                output.linkify(win.Id),
                get_type_label(win, None),
                u"{:.0f}".format(to_mm(w_ft)) if w_ft > 0.0 else u"\u2014",
                u"{:.0f}".format(to_mm(h_ft)) if h_ft > 0.0 else u"\u2014",
                u"[{}]".format(source),
                status,
            ])

    # ── 6. Output ─────────────────────────────────────────────────────────────
    output.print_md(u"# Window Width & Height \u2014 Parameter Update")
    output.print_md(
        u"**Width \u2192** `{}`   |   **Height \u2192** `{}`".format(
            width_param, height_param)
    )
    output.print_md(u"")
    output.print_md(u"## Host Model Windows")
    output.print_md(
        u"Updated: **{}** \u2022 Skipped: **{}** \u2022 Bbox fallback: **{}**".format(
            updated, skipped, bbox_count)
    )

    if host_rows:
        output.print_table(
            host_rows,
            columns=[u"Element", u"Family : Type",
                     u"Width (mm)", u"Height (mm)",
                     u"Source", u"Status"]
        )

    # ── 7. Linked windows — read-only report ──────────────────────────────────
    if linked_wins:
        output.print_md(u"---")
        output.print_md(u"## Linked Model Windows (read-only)")
        output.print_md(
            u"_{} linked window(s) found. Parameters cannot be written to linked "
            u"documents from the host._".format(len(linked_wins))
        )
        linked_rows = []
        for (win, _xform, link_doc) in linked_wins:
            w_ft, h_ft, source = resolve_width_height(win, link_doc)
            linked_rows.append([
                link_doc.Title,
                get_type_label(win, link_doc),
                u"{:.0f}".format(to_mm(w_ft)) if w_ft > 0.0 else u"\u2014",
                u"{:.0f}".format(to_mm(h_ft)) if h_ft > 0.0 else u"\u2014",
                u"[{}]".format(source),
            ])
        output.print_table(
            linked_rows,
            columns=[u"Link Document", u"Family : Type",
                     u"Width (mm)", u"Height (mm)", u"Source"]
        )

    # ── 8. Footnotes ──────────────────────────────────────────────────────────
    if bbox_count > 0:
        output.print_md(
            u"\n> \u26a0 **Source [bbox]**: No Width/Height parameters found on the window "
            u"family \u2014 bounding-box estimation used. Add Width/Height parameters to "
            u"these families for accurate results."
        )

    output.print_md(u"---")
    output.print_md(u"_Complete._")


main()
