# -*- coding: utf-8 -*-
"""
Script:   AS_Compare Linked Model vs Bound Group (Stable v2)
Desc:     Compares a linked Revit model with a model group (bound copy of that
          link) and classifies every element into MUTUALLY EXCLUSIVE buckets:
              - UNCHANGED   : matched, position within tolerance
              - MOVED       : matched, displaced > tolerance (pure geometry)
              - ORPHAN      : in group but not in link (new binding geometry)
              - MISSING     : in link but not in group (failed to bind)
              - ERROR       : flagged in the binding error report
                              (regardless of position - overrides other buckets)
          An element appears in ONE bucket only. Error-report parsing uses a
          tightened regex that ignores the summary id on error-bearing
          elements so it does not over-match unrelated IDs.
Author:   Aukett Swanke - Digital
Usage:    Open a view. Run the tool, set tolerance, optionally load a Revit
          binding-error report (HTML/TXT), pick the linked model, then the
          model group. Review report, accept highlights and CSV export.
Result:   Colour-coded overrides (Red=error, Gold=moved, Orange=orphan),
          pyRevit report with linkified IDs, optional CSV export.
"""

# ------------------------------------------------------------------------------
# Imports
# ------------------------------------------------------------------------------
from pyrevit import revit, DB
from pyrevit import script, forms
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException
import os
import re
import csv
import gc
import traceback

logger = script.get_logger()
output = script.get_output()
doc    = revit.doc
uidoc  = revit.uidoc

# ------------------------------------------------------------------------------
# Constants - tunable guards against memory / UI stalls
# ------------------------------------------------------------------------------
TOLERANCE_MM           = 5.0
FALLBACK_MATCH_MM      = 50.0
FT_TO_MM               = 304.8

# Hard ceilings to prevent runaway memory on huge models
MAX_LINK_ELEMENTS      = 200000
MAX_GROUP_MEMBERS      = 200000
MAX_REPORT_ROWS        = 500
MAX_OVERRIDES_PER_TX   = 500
MAX_CSV_ROWS           = 50000
MAX_ERROR_FILE_BYTES   = 50 * 1024 * 1024
BUCKET_CELL_MM         = 1000.0
PROGRESS_UPDATE_EVERY  = 250

# Near-zero threshold for spatial-match tie cases (float noise)
NEAR_ZERO_MM           = 0.01

# AUK brand - three-tier colour coding
GOLD   = (232, 167, 53)
RED    = (220, 30, 30)
ORANGE = (230, 120, 0)

TX_GROUP_NAME = "AUK - Compare Link vs Bound Group"
TX_OVERRIDE   = "AUK - Highlight Comparison Results"


def _bic_int(bic_name):
    try:
        return int(getattr(DB.BuiltInCategory, bic_name))
    except Exception:
        return None


_SKIP_BIC_NAMES = [
    "OST_RvtLinks", "OST_IOSModelGroups", "OST_SketchLines",
    "OST_IOSSketchGrid", "OST_Cameras", "OST_IOSAttachedDetailGroups",
    "OST_Views", "OST_Sheets", "OST_Schedules", "OST_Levels",
    "OST_Grids", "OST_ReferenceLines", "OST_CLines",
]
SKIP_BIC_INTS = set([v for v in (_bic_int(n) for n in _SKIP_BIC_NAMES)
                     if v is not None])


# ------------------------------------------------------------------------------
# Version-safe helpers
# ------------------------------------------------------------------------------
def eid_int(element_id):
    if element_id is None:
        return -1
    try:
        return element_id.Value
    except AttributeError:
        try:
            return element_id.IntegerValue
        except Exception:
            return -1


def safe_name(element):
    if element is None:
        return "<none>"
    try:
        return DB.Element.Name.GetValue(element)
    except Exception:
        try:
            return element.Name
        except Exception:
            return "<unnamed>"


def safe_category_id(elem):
    try:
        cat = elem.Category
        if cat is None:
            return -1
        return eid_int(cat.Id)
    except Exception:
        return -1


def safe_category_name(elem):
    try:
        cat = elem.Category
        if cat is None:
            return "<no category>"
        return cat.Name or "<no category>"
    except Exception:
        return "<no category>"


def is_element_valid(elem):
    if elem is None:
        return False
    try:
        return bool(elem.IsValidObject)
    except Exception:
        return False


def safe_start(tx):
    try:
        if not tx.HasStarted():
            tx.Start()
        return True
    except Exception as e:
        logger.error("Transaction start failed: {}".format(str(e)))
        return False


def safe_commit(tx):
    try:
        if tx.HasStarted() and not tx.HasEnded():
            tx.Commit()
            return True
    except Exception as e:
        logger.error("Commit failed: {}".format(str(e)))
        try:
            tx.RollBack()
        except Exception:
            pass
    return False


def safe_rollback(tx):
    try:
        if tx.HasStarted() and not tx.HasEnded():
            tx.RollBack()
    except Exception:
        pass


# ------------------------------------------------------------------------------
# Selection filters
# ------------------------------------------------------------------------------
class LinkInstanceFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            return isinstance(elem, DB.RevitLinkInstance)
        except Exception:
            return False

    def AllowReference(self, ref, pt):
        return False


class GroupInstanceFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            return isinstance(elem, DB.Group)
        except Exception:
            return False

    def AllowReference(self, ref, pt):
        return False


# ------------------------------------------------------------------------------
# Validation
# ------------------------------------------------------------------------------
def validate_environment():
    if doc is None:
        forms.alert("No active document.", exitscript=True)
    try:
        if doc.IsFamilyDocument:
            forms.alert("Cannot run in the Family Editor.", exitscript=True)
    except Exception:
        forms.alert("Invalid document state.", exitscript=True)
    if uidoc is None:
        forms.alert("No active UI document.", exitscript=True)


def get_loaded_link_instances():
    links = []
    try:
        collector = DB.FilteredElementCollector(doc)\
            .OfClass(DB.RevitLinkInstance)\
            .WhereElementIsNotElementType()
        for li in collector:
            try:
                if is_element_valid(li) and li.GetLinkDocument() is not None:
                    links.append(li)
            except Exception:
                continue
    except Exception as e:
        logger.error("Link collection failed: {}".format(str(e)))
    return links


def has_any_group_instances():
    try:
        return DB.FilteredElementCollector(doc)\
            .OfClass(DB.Group)\
            .WhereElementIsNotElementType()\
            .GetElementCount() > 0
    except Exception:
        return False


# ------------------------------------------------------------------------------
# Error report parser - TIGHTENED
# ------------------------------------------------------------------------------
# Revit error HTML rows have this shape:
#   <Discipline> : <Category> : <Type> : <Family> : <TypeName> : id <N>
# We ONLY want IDs that appear with at least one preceding " : " on the same
# line / table cell AND directly after the word "id". This rejects stray
# numeric-looking text and prose.
ERR_ID_LINE_RE = re.compile(
    r":\s*id\s+(\d+)\s*(?:<|$)",
    re.IGNORECASE
)
# Fallback: inside HTML list items separated by <br>, the form "id N<br>" or
# "id N</td>" is also valid; the regex above catches both endings.


def parse_error_report(path):
    """Extract element IDs from a Revit HTML/TXT error report.
    Only matches the "<stuff> : id <N>" pattern that Revit actually emits -
    avoids over-matching prose or unrelated numbers."""
    ids = set()
    if not path:
        return ids
    try:
        if not os.path.isfile(path):
            return ids
        size = os.path.getsize(path)
        if size > MAX_ERROR_FILE_BYTES:
            logger.warning("Error report too large ({} bytes) - skipping."
                           .format(size))
            return ids
        with open(path, "rb") as f:
            raw = f.read()
        try:
            text = raw.decode("utf-8", "replace")
        except Exception:
            try:
                text = raw.decode("latin-1", "replace")
            except Exception:
                return ids
        for m in ERR_ID_LINE_RE.finditer(text):
            try:
                ids.add(int(m.group(1)))
            except (ValueError, TypeError):
                continue
    except Exception as e:
        logger.warning("Error report parse failed: {}".format(str(e)))

    logger.info("Error report parsed: {} unique element IDs extracted."
                .format(len(ids)))
    return ids


# ------------------------------------------------------------------------------
# User picking
# ------------------------------------------------------------------------------
def pick_link_instance():
    if not get_loaded_link_instances():
        forms.alert("No loaded Revit link instances found.", exitscript=True)
        return None
    try:
        with forms.WarningBar(title="Select the LINKED MODEL (source)"):
            ref = uidoc.Selection.PickObject(
                ObjectType.Element, LinkInstanceFilter(),
                "Pick the linked Revit model instance"
            )
    except OperationCanceledException:
        return None
    except Exception as e:
        logger.error("Link pick failed: {}".format(str(e)))
        return None
    if ref is None:
        return None
    try:
        link_inst = doc.GetElement(ref.ElementId)
    except Exception:
        return None
    if not isinstance(link_inst, DB.RevitLinkInstance):
        forms.alert("Selected element is not a linked model.")
        return None
    try:
        if link_inst.GetLinkDocument() is None:
            forms.alert("Selected link is not loaded.")
            return None
    except Exception:
        forms.alert("Could not access the linked document.")
        return None
    return link_inst


def pick_group_instance():
    if not has_any_group_instances():
        forms.alert("No model groups found.", exitscript=True)
        return None
    try:
        with forms.WarningBar(title="Select the MODEL GROUP (bound copy)"):
            ref = uidoc.Selection.PickObject(
                ObjectType.Element, GroupInstanceFilter(),
                "Pick the model group instance"
            )
    except OperationCanceledException:
        return None
    except Exception as e:
        logger.error("Group pick failed: {}".format(str(e)))
        return None
    if ref is None:
        return None
    try:
        grp = doc.GetElement(ref.ElementId)
    except Exception:
        return None
    if not isinstance(grp, DB.Group):
        forms.alert("Selected element is not a model group.")
        return None
    return grp


# ------------------------------------------------------------------------------
# Element filtering
# ------------------------------------------------------------------------------
def is_geometry_element(elem):
    if not is_element_valid(elem):
        return False
    try:
        if isinstance(elem, DB.ElementType):
            return False
    except Exception:
        return False
    cat_int = safe_category_id(elem)
    if cat_int == -1:
        return False
    if cat_int in SKIP_BIC_INTS:
        return False
    try:
        if elem.ViewSpecific:
            return False
    except Exception:
        pass
    try:
        if elem.Category.CategoryType != DB.CategoryType.Model:
            return False
    except Exception:
        return False
    return True


def collect_linked_elements(link_inst):
    try:
        linked_doc = link_inst.GetLinkDocument()
    except Exception:
        return [], None
    if linked_doc is None:
        return [], None
    try:
        transform = link_inst.GetTotalTransform()
    except Exception:
        return [], None

    elems = []
    count = 0
    with forms.ProgressBar(title="Reading linked model...",
                           cancellable=True, step=500) as pb:
        try:
            collector = DB.FilteredElementCollector(linked_doc)\
                .WhereElementIsNotElementType()
            it = collector.GetElementIterator()
            it.Reset()
            while it.MoveNext():
                if pb.cancelled:
                    break
                el = None
                try:
                    el = it.Current
                except Exception:
                    continue
                count += 1
                try:
                    if is_geometry_element(el):
                        elems.append(el)
                except Exception:
                    continue
                if count % PROGRESS_UPDATE_EVERY == 0:
                    try:
                        pb.update_progress(count % 1000, 1000)
                    except Exception:
                        pass
                if len(elems) >= MAX_LINK_ELEMENTS:
                    logger.warning("Link element cap reached ({}); truncating."
                                   .format(MAX_LINK_ELEMENTS))
                    break
        except Exception as e:
            logger.error("Link iteration failed: {}".format(str(e)))
    return elems, transform


def collect_group_members(group):
    members = []
    try:
        member_ids = list(group.GetMemberIds())
    except Exception as e:
        logger.error("GetMemberIds failed: {}".format(str(e)))
        return members

    total = len(member_ids)
    if total > MAX_GROUP_MEMBERS:
        logger.warning("Group has {} members; truncating to {}."
                       .format(total, MAX_GROUP_MEMBERS))
        member_ids = member_ids[:MAX_GROUP_MEMBERS]

    with forms.ProgressBar(title="Reading group members...",
                           cancellable=True, step=200) as pb:
        for i, mid in enumerate(member_ids):
            if pb.cancelled:
                break
            if i % PROGRESS_UPDATE_EVERY == 0:
                try:
                    pb.update_progress(i, max(1, len(member_ids)))
                except Exception:
                    pass
            try:
                el = doc.GetElement(mid)
            except Exception:
                continue
            try:
                if is_geometry_element(el):
                    members.append(el)
            except Exception:
                continue
    return members


# ------------------------------------------------------------------------------
# Geometry / centroid
# ------------------------------------------------------------------------------
def element_centroid(elem, transform=None):
    """Centroid via LocationPoint / LocationCurve midpoint / bbox centre.
    Never calls get_Geometry() (crash vector on curtain panels etc.)."""
    if not is_element_valid(elem):
        return None
    pt = None
    try:
        loc = elem.Location
        if isinstance(loc, DB.LocationPoint):
            try:
                pt = loc.Point
            except Exception:
                pt = None
        elif isinstance(loc, DB.LocationCurve):
            try:
                c = loc.Curve
                if c is not None:
                    p0 = c.GetEndPoint(0)
                    p1 = c.GetEndPoint(1)
                    pt = (p0 + p1) * 0.5
            except Exception:
                pt = None
    except Exception:
        pt = None

    if pt is None:
        try:
            bb = elem.get_BoundingBox(None)
            if bb is not None and bb.Min is not None and bb.Max is not None:
                pt = (bb.Min + bb.Max) * 0.5
        except Exception:
            pt = None

    if pt is None:
        return None
    if transform is not None:
        try:
            pt = transform.OfPoint(pt)
        except Exception:
            return None
    return pt


def distance_mm(p1, p2):
    if p1 is None or p2 is None:
        return None
    try:
        return p1.DistanceTo(p2) * FT_TO_MM
    except Exception:
        return None


def type_key(elem):
    try:
        if not is_element_valid(elem):
            return u""
        tid = elem.GetTypeId()
        if tid is None or eid_int(tid) <= 0:
            return u""
        src_doc = elem.Document
        etype = src_doc.GetElement(tid)
        if etype is None:
            return u""
        fam = u""
        try:
            fam = etype.FamilyName or u""
        except Exception:
            fam = u""
        return u"{}::{}".format(fam, safe_name(etype))
    except Exception:
        return u""


# ------------------------------------------------------------------------------
# Lightweight element records
# ------------------------------------------------------------------------------
class ElemRec(object):
    __slots__ = ("eid_int", "uid", "centroid", "cat_int", "tkey",
                 "cat_name", "name", "matched")

    def __init__(self, elem, transform=None):
        self.eid_int = eid_int(elem.Id) if is_element_valid(elem) else -1
        try:
            self.uid = elem.UniqueId
        except Exception:
            self.uid = None
        self.centroid = element_centroid(elem, transform)
        self.cat_int = safe_category_id(elem)
        self.tkey = type_key(elem)
        self.cat_name = safe_category_name(elem)
        self.name = safe_name(elem)
        self.matched = False


def build_records(elems, transform, title):
    recs = []
    total = len(elems)
    with forms.ProgressBar(title=title, cancellable=True, step=200) as pb:
        for i, el in enumerate(elems):
            if pb.cancelled:
                break
            if i % PROGRESS_UPDATE_EVERY == 0:
                try:
                    pb.update_progress(i, max(1, total))
                except Exception:
                    pass
            try:
                recs.append(ElemRec(el, transform))
            except Exception:
                continue
    return recs


# ------------------------------------------------------------------------------
# Comparison - MUTUALLY EXCLUSIVE buckets
# ------------------------------------------------------------------------------
def build_link_index(link_recs):
    uid_idx = {}
    bucket = {}
    for i, rec in enumerate(link_recs):
        if rec.uid:
            uid_idx[rec.uid] = i
        if rec.centroid is not None:
            try:
                gx = int(round(rec.centroid.X * FT_TO_MM / BUCKET_CELL_MM))
                gy = int(round(rec.centroid.Y * FT_TO_MM / BUCKET_CELL_MM))
                gz = int(round(rec.centroid.Z * FT_TO_MM / BUCKET_CELL_MM))
            except Exception:
                continue
            key = (rec.cat_int, rec.tkey, gx, gy, gz)
            bucket.setdefault(key, []).append(i)
    return uid_idx, bucket


def fallback_find(grec, link_recs, bucket):
    if grec.centroid is None:
        return -1, None
    try:
        gx = int(round(grec.centroid.X * FT_TO_MM / BUCKET_CELL_MM))
        gy = int(round(grec.centroid.Y * FT_TO_MM / BUCKET_CELL_MM))
        gz = int(round(grec.centroid.Z * FT_TO_MM / BUCKET_CELL_MM))
    except Exception:
        return -1, None

    best_idx = -1
    best_d = None
    for dx in (-1, 0, 1):
        for dy in (-1, 0, 1):
            for dz in (-1, 0, 1):
                key = (grec.cat_int, grec.tkey, gx + dx, gy + dy, gz + dz)
                idx_list = bucket.get(key)
                if not idx_list:
                    continue
                for li in idx_list:
                    lrec = link_recs[li]
                    if lrec.matched:
                        continue
                    d = distance_mm(lrec.centroid, grec.centroid)
                    if d is None:
                        continue
                    if d <= FALLBACK_MATCH_MM and \
                       (best_d is None or d < best_d):
                        best_idx = li
                        best_d = d
    return best_idx, best_d


def compare(link_recs, group_recs, tolerance_mm, error_ids):
    """Classify every group element into ONE bucket.
    Priority: ERROR > ORPHAN > MOVED > UNCHANGED.
    ERROR bucket preserves the matched link record (if any) for reporting."""
    errors = []
    moved = []
    unchanged = 0
    orphans = []

    with forms.ProgressBar(title="Comparing elements...",
                           cancellable=True, step=200) as pb:
        total = len(group_recs)
        uid_idx, bucket = build_link_index(link_recs)

        for i, grec in enumerate(group_recs):
            if pb.cancelled:
                break
            if i % PROGRESS_UPDATE_EVERY == 0:
                try:
                    pb.update_progress(i, max(1, total))
                except Exception:
                    pass

            is_error = grec.eid_int in error_ids

            # --- Resolve a match first (UID, then spatial) ---
            link_idx = -1
            match_type = None

            if grec.uid and grec.uid in uid_idx:
                li = uid_idx[grec.uid]
                if not link_recs[li].matched:
                    link_idx = li
                    match_type = "uid"

            if link_idx < 0:
                li, _ = fallback_find(grec, link_recs, bucket)
                if li >= 0:
                    link_idx = li
                    match_type = "spatial"

            matched = link_idx >= 0
            lrec = link_recs[link_idx] if matched else None
            if matched:
                lrec.matched = True

            d = distance_mm(lrec.centroid, grec.centroid) if matched else None

            # --- Classify into ONE bucket ---
            if is_error:
                errors.append({
                    "grec": grec,
                    "lrec": lrec,
                    "dist_mm": d,
                    "match_type": match_type,
                    "matched": matched,
                })
                continue

            if not matched:
                orphans.append({"grec": grec})
                continue

            # Matched - decide moved vs unchanged purely on geometry
            if d is None:
                # Could not compute - conservatively unchanged rather than
                # polluting the moved bucket. User gets a diagnostic count.
                unchanged += 1
                continue

            if d <= tolerance_mm or d < NEAR_ZERO_MM:
                unchanged += 1
            else:
                moved.append({
                    "grec": grec,
                    "lrec": lrec,
                    "dist_mm": d,
                    "match_type": match_type,
                })

    missing = [rec for rec in link_recs if not rec.matched]

    return {
        "errors": errors,
        "moved": moved,
        "unchanged": unchanged,
        "orphans": orphans,
        "missing": missing,
    }


# ------------------------------------------------------------------------------
# View overrides
# ------------------------------------------------------------------------------
def get_solid_fill_pattern_id():
    try:
        for fp in DB.FilteredElementCollector(doc)\
                .OfClass(DB.FillPatternElement).ToElements():
            try:
                pat = fp.GetFillPattern()
                if pat is not None and pat.IsSolidFill:
                    return fp.Id
            except Exception:
                continue
    except Exception:
        pass
    return DB.ElementId.InvalidElementId


def build_ogs(rgb, solid_id):
    ogs = DB.OverrideGraphicSettings()
    try:
        r, g, b = rgb
        color = DB.Color(r, g, b)
        ogs.SetProjectionLineColor(color)
        ogs.SetCutLineColor(color)
        ogs.SetSurfaceForegroundPatternColor(color)
        ogs.SetSurfaceForegroundPatternVisible(True)
        ogs.SetCutForegroundPatternColor(color)
        ogs.SetCutForegroundPatternVisible(True)
        if solid_id != DB.ElementId.InvalidElementId:
            ogs.SetSurfaceForegroundPatternId(solid_id)
            ogs.SetCutForegroundPatternId(solid_id)
    except Exception as e:
        logger.warning("OGS build warning: {}".format(str(e)))
    return ogs


def apply_overrides_chunked(view, override_jobs):
    if not override_jobs:
        return 0
    applied = 0
    total = len(override_jobs)
    chunks = [override_jobs[i:i + MAX_OVERRIDES_PER_TX]
              for i in range(0, total, MAX_OVERRIDES_PER_TX)]

    with forms.ProgressBar(title="Applying overrides...",
                           cancellable=True, step=1) as pb:
        for ci, chunk in enumerate(chunks):
            if pb.cancelled:
                break
            try:
                pb.update_progress(ci + 1, len(chunks))
            except Exception:
                pass
            tx = DB.Transaction(doc,
                                "{} ({}/{})".format(TX_OVERRIDE,
                                                    ci + 1, len(chunks)))
            if not safe_start(tx):
                continue
            try:
                for eid_val, ogs in chunk:
                    try:
                        view.SetElementOverrides(DB.ElementId(eid_val), ogs)
                        applied += 1
                    except Exception:
                        continue
                if not safe_commit(tx):
                    safe_rollback(tx)
            except Exception as e:
                logger.error("Chunk override failed: {}".format(str(e)))
                safe_rollback(tx)
    return applied


def apply_overrides(view, results):
    """Buckets are mutually exclusive - no dedup priority logic needed."""
    if view is None:
        return 0
    solid_id = get_solid_fill_pattern_id()
    ogs_red    = build_ogs(RED,    solid_id)
    ogs_gold   = build_ogs(GOLD,   solid_id)
    ogs_orange = build_ogs(ORANGE, solid_id)

    jobs = []
    for row in results["errors"]:
        eid_val = row["grec"].eid_int
        if eid_val > 0:
            jobs.append((eid_val, ogs_red))
    for row in results["moved"]:
        eid_val = row["grec"].eid_int
        if eid_val > 0:
            jobs.append((eid_val, ogs_gold))
    for orp in results["orphans"]:
        eid_val = orp["grec"].eid_int
        if eid_val > 0:
            jobs.append((eid_val, ogs_orange))
    return apply_overrides_chunked(view, jobs)


# ------------------------------------------------------------------------------
# Reporting
# ------------------------------------------------------------------------------
def format_point(p):
    if p is None:
        return "-"
    try:
        return "{:.0f}, {:.0f}, {:.0f}".format(
            p.X * FT_TO_MM, p.Y * FT_TO_MM, p.Z * FT_TO_MM)
    except Exception:
        return "-"


def linkify_safe(eid_val):
    try:
        return output.linkify(DB.ElementId(eid_val))
    except Exception:
        return str(eid_val)


def print_report(link_inst, group, results, tolerance_mm, error_ids_count):
    errors    = results["errors"]
    moved     = results["moved"]
    unchanged = results["unchanged"]
    orphans   = results["orphans"]
    missing   = results["missing"]

    output.print_md("# AUK - Link vs Bound Group Comparison")
    output.print_md("**Linked model:** {}".format(safe_name(link_inst)))
    output.print_md("**Group instance:** {} (Id {})"
                    .format(safe_name(group), eid_int(group.Id)))
    output.print_md("**Displacement tolerance:** {:.2f} mm"
                    .format(tolerance_mm))
    if error_ids_count > 0:
        output.print_md("**Binding error report loaded:** {} flagged IDs"
                        .format(error_ids_count))
    output.print_md("---")
    output.print_md("## Summary (mutually exclusive buckets)")
    output.print_md("- **Unchanged:** {}".format(unchanged))
    output.print_md("- **Moved (> tolerance, not error-flagged):** {}"
                    .format(len(moved)))
    output.print_md("- **Orphans (new geometry, not error-flagged):** {}"
                    .format(len(orphans)))
    output.print_md("- **Error-flagged (from report):** {}".format(len(errors)))
    output.print_md("- **Missing (in link but not in group):** {}"
                    .format(len(missing)))
    output.print_md("---")

    # Errors
    if errors:
        rows = []
        for row in errors[:MAX_REPORT_ROWS]:
            grec = row["grec"]
            lrec = row.get("lrec")
            d = row.get("dist_mm")
            state = "matched" if row.get("matched") else "orphan"
            rows.append([
                state,
                linkify_safe(grec.eid_int),
                grec.cat_name,
                grec.name,
                "-" if d is None else "{:.1f}".format(d),
                format_point(lrec.centroid) if lrec else "-",
                format_point(grec.centroid),
            ])
        output.print_table(
            table_data=rows,
            title="Error-flagged elements (from binding report)",
            columns=["State", "Element Id", "Category", "Name",
                     "Distance (mm)", "Link pos (mm)", "Group pos (mm)"]
        )

    # Moved
    if moved:
        moved_sorted = sorted(
            moved,
            key=lambda r: -(r["dist_mm"] or 0.0)
        )
        rows = []
        for row in moved_sorted[:MAX_REPORT_ROWS]:
            grec = row["grec"]
            lrec = row["lrec"]
            d = row["dist_mm"]
            tag = "MOVED"
            if row.get("match_type") == "spatial":
                tag += "*"
            rows.append([
                tag,
                linkify_safe(grec.eid_int),
                grec.cat_name,
                grec.name,
                "{:.1f}".format(d),
                format_point(lrec.centroid),
                format_point(grec.centroid),
            ])
        output.print_table(
            table_data=rows,
            title="Moved elements (pure position change)",
            columns=["Type", "Element Id", "Category", "Name",
                     "Distance (mm)", "Link pos (mm)", "Group pos (mm)"]
        )
        if len(moved_sorted) > MAX_REPORT_ROWS:
            output.print_md("_Showing top {} of {} rows._"
                            .format(MAX_REPORT_ROWS, len(moved_sorted)))
        output.print_md("_* = matched by spatial fallback (UID differed)._")

    # Orphans
    if orphans:
        rows = []
        for orp in orphans[:MAX_REPORT_ROWS]:
            grec = orp["grec"]
            rows.append([
                linkify_safe(grec.eid_int),
                grec.cat_name,
                grec.name,
                format_point(grec.centroid),
            ])
        output.print_table(
            table_data=rows,
            title="Orphan group members (no counterpart in link)",
            columns=["Element Id", "Category", "Name", "Group pos (mm)"]
        )
        if len(orphans) > MAX_REPORT_ROWS:
            output.print_md("_Showing top {} of {} orphans._"
                            .format(MAX_REPORT_ROWS, len(orphans)))

    # Missing
    if missing:
        rows = []
        for mrec in missing[:MAX_REPORT_ROWS]:
            rows.append([mrec.eid_int, mrec.cat_name, mrec.name])
        output.print_table(
            table_data=rows,
            title="Missing - in link but not in group",
            columns=["Linked Element Id", "Category", "Name"]
        )
        if len(missing) > MAX_REPORT_ROWS:
            output.print_md("_Showing top {} of {} missing._"
                            .format(MAX_REPORT_ROWS, len(missing)))


# ------------------------------------------------------------------------------
# CSV export
# ------------------------------------------------------------------------------
def _safe_encode(s):
    try:
        if s is None:
            return ""
        if isinstance(s, str):
            return s
        return s.encode("utf-8", "replace")
    except Exception:
        return ""


def _write_row(writer, classification, grec, lrec=None, d=None,
               match_type="", is_error=False):
    lp = lrec.centroid if lrec else None
    gp = grec.centroid if grec else None
    try:
        writer.writerow([
            classification,
            "Y" if is_error else "N",
            match_type or "",
            grec.eid_int if grec else "",
            lrec.eid_int if lrec else "",
            _safe_encode(grec.cat_name if grec else ""),
            _safe_encode(grec.name if grec else ""),
            "" if d is None else "{:.2f}".format(d),
            "" if lp is None else "{:.2f}".format(lp.X * FT_TO_MM),
            "" if lp is None else "{:.2f}".format(lp.Y * FT_TO_MM),
            "" if lp is None else "{:.2f}".format(lp.Z * FT_TO_MM),
            "" if gp is None else "{:.2f}".format(gp.X * FT_TO_MM),
            "" if gp is None else "{:.2f}".format(gp.Y * FT_TO_MM),
            "" if gp is None else "{:.2f}".format(gp.Z * FT_TO_MM),
        ])
        return True
    except Exception:
        return False


def export_csv(results, link_inst, group):
    try:
        default_name = "AS_LinkVsGroup_{}_Grp{}.csv".format(
            safe_name(link_inst).replace(" ", "_").replace(".rvt", ""),
            eid_int(group.Id)
        )
        save_path = forms.save_file(
            file_ext="csv",
            default_name=default_name,
            title="Export comparison to CSV"
        )
        if not save_path:
            return None

        written = 0
        with open(save_path, "wb") as f:
            writer = csv.writer(f)
            writer.writerow([
                "Classification", "IsErrorFlagged", "MatchType",
                "GroupElementId", "LinkElementId",
                "Category", "Name",
                "DistanceMM",
                "LinkX_mm", "LinkY_mm", "LinkZ_mm",
                "GroupX_mm", "GroupY_mm", "GroupZ_mm",
            ])
            # ERROR (error bucket is always flagged)
            for row in results["errors"]:
                if written >= MAX_CSV_ROWS:
                    break
                if _write_row(writer, "ERROR",
                              row["grec"], row.get("lrec"),
                              row.get("dist_mm"),
                              row.get("match_type") or "",
                              is_error=True):
                    written += 1
            # MOVED (never flagged - buckets are mutually exclusive)
            for row in results["moved"]:
                if written >= MAX_CSV_ROWS:
                    break
                if _write_row(writer, "MOVED",
                              row["grec"], row.get("lrec"),
                              row.get("dist_mm"),
                              row.get("match_type") or "",
                              is_error=False):
                    written += 1
            # ORPHAN (never flagged - buckets are mutually exclusive)
            for orp in results["orphans"]:
                if written >= MAX_CSV_ROWS:
                    break
                if _write_row(writer, "ORPHAN",
                              orp["grec"], None, None, "",
                              is_error=False):
                    written += 1
            # MISSING (link-side only)
            for mrec in results["missing"]:
                if written >= MAX_CSV_ROWS:
                    break
                try:
                    writer.writerow([
                        "MISSING", "N", "",
                        "", mrec.eid_int,
                        _safe_encode(mrec.cat_name),
                        _safe_encode(mrec.name),
                        "", "", "", "", "", "", "",
                    ])
                    written += 1
                except Exception:
                    continue
        return save_path
    except Exception as e:
        logger.error("CSV export failed: {}".format(str(e)))
        forms.alert("CSV export failed: {}".format(str(e)))
        return None


# ------------------------------------------------------------------------------
# UI prompts
# ------------------------------------------------------------------------------
def ask_tolerance():
    try:
        raw = forms.ask_for_string(
            default=str(TOLERANCE_MM),
            prompt="Displacement tolerance in mm:",
            title="AUK - Link vs Group Comparison"
        )
        if raw is None:
            return None
        raw = raw.strip()
        if not raw:
            return None
        val = float(raw)
        return max(0.0, min(val, 100000.0))
    except ValueError:
        forms.alert("Invalid number. Using default {} mm."
                    .format(TOLERANCE_MM))
        return TOLERANCE_MM
    except Exception:
        return TOLERANCE_MM


def ask_error_report():
    try:
        if not forms.alert(
            "Load a Revit binding error report (HTML/TXT)?\n\n"
            "Optional - helps pre-flag elements that failed binding.",
            yes=True, no=True
        ):
            return None
    except Exception:
        return None
    for ext in ("html", "htm", "txt"):
        try:
            path = forms.pick_file(
                file_ext=ext,
                title="Pick the binding error report ({})".format(ext)
            )
            if path:
                return path
        except Exception:
            continue
    return None


# ------------------------------------------------------------------------------
# Main
# ------------------------------------------------------------------------------
def main():
    validate_environment()

    tolerance_mm = ask_tolerance()
    if tolerance_mm is None:
        return

    report_path = ask_error_report()
    error_ids = parse_error_report(report_path) if report_path else set()

    link_inst = pick_link_instance()
    if link_inst is None:
        return

    group = pick_group_instance()
    if group is None:
        return

    linked_elems, link_transform = collect_linked_elements(link_inst)
    if not linked_elems or link_transform is None:
        forms.alert("Linked model has no comparable elements.")
        return
    link_recs = build_records(linked_elems, link_transform,
                              "Indexing linked elements...")
    del linked_elems
    gc.collect()

    group_members = collect_group_members(group)
    if not group_members:
        forms.alert("Selected group has no comparable members.")
        return
    group_recs = build_records(group_members, None,
                               "Indexing group members...")
    del group_members
    gc.collect()

    try:
        results = compare(link_recs, group_recs, tolerance_mm, error_ids)
    except Exception as e:
        logger.error("Comparison failed: {}\n{}"
                     .format(str(e), traceback.format_exc()))
        forms.alert("Comparison failed:\n{}".format(str(e)))
        return

    print_report(link_inst, group, results, tolerance_mm, len(error_ids))

    total_issues = (len(results["errors"]) + len(results["moved"])
                    + len(results["orphans"]) + len(results["missing"]))
    if total_issues == 0:
        forms.alert("Comparison complete.\n\nNo issues found.\n"
                    "Unchanged: {}".format(results["unchanged"]))
        return

    active_view = None
    try:
        active_view = doc.ActiveView
    except Exception:
        active_view = None

    can_override = (active_view is not None
                    and not active_view.IsTemplate
                    and (results["errors"] or results["moved"]
                         or results["orphans"]))

    if can_override:
        try:
            do_it = forms.alert(
                "Comparison complete:\n"
                "  {} error-flagged\n"
                "  {} moved\n"
                "  {} orphan\n"
                "  {} missing\n\n"
                "Apply highlights in active view?\n"
                "(Red=error, Gold=moved, Orange=orphan)".format(
                    len(results["errors"]),
                    len(results["moved"]),
                    len(results["orphans"]),
                    len(results["missing"]),
                ),
                yes=True, no=True
            )
        except Exception:
            do_it = False
        if do_it:
            n = apply_overrides(active_view, results)
            output.print_md("Applied **{}** overrides in view _{}_."
                            .format(n, safe_name(active_view)))

    try:
        if forms.alert("Export full comparison to CSV?",
                       yes=True, no=True):
            p = export_csv(results, link_inst, group)
            if p:
                output.print_md("CSV written: `{}`".format(p))
    except Exception:
        pass


# ------------------------------------------------------------------------------
# Entry point
# ------------------------------------------------------------------------------
if __name__ == "__main__":
    tg = None
    try:
        tg = DB.TransactionGroup(doc, TX_GROUP_NAME)
        try:
            tg.Start()
        except Exception as e:
            logger.error("TG start failed: {}".format(str(e)))
            tg = None

        try:
            main()
            if tg is not None and tg.HasStarted() and not tg.HasEnded():
                try:
                    tg.Assimilate()
                except Exception as e:
                    logger.error("Assimilate failed: {}".format(str(e)))
                    try:
                        tg.RollBack()
                    except Exception:
                        pass
        except OperationCanceledException:
            if tg is not None and tg.HasStarted() and not tg.HasEnded():
                try:
                    tg.RollBack()
                except Exception:
                    pass
        except Exception as e:
            logger.error("Unhandled: {}\n{}"
                         .format(str(e), traceback.format_exc()))
            if tg is not None and tg.HasStarted() and not tg.HasEnded():
                try:
                    tg.RollBack()
                except Exception:
                    pass
            try:
                forms.alert("Unexpected error:\n{}".format(str(e)))
            except Exception:
                pass
    except Exception as e:
        logger.error("Fatal: {}".format(str(e)))
        try:
            forms.alert("Fatal error:\n{}".format(str(e)))
        except Exception:
            pass
    finally:
        try:
            gc.collect()
        except Exception:
            pass