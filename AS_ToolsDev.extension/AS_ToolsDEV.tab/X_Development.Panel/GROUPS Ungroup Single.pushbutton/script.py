# -*- coding: utf-8 -*-
"""
Script:   Ungroup Single-Instance Groups
Desc:     Finds model and/or detail groups with exactly one instance in the
          active document, lets the user confirm which to ungroup, then
          dissolves them in a single undoable transaction.
Author:   Aukett Swanke Digital
Usage:    Open any view, click button. No pre-selection required.
Result:   Selected single-instance groups are ungrouped. Results reported in
          the output panel.
"""
from pyrevit import revit, DB, script, forms

logger = script.get_logger()
output = script.get_output()
doc    = revit.doc

# ── Constants ─────────────────────────────────────────────────────────────────
TITLE = "AUK Ungroup Singles"

# BuiltInCategory integer values for group categories
CAT_MODEL_ID  = int(DB.BuiltInCategory.OST_IOSModelGroups)
CAT_DETAIL_ID = int(DB.BuiltInCategory.OST_IOSDetailGroups)

SCOPE_OPTIONS = [
    "Model Groups only",
    "Detail Groups only",
    "Both Model and Detail Groups",
]


# ── Helpers ───────────────────────────────────────────────────────────────────
def eid_int(element_id):
    """Return ElementId value as int, compatible with Revit 2023-2026."""
    try:
        return int(element_id.Value)          # Revit 2026+
    except AttributeError:
        return int(element_id.IntegerValue)   # Revit 2023-2025


def safe_name(element):
    """Return a display name for an element without raising."""
    try:
        name = DB.Element.Name.GetValue(element)
        if name:
            return name
    except Exception:
        pass
    return "Unnamed"


def get_group_category_label(group):
    """Return 'Model' or 'Detail' string for a group element."""
    try:
        if eid_int(group.GroupType.Category.Id) == CAT_DETAIL_ID:
            return "Detail"
    except Exception:
        pass
    return "Model"


def get_member_count(group):
    """Return the number of member elements in a group, or 0 on failure."""
    try:
        ids = group.GetMemberIds()
        return len(ids) if ids is not None else 0
    except Exception:
        return 0


def get_display_name(group):
    """Return the user-assigned group name.

    Priority:
      1. SYMBOL_NAME_PARAM on the GroupType element (most reliable)
      2. GroupType.Name direct property access
      3. DB.Element.Name.GetValue() fallback
      4. "Unnamed" if all else fails
    """
    # 1 – SYMBOL_NAME_PARAM on the GroupType (user-visible name in Revit UI)
    try:
        type_elem = doc.GetElement(group.GetTypeId())
        if type_elem is not None:
            param = type_elem.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if param and param.HasValue:
                n = param.AsString()
                if n:
                    return n
    except Exception:
        pass

    # 2 – Direct GroupType.Name property
    try:
        n = group.GroupType.Name
        if n:
            return n
    except Exception:
        pass

    # 3 – Static Element.Name helper
    return safe_name(group)


# ── Collection / filtering ────────────────────────────────────────────────────
def collect_groups(target_cat_ids):
    """Return all Group instances whose category ID is in target_cat_ids."""
    groups = []
    try:
        collector = (DB.FilteredElementCollector(doc)
                     .OfClass(DB.Group)
                     .WhereElementIsNotElementType())
        for g in collector:
            try:
                if g.GroupType is None:
                    continue
                cat = g.GroupType.Category
                if cat is None:
                    continue
                if eid_int(cat.Id) in target_cat_ids:
                    groups.append(g)
            except Exception as e:
                logger.warning("Skipped group {}: {}".format(g.Id, e))
    except Exception as e:
        logger.error("Collector failed: {}".format(e))
    return groups


def build_type_map(groups):
    """Return dict: GroupType integer ID -> list of Group instances."""
    type_map = {}
    for g in groups:
        try:
            tid = g.GetTypeId()
            if tid is None or tid == DB.ElementId.InvalidElementId:
                continue
            key = eid_int(tid)
            type_map.setdefault(key, []).append(g)
        except Exception as e:
            logger.warning("Could not get TypeId for group {}: {}".format(g.Id, e))
    return type_map


# ── Selection list wrapper ────────────────────────────────────────────────────
class GroupItem(forms.TemplateListItem):
    """Wraps a group info dict for display in forms.SelectFromList."""

    @property
    def name(self):
        return "[{}]  {}  ({} elements)  [ID: {}]".format(
            self.item["category"],
            self.item["name"],
            self.item["members"],
            eid_int(self.item["group"].Id),
        )


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    # Step 1 – ask which group category scope to process
    scope = forms.SelectFromList.show(
        SCOPE_OPTIONS,
        title=TITLE,
        header="Which group types should be checked?",
        button_name="Continue",
    )
    if not scope:
        script.exit()

    # Resolve category IDs for the chosen scope
    target_ids = set()
    if "Model" in scope:
        target_ids.add(CAT_MODEL_ID)
    if "Detail" in scope:
        target_ids.add(CAT_DETAIL_ID)

    # Step 2 – collect and filter to single-instance types
    all_groups = collect_groups(target_ids)
    if not all_groups:
        forms.alert(
            "No groups found for the selected scope.",
            title=TITLE,
            warn_icon=True,
        )
        return

    type_map   = build_type_map(all_groups)
    singles    = []

    for _tid, instances in type_map.iteritems():
        if len(instances) != 1:
            continue
        g = instances[0]
        singles.append({
            "group":    g,
            "name":     get_display_name(g),
            "members":  get_member_count(g),
            "category": get_group_category_label(g),
        })

    if not singles:
        forms.alert(
            "No single-instance groups found.\n\n"
            "All {} group(s) in scope have multiple instances.".format(len(all_groups)),
            title=TITLE,
        )
        return

    # Sort by category then name (case-insensitive)
    singles.sort(key=lambda x: (x["category"], x["name"].lower()))

    # Step 3 – let the user confirm which groups to ungroup
    selectable = [GroupItem(item, checked=True) for item in singles]

    selected = forms.SelectFromList.show(
        selectable,
        title=TITLE,
        header="Select groups to ungroup  ({} single-instance group(s) found)".format(
            len(singles)),
        multiselect=True,
        checked_only=True,
        button_name="Ungroup Selected",
    )
    if not selected:
        script.exit()

    # Step 4 – ungroup inside a single transaction
    ungrouped    = 0
    failed       = 0
    failed_items = []

    t = DB.Transaction(doc, "Ungroup Single-Instance Groups")
    try:
        t.Start()
        for item in selected:
            g    = item["group"]
            name = item["name"]
            try:
                # Validate element is still in the document before acting
                if doc.GetElement(g.Id) is None:
                    raise Exception("Element no longer in document")
                g.UngroupMembers()
                ungrouped += 1
            except Exception as e:
                failed += 1
                failed_items.append((name, str(e)))
                logger.warning("Failed to ungroup '{}': {}".format(name, e))

        t.Commit()

    except Exception as e:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        logger.error("Transaction error: {}".format(e))
        forms.alert(
            "Transaction failed — no changes were made.\n\n{}".format(str(e)),
            title=TITLE,
            warn_icon=True,
        )
        return

    # Step 5 – report results
    output.print_md("## {} — Results".format(TITLE))
    output.print_md("**Ungrouped:** {}".format(ungrouped))

    if failed:
        output.print_md("**Failed:** {}".format(failed))
        output.print_md("### Failed items")
        for name, reason in failed_items:
            output.print_md("- **{}**: {}".format(name, reason))

    summary = "Ungrouped: {} group(s)".format(ungrouped)
    if failed:
        summary += "\nFailed: {} group(s)  \u2014 see output panel for details.".format(failed)

    forms.alert(summary, title=TITLE)


main()
