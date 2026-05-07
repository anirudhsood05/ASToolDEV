# -*- coding: utf-8 -*-
"""
DEBUG SCRIPT - Tagless Shamelist Tag Detection
Run this in the active view that contains your test walls (2 tagged, 1 untagged).
Paste the FULL output panel text into a new Claude conversation.
"""
from pyrevit import revit, DB, script

output = script.get_output()
doc    = revit.doc
uidoc  = revit.uidoc
view   = uidoc.ActiveView

output.print_md("# Tagless Shamelist - Tag Detection Debug")
output.print_md("**Revit Version:** {}".format(doc.Application.VersionNumber))
output.print_md("**Active View:** {} (ID: {})".format(view.Name, view.Id.IntegerValue))
output.print_md("**View Type:** {}".format(str(view.ViewType)))
output.print_md("---")

# ── STEP 1: Walls visible in view ─────────────────────────────────────────────
output.print_md("## STEP 1: Walls collected via view-scoped FilteredElementCollector")
try:
    walls = (
        DB.FilteredElementCollector(doc, view.Id)
        .OfCategory(DB.BuiltInCategory.OST_Walls)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    output.print_md("**Wall count in view:** {}".format(len(walls)))
    for w in walls:
        try:
            type_name = doc.GetElement(w.GetTypeId()).Name
        except Exception:
            type_name = "(could not read type)"
        output.print_md("- Wall ID: {}  |  Type: {}".format(w.Id.IntegerValue, type_name))
except Exception as e:
    output.print_md("**ERROR collecting walls:** {}".format(str(e)))

output.print_md("---")

# ── STEP 2: All tags in view via OfClass(IndependentTag) ──────────────────────
output.print_md("## STEP 2: Tags collected via OfClass(IndependentTag)")
try:
    all_tags = (
        DB.FilteredElementCollector(doc, view.Id)
        .OfClass(DB.IndependentTag)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    output.print_md("**Total IndependentTag count in view:** {}".format(len(all_tags)))
    for tag in all_tags:
        output.print_md("- Tag ID: {}  |  Category: {}".format(
            tag.Id.IntegerValue,
            tag.Category.Name if tag.Category else "(no category)"
        ))
except Exception as e:
    output.print_md("**ERROR collecting tags:** {}".format(str(e)))

output.print_md("---")

# ── STEP 3: Extract tagged element IDs - try both API paths ───────────────────
output.print_md("## STEP 3: Extracting tagged element IDs from each tag")
tagged_ids_multi  = set()
tagged_ids_single = set()

try:
    tags = (
        DB.FilteredElementCollector(doc, view.Id)
        .OfClass(DB.IndependentTag)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    for tag in tags:
        output.print_md("### Tag ID: {}".format(tag.Id.IntegerValue))

        # Path A: GetTaggedLocalElementIds (Revit 2023+)
        try:
            multi = tag.GetTaggedLocalElementIds()
            ids_found = []
            for link_eid in multi:
                try:
                    host_id = link_eid.HostElementId.IntegerValue
                    ids_found.append(host_id)
                    tagged_ids_multi.add(host_id)
                except Exception as inner:
                    output.print_md("  - HostElementId read error: {}".format(str(inner)))
            output.print_md("  Path A (GetTaggedLocalElementIds): {}".format(ids_found))
        except Exception as e:
            output.print_md("  Path A FAILED: {}".format(str(e)))

        # Path B: TaggedLocalElementId (Revit 2021/2022)
        try:
            eid = tag.TaggedLocalElementId
            if eid and eid != DB.ElementId.InvalidElementId:
                tagged_ids_single.add(eid.IntegerValue)
                output.print_md("  Path B (TaggedLocalElementId): {}".format(eid.IntegerValue))
            else:
                output.print_md("  Path B: InvalidElementId or None")
        except Exception as e:
            output.print_md("  Path B FAILED: {}".format(str(e)))

        # Path C: GetTaggedReferences (alternative API)
        try:
            refs = tag.GetTaggedReferences()
            ref_ids = []
            for r in refs:
                try:
                    ref_ids.append(r.ElementId.IntegerValue)
                except Exception:
                    pass
            output.print_md("  Path C (GetTaggedReferences): {}".format(ref_ids))
        except Exception as e:
            output.print_md("  Path C FAILED: {}".format(str(e)))

except Exception as e:
    output.print_md("**ERROR in Step 3:** {}".format(str(e)))

output.print_md("---")

# ── STEP 4: Cross-reference walls vs tagged IDs ───────────────────────────────
output.print_md("## STEP 4: Cross-reference - which walls are detected as tagged/untagged")

output.print_md("**tagged_ids via Path A:** {}".format(sorted(tagged_ids_multi)))
output.print_md("**tagged_ids via Path B:** {}".format(sorted(tagged_ids_single)))
combined = tagged_ids_multi | tagged_ids_single
output.print_md("**Combined tagged IDs:** {}".format(sorted(combined)))

try:
    walls = (
        DB.FilteredElementCollector(doc, view.Id)
        .OfCategory(DB.BuiltInCategory.OST_Walls)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    for w in walls:
        wid = w.Id.IntegerValue
        in_multi  = wid in tagged_ids_multi
        in_single = wid in tagged_ids_single
        in_either = wid in combined
        output.print_md(
            "- Wall {} | PathA={} | PathB={} | Combined={} | VERDICT: {}".format(
                wid, in_multi, in_single, in_either,
                "TAGGED (would be skipped)" if in_either else "UNTAGGED (should appear in results)"
            )
        )
except Exception as e:
    output.print_md("**ERROR in Step 4:** {}".format(str(e)))

output.print_md("---")

# ── STEP 5: Tags via OfCategory(OST_WallTags) for comparison ──────────────────
output.print_md("## STEP 5: Tags via OfCategory(OST_WallTags) - legacy method comparison")
try:
    wall_tags = (
        DB.FilteredElementCollector(doc, view.Id)
        .OfCategory(DB.BuiltInCategory.OST_WallTags)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    output.print_md("**Wall tag count via OfCategory(OST_WallTags):** {}".format(len(wall_tags)))
    for t in wall_tags:
        output.print_md("- Tag ID: {}".format(t.Id.IntegerValue))
except Exception as e:
    output.print_md("**ERROR:** {}".format(str(e)))

output.print_md("---")
output.print_md("## END OF DEBUG OUTPUT - paste everything above into a new Claude conversation")