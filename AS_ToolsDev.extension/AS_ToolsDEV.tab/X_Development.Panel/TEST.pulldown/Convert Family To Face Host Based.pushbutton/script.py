# -*- coding: utf-8 -*-
"""
Script:   Convert Family to Face-Hosted
Desc:     Converts selected family instances to face-hosted (face-based) families
Author:   AUK
Usage:    Select one or more family instances in the model, then run the script
Result:   Selected families are converted to face-hosted; all instances are deleted first
"""
# pylint: disable=import-error,invalid-name,broad-except
from pyrevit import revit, DB
from pyrevit import script
from pyrevit import forms

logger = script.get_logger()
output = script.get_output()
doc = revit.doc


# ── Helpers ───────────────────────────────────────────────────────────────────

def get_family_instances(family):
    """Return list of ElementIds for all instances of all symbols in family."""
    instance_ids = []
    try:
        for symbol_id in family.GetFamilySymbolIds():
            collector = DB.FilteredElementCollector(doc)\
                          .WherePasses(DB.FamilyInstanceFilter(doc, symbol_id))\
                          .ToElementIds()
            instance_ids.extend(list(collector))
    except Exception as e:
        logger.warning("Could not collect instances for family '{}': {}".format(
            family.Name, str(e)))
    return instance_ids


def delete_all_instances(family):
    """Delete all instances of the given family. Must be called inside a transaction."""
    instance_ids = get_family_instances(family)
    if not instance_ids:
        logger.info("No instances found for family '{}'.".format(family.Name))
        return 0
    for eid in instance_ids:
        try:
            doc.Delete(eid)
        except Exception as e:
            logger.warning("Could not delete instance {}: {}".format(eid, str(e)))
    return len(instance_ids)


# ── Validation ────────────────────────────────────────────────────────────────

def get_valid_families():
    """
    Resolve selected elements to unique Family objects.
    Returns list of DB.Family or exits with alert if none found.
    """
    selection = revit.get_selection()
    elements = selection.elements

    if not elements:
        forms.alert(
            "No elements selected.\n\nSelect one or more family instances and try again.",
            exitscript=True
        )

    seen_ids = set()
    families = []

    for el in elements:
        # Must be a FamilyInstance to access .Symbol.Family
        if not isinstance(el, DB.FamilyInstance):
            logger.warning("Skipping element {} — not a FamilyInstance.".format(el.Id))
            continue

        symbol = el.Symbol
        if symbol is None:
            logger.warning("Skipping element {} — Symbol is None.".format(el.Id))
            continue

        family = symbol.Family
        if family is None:
            logger.warning("Skipping element {} — Family is None.".format(el.Id))
            continue

        # Deduplicate by family Id
        try:
            fid = family.Id.Value
        except AttributeError:
            # Revit 2023/2024 fallback
            fid = family.Id.IntegerValue

        if fid not in seen_ids:
            seen_ids.add(fid)
            families.append(family)

    if not families:
        forms.alert(
            "No valid family instances found in selection.\n\n"
            "Select at least one placed family instance and try again.",
            exitscript=True
        )

    return families


# ── Main logic ────────────────────────────────────────────────────────────────

def main():
    families = get_valid_families()

    # Pre-check: verify all families can be converted before touching anything
    non_convertible = []
    for family in families:
        try:
            if not DB.FamilyUtils.FamilyCanConvertToFaceHostBased(doc, family.Id):
                non_convertible.append(family.Name)
        except Exception as e:
            logger.warning("Could not check convertibility for '{}': {}".format(
                family.Name, str(e)))
            non_convertible.append(family.Name)

    if non_convertible:
        forms.alert(
            "The following families cannot be converted to face-hosted "
            "by the Revit API:\n\n{}\n\nNo changes have been made.".format(
                "\n".join(u"\u2022 " + n for n in non_convertible)
            ),
            exitscript=True
        )

    family_names = [f.Name for f in families]
    if not forms.alert(
        "The following families will be converted to face-hosted.\n\n"
        "{}\n\nAll existing instances will be permanently deleted.\n\n"
        "This cannot be undone. Continue?".format(
            "\n".join(u"\u2022 " + n for n in family_names)
        ),
        yes=True, cancel=True
    ):
        script.exit()

    succeeded = []
    failed = []

    for family in families:
        fname = family.Name
        try:
            # Step 1: delete all instances (required before conversion)
            with revit.Transaction("Delete instances: {}".format(fname)):
                deleted = delete_all_instances(family)
            logger.info("Deleted {} instance(s) of '{}'.".format(deleted, fname))

            # Step 2: convert family to face-hosted
            with revit.Transaction("Convert to face-hosted: {}".format(fname)):
                DB.FamilyUtils.ConvertFamilyToFaceHostBased(doc, family.Id)

            succeeded.append(fname)
            logger.info("Successfully converted '{}'.".format(fname))

        except Exception as e:
            failed.append((fname, str(e)))
            logger.error("Failed to convert '{}': {}".format(fname, str(e)))

    # ── Summary ───────────────────────────────────────────────────────────────
    output.print_md("## Convert to Face-Hosted — Results")

    if succeeded:
        output.print_md("**Converted successfully ({}):**".format(len(succeeded)))
        for name in succeeded:
            output.print_md(u"- {}".format(name))

    if failed:
        output.print_md("**Failed ({}):**".format(len(failed)))
        for name, err in failed:
            output.print_md(u"- **{}**: `{}`".format(name, err))

    if not failed:
        forms.alert(
            "Conversion complete.\n{} family/families converted successfully.".format(
                len(succeeded)
            )
        )
    else:
        forms.alert(
            "{} converted, {} failed.\nSee the output panel for details.".format(
                len(succeeded), len(failed)
            )
        )


main()