# -*- coding: utf-8 -*-
# config.py - Purge Families configuration utility
# Allows user to pick temporary directory for family purge processing

try:
    from pyrevit.versionmgr import PYREVIT_VERSION
except:
    from pyrevit import versionmgr
    PYREVIT_VERSION = versionmgr.get_pyrevit_version()

pyRevitNewer44 = (PYREVIT_VERSION.major, PYREVIT_VERSION.minor) >= (4, 5)

if pyRevitNewer44:
    from pyrevit import script
    from pyrevit.forms import WPFWindow, alert, pick_folder
    my_config = script.get_config()
else:
    from scriptutils import this_script as script
    from scriptutils.userinput import WPFWindow, pick_folder
    my_config = script.config
    from Autodesk.Revit.UI import TaskDialog
    def alert(msg):
        TaskDialog.Show('pyrevit', msg)

import os
import tempfile


# Default temporary directory (inlined to remove pyapex_utils / purge_families_defaults dependency)
DEFAULT_TEMP_DIR = os.path.join(tempfile.gettempdir(), "pyApex_purgeFamilies")


def list2str(v):
    """Convert list to string - first element if list, else as-is"""
    if v is None:
        return ""
    if isinstance(v, list):
        if len(v) > 0:
            return str(v[0])
        return ""
    return str(v)


def str2list(v):
    """Convert string to single-item list"""
    if v is None:
        return [""]
    if isinstance(v, list):
        return v
    return [str(v)]


def safe_get_config_value(cfg, key, default):
    """Safely read a config value, fall back to default on any error"""
    try:
        v = getattr(cfg, key)
        if v is None:
            return default
        return v
    except:
        return default


class PurgeFamiliesConfigWindow(WPFWindow):
    """Configuration dialog for Purge Families temp directory"""

    def __init__(self, xaml_file_name):
        try:
            WPFWindow.__init__(self, xaml_file_name)
        except Exception as e:
            alert("Cannot load configuration window:\n%s" % str(e))
            raise

        # Load current temp_dir from config, fall back to default
        current = safe_get_config_value(my_config, 'temp_dir', DEFAULT_TEMP_DIR)
        try:
            self.temp_dir.Text = list2str(current)
        except:
            try:
                self.temp_dir.Text = DEFAULT_TEMP_DIR
            except:
                pass

    def restore_defaults(self, p1=None, p2=None, *args):
        """Restore default values for config fields"""
        try:
            if len(args) == 0 or "temp_dir" in args:
                self.temp_dir.Text = list2str(DEFAULT_TEMP_DIR)
        except Exception as e:
            alert("Cannot restore defaults:\n%s" % str(e))

    # noinspection PyUnusedLocal
    # noinspection PyMethodMayBeStatic
    def save_options(self, sender, args):
        """Validate and save configuration"""
        errors = []

        try:
            directory = self.temp_dir.Text
        except:
            alert("Cannot read directory from form.")
            return

        if not directory or not directory.strip():
            alert("Please specify a temporary directory.")
            return

        directory = directory.strip()

        # Validate / create directory
        try:
            if not os.path.exists(directory):
                try:
                    os.makedirs(directory)
                except Exception as e:
                    errors.append("Cannot create directory: %s" % str(e))

            if not errors and not os.path.isdir(directory):
                errors.append("Specified path is not a directory.")

            # Test write access
            if not errors:
                try:
                    test_file = os.path.join(directory, ".write_test")
                    with open(test_file, 'w') as f:
                        f.write("test")
                    os.remove(test_file)
                except Exception as e:
                    errors.append("Directory is not writable: %s" % str(e))

            if not errors:
                try:
                    my_config.temp_dir = str2list(directory)
                except Exception as e:
                    errors.append("Cannot write to config: %s" % str(e))

        except Exception as e:
            errors.append("Specified path is invalid.\n%s" % str(e))

        if errors:
            alert("Can't save config.\n" + "\n".join(errors))
            return

        try:
            script.save_config()
        except Exception as e:
            alert("Config validated but could not be saved:\n%s" % str(e))
            return

        try:
            self.Close()
        except:
            pass

    # noinspection PyUnusedLocal
    # noinspection PyMethodMayBeStatic
    def pick(self, sender, args):
        """Open folder picker dialog"""
        try:
            path = pick_folder()
            if path:
                self.temp_dir.Text = path
        except Exception as e:
            alert("Cannot open folder picker:\n%s" % str(e))


if __name__ == '__main__':
    try:
        PurgeFamiliesConfigWindow('PurgeFamiliesConfig.xaml').ShowDialog()
    except Exception as e:
        try:
            alert("Cannot open Purge Families config:\n%s" % str(e))
        except:
            pass