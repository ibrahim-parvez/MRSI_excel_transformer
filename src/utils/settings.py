import json
import copy
import datetime

_SETTINGS_CONFIG = {
    # UI Toggles
    "ENABLE_FREEZE_PANE": True,
    "CO2_TEMP_MODE": "25 °C",
    #"ENABLE_YIELD_CALC": False,
    #"ENABLE_CO2_CALC": False,

    # Toggle for Stdev
    "STDEV_THRESHOLD_ENABLED": False,
    "STDEV_THRESHOLD": 0.08,
    
    # Outlier Configuration
    "OUTLIER_SIGMA": 2,                        
    "OUTLIER_EXCLUSION_MODE": "Individual",   
    
    # Split Calculation Modes
    "CALC_MODE_STEP3": "Last 6",               
    "CALC_MODE_STEP7": "All Values",           

    # Materials split by Type
    "REFERENCE_MATERIALS": {
        "Carbonate": [
            {"col_c": "IAEA 603", "col_d": "", "col_e": "", "col_f": "2.46", "col_g": "-2.37", "col_h": "", "color": "green"},
            {"col_c": "LSVEC",    "col_d": "", "col_e": "", "col_f": "-46.6", "col_g": "-26.7", "col_h": "", "color": "lightblue"},
            {"col_c": "NBS 18",   "col_d": "", "col_e": "", "col_f": "-5.01", "col_g": "-23.01", "col_h": "", "color": "red"},
            {"col_c": "NBS 19",   "col_d": "", "col_e": "", "col_f": "1.95",  "col_g": "-2.20",  "col_h": "", "color": "darkblue"}
        ],
        "Water": [
            {"col_c": "MRSI-STD-W1", "col_d": "", "col_e": "", "col_f": "-3.52", "col_g": "-0.58", "col_h": "", "color": "red"},
            {"col_c": "MRSI-STD-W2",  "col_d": "", "col_e": "", "col_f": "-214.79", "col_g": "-28.08", "col_h": "", "color": "darkblue"},
            {"col_c": "USGS W-67400",  "col_d": "", "col_e": "", "col_f": "1.25", "col_g": "-1.97", "col_h": "", "color": "orange"},
            {"col_c": "USGS W-64444",  "col_d": "", "col_e": "", "col_f": "-399.1", "col_g": "-51.14", "col_h": "", "color": "green"}
        ],
        "CO2": [
            {"col_c": "IAEA 603", "col_d": "", "col_e": "", "col_f": "2.46", "col_g": "7.86", "col_h": "", "color": "green"},
            {"col_c": "LSVEC",    "col_d": "", "col_e": "", "col_f": "-46.6", "col_g": "", "col_h": "", "color": "lightblue"},
            {"col_c": "NBS 18",   "col_d": "", "col_e": "", "col_f": "-5.01", "col_g": "-13.00", "col_h": "", "color": "red"},
            {"col_c": "NBS 19",   "col_d": "", "col_e": "", "col_f": "1.95",  "col_g": "8.03",  "col_h": "", "color": "darkblue"}
        ]
    },

    # Slope Groups split by Type
    "SLOPE_INTERCEPT_GROUPS": {
        "Carbonate": [
            ["NBS 18", "NBS 19"],
            ["NBS 18", "NBS 19", "IAEA 603"]
        ],
        "Water": [
            ["MRSI-STD-W1", "MRSI-STD-W2"],
            ["USGS W-67400", "USGS W-64444"]
        ],
        "Yield": [
            ["NBS 18", "NBS 19"],
            ["Carrara"]
        ],
        "CO2": [
            ["NBS 18", "NBS 19"],
            ["NBS 18", "NBS 19", "IAEA 603"]
        ],
    }
}

def get_setting(key, sub_key=None):
    """
    Returns the current value. 
    If sub_key is provided (e.g. 'Carbonate'), returns that specific subset.
    """
    val = _SETTINGS_CONFIG.get(key)
    
    # Return deep copies to prevent accidental reference mutation
    if key in ["REFERENCE_MATERIALS", "SLOPE_INTERCEPT_GROUPS"]:
        if sub_key and isinstance(val, dict):
            return [item[:] if isinstance(item, list) else item.copy() for item in val.get(sub_key, [])]
        return val # Return whole dict if no sub_key
    return val

def set_setting(key, value, sub_key=None):
    """
    Sets the new value. 
    If sub_key is provided (e.g. 'Carbonate'), updates only that entry in the dictionary.
    """
    if key == "STDEV_THRESHOLD_ENABLED":
        _SETTINGS_CONFIG[key] = bool(value)
        return True, "Updated"

    elif key == "STDEV_THRESHOLD":
        try:
            new_value = float(value)
            if new_value <= 0: return False, "Must be positive."
            _SETTINGS_CONFIG[key] = new_value
            return True, "Updated"
        except ValueError:
            return False, "Invalid number"

    elif key == "OUTLIER_SIGMA":
        if value in [1, 2, 3]:
            _SETTINGS_CONFIG[key] = value
            return True, "Updated"
        return False, "Invalid Sigma"
    
    elif key == "OUTLIER_EXCLUSION_MODE":
        _SETTINGS_CONFIG[key] = value
        return True, "Updated"

    elif key in ["CALC_MODE_STEP3", "CALC_MODE_STEP7"]:
        _SETTINGS_CONFIG[key] = value
        return True, "Updated"

    elif key in ["REFERENCE_MATERIALS", "SLOPE_INTERCEPT_GROUPS"]:
        if sub_key:
            if key not in _SETTINGS_CONFIG: _SETTINGS_CONFIG[key] = {}
            _SETTINGS_CONFIG[key][sub_key] = value
            return True, f"Updated {sub_key}"
        else:
            _SETTINGS_CONFIG[key] = value
            return True, "Updated all"

    # Fallback
    _SETTINGS_CONFIG[key] = value
    return True, "Updated"

def get_reference_names(material_type="Carbonate"):
    """Helper to get list of names for a specific material type."""
    mats = get_setting("REFERENCE_MATERIALS", sub_key=material_type)
    return [m["col_c"] for m in mats if m.get("col_c")]

# --- AUTO-SAVE & FILE MANAGEMENT LOGIC ---

_AUTO_SAVES = []

def auto_save():
    """Creates a timestamped auto-save of the current settings."""
    global _AUTO_SAVES
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
    # Insert at the beginning of the list
    _AUTO_SAVES.insert(0, {"timestamp": timestamp, "data": copy.deepcopy(_SETTINGS_CONFIG)})
    # Keep only the last 5 saves to avoid memory bloat
    if len(_AUTO_SAVES) > 5:
        _AUTO_SAVES.pop()

def get_auto_saves():
    """Returns a list of timestamps for the auto-saves."""
    return [qs["timestamp"] for qs in _AUTO_SAVES]

def restore_auto_save(index):
    """Restores the settings from a specific auto-save index."""
    global _SETTINGS_CONFIG
    if 0 <= index < len(_AUTO_SAVES):
        _SETTINGS_CONFIG.update(copy.deepcopy(_AUTO_SAVES[index]["data"]))
        return True
    return False

def export_to_file(filepath):
    """Exports the current settings dictionary to a JSON file."""
    try:
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(_SETTINGS_CONFIG, f, indent=4)
        return True, "Settings exported successfully."
    except Exception as e:
        return False, f"Export failed: {str(e)}"

def import_from_file(filepath):
    """Imports settings from a JSON file and updates the current configuration."""
    global _SETTINGS_CONFIG
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        if isinstance(data, dict):
            _SETTINGS_CONFIG.update(data)
            return True, "Settings imported successfully."
        return False, "Invalid file format. Expected a JSON dictionary."
    except Exception as e:
        return False, f"Import failed: {str(e)}"
    
def get_auto_save_data(row_index):
    """
    Returns the dictionary of settings associated with a specific auto-save.
    """
    try:
        return _AUTO_SAVES[row_index].get("data", {}) 
    except (IndexError, KeyError):
        return None