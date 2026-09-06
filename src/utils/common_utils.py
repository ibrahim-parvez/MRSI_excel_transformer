import errno
import os
import shutil
import tempfile

from openpyxl.comments import Comment
from openpyxl.styles import Font, Alignment
import utils.settings as settings


def save_workbook_atomic(wb, file_path):
    """
    Save `wb` to `file_path` without ever leaving a half-written file behind.

    openpyxl's own save opens the destination with ZipFile(..., "w"), which
    truncates it immediately. If the write then fails - disk full, a USB or
    network drive disappearing, a crash, a force-quit - the user's workbook is
    left truncated and unopenable.

    Writing to a temporary file in the SAME directory and then os.replace()-ing
    it into position makes the swap atomic: the destination is either the old
    file or the complete new one, never a partial one. Same directory matters,
    because os.replace is only atomic within one filesystem.
    """
    file_path = os.path.abspath(file_path)
    folder = os.path.dirname(file_path) or "."

    # os.replace() only needs write permission on the DIRECTORY, so without
    # this check a read-only workbook would be silently overwritten - which the
    # plain openpyxl save could never do. Fail the same way it used to.
    if os.path.exists(file_path) and not os.access(file_path, os.W_OK):
        raise PermissionError(errno.EACCES, "Permission denied", file_path)

    fd, tmp_path = tempfile.mkstemp(prefix=".dnt_save_", suffix=".xlsx", dir=folder)
    os.close(fd)
    try:
        wb.save(tmp_path)
        # mkstemp creates the file 0600; keep whatever permissions the
        # original had so the saved file does not become owner-only.
        if os.path.exists(file_path):
            try:
                shutil.copymode(file_path, tmp_path)
            except OSError:
                pass
        os.replace(tmp_path, file_path)
    except BaseException:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        raise

# Map the UI labels to their actual backend values
#sigma_opts = [("1σ", 1), ("2σ", 2), ("3σ", 3)]

def embed_settings_popup(ws, cell_coordinate="AB1", show_popup=True):
    """
    Embeds specific calculation settings into a cell as a clean, hoverable Excel comment.
    Uses safe ASCII characters to ensure line-heights render identically on Windows and Mac.
    Dynamically appends Yield and CO2 settings if they are enabled.
    """
    if not show_popup:
        return 

    config = settings._SETTINGS_CONFIG
    
    def format_opts(options, selected_val):
        return "\n".join([f"  {'[x]' if opt_val == selected_val else '[  ]'} {opt_label}" 
                          for opt_label, opt_val in options])

    # Map the UI labels to their actual backend values
    sigma_opts = [("1σ", 1), ("2σ", 2), ("3σ", 3)]
    
    excl_opts = [
        ("Individual (Keep Valid C or O)", "Individual"), 
        ("Exclude Entire Row", "Exclude Row")
    ]
    
    step3_opts = [
        ("Last 6", "Last 6"), 
        ("Last 6 Outliers Excluded", "Last 6 Outliers Excluded")
    ]
    
    step7_opts = [
        ("All Values", "All Values"), 
        ("Outliers Excluded", "Outliers Excluded")
    ]
    
    # Format Stdev Threshold based on Enable/Disable toggle ---
    stdev_enabled = config.get('STDEV_THRESHOLD_ENABLED', True)
    stdev_val = config.get('STDEV_THRESHOLD')
    stdev_display = f"{stdev_val}" if stdev_enabled else "Disabled"
    
    clean_text = (
        "--- Run Settings ---\n\n"
        f"Stdev Threshold: {stdev_display}\n\n"
        
        "Outlier Calculation (Sigma):\n"
        f"{format_opts(sigma_opts, config.get('OUTLIER_SIGMA'))}\n\n"
        
        "Exclusion Logic:\n"
        f"{format_opts(excl_opts, config.get('OUTLIER_EXCLUSION_MODE'))}\n\n"
        
        "Measured Delta values (Step 3):\n"
        f"{format_opts(step3_opts, config.get('CALC_MODE_STEP3'))}\n\n"
        
        "Average for RM (Step 7):\n"
        f"{format_opts(step7_opts, config.get('CALC_MODE_STEP7'))}"
    )

    # --- DYNAMIC SECTIONS ---
    calc_yield = config.get('CALC_YIELD', False)
    calc_co2 = config.get('CALC_CO2', False)
    
    added_height = 0
    
    if calc_yield:
        # Fetch compound settings, falling back to defaults if not found
        yield_comps = config.get('YIELD_COMPOUNDS', {"ref": "CaCO3", "samp": "MnCO3"})
        clean_text += (
            "\n\n--- Yield Settings ---\n"
            f"  • Ref Compound: {yield_comps.get('ref', 'Unknown')}\n"
            f"  • Sample Compound: {yield_comps.get('samp', 'Unknown')}"
        )
        added_height += 65 # Increase box height to fit this text
        
    if calc_co2:
        co2_temp = config.get('CO2_TEMP_MODE', 'Custom')
        if co2_temp == "Custom":
            custom_label = str(config.get('CO2_CUSTOM_LABEL') or "").strip()
            if custom_label:
                co2_temp = f"Custom ({custom_label})"
        clean_text += (
            "\n\n--- CO2 Settings ---\n"
            f"  • Temp Mode: {co2_temp}"
        )
        added_height += 50 # Increase box height to fit this text

    total_height = 320 + added_height
    
    # Windows Excel VML ignores px and collapses the box. pt (points) works cross-platform.
    settings_comment = Comment(
        text=clean_text, 
        author="DNT", 
        width="250pt",  
        height=f"{total_height}pt"  
    )
    
    # Target the cell, set the text, and attach the comment
    target_cell = ws[cell_coordinate]
    target_cell.value = "⚙️ Settings"
    target_cell.comment = settings_comment
    
    # Style the cell so it looks distinct (blue, bold, centered)
    target_cell.font = Font(color="0052cc", bold=True)
    target_cell.alignment = Alignment(horizontal="center", vertical="center")

def normalize_name(s):
    if s is None:
        return ''
    return ' '.join(str(s).split()).lower()