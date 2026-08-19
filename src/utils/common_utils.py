from openpyxl.comments import Comment
from openpyxl.styles import Font, Alignment
import utils.settings as settings

def embed_settings_popup(ws, cell_coordinate="AB1", show_popup=True):
    """
    Embeds specific calculation settings into a cell as a clean, hoverable Excel comment.
    If show_popup is False, the function exits without doing anything.
    """
    if not show_popup:
        return 

    config = settings._SETTINGS_CONFIG
    
    # Helper to format options like a radio button list
    def format_opts(options, selected_val):
        return "\n".join([f"  {'◉' if opt_val == selected_val else '○'} {opt_label}" 
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
        
        "Measured 𝛅 values (Step 3):\n"
        f"{format_opts(step3_opts, config.get('CALC_MODE_STEP3'))}\n\n"
        
        "Average for RM (Step 7):\n"
        f"{format_opts(step7_opts, config.get('CALC_MODE_STEP7'))}"
    )
    
    # CRITICAL FIX: Use "pt" instead of "px". 
    # Windows Excel VML ignores px and collapses the box. pt (points) works cross-platform.
    settings_comment = Comment(
        text=clean_text, 
        author="DNT", 
        width="250pt",  
        height="320pt"  
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