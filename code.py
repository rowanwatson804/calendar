# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk
from tkinter import font as tkFont # Import font module
from tkinter import messagebox
from datetime import datetime, date, time, timedelta
import calendar
import json
import os

# Attempt to import tkcalendar components
try:
    from tkcalendar import DateEntry, Calendar # Import Calendar widget too
    tkcalendar_AVAILABLE = True
except ImportError:
    tkcalendar_AVAILABLE = False
    print("Warning: 'tkcalendar' library not found. Some features disabled.")
    print("         Install with: pip install tkcalendar")

# Set first day of the week (optional, Monday is default in many locales)
# calendar.setfirstweekday(calendar.SUNDAY) # Or MONDAY
# Let's default to Monday for consistency in the view
calendar.setfirstweekday(calendar.MONDAY)


# --- Constants ---
DATE_FORMAT = "%B %d, %Y" # For display in treeview, status messages etc.
DATE_FORMAT_SHORT = "%a, %b %d, %Y" # For Week View day headers
TIME_FORMAT = "%H:%M:%S"
DATETIME_ISO_FORMAT = "%Y-%m-%dT%H:%M:%S" # For saving/loading
DATETIME_DISPLAY_FORMAT = "%b %d, %Y %H:%M:%S" # For treeview display
DEFAULT_TIME = "00:00:00"
UPDATE_INTERVAL_MS = 1000
SAVE_FILENAME = "event_tracker_data.json" # Renamed to reflect content
STATUS_CLEAR_DELAY_MS = 4000

# Sort Options
SORT_ALPHA = "Alphabetical (A-Z)"
SORT_ALPHA_REV = "Alphabetical (Z-A)"
SORT_CLOSEST = "Closest First"
SORT_OPTIONS = [SORT_CLOSEST, SORT_ALPHA, SORT_ALPHA_REV]
DEFAULT_SORT = SORT_CLOSEST

# Indicator Symbols & Thresholds
INDICATOR_PAST = '✅'; INDICATOR_URGENT = '🔥'; INDICATOR_SOON = '⏳'
INDICATOR_NEAR = '🗓️'; INDICATOR_FAR = '•'
INDICATOR_THRESHOLD_URGENT = 1; INDICATOR_THRESHOLD_SOON = 7; INDICATOR_THRESHOLD_NEAR = 30

# Color Palettes (Added text widget colors)
LIGHT_COLORS = {
    "bg": "#F0F0F0", "fg": "#000000",
    "frame_bg": "#F0F0F0", "frame_fg": "#000000",
    "btn_bg": "#E1E1E1", "btn_fg": "#000000", "btn_active": "#ECECEC",
    "entry_bg": "#FFFFFF", "entry_fg": "#000000",
    "tree_bg": "#FFFFFF", "tree_fg": "#000000", "tree_odd": "#F0F0F0", "tree_even": "#FFFFFF",
    "tree_head_bg": "#E1E1E1", "tree_head_fg": "#000000",
    "tree_select_bg": "#0078D7", "tree_select_fg": "#FFFFFF",
    "status_bg": "#F0F0F0", "status_fg": "#000000",
    "tab_bg": "#D1D1D1", "tab_fg": "#000000", "tab_active_bg": "#F0F0F0", "tab_active_fg": "#000000",
    "text_bg": "#FFFFFF", "text_fg": "#000000", # For Text widget
    "text_bold_fg": "#0000AA", # Bold text color (e.g., dates)
    # Calendar specific
    "cal_normal_bg": "white", "cal_normal_fg": "black",
    "cal_weekend_bg": "white", "cal_weekend_fg": "black",
    "cal_header_bg": "#E1E1E1", "cal_header_fg": "black",
    "cal_select_bg": "lightblue", "cal_select_fg": "black",
    "cal_event_marker_bg": "lightblue", "cal_event_marker_fg": "black",
    "cal_border": "grey",
    "cal_entry_select_bg": "lightblue", "cal_entry_select_fg": "black",
}

DARK_COLORS = {
    "bg": "#2E2E2E", "fg": "#EAEAEA",
    "frame_bg": "#2E2E2E", "frame_fg": "#EAEAEA",
    "btn_bg": "#505050", "btn_fg": "#EAEAEA", "btn_active": "#6A6A6A",
    "entry_bg": "#3C3C3C", "entry_fg": "#EAEAEA",
    "tree_bg": "#3C3C3C", "tree_fg": "#EAEAEA", "tree_odd": "#303030", "tree_even": "#3C3C3C",
    "tree_head_bg": "#505050", "tree_head_fg": "#EAEAEA",
    "tree_select_bg": "#005A9E", "tree_select_fg": "#EAEAEA",
    "status_bg": "#2E2E2E", "status_fg": "#EAEAEA",
    "tab_bg": "#404040", "tab_fg": "#EAEAEA", "tab_active_bg": "#2E2E2E", "tab_active_fg": "#EAEAEA",
    "text_bg": "#3C3C3C", "text_fg": "#EAEAEA", # For Text widget
    "text_bold_fg": "#87CEEB", # Bold text color (e.g., dates) - sky blue
     # Calendar specific
    "cal_normal_bg": "#3C3C3C", "cal_normal_fg": "white",
    "cal_weekend_bg": "#3C3C3C", "cal_weekend_fg": "white",
    "cal_header_bg": "#505050", "cal_header_fg": "white",
    "cal_select_bg": "#005A9E", "cal_select_fg": "white",
    "cal_event_marker_bg": "#005A9E", "cal_event_marker_fg": "white",
    "cal_border": "#6A6A6A",
    "cal_entry_select_bg": "#005A9E", "cal_entry_select_fg": "white",
}

# INITIAL_EVENTS_DATA (Ensure this is defined)
INITIAL_EVENTS_DATA = {
    "New Year's Day": (1, 1), "Martin Luther King, Jr. Day (Approx)": (1, 15),
    "Groundhog Day": (2, 2), "My Birthday": (2, 5), "Valentine's Day": (2, 14),
    "Presidents' Day (Approx)": (2, 19), "St. Patrick's Day": (3, 17),
    "April Fools' Day": (4, 1), "Memorial Day (Approx)": (5, 27),
    "Juneteenth": (6, 19), "Independence Day": (7, 4),
    "Labor Day (Approx)": (9, 2), "Columbus Day (Approx)": (10, 14),
    "Halloween": (10, 31), "Veterans Day": (11, 11),
    "Thanksgiving Day (Approx)": (11, 28), "Christmas Day": (12, 25),
    "New Year's Eve": (12, 31),
}


# --- Data Store ---
tracked_events = []
settings = {"dark_mode": False} # Default to light mode

# --- Global Widget References ---
date_input_widget = None; status_label = None; remove_button = None
status_clear_job = None; update_job_id = None; event_tree = None
location_entry_var = None; label_entry_var = None; date_entry_var = None
time_entry_var = None; sort_var = None
notebook = None
calendar_widget = None
selected_date_event_label = None # Label widget itself
selected_date_event_var = None # StringVar for the label
dark_mode_var = None # BooleanVar for the checkbutton
root = None # Define root globally
style = None # Style object
input_frame = None # Make input_frame global for styling

# Week View Globals
week_view_tab_frame = None
week_view_label_var = None
week_view_text = None
current_week_start_date = None # Stores the date object of the start of the displayed week

# --- Helper Functions ---

def get_week_start(ref_date):
    """Calculates the start date (Monday) of the week containing ref_date."""
    # weekday() returns 0 for Monday, 6 for Sunday
    start_delta = timedelta(days=ref_date.weekday())
    return ref_date - start_delta

def calculate_next_occurrence(month, day):
    today = date.today(); year = today.year; target_date_this_year = None
    try: target_date_this_year = date(year, month, day)
    except ValueError: # Handle invalid dates like Feb 30
        if month == 2 and day == 29: # Handle Leap Day specifically
            temp_year = year;
            # Find the next leap year starting from the current year
            while not calendar.isleap(temp_year): temp_year += 1
            # If today is Mar 1st or later in a leap year, find the *next* leap year
            if today >= date(year, 3, 1) and calendar.isleap(year):
                 year = temp_year + 4;
                 while not calendar.isleap(year): year += 4 # Ensure it's a leap year
            else: # Otherwise, use the first leap year found
                year = temp_year
            try: return date(year, month, day)
            except ValueError: return None # Should not happen if logic is correct
        else: return None # Other invalid dates (e.g., Apr 31)
    # If valid date this year
    if target_date_this_year and target_date_this_year < today:
        # Target date has passed this year, move to next year
        year += 1
        try:
            # Special handling if next year isn't a leap year for Feb 29
            if month == 2 and day == 29 and not calendar.isleap(year):
                year += 1 # Go to the year after next
                while not calendar.isleap(year): year += 1 # Find the next leap year
            return date(year, month, day)
        except ValueError: return None # Should not happen
    else:
        # Target date is today or in the future this year
        return target_date_this_year


def format_timedelta(delta):
    total_seconds = int(delta.total_seconds()); prefix = "In: " if total_seconds >= 0 else "Ago: "
    total_seconds = abs(total_seconds); sign = 1 if prefix == "In: " else -1
    days, rem_secs = divmod(total_seconds, 86400); hours, rem_secs = divmod(rem_secs, 3600); minutes, seconds = divmod(rem_secs, 60)
    if sign == -1 and days == 0 and hours == 0 and minutes == 0 and seconds < 60: return "Just now or Past"
    parts = [];
    if days > 0: parts.append(f"{days}d");
    if hours > 0: parts.append(f"{hours}h")
    if minutes > 0: parts.append(f"{minutes}m")
    if (days == 0 and total_seconds > 0) or not parts : parts.append(f"{seconds}s")
    if not parts: return "Now"
    return prefix + " ".join(parts)

def calculate_indicator_symbol(target_dt):
    now = datetime.now(); delta = target_dt - now; days_left = delta.total_seconds() / 86400.0
    if days_left < 0: return INDICATOR_PAST
    elif days_left <= INDICATOR_THRESHOLD_URGENT: return INDICATOR_URGENT
    elif days_left <= INDICATOR_THRESHOLD_SOON: return INDICATOR_SOON
    elif days_left <= INDICATOR_THRESHOLD_NEAR: return INDICATOR_NEAR
    else: return INDICATOR_FAR

# --- Status Bar Function ---
def update_status(message, clear_after=True):
    global status_label, status_clear_job, root
    if status_label and root:
        try:
            if status_label.winfo_exists(): # Check if widget exists
                status_label.config(text=message)
                if status_clear_job:
                    try: root.after_cancel(status_clear_job)
                    except: pass # Ignore errors if job already cancelled
                    status_clear_job = None
                if clear_after and message:
                    # Schedule clearing only if message is not empty
                    if root.winfo_exists(): # Check root again before scheduling
                        status_clear_job = root.after(STATUS_CLEAR_DELAY_MS, lambda: update_status("", clear_after=False))
            # else: print(f"Status label destroyed, cannot update: {message}") # Optional debug
        except tk.TclError: pass # Handle cases where widget might be destroyed mid-call
    # else: print(f"Status (Label/Root not ready): {message}") # Optional debug

# --- Save/Load Functions ---
def save_data(): # Saves both events and settings
    global settings # Access global settings dictionary
    custom_events_to_save = []
    for event in tracked_events:
        if event.get('is_custom', False):
            event_data = {'label': event['label'],'target_dt_iso': event['target_dt'].strftime(DATETIME_ISO_FORMAT)}
            if event.get('location'): event_data['location'] = event['location']
            custom_events_to_save.append(event_data)
    data_to_save = { "settings": settings, "events": custom_events_to_save }
    try:
        with open(SAVE_FILENAME, 'w') as f: json.dump(data_to_save, f, indent=4)
        update_status(f"Saved settings and {len(custom_events_to_save)} custom events.")
    except Exception as e: update_status(f"Error saving data: {e}"); print(f"Error saving data to {SAVE_FILENAME}: {e}")

def load_data(): # Loads both settings and events
    global settings # Modify global settings dictionary
    loaded_events_data = []
    if not os.path.exists(SAVE_FILENAME): update_status("No data file found. Using defaults.", clear_after=False); return loaded_events_data
    try:
        with open(SAVE_FILENAME, 'r') as f: loaded_data = json.load(f)
        loaded_settings = loaded_data.get("settings", {}); settings["dark_mode"] = loaded_settings.get("dark_mode", False)
        loaded_events_raw = loaded_data.get("events", []); loaded_count = 0
        for item in loaded_events_raw:
            try:
                target_dt = datetime.strptime(item['target_dt_iso'], DATETIME_ISO_FORMAT)
                if 'label' not in item or not item['label']: continue
                loaded_event = {'label': item['label'], 'target_dt': target_dt, 'is_custom': True, 'tree_id': None, 'location': item.get('location', None)}
                loaded_events_data.append(loaded_event); loaded_count += 1
            except (KeyError, ValueError, TypeError) as e: print(f"Skipping invalid event item during load: {item}. Error: {e}")
        update_status(f"Loaded settings and {loaded_count} custom events.")
        return loaded_events_data
    except json.JSONDecodeError as e: update_status(f"Error: Corrupted data in {SAVE_FILENAME}. Using defaults."); print(f"Error parsing JSON: {e}"); settings = {"dark_mode": False}; return []
    except Exception as e: update_status(f"Error loading data file: {e}. Using defaults."); print(f"Error loading file: {e}"); settings = {"dark_mode": False}; return []


# --- Core Logic & UI Handlers ---

# Function to apply theme colors (CORRECTED VERSION)
def apply_styles():
    global root, style, settings, calendar_widget, date_input_widget, status_label
    global selected_date_event_label, event_tree, input_frame, week_view_text
    is_dark = settings.get("dark_mode", False)
    colors = DARK_COLORS if is_dark else LIGHT_COLORS

    try:
        if not style: print("Style object not ready"); return

        # --- Configure ttk Styles ---
        style.configure('.', background=colors['bg'], foreground=colors['fg'])
        style.configure('TLabel', background=colors['bg'], foreground=colors['fg'])
        style.configure('TButton', background=colors['btn_bg'], foreground=colors['btn_fg'])
        style.map('TButton', background=[('active', colors['btn_active'])])
        style.configure('TEntry', fieldbackground=colors['entry_bg'], foreground=colors['entry_fg'], insertcolor=colors['fg']) # insertcolor for cursor
        style.configure('Treeview', background=colors['tree_bg'], foreground=colors['tree_fg'], fieldbackground=colors['tree_bg'])
        style.map('Treeview', background=[('selected', colors['tree_select_bg'])], foreground=[('selected', colors['tree_select_fg'])])
        style.configure('Treeview.Heading', background=colors['tree_head_bg'], foreground=colors['tree_head_fg'], relief=tk.FLAT)
        style.configure('oddrow.Treeview', background=colors['tree_odd'], foreground=colors['tree_fg'])
        style.configure('evenrow.Treeview', background=colors['tree_even'], foreground=colors['tree_fg'])
        style.configure('TCheckbutton', background=colors['bg'], foreground=colors['fg'])
        style.map('TCheckbutton', indicatorcolor=[('selected', colors['fg']), ('!selected', colors['entry_bg'])], background=[('active', colors['bg'])])
        style.configure('TNotebook', background=colors['bg'])
        style.configure('TNotebook.Tab', background=colors['tab_bg'], foreground=colors['tab_fg'], padding=[5, 2])
        style.map('TNotebook.Tab', background=[('selected', colors['tab_active_bg'])], foreground=[('selected', colors['tab_active_fg'])])
        # Configure TLabelframe style (affects input_frame)
        style.configure('TLabelframe', background=colors['frame_bg'], bordercolor=colors['fg'])
        style.configure('TLabelframe.Label', background=colors['frame_bg'], foreground=colors['frame_fg']) # Styles the label *text* of the frame
        style.configure('Vertical.TScrollbar', background=colors['btn_bg'], troughcolor=colors['bg'])
        style.map('Vertical.TScrollbar', background=[('active', colors['btn_active'])])
        style.configure('Horizontal.TScrollbar', background=colors['btn_bg'], troughcolor=colors['bg'])
        style.map('Horizontal.TScrollbar', background=[('active', colors['btn_active'])])

        # --- Configure Specific Non-ttk or Complex Widgets ---
        if root: root.config(background=colors['bg']) # Root window is standard Tk

        # input_frame (ttk.LabelFrame) is now styled via TLabelframe style above
        # status_label (ttk.Label) is now styled via TLabel style above
        # selected_date_event_label (ttk.Label) is now styled via TLabel style above
        # Re-apply non-style config options if they get overridden or aren't part of the style
        if status_label and status_label.winfo_exists():
             status_label.config(relief=tk.SUNKEN) # Ensure sunken relief is maintained
        if selected_date_event_label and selected_date_event_label.winfo_exists():
             selected_date_event_label.config(relief=tk.GROOVE, borderwidth=1) # Ensure groove relief is maintained


        # Configure Week View Text Widget (tk.Text is standard Tk, so .config is OK)
        if week_view_text and week_view_text.winfo_exists():
            week_view_text.config(background=colors['text_bg'], foreground=colors['text_fg'],
                                  insertbackground=colors['fg']) # Cursor color
            # Configure tags used in week view
            week_view_text.tag_configure("bold_date", font=('Helvetica', 10, 'bold'), foreground=colors['text_bold_fg'])
            week_view_text.tag_configure("event_item", lmargin1=10, lmargin2=10) # Indent event lines

        # --- Configure tkcalendar Widgets ---
        if tkcalendar_AVAILABLE:
            cal_options = {
                'background': colors['cal_normal_bg'], 'foreground': colors['cal_normal_fg'],
                'bordercolor': colors['cal_border'], 'headersbackground': colors['cal_header_bg'],
                'headersforeground': colors['cal_header_fg'], 'normalbackground': colors['cal_normal_bg'],
                'normalforeground': colors['cal_normal_fg'], 'weekendbackground': colors['cal_weekend_bg'],
                'weekendforeground': colors['cal_weekend_fg'], 'selectbackground': colors['cal_select_bg'],
                'selectforeground': colors['cal_select_fg'],
            }
            if calendar_widget and calendar_widget.winfo_exists():
                calendar_widget.config(**cal_options)
                calendar_widget.tag_config('event_marker',
                                           background=colors['cal_event_marker_bg'],
                                           foreground=colors['cal_event_marker_fg'])
            if date_input_widget and isinstance(date_input_widget, DateEntry) and date_input_widget.winfo_exists():
                 try:
                    date_input_widget.config(**cal_options)
                    # These selectbackground/foreground are specific DateEntry config options
                    date_input_widget.config(selectbackground=colors['cal_entry_select_bg'],
                                             selectforeground=colors['cal_entry_select_fg'])
                    # Rely on TEntry style for the entry part, avoid direct config if possible
                 except tk.TclError as e: print(f"TclError configuring DateEntry styles: {e}")


        # --- Refresh Treeview Row Tags ---
        if event_tree and event_tree.winfo_exists():
            for i, item_id in enumerate(event_tree.get_children()):
                tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                current_tags = list(event_tree.item(item_id, 'tags'))
                # Remove old tags cleanly before adding new one
                if 'oddrow' in current_tags: current_tags.remove('oddrow')
                if 'evenrow' in current_tags: current_tags.remove('evenrow')
                if tag not in current_tags: current_tags.append(tag)
                event_tree.item(item_id, tags=tuple(current_tags))

        # --- Refresh week view ONLY if data changes, not just style ---
        # The update will be triggered by on_tab_changed or sort_and_redisplay when needed.
        # *** DO NOT CALL update_week_view() from here ***
        # if notebook and week_view_tab_frame:
        #     try:
        #         if notebook.winfo_exists() and week_view_tab_frame.winfo_exists():
        #             current_tab_path = notebook.select()
        #             if current_tab_path == str(week_view_tab_frame): # Compare widget path strings
        #                 update_week_view() # <-- REMOVED THIS CALL
        #     except tk.TclError: pass


    except tk.TclError as e: print(f"TclError applying styles: {e}")
    except Exception as e: print(f"Unexpected error applying styles: {e}")


# Function called by the Checkbutton
def toggle_dark_mode():
    global dark_mode_var, settings
    if dark_mode_var:
        settings["dark_mode"] = dark_mode_var.get()
        apply_styles() # Apply styles will trigger week view update if visible
        # Trigger week view update IF it's the selected tab after style change
        on_tab_changed(None)
        update_status(f"Switched to {'Dark' if settings['dark_mode'] else 'Light'} Mode.", clear_after=True)
    else: print("Error: dark_mode_var not initialized.")

# --- Week View Functions ---

def update_week_view():
    """Refreshes the content of the week view text widget."""
    global week_view_text, week_view_label_var, current_week_start_date, tracked_events
    if not all([week_view_text, week_view_label_var, current_week_start_date]):
        print("Week view components not ready for update.") # Debug
        return # Components not ready

    try:
        if not week_view_text.winfo_exists():
            return # Widget destroyed
    except tk.TclError:
        return # Widget destroyed (TclError)

    week_end_date = current_week_start_date + timedelta(days=6)
    week_range_str = f"Week: {current_week_start_date.strftime(DATE_FORMAT_SHORT)} - {week_end_date.strftime(DATE_FORMAT_SHORT)}"
    week_view_label_var.set(week_range_str)

    # Filter events for the current week
    events_this_week = []
    for event in tracked_events:
        if 'target_dt' not in event: continue # Skip if event data is incomplete
        event_date = event['target_dt'].date()
        if current_week_start_date <= event_date <= week_end_date:
            events_this_week.append(event)

    # Sort events by datetime
    events_this_week.sort(key=lambda ev: ev['target_dt'])

    # Enable text widget for update, clear, disable after
    try:
        week_view_text.config(state=tk.NORMAL)
        week_view_text.delete('1.0', tk.END)

        # Group events by day and format
        for i in range(7):
            day_date = current_week_start_date + timedelta(days=i)
            day_events = [ev for ev in events_this_week if ev['target_dt'].date() == day_date]

            # Add day header
            day_header = day_date.strftime(DATE_FORMAT_SHORT) # e.g., "Mon, Jan 01, 2024"
            week_view_text.insert(tk.END, f"\n{day_header}\n", "bold_date")

            if day_events:
                for ev in day_events:
                    time_str = ev['target_dt'].strftime(TIME_FORMAT)
                    loc_str = f" ({ev.get('location')})" if ev.get('location') else ""
                    event_line = f"  {time_str} - {ev['label']}{loc_str}\n"
                    week_view_text.insert(tk.END, event_line, "event_item")
            else:
                 week_view_text.insert(tk.END, "  No events scheduled\n", "event_item") # Use same tag for consistent indent

        # Remove leading newline if present
        if week_view_text.get("1.0", "1.1") == "\n":
             week_view_text.delete("1.0", "1.1")

        week_view_text.config(state=tk.DISABLED) # Make read-only again
        week_view_text.see("1.0") # Scroll to top

    except tk.TclError as e:
        print(f"Error updating week view text widget: {e}")
        try:
            # Ensure it's disabled even if error occurred mid-update
             if week_view_text.winfo_exists():
                 week_view_text.config(state=tk.DISABLED)
        except tk.TclError: pass # Widget might be gone


def show_previous_week():
    global current_week_start_date
    if current_week_start_date:
        current_week_start_date -= timedelta(weeks=1)
        update_week_view()

def show_next_week():
    global current_week_start_date
    if current_week_start_date:
        current_week_start_date += timedelta(weeks=1)
        update_week_view()

def show_current_week():
    global current_week_start_date
    today = date.today()
    current_week_start_date = get_week_start(today)
    update_week_view()

def on_tab_changed(event):
    """Called when the notebook tab is changed."""
    global notebook, week_view_tab_frame
    try:
        if not notebook or not notebook.winfo_exists(): return

        selected_tab_widget_path = notebook.select()
        # Check if the selected tab is the week view tab
        if week_view_tab_frame and week_view_tab_frame.winfo_exists() and selected_tab_widget_path == str(week_view_tab_frame):
             update_week_view() # Refresh week view when it becomes visible
        # Check if selected tab is Calendar View and update markers/selected date info
        elif calendar_tab_frame and calendar_tab_frame.winfo_exists() and selected_tab_widget_path == str(calendar_tab_frame):
            update_calendar_markers() # Refresh markers
            show_events_for_selected_date() # Refresh label below calendar

    except tk.TclError:
        pass # Notebook might not exist yet or tab frame gone
    except Exception as e:
        print(f"Error in on_tab_changed: {e}")


# --- (Other core logic/UI handlers remain largely the same) ---

def update_calendar_markers():
    global calendar_widget, tracked_events, settings # Need settings for colors
    if not tkcalendar_AVAILABLE or not calendar_widget: return
    is_dark = settings.get("dark_mode", False)
    colors = DARK_COLORS if is_dark else LIGHT_COLORS
    try:
        if not calendar_widget.winfo_exists(): return
        calendar_widget.calevent_remove(tag='event_marker')
    except tk.TclError: return
    event_dates = set(event['target_dt'].date() for event in tracked_events if 'target_dt' in event)
    for event_date in event_dates:
        try:
             if calendar_widget.winfo_exists():
                 calendar_widget.calevent_create(event_date, 'Event', tags='event_marker')
             else: break
        except tk.TclError: break
    try:
        if calendar_widget.winfo_exists():
            # Ensure marker style uses current theme colors
            calendar_widget.tag_config('event_marker', background=colors['cal_event_marker_bg'], foreground=colors['cal_event_marker_fg'])
    except tk.TclError: pass


def sort_and_redisplay():
    global tracked_events, event_tree, root, sort_var
    if not event_tree or not sort_var or not root: return
    try:
        sort_method = sort_var.get(); now = datetime.now()
        # Filter out any potential malformed events before sorting
        valid_events = [ev for ev in tracked_events if 'target_dt' in ev and 'label' in ev]

        if sort_method == SORT_ALPHA: valid_events.sort(key=lambda event: event['label'].lower())
        elif sort_method == SORT_ALPHA_REV: valid_events.sort(key=lambda event: event['label'].lower(), reverse=True)
        else: valid_events.sort(key=lambda event: abs(event['target_dt'] - now)) # Closest first (past or future)

        tracked_events = valid_events # Update the main list with sorted valid events

        if event_tree.winfo_exists():
            # Clear existing items before re-inserting
            for item in event_tree.get_children(): event_tree.delete(item)
        else: return # Stop if treeview is gone

        # Re-populate the treeview
        for i, event in enumerate(tracked_events):
            target_dt_str = event['target_dt'].strftime(DATETIME_DISPLAY_FORMAT)
            difference = event['target_dt'] - now; formatted_diff = format_timedelta(difference)
            tag = 'evenrow' if i % 2 == 0 else 'oddrow' # Determine tag based on index
            indicator = calculate_indicator_symbol(event['target_dt'])
            location_str = event.get('location', '')
            if event_tree.winfo_exists():
                # Apply the correct row tag during insert
                new_tree_id = event_tree.insert('', tk.END, values=(indicator, event['label'], location_str, target_dt_str, formatted_diff), tags=(tag,))
                # Update the event dict with its new Treeview ID
                tracked_events[i]['tree_id'] = new_tree_id # Use index i which corresponds to the sorted list
            else: break # Stop if treeview destroyed mid-loop

        update_remove_button_state()
        update_calendar_markers() # Updates markers with current theme color
        # Update week view if it's currently visible
        on_tab_changed(None) # Trigger update if week view is selected

    except tk.TclError as e: print(f"Error during sort/redisplay (widget likely destroyed): {e}")
    except Exception as e: print(f"Unexpected error during sort/redisplay: {e}")


def update_display():
    global tracked_events, root, event_tree, update_job_id
    # Check if essential widgets exist
    if not event_tree or not root or not root.winfo_exists():
        # Cancel timer if root or treeview is gone
        if update_job_id and root:
            try: root.after_cancel(update_job_id)
            except: pass # Ignore error if already cancelled or root gone
        update_job_id = None; return # Stop the update loop

    now = datetime.now(); items_to_remove_from_tracking = []
    current_tracked_tree_ids = set() # Keep track of IDs currently in our tracked_events list

    try:
        # Get visible items safely
        visible_items_in_treeview = set(event_tree.get_children())
    except tk.TclError:
        # Treeview was destroyed between checks, cancel timer and stop
        if update_job_id and root:
            try: root.after_cancel(update_job_id)
            except: pass
        update_job_id = None; return

    # Iterate over a copy in case tracked_events is modified elsewhere
    current_tracked_events = tracked_events[:]
    for event in current_tracked_events:
        if 'target_dt' not in event: continue # Skip malformed event

        tree_id = event.get('tree_id')
        if tree_id:
            current_tracked_tree_ids.add(tree_id) # Record the ID we expect to see

            # Check if the event *should* be in the treeview and *is* actually there
            if tree_id in visible_items_in_treeview:
                difference = event['target_dt'] - now; formatted_diff = format_timedelta(difference)
                indicator = calculate_indicator_symbol(event['target_dt'])
                try:
                    # Update existing item if it's still there
                    event_tree.set(tree_id, column='diff', value=formatted_diff)
                    event_tree.set(tree_id, column='indicator', value=indicator)
                except tk.TclError:
                    # Item was deleted from treeview between get_children and set, mark for removal from tracking
                    items_to_remove_from_tracking.append(event)
                    visible_items_in_treeview.discard(tree_id) # No longer visible
            else:
                 # Event has a tree_id but it's NOT in visible_items_in_treeview
                 # This means it was likely deleted from the tree (e.g., by remove_selected_event)
                 # Mark it for removal from our internal list to maintain consistency
                 items_to_remove_from_tracking.append(event)
        # else: # Event doesn't have a tree_id yet (maybe just added, will be handled by sort_and_redisplay)
             # pass

    # Clean up internal tracking list if necessary
    if items_to_remove_from_tracking:
        tracked_events = [ev for ev in tracked_events if ev not in items_to_remove_from_tracking]

    # Schedule next update, checking root existence again
    try:
        if root.winfo_exists():
            update_job_id = root.after(UPDATE_INTERVAL_MS, update_display)
        else:
            update_job_id = None # Root gone, stop loop
    except tk.TclError: update_job_id = None # Window destroyed during check
    except Exception as e: print(f"Unexpected error scheduling next update: {e}"); update_job_id = None


def add_event_to_tracker(label, target_dt, location=None, is_custom=False):
    # Check for duplicate label before adding
    if any(e.get('label','').lower() == label.lower() for e in tracked_events):
        messagebox.showwarning("Duplicate Label", f"An event with the label '{label}' already exists.")
        return None # Indicate failure due to duplicate

    if not target_dt:
        messagebox.showerror("Internal Error", "Target datetime missing for add_event_to_tracker.")
        return None # Indicate failure

    new_event = {
        'label': label,
        'target_dt': target_dt,
        'location': location if location else None,
        'is_custom': is_custom,
        'tree_id': None # Will be set when added to treeview by sort_and_redisplay
    }
    tracked_events.append(new_event)
    return new_event # Return the added event


def add_custom_event(event=None): # Accept event argument for binding
    global date_input_widget, location_entry_var, label_entry_var, date_entry_var, time_entry_var, label_entry, time_entry # Added global refs for focus
    # Ensure all input widgets are initialized
    if not all([label_entry_var, location_entry_var, date_entry_var, time_entry_var, date_input_widget, label_entry, time_entry]):
        messagebox.showerror("Error", "Input fields not ready.")
        return

    label = label_entry_var.get().strip()
    time_str = time_entry_var.get().strip()
    location = location_entry_var.get().strip() # Optional
    target_date_obj = None

    if not label:
        messagebox.showerror("Input Error", "Please enter an event label.")
        label_entry.focus_set() # Focus the label entry
        return

    # --- Get Date ---
    try:
        if tkcalendar_AVAILABLE and isinstance(date_input_widget, DateEntry):
            target_date_obj = date_input_widget.get_date() # Returns a date object
        else: # Using standard ttk.Entry for date
            date_str = date_entry_var.get()
            target_date_obj = datetime.strptime(date_str, DATE_FORMAT).date()
    except ValueError:
         # Determine the correct expected format for the error message
         date_fmt_display = DATE_FORMAT
         if tkcalendar_AVAILABLE and isinstance(date_input_widget, DateEntry):
             # Use the pattern set for DateEntry, converting tkcalendar format to upper for display
             try:
                 tk_pattern = date_input_widget.cget('date_pattern') # e.g., 'mm/dd/yyyy'
                 date_fmt_display = tk_pattern.upper() # e.g., 'MM/DD/YYYY'
             except tk.TclError: pass # Widget might be gone
         messagebox.showerror("Input Error", f"Invalid date format. Please use {date_fmt_display}.")
         try:
            if date_input_widget.winfo_exists(): date_input_widget.focus_set()
         except tk.TclError: pass
         return
    except Exception as e: # Catch other potential errors during date parsing
        messagebox.showerror("Input Error", f"Could not parse date: {e}")
        try:
            if date_input_widget.winfo_exists(): date_input_widget.focus_set()
        except tk.TclError: pass
        return

    # --- Get Time ---
    try:
        target_time_obj = datetime.strptime(time_str, TIME_FORMAT).time()
    except ValueError:
        messagebox.showerror("Input Error", f"Invalid time format. Please use HH:MM:SS (e.g., {DEFAULT_TIME}).")
        try:
            if time_entry.winfo_exists():
                time_entry.focus_set()
                time_entry.selection_range(0, tk.END)
        except tk.TclError: pass
        return

    # --- Combine Date and Time ---
    if target_date_obj and target_time_obj:
        target_dt = datetime.combine(target_date_obj, target_time_obj)

        # --- Add to Tracker (handles duplicate label check) ---
        new_event = add_event_to_tracker(label, target_dt, location=location, is_custom=True)

        if new_event: # If adding was successful (no duplicate label)
            # --- Clear Input Fields ---
            label_entry_var.set("")
            # Reset date field carefully
            try:
                if tkcalendar_AVAILABLE and isinstance(date_input_widget, DateEntry):
                    if date_input_widget.winfo_exists(): date_input_widget.set_date(date.today()) # Reset to today
                else:
                    if date_input_widget.winfo_exists(): date_entry_var.set(date.today().strftime(DATE_FORMAT))
            except tk.TclError: pass # Widget might be destroyed
            except Exception as e: print(f"Error resetting date field: {e}") # Log other errors

            time_entry_var.set(DEFAULT_TIME) # Reset time
            location_entry_var.set("") # Clear location

            update_status(f"Added custom event: {label}")
            sort_and_redisplay() # Sort includes redisplay and updates calendar/week view
            try:
                if label_entry.winfo_exists(): label_entry.focus_set() # Set focus back to label for next entry
            except tk.TclError: pass

        # If new_event is None, it means add_event_to_tracker showed an error (duplicate label)
        # No need to show another error here.

    else:
        # This case should theoretically not be reached if date/time parsing is successful
        messagebox.showerror("Input Error", "Failed to determine target date or time.")


def remove_selected_event():
    global tracked_events, remove_button, event_tree
    if not event_tree or not remove_button: return # Widgets not ready

    try:
        if not event_tree.winfo_exists(): return
        selected_tree_ids = event_tree.selection() # Get tuple of selected item IDs
        if not selected_tree_ids:
            update_status("No events selected to remove.", clear_after=True)
            return
    except tk.TclError: return # Treeview likely destroyed

    # Confirmation dialog
    num_selected = len(selected_tree_ids)
    confirm_msg = f"Remove {num_selected} selected event(s)?" if num_selected > 1 else f"Remove selected event?"
    if not messagebox.askyesno("Confirm Removal", confirm_msg):
        return

    removed_count = 0
    tree_ids_to_remove = set(selected_tree_ids) # Use set for efficient lookup
    new_tracked_events = [] # Build a new list of events to keep

    # Iterate through the *current* tracked_events list
    for event in tracked_events:
        tree_id = event.get('tree_id')
        if tree_id in tree_ids_to_remove:
            # Event is selected for removal
            removed_count += 1
            # Try removing from Treeview widget if it still exists
            try:
                if event_tree.exists(tree_id):
                    event_tree.delete(tree_id)
            except tk.TclError: pass # Ignore error if item already gone from treeview
            # Do *not* add this event to new_tracked_events
        else:
            # Event is not selected, keep it
            new_tracked_events.append(event)

    if removed_count > 0:
        tracked_events = new_tracked_events # Update the main list
        update_status(f"Removed {removed_count} event(s).")
        update_calendar_markers() # Update calendar dots
        on_tab_changed(None) # Trigger week view update if visible
        update_remove_button_state() # Disable remove button if selection is now empty
    else:
        # This might happen if the selected item was somehow removed between selection and button press
        update_status("Could not remove selected event(s) (already gone?).")


def update_remove_button_state(event=None): # Accept event for binding
    global remove_button, event_tree;
    if remove_button and event_tree:
         try:
             if not remove_button.winfo_exists() or not event_tree.winfo_exists(): return
             # Enable button only if there's a selection in the treeview
             state = tk.NORMAL if event_tree.selection() else tk.DISABLED
             remove_button.config(state=state)
         except tk.TclError:
              # Handle case where treeview or button is destroyed during check
              try:
                  if remove_button.winfo_exists():
                      remove_button.config(state=tk.DISABLED)
              except: pass # Ignore if button is also gone


def on_sort_change(event=None): # Accept event for binding
    global sort_var
    if sort_var:
        update_status(f"Sorting by {sort_var.get()}...");
    sort_and_redisplay(); # This handles the actual sorting and display update
    update_status("Sort complete.", clear_after=True) # Give feedback after sorting


def scroll_calendar_month(event):
    """ Handles mouse wheel scrolling over the main calendar widget """
    global calendar_widget
    if not tkcalendar_AVAILABLE or not calendar_widget: return

    try:
        if not calendar_widget.winfo_exists(): return # Check if widget exists

        # Determine scroll direction (platform differences)
        if event.num == 4 or event.delta > 0: # Linux scroll up / Windows scroll up
            calendar_widget.prev_month()
        elif event.num == 5 or event.delta < 0: # Linux scroll down / Windows scroll down
            calendar_widget.next_month()

    except tk.TclError:
        # print("Warning: TclError during calendar scroll (widget likely destroyed)") # Optional debug
        pass
    except Exception as e:
        # Catch any other unexpected errors during scrolling
        print(f"Error during calendar scroll: {e}")


def show_events_for_selected_date(event=None): # Accept event argument
    """ Updates the label below the calendar with events for the selected date """
    global calendar_widget, tracked_events, selected_date_event_var, selected_date_event_label
    if not tkcalendar_AVAILABLE or not calendar_widget or not selected_date_event_var or not selected_date_event_label:
        return # Components not ready

    try:
        if not calendar_widget.winfo_exists() or not selected_date_event_label.winfo_exists():
            return # Widget destroyed
        # Get the selected date as a date object
        selected_date_obj = calendar_widget.selection_get()
        if selected_date_obj is None:
             # No date explicitly selected, maybe use current date or clear label
             try:
                 selected_date_event_var.set("Select a date to view events.")
             except tk.TclError: pass # Label might be gone
             return # No date selected, nothing to show
        selected_date_str = selected_date_obj.strftime(DATE_FORMAT) # Format for display
    except AttributeError:
        print("Error: Could not get selected date (AttributeError).")
        try: selected_date_event_var.set("Error retrieving date.")
        except: pass
        return
    except tk.TclError: return # Widget destroyed
    except Exception as e: # Catch other potential errors
        print(f"Error getting selected date: {e}")
        try: selected_date_event_var.set(f"Error: {e}")
        except: pass
        return

    # Filter events for the selected date
    matching_events = [
        event for event in tracked_events
        if 'target_dt' in event and event['target_dt'].date() == selected_date_obj
    ]
    # Sort events by time for the selected day
    matching_events.sort(key=lambda ev: ev['target_dt'].time())

    # Format the display text
    if not matching_events:
        display_text = f"No events scheduled for {selected_date_str}."
    else:
        event_lines = [f"Events for {selected_date_str}:"]
        for ev in matching_events:
            time_str = ev['target_dt'].strftime(TIME_FORMAT)
            loc_str = f" ({ev.get('location', '')})" if ev.get('location') else "" # Add location if present
            event_lines.append(f"- {ev['label']} at {time_str}{loc_str}")
        display_text = "\n".join(event_lines)

    # Update the label text, ensuring label widget exists
    try:
        if selected_date_event_label.winfo_exists():
            selected_date_event_var.set(display_text)
            # Dynamically adjust wraplength based on current label width (safer check)
            selected_date_event_label.update_idletasks() # Ensure width is calculated
            current_width = selected_date_event_label.winfo_width()
            if current_width > 10: # Ensure width is sensible before setting wraplength
                wraplen = max(200, current_width - 10) # Min width 200
                selected_date_event_label.config(wraplength=wraplen)
    except tk.TclError: pass # Widget likely destroyed
    except Exception as e:
        print(f"Error updating selected date label: {e}")


def on_closing():
    global root, status_clear_job, update_job_id
    print("Closing application...")
    # Cancel scheduled tasks first
    if update_job_id and root:
        try: root.after_cancel(update_job_id); print("Cancelled update loop.")
        except Exception as e: print(f"Error cancelling update: {e}");
    update_job_id = None
    if status_clear_job and root:
        try: root.after_cancel(status_clear_job); print("Cancelled status clear.")
        except Exception as e: print(f"Error cancelling status: {e}");
    status_clear_job = None

    # Save data
    save_data()

    # Destroy window
    print("Destroying main window...")
    try:
        if root and root.winfo_exists():
            root.destroy()
    except tk.TclError as e:
        # This can happen if the window is already being destroyed
        print(f"Error destroying window (may already be closing): {e}");
    root = None # Ensure root is cleared

# --- Create Main Window ---
root = tk.Tk(); root.title("Event Time Tracker"); root.geometry("950x700")
root.minsize(600, 450) # Set a minimum size (increased height slightly for week view)

# --- Style ---
style = ttk.Style(root)
# Try to use a modern theme if available, fallback to clam
try:
    available_themes = style.theme_names()
    if 'vista' in available_themes: style.theme_use('vista')
    elif 'aqua' in available_themes: style.theme_use('aqua') # macOS
    elif 'winxpnative' in available_themes: style.theme_use('winxpnative') # Older Windows
    else: style.theme_use('clam') # Good fallback
except tk.TclError:
     style.theme_use('clam') # Fallback

# Initial config (overridden by apply_styles later)
style.configure("TLabel", padding=3); style.configure("TButton", padding=5); style.configure("TEntry", padding=3)
style.configure("Treeview.Heading", font=('Helvetica', 10, 'bold')); style.configure("Treeview", rowheight=25);

# --- Input Frame (Stays at Top) ---
input_frame = ttk.LabelFrame(root, text="Add Custom Event", padding=(10, 5, 10, 10));
input_frame.pack(pady=(10,5), padx=10, fill=tk.X, side=tk.TOP, anchor='n'); # Ensure it stays top
input_frame.columnconfigure(1, weight=1) # Allow label/location entry to expand

# Row 0: Label
ttk.Label(input_frame, text="Label:").grid(row=0, column=0, padx=5, pady=6, sticky=tk.W);
label_entry_var = tk.StringVar(root);
label_entry = ttk.Entry(input_frame, textvariable=label_entry_var, width=40);
label_entry.grid(row=0, column=1, columnspan=3, padx=5, pady=6, sticky=tk.EW);
label_entry.bind("<Return>", add_custom_event) # Bind Enter key

# Row 1: Location
ttk.Label(input_frame, text="Location:").grid(row=1, column=0, padx=5, pady=6, sticky=tk.W);
location_entry_var = tk.StringVar(root);
location_entry = ttk.Entry(input_frame, textvariable=location_entry_var, width=40);
location_entry.grid(row=1, column=1, columnspan=3, padx=5, pady=6, sticky=tk.EW);
location_entry.bind("<Return>", add_custom_event) # Bind Enter key

# Row 2: Date
ttk.Label(input_frame, text="Date:").grid(row=2, column=0, padx=5, pady=6, sticky=tk.W);
date_entry_var = tk.StringVar(root)
date_format_helper_text = ""
if tkcalendar_AVAILABLE:
    # Common patterns: 'm/d/yy', 'mm/dd/yyyy', 'd-M-Y', 'dd-Mon-YYYY', etc.
    # Let's use a locale-aware default if possible, otherwise a common one.
    try:
        # Use locale's short date format if possible
        import locale
        locale.setlocale(locale.LC_TIME, '') # Use system locale
        tkcalendar_pattern = locale.nl_langinfo(locale.D_FMT).replace('%','').lower()
        # Basic conversion for common strftime codes to tkcalendar codes
        tkcalendar_pattern = tkcalendar_pattern.replace('y', 'yy', 1) # %y -> yy (first occurrence only)
        tkcalendar_pattern = tkcalendar_pattern.replace('yy', 'yyyy') # Ensure %Y -> yyyy
        # Handle potential leading zeros if needed based on locale (tricky)
        # For simplicity, stick to mm/dd/yyyy as fallback
        if not ('d' in tkcalendar_pattern and 'm' in tkcalendar_pattern and 'y' in tkcalendar_pattern):
            tkcalendar_pattern = 'mm/dd/yyyy' # Fallback
    except Exception:
        tkcalendar_pattern = 'mm/dd/yyyy' # Default fallback

    date_input_widget = DateEntry(input_frame, textvariable=date_entry_var,
                                  date_pattern=tkcalendar_pattern, width=18, borderwidth=2,
                                  state="readonly", # Forces calendar use
                                  )
    date_input_widget.set_date(date.today()) # Set initial value
    date_format_helper_text = tkcalendar_pattern.upper() # e.g., MM/DD/YYYY
else:
    # Fallback to standard Entry if tkcalendar not installed
    date_input_widget = ttk.Entry(input_frame, textvariable=date_entry_var, width=20)
    default_date_str = date.today().strftime(DATE_FORMAT); date_entry_var.set(default_date_str)
    date_format_helper_text = DATE_FORMAT # e.g., Month DD, YYYY
date_input_widget.grid(row=2, column=1, padx=5, pady=6, sticky=tk.W)
# Add a subtle label showing the expected format
date_hint_label = ttk.Label(input_frame, text=f"({date_format_helper_text})", foreground="grey")
date_hint_label.grid(row=2, column=2, padx=5, pady=6, sticky=tk.W)

# Row 3: Time
ttk.Label(input_frame, text="Time:").grid(row=3, column=0, padx=5, pady=6, sticky=tk.W);
time_entry_var = tk.StringVar(root, value=DEFAULT_TIME);
time_entry = ttk.Entry(input_frame, textvariable=time_entry_var, width=10);
time_entry.grid(row=3, column=1, padx=5, pady=6, sticky=tk.W);
time_hint_label = ttk.Label(input_frame, text="(HH:MM:SS)", foreground="grey")
time_hint_label.grid(row=3, column=2, padx=5, pady=6, sticky=tk.W)
time_entry.bind("<Return>", add_custom_event) # Bind Enter key

# Row 4: Add Button (Centered)
add_button_frame = ttk.Frame(input_frame); # Use a frame to center the button
add_button_frame.grid(row=4, column=0, columnspan=4, pady=(10, 5));
add_button = ttk.Button(add_button_frame, text="Add Custom Event", command=add_custom_event);
add_button.pack() # Pack inside the frame to center it

# --- Notebook for Tabs ---
notebook = ttk.Notebook(root, padding=(0, 5, 0, 0));
notebook.pack(pady=(0,5), padx=10, fill=tk.BOTH, expand=True, side=tk.TOP, anchor='n') # Fill remaining top space
# Bind tab change event AFTER notebook is created
notebook.bind("<<NotebookTabChanged>>", on_tab_changed)


# --- Tab 1: Event List ---
list_tab_frame = ttk.Frame(notebook, padding=10);
list_tab_frame.pack(fill=tk.BOTH, expand=True) # Ensure frame fills tab space
notebook.add(list_tab_frame, text=' Event List ')

# Sort controls within list tab
sort_control_frame = ttk.Frame(list_tab_frame);
sort_control_frame.pack(fill=tk.X, pady=(0, 8));
ttk.Label(sort_control_frame, text="Sort by:").pack(side=tk.LEFT, padx=(0, 5));
sort_var = tk.StringVar(root, value=DEFAULT_SORT);
sort_combo = ttk.Combobox(sort_control_frame, textvariable=sort_var, values=SORT_OPTIONS, state="readonly", width=18);
sort_combo.pack(side=tk.LEFT);
sort_combo.bind("<<ComboboxSelected>>", on_sort_change) # Update on selection change

# Treeview container (allows scrollbar placement)
tree_container = ttk.Frame(list_tab_frame);
tree_container.pack(fill=tk.BOTH, expand=True)

columns = ('indicator', 'label', 'location', 'target_dt_str', 'diff');
event_tree = ttk.Treeview(tree_container, columns=columns, show='headings', selectmode='extended')

# Define headings
event_tree.heading('indicator', text='', anchor=tk.CENTER);
event_tree.heading('label', text='Label', anchor=tk.W);
event_tree.heading('location', text='Location', anchor=tk.W);
event_tree.heading('target_dt_str', text='Next Occurrence', anchor=tk.CENTER);
event_tree.heading('diff', text='Time Difference', anchor=tk.W)

# Define column properties
event_tree.column('indicator', width=40, minwidth=30, anchor=tk.CENTER, stretch=tk.NO);
event_tree.column('label', width=200, minwidth=150, anchor=tk.W, stretch=tk.YES); # Allow stretch
event_tree.column('location', width=150, minwidth=100, anchor=tk.W, stretch=tk.YES); # Allow stretch
event_tree.column('target_dt_str', width=180, minwidth=160, anchor=tk.CENTER, stretch=tk.NO);
event_tree.column('diff', width=180, minwidth=120, anchor=tk.W, stretch=tk.NO)

# Scrollbar
scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=event_tree.yview, style='Vertical.TScrollbar');
event_tree.configure(yscrollcommand=scrollbar.set);

# Grid layout for treeview and scrollbar
tree_container.grid_columnconfigure(0, weight=1); # Treeview expands horizontally
tree_container.grid_rowconfigure(0, weight=1); # Treeview expands vertically
event_tree.grid(row=0, column=0, sticky='nsew');
scrollbar.grid(row=0, column=1, sticky='ns');

# Bind selection change to update button state
event_tree.bind('<<TreeviewSelect>>', update_remove_button_state)

# --- Tab 2: Calendar View ---
calendar_tab_frame = ttk.Frame(notebook, padding=10);
calendar_tab_frame.pack(fill=tk.BOTH, expand=True)
notebook.add(calendar_tab_frame, text=' Calendar View ')

if tkcalendar_AVAILABLE:
    # Use the same pattern determined for DateEntry for consistency
    calendar_pattern = tkcalendar_pattern if 'tkcalendar_pattern' in globals() else 'mm/dd/yyyy'
    calendar_widget = Calendar(calendar_tab_frame, selectmode='day',
                               year=datetime.now().year, month=datetime.now().month, day=datetime.now().day,
                               showweeknumbers=False, date_pattern=calendar_pattern,
                               showothermonthdays=False, # Cleaner look
                               # Styles are applied in apply_styles
                               )
    calendar_widget.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    # Event marker tag is configured during apply_styles / update_calendar_markers

    # Bind events
    calendar_widget.bind("<<CalendarSelected>>", show_events_for_selected_date)
    calendar_widget.bind("<MouseWheel>", scroll_calendar_month) # Windows/macOS
    calendar_widget.bind("<Button-4>", scroll_calendar_month) # Linux scroll up
    calendar_widget.bind("<Button-5>", scroll_calendar_month) # Linux scroll down

    # Label to show events for the selected date
    selected_date_event_var = tk.StringVar(root, value="Click a date on the calendar to see its events.")
    # Initial wraplength calculation moved to show_events_for_selected_date for dynamic update
    selected_date_event_label = ttk.Label(calendar_tab_frame,
                                         textvariable=selected_date_event_var,
                                         wraplength=300, # Initial placeholder, will be adjusted
                                         padding=(5, 5), relief=tk.GROOVE, borderwidth=1, # Keep relief for visibility
                                         justify=tk.LEFT, anchor='nw')
    selected_date_event_label.pack(fill=tk.X, pady=(10, 0), padx=5, side=tk.BOTTOM)

else: # tkcalendar not available
    ttk.Label(calendar_tab_frame,
              text="Calendar view requires 'tkcalendar'.\nInstall using pip:\npip install tkcalendar",
              wraplength=300, justify=tk.CENTER, style='TLabel' # Use default style
             ).pack(pady=50, padx=20, expand=True, fill=tk.BOTH)

# --- Tab 3: Week View ---
week_view_tab_frame = ttk.Frame(notebook, padding=10)
week_view_tab_frame.pack(fill=tk.BOTH, expand=True)
notebook.add(week_view_tab_frame, text=' Week View ')

# Navigation controls for Week View
week_nav_frame = ttk.Frame(week_view_tab_frame)
week_nav_frame.pack(fill=tk.X, pady=(0, 5))

prev_week_btn = ttk.Button(week_nav_frame, text="◀ Prev Week", command=show_previous_week, width=12)
prev_week_btn.pack(side=tk.LEFT, padx=(0, 5))

today_week_btn = ttk.Button(week_nav_frame, text="Today", command=show_current_week, width=8)
today_week_btn.pack(side=tk.LEFT, padx=5)

next_week_btn = ttk.Button(week_nav_frame, text="Next Week ▶", command=show_next_week, width=12)
next_week_btn.pack(side=tk.LEFT, padx=(5, 10))

week_view_label_var = tk.StringVar(root, value="Week: Loading...")
week_label = ttk.Label(week_nav_frame, textvariable=week_view_label_var, anchor=tk.W)
week_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

# Text area for displaying week events
week_text_frame = ttk.Frame(week_view_tab_frame) # Frame for text + scrollbar
week_text_frame.pack(fill=tk.BOTH, expand=True)

week_view_text = tk.Text(week_text_frame, wrap=tk.WORD, height=15, width=70,
                         padx=5, pady=5, state=tk.DISABLED) # Start disabled (read-only)
                         # Font/colors set by apply_styles

week_scrollbar_y = ttk.Scrollbar(week_text_frame, orient=tk.VERTICAL, command=week_view_text.yview, style='Vertical.TScrollbar')
week_view_text.configure(yscrollcommand=week_scrollbar_y.set)

# Use grid within the text_frame for layout flexibility
week_text_frame.grid_rowconfigure(0, weight=1)
week_text_frame.grid_columnconfigure(0, weight=1)
week_view_text.grid(row=0, column=0, sticky="nsew")
week_scrollbar_y.grid(row=0, column=1, sticky="ns")

# Define tags used in the text widget (colors/fonts applied in apply_styles)
base_font = tkFont.nametofont("TkDefaultFont") # Get default font for sizing
week_view_text.tag_configure("bold_date", font=(base_font.actual()['family'], base_font.actual()['size'], 'bold'))
week_view_text.tag_configure("event_item", lmargin1=15, lmargin2=15) # Indent event lines


# --- Bottom Controls (Below Notebook) ---
bottom_frame = ttk.Frame(root, padding=(10, 5, 10, 10));
bottom_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(5, 0)) # Pack at bottom, fill horizontally

# Status Label (Expands to fill space)
status_label = ttk.Label(bottom_frame, text="Initializing...", anchor=tk.W, relief=tk.SUNKEN, padding=3); # Keep relief
status_label.pack(side=tk.LEFT, fill=tk.X, expand=True) # Expand status bar

# Dark Mode Checkbox (Right side, but before remove button)
dark_mode_var = tk.BooleanVar(root) # Value set during initialization
dark_mode_check = ttk.Checkbutton(bottom_frame, text="Dark Mode", variable=dark_mode_var, command=toggle_dark_mode)
dark_mode_check.pack(side=tk.LEFT, padx=(10, 5)) # Add padding

# Remove Button (Rightmost)
remove_button = ttk.Button(bottom_frame, text="Remove Selected", command=remove_selected_event, state=tk.DISABLED);
remove_button.pack(side=tk.RIGHT, padx=(5, 0)) # Pack to the right


# --- Initial Setup & Start ---
def initialize_app():
    global dark_mode_var, settings, tracked_events, current_week_start_date # Added week start date
    update_status("Loading data...")
    loaded_custom_events = load_data() # Loads settings AND events now

    # Set dark mode checkbox based on loaded settings *before* applying styles
    if dark_mode_var:
        try:
            dark_mode_var.set(settings.get("dark_mode", False))
        except tk.TclError: print("Error: dark_mode_var (BooleanVar) not ready during init.")
    else: print("Error: dark_mode_var (widget reference) not ready during init.")

    # Apply initial style based on loaded settings
    apply_styles()

    # Add loaded custom events first
    tracked_events.extend(loaded_custom_events)

    # Add built-in events (avoiding duplicates with loaded custom ones)
    update_status("Adding built-in events...")
    loaded_custom_labels = {ev.get('label','').lower() for ev in tracked_events if ev.get('is_custom', False)};
    built_in_added = 0
    for label, (month, day) in INITIAL_EVENTS_DATA.items():
        # Only add if a custom event with the same label doesn't exist
        if label.lower() not in loaded_custom_labels:
            next_date = calculate_next_occurrence(month, day)
            if next_date:
                # Add built-in event with default time 00:00:00
                add_event_to_tracker(label, datetime.combine(next_date, datetime.min.time()),
                                     location=None, is_custom=False)
                built_in_added += 1

    if built_in_added > 0: update_status(f"Added {built_in_added} built-in events.")
    else: update_status("No new built-in events added (might exist as custom).") # More informative status

    # Set initial week view date
    current_week_start_date = get_week_start(date.today())

    # Initial sort and display (calls update_calendar_markers and potentially week view update via on_tab_changed)
    update_status(f"Performing initial sort ({DEFAULT_SORT})...");
    sort_and_redisplay() # This now assigns tree_id to events

    # Call show_events_for_selected_date once initially for calendar label
    if tkcalendar_AVAILABLE:
        show_events_for_selected_date()

    # Start the periodic update loop for time differences in the list view
    update_display();

    update_status("Ready.") # Final status

# Schedule initialization after the main loop starts
root.after(100, initialize_app)
root.protocol("WM_DELETE_WINDOW", on_closing);
root.mainloop()