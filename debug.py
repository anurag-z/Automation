import time

import time
from pywinauto import win32functions

def safe_type(window, keystrokes, wait_time=0.2):
    """
    Checks if FADS is the active Windows application before typing. 
    If you clicked away, it forces FADS back to the front first!
    """
    try:
        # 1. Get the real, underlying window element
        real_window = window.wrapper_object()
        
        # 2. Check if the active window in Windows matches our FADS window
        if win32functions.GetForegroundWindow() != real_window.handle:
            print(f"⚠️ Focus lost! Forcing FADS back to the front before typing '{keystrokes}'...")
            real_window.set_focus()
            time.sleep(0.3) # Give Windows a split second to pull it forward
            
        # 3. Safe to type!
        window.type_keys(keystrokes)
        
        # 4. Optional standard pause after typing
        if wait_time > 0:
            time.sleep(wait_time)
            
    except Exception as e:
        print(f"❌ CRITICAL: Could not type '{keystrokes}'. Error: {e}")
        raise e


def navigate_field(area, fieldname, row, column, window):
    try:
        print(f"\n[+] Processing: {area} -> {fieldname}")
        
        # Bring it up initially
        window.set_focus()
        window.maximize()
        
        # --- Using our new safe_type method ---
        safe_type(window, "{F3}{ESC}")
        
        for key in "MAS":
            safe_type(window, key, wait_time=0.2)
            
        safe_type(window, area + "{ENTER}", wait_time=0.5)
        
        safe_type(window, "{ESC}{ESC}")
        
        for key in "FFS":
            safe_type(window, key, wait_time=0.2)
            
        safe_type(window, fieldname + "{ENTER}", wait_time=0.8)
        
        safe_type(window, "{F9}AA", wait_time=1.2)
        
        # Extract the data using your separated logic from earlier
        field_values_list = get_field_values(window)
        
        # Back out safely
        safe_type(window, "{ESC}")
        
    except Exception as e:
        print(f"[navigate_field] Error navigating to {area} -> {fieldname}: {e}")
