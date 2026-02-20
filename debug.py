import time

def safe_type(window, keystrokes, wait_time=0.2):
    """
    Checks focus before sending ANY keystrokes. 
    If focus is lost, it tries to restore it before typing.
    """
    try:
        # 1. Check if the window currently has focus
        if not window.has_focus():
            print(f"⚠️ Focus lost! Restoring focus before typing: {keystrokes}")
            window.set_focus()
            time.sleep(0.3) # Give Windows time to bring it front
            
            # Double-check if we actually got focus back
            if not window.has_focus():
                raise RuntimeError("OS blocked focus transfer.")

        # 2. Safe to type!
        window.type_keys(keystrokes)
        
        # 3. Optional standard pause after typing
        if wait_time > 0:
            time.sleep(wait_time)
            
        return True

    except Exception as e:
        # Stop the script or handle the error so it DOESN'T type in another app
        print(f"❌ CRITICAL: Could not secure window focus to type '{keystrokes}'. Error: {e}")
        raise e # We raise the error to immediately stop the bot from wreaking havoc



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
