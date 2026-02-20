import time
import win32gui

def safe_type(window, keystrokes, wait_time=0.2):
    """
    Acts as a Pause/Play button. If you click into another app, 
    the script pauses and waits until FADS is the active window again before typing.
    """
    try:
        real_window = window.wrapper_object()
        fads_handle = real_window.handle
        
        # --- THE FOCUS LOCK LOOP ---
        attempts = 0
        while win32gui.GetForegroundWindow() != fads_handle:
            print(f"⚠️ Bot Paused: Waiting for FADS to be the active window to type '{keystrokes}'...")
            
            # Try to politely ask Windows to bring it to the front
            try:
                real_window.set_focus()
            except:
                pass
                
            time.sleep(1.0) # Wait 1 full second, then check again
            
            attempts += 1
            if attempts > 15: # If it's blocked for 15 seconds, give up to prevent infinite loops
                raise Exception("Focus timeout. Windows refused to bring FADS to the front.")

        # The absolute split-second the loop confirms FADS is in front, we fire the keys!
        window.type_keys(keystrokes)
        
        # Optional pause after typing
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


import time
import win32gui

def safe_get_field_values(window):
    """
    Waits until FADS is the active foreground window, then safely captures 
    the screen and extracts the OCR data.
    """
    try:
        real_window = window.wrapper_object()
        fads_handle = real_window.handle
        
        # --- FOCUS LOCK FOR SCREENSHOT ---
        attempts = 0
        while win32gui.GetForegroundWindow() != fads_handle:
            print("📸 OCR Paused: Waiting for FADS to be fully visible before taking screenshot...")
            
            try:
                real_window.set_focus()
            except:
                pass
                
            time.sleep(1.0) # Check again in 1 second
            
            attempts += 1
            if attempts > 15:
                raise Exception("Focus timeout. Windows blocked FADS from coming to the front.")

        # ---> CRITICAL PAUSE <---
        # Windows has a brief animation when a window comes to the front. 
        # If we screenshot instantly, the text might be blurry or half-transparent.
        # We wait half a second for the UI to completely settle.
        time.sleep(0.5) 

        # ---------------------------------------------------------
        # THE COAST IS CLEAR - TAKE THE PICTURE!
        # ---------------------------------------------------------
        print("Taking safe screenshot...")
        
        # Call your existing capture method here
        screen_data = capture_and_read(window) 
        
        if not screen_data or screen_data.strip() == "":
            print("❌ No data captured from screen.")
            return []
            
        # Parse out the last line exactly as you had it
        lines = screen_data.strip().split('\n')
        last_line = lines[-1]
        
        # Split into your list
        ls = last_line.split()
        for i in ls:
            print("Values", i)
            
        return ls

    except Exception as e:
        print(f"❌ Error extracting field values safely: {e}")
        return []
