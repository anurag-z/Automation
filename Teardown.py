@pytest.fixture(scope="session")
def fads_window():
    """Launches the app once, and cleanly closes it after all tests are done."""
    # ==========================================
    # ⬆️ SESSION SETUP (Runs once at the beginning)
    # ==========================================
    app = launch_fads() 
    window = Desktop(backend="uia").window(title_re=".*FADS PRIME.*")
    window.wait('ready', timeout=10)
    
    # Pass the window to your 100+ tests
    yield window 
    
    # ==========================================
    # ⬇️ SESSION TEARDOWN (Runs once at the very end)
    # ==========================================
    print("\n[🏁 Bulk Execution Complete] Closing FADS and Command Prompt...")
    
    try:
        # Option A: Graceful Exit (Highly Recommended for DOS apps)
        # Ensure we are at the main menu first
        for _ in range(5):
            safe_type(window, "{ESC}", wait_time=0.2)
        
        # Press the specific key to Exit FADS (e.g., F6, or ESC and Y)
        # UPDATE THESE KEYS to match exactly how you manually quit FADS!
        safe_type(window, "{ESC}Y", wait_time=1.0) 
        
        # If quitting FADS drops you into a standard "C:\>" DOS prompt, 
        # type 'exit' to close the black CMD window completely.
        safe_type(window, "exit{ENTER}", wait_time=0.5)
        
    except Exception as e:
        print(f"⚠️ Graceful exit failed: {e}. Forcing application to close...")
        
    finally:
        # Option B: The "Nuke" Option
        # Just in case the keystrokes failed, we force-kill the process
        try:
            app.kill()
            print("✅ FADS process killed.")
        except:
            pass
