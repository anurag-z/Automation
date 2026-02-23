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





from PIL import ImageGrab, Image, ImageOps

def read_bottom_status_bar(window):
    """
    Specifically grabs the bottom 40 pixels of the original FADS window.
    Optimized for black text on a blue background.
    """
    try:
        rect = window.wrapper_object().rectangle()
        # Capture the whole window
        img = ImageGrab.grab(bbox=(rect.left, rect.top, rect.right, rect.bottom))
        width, height = img.size
        
        # 1. Crop ONLY the bottom 40 pixels (the blue bar)
        crop_h = 40 
        cropped_img = img.crop((0, height - crop_h, width, height))
        
        # 2. Resize 3x to give Tesseract more pixels to work with
        img_resized = cropped_img.resize((width * 3, crop_h * 3), resample=Image.Resampling.LANCZOS)
        
        # 3. Convert to Grayscale
        gray = img_resized.convert('L')
        
        # 4. Smart Thresholding: 
        # Black text is very dark (close to 0). Blue background is lighter (~100-150).
        # We force anything darker than 80 to become pure black (0), and everything else to pure white (255).
        bw_img = gray.point(lambda x: 0 if x < 80 else 255, '1')
        
        # 5. Save the debug image so you can physically verify the crop and color
        ensure_dir_exists(FADS_DEBUG_DIR)
        bw_img.save(os.path.join(FADS_DEBUG_DIR, "debug_status_bar.png"))
        
        # 6. OCR using PSM 7 (Single Line Mode)
        custom_config = r'--psm 7 -c preserve_interword_spaces=1'
        text = pytesseract.image_to_string(bw_img, config=custom_config)
        
        # Clean up text
        clean_text = text.strip().replace("‘", "").replace("'", "")
        return clean_text

    except Exception as e:
        print(f"❌ [read_bottom_status_bar] Failed: {e}")
        return ""
