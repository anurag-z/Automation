// ... inside your Tesseract block ...
using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
{
    // 1. Save to a temporary file
    string tempFile = "temp_ocr.tif";
    try 
    {
        enlarged.Save(tempFile, System.Drawing.Imaging.ImageFormat.Tiff);

        // 2. Load from File (This method ALWAYS exists)
        using (var img = Pix.LoadFromFile(tempFile)) 
        {
            using (var page = engine.Process(img, PageSegMode.SparseText))
            {
                string text = page.GetText();
                Console.WriteLine("Found: " + text);
                
                if (text.Contains("F3")) Console.WriteLine("[PASS] F3 Found");
            }
        }
    }
    finally 
    {
        // 3. Cleanup: Delete the file so we don't leave junk
        if (System.IO.File.Exists(tempFile))
            System.IO.File.Delete(tempFile);
    }
}
