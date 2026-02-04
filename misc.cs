// ... (Inside Main, after capturing 'original') ...

            // 3. PRE-PROCESS IMAGE (The "Smart" Way)
            Console.WriteLine("Optimizing image for OCR...");
            
            // A. Scale 2x (Not 3x, keeps it manageable)
            int scale = 2;
            Bitmap processed = new Bitmap(original.Width * scale, original.Height * scale);

            using (Graphics g = Graphics.FromImage(processed))
            {
                // CRITICAL CHANGE: Use 'NearestNeighbor' to keep DOS text crisp, not blurry
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Half;
                g.DrawImage(original, 0, 0, processed.Width, processed.Height);
            }

            // B. Invert Colors (Make it Black Text on White Background)
            // This loop is slow but simple and requires no extra libraries
            for (int y = 0; y < processed.Height; y++)
            {
                for (int x = 0; x < processed.Width; x++)
                {
                    Color pixel = processed.GetPixel(x, y);
                    
                    // Simple Invert: 255 - CurrentValue
                    Color inverted = Color.FromArgb(255 - pixel.R, 255 - pixel.G, 255 - pixel.B);
                    processed.SetPixel(x, y, inverted);
                }
            }

            // Save it to check (It should look like a printed document now)
            processed.Save("debug_ready_for_ocr.png", ImageFormat.Png);

            // 4. READ TEXT
            Console.WriteLine("Reading text...");
            
            using (var engine = new TesseractEngine(@"./tessdata", "eng", EngineMode.Default))
            {
                // Use Option 2 (Temp File) since it is the most reliable for you
                string tempFile = "temp_ocr.tif";
                processed.Save(tempFile, ImageFormat.Tiff);

                using (var img = Pix.LoadFromFile(tempFile)) 
                {
                    // Use 'SingleBlock' mode if 'SparseText' missed things
                    using (var page = engine.Process(img, PageSegMode.SingleBlock))
                    {
                        string text = page.GetText();
                        Console.WriteLine("--- FOUND TEXT ---");
                        Console.WriteLine(text);
                        Console.WriteLine("------------------");
                    }
                }
            }
