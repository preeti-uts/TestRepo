using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using AutoIt;

namespace selenium_framework
{
    public static class PDF_Framework
    {
        private static double pdf_error_rate = 0.001; //  0.1%

        //
        // PDF_Framework - constructor
        // Params: pdf file, output folder, page count
        // Will only capture the number of page images up to the count
        //
        // Start acrobat reader the first few times and after install and answer pop-ups
        // Set view size to ~100% and close tool menu
        //

        public static bool pdfSample(string file, string output, int pages)
        {
            spawn_app(file);

            acrobat_activate();

            // The whole document goes as one txt file
            acrobat_copytext();
            if (!process_PDFText(output, file))
                return false;

            // Change to image mode
            acrobat_snapshot();

            for (int i = 1; i <= pages; i++)
            {
                acrobat_copyimage();
                if (!process_PDFImage(output, file, i))
                    return false;
            }

            acrobat_close();
            return true;
        }

        //
        // process_PDFImages
        // Writes output as filename(p) where p is the page number
        //
        private static bool process_PDFImage(string output, string file, int page)
        {
            Image snapshot;

            // Create empty txt file
            string imagefile = Path.GetFileNameWithoutExtension(file) + "(" + page.ToString() + ").png";
            int i = 0;

            while (i < 3)
            {
                if (Clipboard.ContainsImage())
                {
                    snapshot = Clipboard.GetImage();
                    snapshot.Save(Path.Combine(output, imagefile));
                    Console.WriteLine("PDF_Framework: Page " + page + " clipboard image saved");
                    return true;
                }
                else
                {
                    i++;
                    sleep(2);
                }
            }

            Console.WriteLine("PDF_Framework: Clipboard has NO image");
            return false;

        } // process_PDFImage

        //
        // process_PDFText
        // Writes text file to output folder
        //
        private static bool process_PDFText(string output, string file)
        {
            // Create empty txt file
            string textfile = Path.GetFileNameWithoutExtension(file) + ".txt";
            int i = 0;

            // Write clipboard to text file
            while (i < 3)
            {
                if (Clipboard.ContainsText())
                {
                    var text = Clipboard.GetText();
                    File.WriteAllText(Path.Combine(output, textfile), text);
                    Console.WriteLine("PDF_Framework: Clipboard text saved");
                    return true;
                }
                else
                {
                    i++;
                    sleep(2);
                }
            }

            Console.WriteLine("PDF_Framework: Clipboard has NO text");
            return false;

        } // process_PDFText

        //
        // spawn_app
        //
        private static void spawn_app(string file)
        {
            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;

            startInfo.FileName = file;
            process.StartInfo = startInfo;
            process.Start();

        } // spawn_app

        //
        // maskedText
        //
        public static bool maskedText(string newfile, string basefile)
        {
            StreamReader nfile = new StreamReader(newfile);
            StreamReader bfile = new StreamReader(basefile);

            string nline = "";
            string bline = "";

            int lc = 0;

            //
            // Take each base line
            // If its !! then interpret as regex
            // Read new line and do compare, regex or not
            //

            Console.WriteLine("PDF_Framework: Begin masked text");

            while ((bline = bfile.ReadLine()) != null)
            {
                lc++;

                // Read new line
                if ((nline = nfile.ReadLine()) == null)
                {
                    // Must have the same number of line
                    Console.WriteLine("PDF_Framework: New file - read fails");
                    return false;
                }

                if (bline.StartsWith("!!"))
                {
                    // regex compare
                    Regex regex = new Regex(bline.Substring(2));
                    Match match = regex.Match(nline);
                    if (match.Success)
                    {
                        continue;
                    }
                    else
                    {
                        Console.WriteLine(lc + ":" + bline);
                        Console.WriteLine(lc + "?" + nline);
                        Console.WriteLine("PDF_Framework: Masked text error");
                        return false;
                    }
                }
                else
                {
                    // std compare
                    if (bline.Equals(nline))
                    {
                        continue;
                    }
                    else
                    {
                        Console.WriteLine(lc + ":" + bline);
                        Console.WriteLine(lc + "?" + nline);
                        Console.WriteLine("PDF_Framework: Masked text error");
                        return false;
                    }
                }
            }

            // Read new line one more time - it should fail
            if ((nline = nfile.ReadLine()) != null)
            {
                // Must have the same number of lines
                Console.WriteLine("PDF_Framework: New file - too many lines");
                return false;
            }

            Console.WriteLine("PDF_Framework: Text matched - " + lc + " lines read");

            nfile.Close();
            bfile.Close();

            return true;

        } // maskedText

        //
        // maskedImage
        //
        public static bool maskedImage(string newfile, string basefile)
        {
            int badpixel = 0;
            int maskpixel = 0;
            int bytes;

            Color RED = Color.FromArgb(255, 255, 0, 0);
            Color bcolor = Color.FromArgb(0, 0, 0, 0);
            Color ncolor = Color.FromArgb(0, 0, 0, 0);

            //
            // Load the basefile bitmap
            //

            Bitmap bimage = (Bitmap)Image.FromFile(basefile);

            //Get the bitmap data
            var bbitmapData = bimage.LockBits(
                new Rectangle(0, 0, bimage.Width, bimage.Height),
                ImageLockMode.ReadWrite,
                bimage.PixelFormat
            );

            //Initialize an array for all the image data
            byte[] bimageBytes = new byte[bbitmapData.Stride * bimage.Height];

            //Copy the bitmap data to the local array
            Marshal.Copy(bbitmapData.Scan0, bimageBytes, 0, bimageBytes.Length);

            //Unlock the bitmap
            bimage.UnlockBits(bbitmapData);

            //Find pixelsize
            int bpixelSize = Image.GetPixelFormatSize(bimage.PixelFormat);

            Console.WriteLine("PDF_Framework: Baseline image loaded: " + basefile);
            //debug(TRACE, "Width=" + bimage.Width + " Height=" + bimage.Height + " Stride=" + bbitmapData.Stride + " pixelsize=" + bpixelSize);

            //
            // Load the newfile bitmap
            //

            Bitmap nimage = (Bitmap)Image.FromFile(newfile);

            //Get the bitmap data
            var nbitmapData = nimage.LockBits(
                new Rectangle(0, 0, nimage.Width, nimage.Height),
                ImageLockMode.ReadWrite,
                nimage.PixelFormat
            );

            //Initialize an array for all the image data
            byte[] nimageBytes = new byte[nbitmapData.Stride * nimage.Height];

            //Copy the bitmap data to the local array
            Marshal.Copy(nbitmapData.Scan0, nimageBytes, 0, nimageBytes.Length);

            //Unlock the bitmap
            nimage.UnlockBits(nbitmapData);

            //Find pixelsize
            int npixelSize = Image.GetPixelFormatSize(nimage.PixelFormat);

            Console.WriteLine("PDF_Framework: Newfile image loaded: " + newfile);

            //
            // Compare the images
            //

            if (bimageBytes.Length != nimageBytes.Length)
            {
                Console.WriteLine("PDF_Framework: Mismatched image length");
                return false;
            }

            bytes = bpixelSize / 8;

            // Loop pixels
            for (int i = 0; i < bimageBytes.Length; i += bytes)
            {
                // Copy the bits into a local array
                var bpixelData = new byte[bytes];
                var npixelData = new byte[bytes];

                // Note: reverses byte order
                Array.Copy(bimageBytes, i, bpixelData, 0, bytes);
                Array.Copy(nimageBytes, i, npixelData, 0, bytes);

                // Test the color of the baseline pixel
                // PNG four (4) byte order is rgbA - ignore Alpha, automaticall set to 255

                bcolor = Color.FromArgb(bpixelData[2], bpixelData[1], bpixelData[0]);
                ncolor = Color.FromArgb(npixelData[2], npixelData[1], npixelData[0]);

                if (bcolor == RED)
                {
                    // Mask RED
                    maskpixel++;
                }
                else
                {
                    // Not mask RED
                    if (bcolor.Equals(ncolor))
                        continue;
                    else
                        badpixel++;
                }
            }

            // End of comparison
            // Error tollerance is 0.1%
            double tollerance = badpixel / (double)(bimageBytes.Length / (bpixelSize / 8));
            //debug(TRACE, "badpixel%: " + (tollerance * 100.0).ToString("F") + "%");
            //debug(TRACE, "badpixel=" + badpixel + " maskpixel=" + maskpixel);

            if (tollerance > pdf_error_rate)
            {
                Console.WriteLine("PDF_Framework: Too many bad pixels");
                return false;
            }

            Console.WriteLine("PDF_Framework: Image matched");
            return true;

        } // compare_maskedImage

        //
        // Adobe Reader Interface Drivers
        //

        private static string wClass = "[CLASS:AcrobatSDIWindow]";

        public static void acrobat_activate()
        {
            AutoItX.WinWaitActive(wClass);
            sleep(1);
        }

        public static void acrobat_copytext()
        {
            AutoItX.Send("^a");
            AutoItX.Send("^c");
            sleep(5);
        }

        public static void acrobat_copyimage()
        {
            AutoItX.Send("^a");
            AutoItX.Send("^c");
            sleep(2);
            AutoItX.MouseClick("left", 588, 83);    // down page control
            sleep(1);
        }

        public static void acrobat_snapshot()
        {
            AutoItX.MouseClick("left", 205, 55);
            AutoItX.Send("a");
            sleep(1);
            AutoItX.MouseClick("left", 588, 243);   // activate snapshot mode - click doc
            sleep(1);
        }

        public static void acrobat_close()
        {
            AutoItX.WinClose(wClass);
        }

        //
        // Utility
        //
        public static void sleep(int t)
        {
            System.Threading.Thread.Sleep(t * 1000);
        } // sleep
    }
}
