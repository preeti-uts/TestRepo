using OpenCvSharp.Extensions;
using OpenCvSharp;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Windows;
using System.Management;
using Point = OpenCvSharp.Point;
using System.IO;
using System.Drawing.Imaging;
using Tesseract;
using System.Text.RegularExpressions;
using System.Linq;
using Rect = OpenCvSharp.Rect;
using Size = OpenCvSharp.Size;

namespace selenium_framework
{
    public partial class Selenium_Framework
    {
        // Mouse methods from user32 dll
        static readonly int SM_CXSCREEN = 0;
        static readonly int SM_CYSCREEN = 1;

        [DllImport("user32.dll")]
        static extern int GetSystemMetrics(int nIndex);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);
        //Mouse actions
        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;
        private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
        private const uint MOUSEEVENTF_RIGHTUP = 0x0010;
        private const uint MOUSEEVENTF_MOVE = 0x0001;
        private const uint MOUSEEVENTF_ABSOLUTE = 0x8000;

        // Left mouse click
        public static void leftClick(int xPoint, int yPoint)
        {
            leftClick(xPoint, yPoint, 0);
        }
        public static void leftClick(int xPoint, int yPoint, int offset)
        {
            Console.WriteLine("Trace: leftClick");

            var toClickX = xPoint + offset;
            var toClickY = yPoint + offset;

            // Convert to screen pixel location
            int sx = GetSystemMetrics(SM_CXSCREEN);
            int sy = GetSystemMetrics(SM_CYSCREEN);

            int screenLocationX = toClickX * 65536 / sx;
            int screenLocationY = toClickY * 65536 / sy;

            // click the mouse
            mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_MOVE | MOUSEEVENTF_LEFTDOWN, (uint)screenLocationX, (uint)screenLocationY, 0, 0);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_MOVE | MOUSEEVENTF_LEFTUP, (uint)screenLocationX, (uint)screenLocationY, 0, 0);
        }

        // Right mouse click
        public static void rightClick(int xPoint, int yPoint)
        {
            rightClick(xPoint, yPoint, 0);
        }
        public static void rightClick(int xPoint, int yPoint, int offset)
        {
            Console.WriteLine("Trace: rightClick");

            var toClickX = xPoint + offset;
            var toClickY = yPoint + offset;

            // Convert to screen pixel location
            int sx = GetSystemMetrics(SM_CXSCREEN);
            int sy = GetSystemMetrics(SM_CYSCREEN);

            int screenLocationX = toClickX * 65536 / sx;
            int screenLocationY = toClickY * 65536 / sy;

            // click the mouse
            mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_RIGHTDOWN | MOUSEEVENTF_MOVE, (uint)screenLocationX, (uint)screenLocationY, 0, 0);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_RIGHTUP | MOUSEEVENTF_MOVE, (uint)screenLocationX, (uint)screenLocationY, 0, 0);
        }

        // Double click left mouse button
        public static void doubleClick(int xPoint, int yPoint)
        {
            doubleClick(xPoint, yPoint, 0);
        }

        public static void doubleClick(int xPoint, int yPoint, int offset)
        {
            Console.WriteLine("Trace: doubleClick");

            var toClickX = xPoint + offset;
            var toClickY = yPoint + offset;

            // Convert to screen pixel location
            int sx = GetSystemMetrics(SM_CXSCREEN);
            int sy = GetSystemMetrics(SM_CYSCREEN);

            int screenLocationX = toClickX * 65536 / sx;
            int screenLocationY = toClickY * 65536 / sy;

            // click the mouse
            mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_MOVE, (uint)screenLocationX, (uint)screenLocationY, 0, 0);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_LEFTUP | MOUSEEVENTF_MOVE, (uint)screenLocationX, (uint)screenLocationY, 0, 0);
            mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_MOVE, (uint)screenLocationX, (uint)screenLocationY, 0, 0);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_LEFTUP | MOUSEEVENTF_MOVE, (uint)screenLocationX, (uint)screenLocationY, 0, 0);
        }

        // Random mouse move - needed for some screens to refresh on load
        public void randomMouseMove()
        {
            Random myRandom = new Random();
            int mouseRandom = myRandom.Next(400, 400);

            var toMoveX = mouseRandom + myRandom.Next(100);
            var toMoveY = mouseRandom + myRandom.Next(100);
            System.Windows.Forms.Cursor.Position = new System.Drawing.Point(toMoveX, toMoveY);

            /*var toClickX = mouseRandom;
            var toClickY = mouseRandom;

            // Convert to screen pixel location
            int sx = GetSystemMetrics(SM_CXSCREEN);
            int sy = GetSystemMetrics(SM_CYSCREEN);

            int screenLocationX = toClickX * 65536 / sx;
            int screenLocationY = toClickY * 65536 / sy;

            // Move the mouse
            mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_MOVE, (uint)screenLocationX, (uint)screenLocationY, 0, 0);
            Thread.Sleep(100);
            //mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_MOVE, (uint)screenLocationX, (uint)screenLocationY, 0, 0);
            */
        }

        // Mouce the mouse to a specific location
        public void MouseMove(int toMoveX, int toMoveY)
        {
            Console.WriteLine("Trace: MouseMove");

            System.Windows.Forms.Cursor.Position = new System.Drawing.Point(toMoveX, toMoveY);
        }

        // Move the mouse cursor relative to its current position
        public void MouseMoveRelative(int offsetX, int offsetY)
        {
            Console.WriteLine("Trace: MouseMoveRelative");

            // Get the current cursor position
            var currentPosition = System.Windows.Forms.Cursor.Position;

            // Calculate the new position
            var newPosition = new System.Drawing.Point(currentPosition.X + offsetX, currentPosition.Y + offsetY);

            // Set the cursor to the new position
            System.Windows.Forms.Cursor.Position = newPosition;
        }

        // Check if file exists
        public bool file_exists(string pathFilename)
        {
            return File.Exists(pathFilename);
        }

        // Captures the primary desktop screen 
        private Bitmap captureDesktop()
        {

            //using (
            Bitmap bmpScreenshot = new Bitmap(Screen.PrimaryScreen.Bounds.Width,
                                        Screen.PrimaryScreen.Bounds.Height, PixelFormat.Format24bppRgb); //)
                                                                                                         //{
            using (Graphics g = Graphics.FromImage(bmpScreenshot))
            {
                g.CopyFromScreen(Screen.PrimaryScreen.Bounds.X,
                                 Screen.PrimaryScreen.Bounds.Y,
                                 0, 0,
                                 Screen.PrimaryScreen.Bounds.Size,
                                 CopyPixelOperation.SourceCopy);
                //return bmpScreenshot;
            }
            return bmpScreenshot;
            //}

            /*                            //Create a new bitmap.
                                        Bitmap bmpScreenshot = new Bitmap(Screen.PrimaryScreen.Bounds.Width,
                                                                                   Screen.PrimaryScreen.Bounds.Height,
                                                                                   PixelFormat.Format24bppRgb);

                                    // Create a graphics object from the bitmap.
                                    Graphics currentDesktop = Graphics.FromImage(bmpScreenshot);

                                    // Take the screenshot from the upper left corner to the right bottom corner.    
                                    currentDesktop.CopyFromScreen(Screen.PrimaryScreen.Bounds.X,
                                                             Screen.PrimaryScreen.Bounds.Y,
                                                             0,
                                                             0,
                                                             Screen.PrimaryScreen.Bounds.Size,
                                                             CopyPixelOperation.SourceCopy);

                                    currentDesktop.Dispose();
                                    return bmpScreenshot;*/
        }

        // Wait for a tempalte image to appear on screen
        public void imageWait(string templateLocation, int timeOutVal)
        {
            imageWait(templateLocation, timeOutVal, false, true, 0.95);
        }

        public void imageWait(string templateLocation, int timeOutVal, bool moveMouse)
        {
            imageWait(templateLocation, timeOutVal, moveMouse, false, 0.95);
        }
        public void imageWait(string templateLocation, int timeOutVal, bool moveMouse, bool checkTimeout)
        {
            imageWait(templateLocation, timeOutVal, moveMouse, checkTimeout, 0.95);
        }
        public void imageWait(string templateLocation, int timeOutVal, bool moveMouse, bool checkTimeout, double imageConfidence)
        {
            debug(TRACE, "imageWait: start. Template : " + templateLocation);

            if (!file_exists(templateLocation))
            {
                // Try and find the image in the images folder / sub folder
                string origtemplateLocation = templateLocation;
                templateLocation = returnImagePath(templateLocation);
                if (templateLocation == null || templateLocation == "")
                {
                    debug(ERROR, "myImageWait: File does not exist. File : " + origtemplateLocation);
                }
            }

            if (imageConfidence < 0 || imageConfidence > 1)
            {
                debug(ERROR, "myImageWait: Image confidence is < 0 or > 1. Value used was: " + imageConfidence);
            }

            Bitmap theDesktop = null;

            Mat myScreen = null;
            Mat imageToFind = null;
            Mat matchResult = null;

            double minVal;
            double maxVal = 0.0;

            Point minloc;
            Point maxloc;

            int counter = 0;

            while (maxVal < imageConfidence && counter < timeOutVal)
            {
                // Capture desktop image and convert to MAT
                theDesktop = captureDesktop();
                myScreen = theDesktop.ToMat();

                // Load template from file
                imageToFind = new Mat(templateLocation);

                // Find the template on the desktop
                try
                {
                    matchResult = myScreen.MatchTemplate(imageToFind, TemplateMatchModes.CCoeffNormed);
                }
                catch (Exception e)
                {
                    //debug(1, String.Format("myImageWait: failed with error : {0}.", e.Message));
                    Debug.WriteLine("myImageWait: failed with error : {0}.", e.Message);
                }


                matchResult.MinMaxLoc(out minVal, out maxVal, out minloc, out maxloc);
                Debug.WriteLine("minval : " + minVal + " maxval : " + maxVal);

                Thread.Sleep(1000);

                if (moveMouse)
                {
                    randomMouseMove();
                }

                counter++;

                // clean up before next loop
                matchResult.Dispose();
                theDesktop.Dispose();
                myScreen.Dispose();
                imageToFind.Dispose();
            }

            Debug.WriteLine("maxval : " + maxVal);
            Debug.WriteLine("time taken : " + counter);

            theDesktop.Dispose();
            imageToFind.Dispose();
            myScreen.Dispose();
            matchResult.Dispose();

            if (checkTimeout)
            {
                if (maxVal < imageConfidence && counter >= timeOutVal)
                {
                    Debug.WriteLine("ImageWait TIMEOUT -  maxval : " + maxVal);
                    debug(ERROR, "imageWait error : image did not appear in specified time. Timeout value : " + timeOutVal + " Counter value : " + counter);
                }
            }

            debug(TRACE, "imageWait: time for image wait : " + counter);

        }

        // Wait for a template image to vanish from the screen
        public void imageVanish(string templateLocation, int timeOutVal)
        {
            imageVanish(templateLocation, timeOutVal, 0.95);
        }

        public void imageVanish(string templateLocation, int timeOutVal, double imageConfidence)
        {
            debug(TRACE, "imageVanish: start. Template : " + templateLocation);

            if (!file_exists(templateLocation))
            {
                // Try and find the image in the images folder / sub folder
                string origtemplateLocation = templateLocation;
                templateLocation = returnImagePath(templateLocation);
                if (templateLocation == null || templateLocation == "")
                {
                    debug(ERROR, "myImageWait: File does not exist. File : " + origtemplateLocation);
                }
            }

            Bitmap theDesktop = null;

            Mat myScreen = null;
            Mat imageToFind = null;
            Mat matchResult = null;

            double minVal;
            double maxVal = 1.0;

            Point minloc;
            Point maxloc;

            int counter = 0;

            while (maxVal > imageConfidence && counter < timeOutVal)
            {
                // Capture desktop image and convert to MAT
                theDesktop = captureDesktop();
                myScreen = theDesktop.ToMat();

                // Load template from file
                imageToFind = new Mat(templateLocation);

                // Find the template on the desktop            
                matchResult = myScreen.MatchTemplate(imageToFind, TemplateMatchModes.CCoeffNormed);

                matchResult.MinMaxLoc(out minVal, out maxVal, out minloc, out maxloc);
                Debug.WriteLine("minval : " + minVal + " maxval : " + maxVal);

                Thread.Sleep(1000);
                counter++;

                // clean up before next loop
                matchResult.Dispose();
                theDesktop.Dispose();
                myScreen.Dispose();
                imageToFind.Dispose();
            }

            Debug.WriteLine("maxval : " + maxVal);
            Debug.WriteLine("time taken : " + counter);

            theDesktop.Dispose();
            imageToFind.Dispose();
            myScreen.Dispose();
            matchResult.Dispose();

            if (maxVal < imageConfidence && counter >= timeOutVal)
            {
                debug(ERROR, "myImageVanish error : image did not vanish in specified time. Timeout value : " + timeOutVal);
            }

            debug(TRACE, "imageVanish: time for image vanish : " + counter);
        }

        // Find and click in the centre of a template image on screen
        public void imageFindClick(string templateLocation)
        {
            imageFindClick(templateLocation, 0, 0);
        }

        public void imageFindClick(string templateLocation, int xOffSet, int yOffset)
        {
            imageFindClick(templateLocation, xOffSet, yOffset, 0.95);
        }

        // test image find and click using openCV
        public void imageFindClick(string templateLocation, int xOffSet, int yOffset, double imageConfidence)
        {
            imageFindClick(templateLocation, xOffSet, yOffset, 0.95, "leftClick");

        }

        public void imageFindClick(string templateLocation, int xOffSet, int yOffset, double imageConfidence, string clickType)
        {
            debug(TRACE, "imageFindClick: start. Template : " + templateLocation);

            if (!file_exists(templateLocation))
            {
                // Try and find the image in the images folder / sub folder
                string origtemplateLocation = templateLocation;
                templateLocation = returnImagePath(templateLocation);
                if (templateLocation == null || templateLocation == "")
                {
                    debug(ERROR, "myImageWait: File does not exist. File : " + origtemplateLocation);
                }
            }

            if (imageConfidence < 0 || imageConfidence > 1)
            {
                debug(ERROR, "myImageWait: Image confidence is < 0 or > 1. Value used was: " + imageConfidence);
            }

            Bitmap theDesktop = null;

            Mat myScreen = null;
            Mat imageToFind = null;
            Mat matchResult = null;

            double minVal;
            double maxVal;

            Point minloc;
            Point maxloc;

            // Capture desktop image and convert to MAT
            theDesktop = captureDesktop();
            myScreen = theDesktop.ToMat();

            // Load template from file
            imageToFind = new Mat(templateLocation);

            // Find the template on the desktop            
            matchResult = myScreen.MatchTemplate(imageToFind, TemplateMatchModes.CCoeffNormed);

            matchResult.MinMaxLoc(out minVal, out maxVal, out minloc, out maxloc);
            Debug.WriteLine("minval : " + minVal + " maxval : " + maxVal);

            // get centre of image.
            //if (maxVal >= imageConfidence)
            //{
            var toClickX = maxloc.X + imageToFind.Width / 2;
            var toClickY = maxloc.Y + imageToFind.Height / 2;


            if (clickType == "leftClick")
            {
                leftClick(toClickX + xOffSet, toClickY + yOffset);
            }
            else if (clickType == "rightClick")
            {
                rightClick(toClickX + xOffSet, toClickY + yOffset);
            }
            else if (clickType == "doubleClick")
            {
                doubleClick(toClickX + xOffSet, toClickY + yOffset);
            }
            //}
            //else
            //{
            //    debug(ERROR, "imageFindClick: Image confidence found was: "+ maxVal +". Expected confidence was: "+ imageConfidence +".");
            //}
            // clean up
            theDesktop.Dispose();
            imageToFind.Dispose();
            imageToFind.Dispose();
            myScreen.Dispose();
            matchResult.Dispose();

        }

        // Class for image location values
        public class LocationValues
        {
            public int X { get; set; }
            public int Y { get; set; }
            public int H { get; set; }
            public int W { get; set; }
        }

        // Returns the location values for an image location
        public LocationValues getImageLocation(string templateLocation)
        {
            debug(TRACE, "getImageLocation: start. Template :" + templateLocation);

            if (!file_exists(templateLocation))
            {
                // Try and find the image in the images folder / sub folder
                string origtemplateLocation = templateLocation;
                templateLocation = returnImagePath(templateLocation);
                if (templateLocation == null || templateLocation == "")
                {
                    debug(ERROR, "myImageWait: File does not exist. File : " + origtemplateLocation);
                }
            }
            Bitmap theDesktop = null;
            Mat myScreen = null;
            Mat imageToFind = null;
            Mat matchResult = null;
            double minVal;
            double maxVal;
            Point minloc;
            Point maxloc;

            // Capture desktop image and convert to MAT
            theDesktop = captureDesktop();
            myScreen = theDesktop.ToMat();

            // Load template from file
            imageToFind = new Mat(templateLocation);

            // Find the template on the desktop
            matchResult = myScreen.MatchTemplate(imageToFind, TemplateMatchModes.CCoeffNormed);
            myScreen.MatchTemplate(imageToFind, TemplateMatchModes.CCoeffNormed);

            matchResult.MinMaxLoc(out minVal, out maxVal, out minloc, out maxloc);
            Debug.WriteLine("minval : " + minVal + " maxval : " + maxVal);

            LocationValues locationValues = new LocationValues() { X = maxloc.X, Y = maxloc.Y, H = imageToFind.Height, W = imageToFind.Width };

            // Clean up
            theDesktop.Dispose();
            imageToFind.Dispose();
            imageToFind.Dispose();
            myScreen.Dispose();
            matchResult.Dispose();

            // results to debug log
            Debug.WriteLine("maxloc.X : " + locationValues.X);
            Debug.WriteLine("maxloc.Y : " + locationValues.Y);
            Debug.WriteLine("matchResult.Height : " + locationValues.H);
            Debug.WriteLine("matchResult.Width : " + locationValues.W);

            return locationValues;
        }


        // old getImageLocation method using array - here for compatability
        public int[] getImageLocation_old(string templateLocation)
        {
            debug(TRACE, "getImageLocation: start. Template :" + templateLocation);

            if (!file_exists(templateLocation))
            {
                // Try and find the image in the images folder / sub folder
                string origtemplateLocation = templateLocation;
                templateLocation = returnImagePath(templateLocation);
                if (templateLocation == null || templateLocation == "")
                {
                    debug(ERROR, "myImageWait: File does not exist. File : " + origtemplateLocation);
                }
            }

            int[] locationValues = new int[4];

            Bitmap theDesktop = null;

            Mat myScreen = null;
            Mat imageToFind = null;
            Mat matchResult = null;

            double minVal;
            double maxVal;

            Point minloc;
            Point maxloc;

            // Capture desktop image and convert to MAT
            theDesktop = captureDesktop();
            myScreen = theDesktop.ToMat();

            // Load template from file
            imageToFind = new Mat(templateLocation);

            // Find the template on the desktop            
            matchResult = myScreen.MatchTemplate(imageToFind, TemplateMatchModes.CCoeffNormed);
            myScreen.MatchTemplate(imageToFind, TemplateMatchModes.CCoeffNormed);

            matchResult.MinMaxLoc(out minVal, out maxVal, out minloc, out maxloc);
            Debug.WriteLine("minval : " + minVal + " maxval : " + maxVal);

            locationValues[0] = maxloc.X;
            locationValues[1] = maxloc.Y;
            locationValues[2] = imageToFind.Height;
            locationValues[3] = imageToFind.Width;

            theDesktop.Dispose();
            imageToFind.Dispose();
            imageToFind.Dispose();
            myScreen.Dispose();
            matchResult.Dispose();

            Debug.WriteLine("maxloc.X : " + locationValues[0]);
            Debug.WriteLine("maxloc.Y : " + locationValues[1]);
            Debug.WriteLine("matchResult.Height : " + locationValues[2]);
            Debug.WriteLine("matchResult.Width : " + locationValues[3]);

            return locationValues;

        }

        // Checks if an image exists on the current screen
        public bool imageExists(string templateLocation)
        {
            return imageExists(templateLocation, 0.95);
        }

        public bool imageExists(string templateLocation, double imageConfidence)
        {
            debug(TRACE, "imageWait: start. Template : " + templateLocation);

            if (!file_exists(templateLocation))
            {
                // Try and find the image in the images folder / sub folder
                string origtemplateLocation = templateLocation;
                templateLocation = returnImagePath(templateLocation);
                if (templateLocation == null || templateLocation == "")
                {
                    debug(ERROR, "myImageWait: File does not exist. File : " + origtemplateLocation);
                }
            }

            Bitmap theDesktop = null;

            Mat myScreen = null;
            Mat imageToFind = null;
            Mat matchResult = null;

            double minVal;
            double maxVal = 0.0;

            Point minloc;
            Point maxloc;

            // Capture desktop image and convert to MAT
            theDesktop = captureDesktop();
            myScreen = theDesktop.ToMat();

            // Load template from file
            imageToFind = new Mat(templateLocation);

            // Find the template on the desktop
            try
            {
                matchResult = myScreen.MatchTemplate(imageToFind, TemplateMatchModes.CCoeffNormed);
            }
            catch (Exception e)
            {
                Debug.WriteLine("imageWait: failed with error : {0}.", e.Message);
            }

            // Get the match results
            matchResult.MinMaxLoc(out minVal, out maxVal, out minloc, out maxloc);
            Debug.WriteLine("minval : " + minVal + " maxval : " + maxVal);

            // Clean up
            theDesktop.Dispose();
            myScreen.Dispose();
            imageToFind.Dispose();
            matchResult.Dispose();

            // Check the value and return the result
            if (maxVal >= imageConfidence)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public bool imageExistsArea(string templateLocation, double imageConfidence, LocationValues searchArea)
        {
            debug(TRACE, "imageWait: start. Template : " + templateLocation);

            if (!file_exists(templateLocation))
            {
                // Try and find the image in the images folder / sub folder
                string origtemplateLocation = templateLocation;
                templateLocation = returnImagePath(templateLocation);
                if (templateLocation == null || templateLocation == "")
                {
                    debug(ERROR, "myImageWait: File does not exist. File : " + origtemplateLocation);
                }
            }

            Bitmap theDesktop = null;

            Mat myScreen = null;
            Mat imageToFind = null;
            Mat matchResult = null;

            double minVal;
            double maxVal = 0.0;

            Point minloc;
            Point maxloc;

            // Capture desktop image and convert to MAT
            theDesktop = captureDesktopArea(searchArea);
            myScreen = theDesktop.ToMat();

            // Load template from file
            imageToFind = new Mat(templateLocation);

            // Find the template on the desktop
            try
            {
                matchResult = myScreen.MatchTemplate(imageToFind, TemplateMatchModes.CCoeffNormed);
            }
            catch (Exception e)
            {
                Debug.WriteLine("imageWait: failed with error : {0}.", e.Message);
            }

            // Get the match results
            matchResult.MinMaxLoc(out minVal, out maxVal, out minloc, out maxloc);
            Debug.WriteLine("minval : " + minVal + " maxval : " + maxVal);

            // Clean up
            theDesktop.Dispose();
            myScreen.Dispose();
            imageToFind.Dispose();
            matchResult.Dispose();

            // Check the value and return the result
            if (maxVal > imageConfidence)
            {
                return true;
            }
            else
            {
                return false;
            }

        }
        // Checks Windows screen scaling setting - framework will not work if scaling is not 100%
        public static int GetWindowsScaling()
        {
            return (int)(100 * Screen.PrimaryScreen.Bounds.Width / SystemParameters.PrimaryScreenWidth);
        }

        // Detect if test is running on a virtual machine
        public bool DetectVirtualMachine()
        {
            bool result = false;
            const string MICROSOFTCORPORATION = "microsoft corporation";

            debug(TRACE, "DetectVirtualMachine: start");

            try
            {
                ManagementObjectSearcher searcher =
                    new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_BaseBoard");

                foreach (ManagementObject queryObj in searcher.Get().Cast<ManagementObject>())
                {
                    result = queryObj["Manufacturer"].ToString().ToLower() == MICROSOFTCORPORATION.ToLower();
                }
                debug(TRACE, "DetectVirtualMachine: result: " + result);
                return result;
            }
            catch (ManagementException ex)
            {
                Debug.WriteLine("TRACE : ManagementException : " + ex.Message);
                debug(TRACE, "DetectVirtualMachine: result: " + result + ". Exception message: " + ex.Message);
                return result;
            }
        }

        // Find and click on a single word of text for a given screen area
        public void findandClickText(LocationValues textArea, string stringToFind)
        {
            ;
            findandClickText(textArea.W, textArea.H, textArea.X, textArea.Y, stringToFind);
        }
        public void findandClickText(int myGetW, int myGetH, int myGetX, int myGetY, string stringToFind)
        {
            findandClickText(myGetW, myGetH, myGetX, myGetY, stringToFind, false);
        }
        public void findandClickText(int myGetW, int myGetH, int myGetX, int myGetY, string stringToFind, bool scale)
        {
            findandClickText(myGetW, myGetH, myGetX, myGetY, stringToFind, scale, 2);
        }
        public void findandClickText(int myGetW, int myGetH, int myGetX, int myGetY, string stringToFind, bool scale, int scaleValue)
        {

            debug(TRACE, "findandClickText: start");

            bool textWasFound = false;

            // save the screen region for the results table for ocr
            tesseractCaptureScreen(myGetW, myGetH, myGetX, myGetY, dirName + @"\cassTemplate_images\tempImages\findandClickText.TIFF");

            // Find the text from the screenshot and click
            debug(TRACE, "findandClickText: tesseractFindAndClickText - incomming values");
            textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\findandClickText.TIFF", "word", scale, scaleValue, myGetX, myGetY, "single", "contains", stringToFind, false);
            if (!textWasFound)
            {
                debug(TRACE, "findandClickText: tesseractFindAndClickText - opposite incomming values");
                textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\findandClickText.TIFF", "word", !scale, scaleValue, myGetX, myGetY, "single", "contains", stringToFind, false);
                if (!textWasFound)
                {
                    debug(TRACE, "findandClickText: tesseractFindAndClickText - no scale, invert image");
                    textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\findandClickText.TIFF", "word", !scale, scaleValue, myGetX, myGetY, "single", "contains", stringToFind, true);
                    if (!textWasFound)
                    {
                        debug(TRACE, "findandClickText: tesseractFindAndClickText - scaled, invert image");
                        textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\findandClickText.TIFF", "word", scale, scaleValue, myGetX, myGetY, "single", "contains", stringToFind, true);
                    }
                }
            }

            // Check result of text find
            if (textWasFound)
            {
                Console.WriteLine("Text Found.");
            }
            else
            {
                Console.WriteLine("Text was not found.");
                debug(ERROR, "Text was not found. expected: " + stringToFind);

            }
        }

        public void findandDoubleClickText(int myGetW, int myGetH, int myGetX, int myGetY, string stringToFind, bool scale, int scaleValue)
        {

            debug(TRACE, "findandDoubleClickText: start");

            bool textWasFound = false;

            // save the screen region for the results table for ocr
            tesseractCaptureScreen(myGetW, myGetH, myGetX, myGetY, dirName + @"\cassTemplate_images\tempImages\findandClickText.TIFF");

            // Find the text from the screenshot and click
            debug(TRACE, "findandDoubleClickText: tesseractFindAndClickText - incomming values");
            textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\findandClickText.TIFF", "word", scale, scaleValue, myGetX, myGetY, "double", "contains", stringToFind, false);
            if (!textWasFound)
            {
                debug(TRACE, "findandDoubleClickText: tesseractFindAndClickText - opposite incomming values");
                textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\findandClickText.TIFF", "word", !scale, scaleValue, myGetX, myGetY, "double", "contains", stringToFind, false);
                if (!textWasFound)
                {
                    debug(TRACE, "findandDoubleClickText: tesseractFindAndClickText - no scale, invert image");
                    textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\findandClickText.TIFF", "word", !scale, scaleValue, myGetX, myGetY, "double", "contains", stringToFind, true);
                    if (!textWasFound)
                    {
                        debug(TRACE, "findandDoubleClickText: tesseractFindAndClickText - scaled, invert image");
                        textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\findandClickText.TIFF", "word", scale, scaleValue, myGetX, myGetY, "double", "contains", stringToFind, true);
                    }
                }
            }

            // Check result of text find
            if (textWasFound)
            {
                Console.WriteLine("Text Found.");
            }
            else
            {
                Console.WriteLine("Text was not found.");
                debug(ERROR, "Text was not found. expected: " + stringToFind);

            }
        }
        // Find text in a screen area and return
        public string findandGetText(LocationValues location)
        {
            int myGetX = location.X;
            int myGetY = location.Y;
            int myGetH = location.H;
            int myGetW = location.W;

            return findandGetText(myGetW, myGetH, myGetX, myGetY, false);
        }
        public string findandGetText(LocationValues location, bool scale)
        {
            int myGetX = location.X;
            int myGetY = location.Y;
            int myGetH = location.H;
            int myGetW = location.W;

            return findandGetText(myGetW, myGetH, myGetX, myGetY, scale);
        }
        public string findandGetText(int myGetW, int myGetH, int myGetX, int myGetY)
        {
            return findandGetText(myGetW, myGetH, myGetX, myGetY, false);
        }
        public string findandGetText(int myGetW, int myGetH, int myGetX, int myGetY, bool scale)
        {
            return findandGetText(myGetW, myGetH, myGetX, myGetY, scale, false);
        }

        public string findandGetText(int myGetW, int myGetH, int myGetX, int myGetY, bool scale, bool invert)
        {
            return findandGetText(myGetW, myGetH, myGetX, myGetY, scale, 2, invert);
        }
        public string findandGetText(int myGetW, int myGetH, int myGetX, int myGetY, bool scale, int scaleValue, bool invert)
        {

            debug(TRACE, "findandGetText: start");

            // save the screen region for the results table for ocr
            tesseractCaptureScreen(myGetW, myGetH, myGetX, myGetY, dirName + @"\cassTemplate_images\tempImages\findandGetText.TIFF");

            using (var engine = new TesseractEngine(dirName + @"\tessdata", "eng", EngineMode.Default))
            {
                using (var img = Pix.LoadFromFile(dirName + @"\cassTemplate_images\tempImages\findandGetText.TIFF"))
                {
                    var imgToOCR = img;
                    if (scale)
                    {
                        imgToOCR = imgToOCR.Scale(scaleValue, scaleValue);
                    }
                    if (invert)
                    {
                        imgToOCR = imgToOCR.ConvertRGBToGray();
                        imgToOCR = imgToOCR.Invert();
                    }
                    imgToOCR.Save(dirName + @"\cassTemplate_images\tempImages\findandGetText_processed.TIFF", Tesseract.ImageFormat.Tiff); // for debug
                    using (var page = engine.Process(imgToOCR))
                    {
                        var text = page.GetText();
                        Console.WriteLine("Text extracted: "+ text);
                        debug(TRACE, "findandGetText: text found : " + text);
                        imgToOCR.Dispose();
                        return text;
                    }

                }
            }
        }
        // Send characters to the currently selected field
        // For special characters see : https://learn.microsoft.com/en-us/dotnet/api/system.windows.forms.sendkeys?view=windowsdesktop-7.0
        public void typeTextInField(string textToType)
        {
            debug(TRACE, "typeTextInField: start");

            try
            {
                Thread.Sleep(500); // add sleep to hopefully improve reliability... maybe?!?!
                SendKeys.SendWait(@textToType);
                Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                debug(ERROR, "typeTextInField: error typing into field : " + ex.Message);
            }
        }

        // Clear characters to the currently selected field
        public void clearTextInField()
        {
            debug(TRACE, "clearTextInField: start");

            try
            {
                // seperate the sendkeys make it more reliable?
                Thread.Sleep(500);
                SendKeys.SendWait(@"{HOME}");
                Thread.Sleep(500);
                SendKeys.SendWait(@"+{END}");
                Thread.Sleep(500);
                SendKeys.SendWait(@"{DEL}");
                Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                debug(ERROR, "clearTextInField: error clearing text from field : " + ex.Message);
            }
        }

        // Check the screen resolution and error out if it does not meet the requires resolution of 1920 x 1080
        public void checkScreenResolution()
        {
            debug(TRACE, "checkScreenResolution: start");

            double height = SystemParameters.FullPrimaryScreenHeight;
            double width = SystemParameters.FullPrimaryScreenWidth;
            double resolution = height * width;
            debug(TRACE, "Screen resolution is:" + width + " x " + height);

            if (width != 1920 && height != 1017)
            {
                Debug.WriteLine("SCREEN: height : " + height + " width : " + width + " resolution : " + resolution);
                debug(ERROR, "SCREEN: width : " + width + " height: " + height + ".\r\nChange the display resolution to 1920 x 1080 in Display settings.");
            }

            // Check window scaling is 100%
            int screenScale = GetWindowsScaling();
            if (screenScale != 100)
            {
                Debug.WriteLine("Screen scale is not 100%, current scale is :" + screenScale);
                debug(ERROR, "Screen scale is not 100%, current scale is :" + screenScale + ".\r\nChange screen scale in Display settings to 100%");
            }
        }

        // Find text in a string from a regular expression
        public string findTextInString(string patternToFind, string stringToCheck)
        {
            debug(TRACE, "findTextInString: start");

            string stringFound;

            string pattern = patternToFind;
            Regex regex = new Regex(pattern);
            MatchCollection matches = regex.Matches(stringToCheck);
            if (matches.Count.Equals(0))
            {
                debug(ERROR, "Did not find a match for " + patternToFind + " in results for  " + patternToFind);
                return null;
            }
            else
            {
                foreach (Match match in matches)
                {
                    // Get the first value ;)
                    stringFound = match.Value;
                    debug(TRACE, "Match found: " + stringFound);

                    return stringFound;
                }
            }

            return null;
        }

        // temp for Sheetal's script
        public void findandClickTextLine_OCRMismatch(int myGetW, int myGetH, int myGetX, int myGetY, string textComparisonType, string stringToFind, bool scale, int scaleValue, string clickType)
        {

            debug(TRACE, "findandClickText: start");

            bool textWasFound = false;

            // save the screen region for the results table for ocr
            tesseractCaptureScreen(myGetW, myGetH, myGetX, myGetY, dirName + @"\cassTemplate_images\tempImages\findandClickTextLine_OCRMismatch.TIFF");

            // Find the text from the screenshot and click
            textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\findandClickTextLine_OCRMismatch.TIFF", "TextLine", scale, scaleValue, myGetX, myGetY, clickType, textComparisonType, stringToFind, false);
            if (!textWasFound)
            {
                textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\findandClickTextLine_OCRMismatch.TIFF", "TextLine", scale, scaleValue, myGetX, myGetY, clickType, textComparisonType, stringToFind, !false);
                
                if (!textWasFound)
                {
                    textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\findandClickTextLine_OCRMismatch.TIFF", "TextLine", !scale, scaleValue, myGetX, myGetY, clickType, textComparisonType, stringToFind, !false);
                    
                    if (!textWasFound)
                    {
                        textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\findandClickTextLine_OCRMismatch.TIFF", "TextLine", scale, 2, myGetX, myGetY, clickType, textComparisonType, stringToFind, !false);
                    }
                }
            }
           
            // Check result of text find
            if (textWasFound)
            {
                Console.WriteLine("Text Found.");
            }
            else
            {
                Console.WriteLine("Text was not found.");
                debug(ERROR, "Text was not found. expected: " + stringToFind);

            }
        }

        // Find text in a line of text and click the value on screen
        public void findandClickTextLine(LocationValues textArea, string stringToFind)
        {
            findandClickTextLine(textArea.W, textArea.H, textArea.X, textArea.Y, stringToFind, false, "single");
        }
        public void findandClickTextLine(int myGetW, int myGetH, int myGetX, int myGetY, string stringToFind, bool scale, string clickType)
        {
            findandClickTextLine(myGetW, myGetH, myGetX, myGetY, "contains", stringToFind, scale, 2, false, clickType);
        }
        public void findandClickTextLine(int myGetW, int myGetH, int myGetX, int myGetY, string stringToFind, bool scale, int scaleValue, string clickType)
        {
            findandClickTextLine(myGetW, myGetH, myGetX, myGetY, "contains", stringToFind, scale, scaleValue, false, clickType);
        }
        public void findandClickTextLine(int myGetW, int myGetH, int myGetX, int myGetY, string textComparisonType, string stringToFind, bool scale, int scaleValue, bool invertImage, string clickType)
        {

            debug(TRACE, "findandClickText: start");

            bool textWasFound = false;

            // save the screen region for the results table for ocr
            tesseractCaptureScreen(myGetW, myGetH, myGetX, myGetY, dirName + @"\cassTemplate_images\tempImages\findandClickTextLine.TIFF");

            // Find the text from the screenshot and click
            debug(TRACE, "findandClickText: text find 1");
            textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\findandClickTextLine.TIFF", "TextLine", scale, scaleValue, myGetX, myGetY, clickType, textComparisonType, stringToFind, invertImage);
            if (!textWasFound)
            {
                debug(TRACE, "findandClickText: text find 2");
                textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\findandClickTextLine.TIFF", "TextLine", scale, scaleValue, myGetX, myGetY, clickType, textComparisonType, stringToFind, !invertImage);
                if (!textWasFound)
                {
                    debug(TRACE, "findandClickText: text find 3");
                    textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\findandClickTextLine.TIFF", "TextLine", !scale, scaleValue, myGetX, myGetY, clickType, textComparisonType, stringToFind, invertImage);
                    if (!textWasFound)
                    {
                        debug(TRACE, "findandClickText: text find 4");
                        textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\findandClickTextLine.TIFF", "TextLine", scale, 3, myGetX, myGetY, clickType, textComparisonType, stringToFind, invertImage);
                        if (!textWasFound)
                        {
                            debug(TRACE, "findandClickText: text find 5");
                            textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\findandClickTextLine.TIFF", "TextLine", scale, 3, myGetX, myGetY, clickType, textComparisonType, stringToFind, !invertImage);
                            if (!textWasFound)
                            {
                                debug(TRACE, "findandClickText: text find 6");
                                textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\findandClickTextLine.TIFF", "TextLine", scale, 4, myGetX, myGetY, clickType, textComparisonType, stringToFind, !invertImage);
                            }
                        }
                    }
                }
            }

            // Check result of text find
            if (textWasFound)
            {
                Console.WriteLine("Text Found.");
            }
            else
            {
                Console.WriteLine("Text was not found.");
                debug(ERROR, "Text was not found. expected: " + stringToFind);

            }
        }

        // Find and return the location of text in a given search area
        public LocationValues findTextLocation(LocationValues searchLocation, string stringToFind, bool scale, int scaleValue)
        {
            return findTextLocation(searchLocation, stringToFind, "word", scale, scaleValue);
        }
        public LocationValues findTextLocation(LocationValues searchLocation, string stringToFind, string pageIterationLevel, bool scale, int scaleValue)
        {

            debug(TRACE, "findTextLocation: start");

            //bool textWasFound = false;
            LocationValues textWasFound = new LocationValues();

            // save the screen region for the results table for ocr
            tesseractCaptureScreen(searchLocation.W, searchLocation.H, searchLocation.X, searchLocation.Y, dirName + @"\cassTemplate_images\tempImages\findTextLocation.TIFF");

            // Find the text from the screenshot and click
            debug(TRACE, "findandClickText: tesseractFindAndClickText - incomming values");
            textWasFound = tesseractFindTextLocation(dirName + @"\cassTemplate_images\tempImages\findTextLocation.TIFF", pageIterationLevel, scale, scaleValue, searchLocation.X, searchLocation.Y, "single", "contains", stringToFind, false);
            if (textWasFound.X == 0)
            {
                debug(TRACE, "findandClickText: tesseractFindAndClickText - opposite incoming values");
                textWasFound = tesseractFindTextLocation(dirName + @"\cassTemplate_images\tempImages\findTextLocation.TIFF", pageIterationLevel, !scale, scaleValue, searchLocation.X, searchLocation.Y, "single", "contains", stringToFind, false);
                if (textWasFound.X == 0)
                {
                    debug(TRACE, "findandClickText: tesseractFindAndClickText - no scale, invert image");
                    textWasFound = tesseractFindTextLocation(dirName + @"\cassTemplate_images\tempImages\findTextLocation.TIFF", pageIterationLevel, !scale, scaleValue, searchLocation.X, searchLocation.Y, "single", "contains", stringToFind, true);
                    if (textWasFound.X == 0)
                    {
                        debug(TRACE, "findandClickText: tesseractFindAndClickText - scaled, invert image");
                        textWasFound = tesseractFindTextLocation(dirName + @"\cassTemplate_images\tempImages\findTextLocation.TIFF", pageIterationLevel, scale, scaleValue, searchLocation.X, searchLocation.Y, "single", "contains", stringToFind, true);
                        if (textWasFound.X == 0)
                        {
                            debug(TRACE, "findandClickText: tesseractFindAndClickText - scaled, invert image");
                            textWasFound = tesseractFindTextLocation(dirName + @"\cassTemplate_images\tempImages\findTextLocation.TIFF", pageIterationLevel, scale, 3, searchLocation.X, searchLocation.Y, "single", "contains", stringToFind, false);
                            if (textWasFound.X == 0)
                            {
                                debug(TRACE, "findandClickText: tesseractFindAndClickText - scaled, invert image");
                                textWasFound = tesseractFindTextLocation(dirName + @"\cassTemplate_images\tempImages\findTextLocation.TIFF", pageIterationLevel, scale, 3, searchLocation.X, searchLocation.Y, "single", "contains", stringToFind, true);
                            }
                        }
                    }
                }
            }

            // Check result of text find
            if (textWasFound.X != 0)
            {
                Console.WriteLine("Text Found.");
                return textWasFound;
            }
            else
            {
                Console.WriteLine("Text was not found.");
                debug(TRACE, "Text was not found. expected: " + stringToFind);
                return textWasFound;
            }
        }

        // Find and return the location of text in a given search area using OCR
        public LocationValues tesseractFindTextLocation(string screenshotToLoad, string pageIteratorLevel, bool scale, int scaleValue, int inGetX, int inGetY, string clickType, string textComparisonType, string stringToFind, bool invertImage)
        {
            debug(TRACE, "tesseractFindTextLocation: start");

            //int clickX = 0;
            //int clickY = 0;
            bool textMatched = false;
            LocationValues textLocations = new LocationValues();

            // set the page iterator, default to word if text unknown
            PageIteratorLevel myLevel;
            pageIteratorLevel = pageIteratorLevel.ToLower();
            switch (pageIteratorLevel)
            {
                case "word":
                    myLevel = PageIteratorLevel.Word;
                    break;
                case "textline":
                    myLevel = PageIteratorLevel.TextLine;
                    break;
                case "block":
                    myLevel = PageIteratorLevel.Block;
                    break;
                case "para":
                    myLevel = PageIteratorLevel.Para;
                    break;
                case "symbol":
                    myLevel = PageIteratorLevel.Symbol;
                    break;
                default:
                    myLevel = PageIteratorLevel.Word;
                    debug(WARNING, "tesseractFindAndClickText: Page Iterator Level not supported defaulted to WORD. Text sent to method: " + pageIteratorLevel);
                    break;
            }

            // Load the tesseract engine and try and find the text
            using (var engine = new TesseractEngine(dirName + @"\tessdata", "eng", EngineMode.Default))
            {
                using (var img = Pix.LoadFromFile(screenshotToLoad))
                {

                    var imgToOCR = img;
                    if (scale)
                    {
                        imgToOCR = imgToOCR.Scale(scaleValue, scaleValue);
                    }

                    if (invertImage)
                    {
                        imgToOCR = imgToOCR.ConvertRGBToGray();
                        imgToOCR = imgToOCR.Invert();
                    }

                    imgToOCR.Save(dirName + @"\cassTemplate_images\tempImages\tesseractFindAndClickText_processed.TIFF", Tesseract.ImageFormat.Tiff); // for debug

                    // process the image for text
                    using (var page = engine.Process(imgToOCR))
                    {
                        var text = page.GetText();
                        debug(TRACE, "tesseractFindAndClickText: All text found : \r\n" + text + "\r\n");

                        // based on the page iterator look at the text returned and try and find the expected value
                        using (var iter = page.GetIterator())
                        {
                            iter.Begin();
                            do
                            {
                                // iterate through all the found text
                                if (iter.TryGetBoundingBox(myLevel, out var rect))
                                {
                                    // 'rect' should containt the location of the text, 'curText' contains the actual text itself
                                    var curText = iter.GetText(myLevel);
                                    Debug.WriteLine(curText);

                                    // remove non standard characters
                                    if (!pageIteratorLevel.ToLower().Equals("block"))
                                    {
                                        curText = removeNonAlphaChars(curText);
                                    }
                                    else
                                    {
                                        curText = curText.Replace("\n", " ");
                                    }

                                    // remove leading and trailing spaces
                                    curText = curText.TrimStart(' ');
                                    curText = curText.TrimEnd(' ');
                                    curText = curText.ToLower();

                                    switch (textComparisonType.ToLower())
                                    {
                                        case "exact":
                                            if (curText.Equals(stringToFind.ToLower()))
                                            {
                                                textMatched = true;
                                            }
                                            break;
                                        case "contains":
                                            if (curText.Contains(stringToFind.ToLower()))
                                            {
                                                textMatched = true;
                                            }
                                            break;
                                        default:
                                            if (curText.Equals(stringToFind.ToLower()))
                                            {
                                                textMatched = true;
                                            }
                                            debug(WARNING, "tesseractFindAndClickText: Text Comparison Type not supported defaulted to EXACT match. Text sent to method: " + textComparisonType);
                                            break;
                                    }

                                    if (textMatched)
                                    {

                                        debug(TRACE, "tesseractFindAndClickText found text: " + stringToFind + " using comparison type: " + textComparisonType);
                                        if (!scale)
                                        {
                                            textLocations.X = inGetX + rect.X1;
                                            textLocations.Y = inGetY + rect.Y1;
                                            textLocations.W = rect.Width;
                                            textLocations.H = rect.Height;
                                            //clickX = inGetX + rect.X1 + (rect.Width / 2);
                                            //clickY = inGetY + rect.Y1 + (rect.Height / 2);

                                        }
                                        else
                                        {
                                            textLocations.X = inGetX + rect.X1 / scaleValue;
                                            textLocations.Y = inGetY + rect.Y1 / scaleValue;
                                            textLocations.W = rect.Width / scaleValue;
                                            textLocations.H = rect.Height / scaleValue;
                                            //clickX = inGetX + (rect.X1 / scaleValue) + (rect.Width / scaleValue);
                                            //clickY = inGetY + (rect.Y1 / scaleValue) + (rect.Height / scaleValue);

                                        }

                                        imgToOCR.Dispose();
                                        return textLocations;
                                    }
                                }

                            } while (iter.Next(myLevel));
                        }
                    }
                    imgToOCR.Dispose();
                }
            }

            return textLocations;
        }

        // Extract text from a screenshot and click the text on screen
        public bool tesseractFindAndClickText(string screenshotToLoad, string pageIteratorLevel, bool scale, int scaleValue, int inGetX, int inGetY, string clickType, string textComparisonType, string stringToFind, bool invertImage)
        {
            debug(TRACE, "tesseractFindAndClickText: start");

            int clickX = 0;
            int clickY = 0;
            bool textMatched = false;

            // set the page iterator, default to word if text unknown
            PageIteratorLevel myLevel;
            pageIteratorLevel = pageIteratorLevel.ToLower();
            switch (pageIteratorLevel)
            {
                case "word":
                    myLevel = PageIteratorLevel.Word;
                    break;
                case "textline":
                    myLevel = PageIteratorLevel.TextLine;
                    break;
                case "block":
                    myLevel = PageIteratorLevel.Block;
                    break;
                case "para":
                    myLevel = PageIteratorLevel.Para;
                    break;
                case "symbol":
                    myLevel = PageIteratorLevel.Symbol;
                    break;
                default:
                    myLevel = PageIteratorLevel.Word;
                    debug(WARNING, "tesseractFindAndClickText: Page Iterator Level not supported defaulted to WORD. Text sent to method: " + pageIteratorLevel);
                    break;
            }

            // Load the tesseract engine and try and find the text
            using (var engine = new TesseractEngine(dirName + @"\tessdata", "eng", EngineMode.Default))
            {
                using (var img = Pix.LoadFromFile(screenshotToLoad))
                {

                    var imgToOCR = img;
                    if (scale)
                    {
                        imgToOCR = imgToOCR.Scale(scaleValue, scaleValue);
                    }

                    if (invertImage)
                    {
                        imgToOCR = imgToOCR.ConvertRGBToGray();
                        imgToOCR = imgToOCR.Invert();
                    }

                    imgToOCR.Save(dirName + @"\cassTemplate_images\tempImages\tesseractFindAndClickText_processed.TIFF", Tesseract.ImageFormat.Tiff); // for debug

                    // process the image for text
                    using (var page = engine.Process(imgToOCR))
                    {
                        var text = page.GetText();
                        debug(TRACE, "tesseractFindAndClickText: All text found : \r\n" + text + "\r\n");

                        // based on the page iterator look at the text returned and try and find the expected value
                        using (var iter = page.GetIterator())
                        {
                            iter.Begin();
                            do
                            {
                                // iterate through all the found text
                                if (iter.TryGetBoundingBox(myLevel, out var rect))
                                {
                                    // 'rect' should containt the location of the text, 'curText' contains the actual text itself
                                    var curText = iter.GetText(myLevel);
                                    Debug.WriteLine(curText);

                                    // remove non standard characters
                                    if (!pageIteratorLevel.ToLower().Equals("block"))
                                    {
                                        curText = removeNonAlphaChars(curText);
                                    }
                                    else
                                    {
                                        curText = curText.Replace("\n", " ");
                                    }

                                    // remove leading and trailing spaces
                                    curText = curText.TrimStart(' ');
                                    curText = curText.TrimEnd(' ');
                                    curText = curText.ToLower();

                                    switch (textComparisonType.ToLower())
                                    {
                                        case "exact":
                                            if (curText.Equals(stringToFind.ToLower()))
                                            {
                                                textMatched = true;
                                            }
                                            break;
                                        case "contains":
                                            if (curText.Contains(stringToFind.ToLower()))
                                            {
                                                textMatched = true;
                                            }
                                            break;
                                        default:
                                            if (curText.Equals(stringToFind.ToLower()))
                                            {
                                                textMatched = true;
                                            }
                                            debug(WARNING, "tesseractFindAndClickText: Text Comparison Type not supported defaulted to EXACT match. Text sent to method: " + textComparisonType);
                                            break;
                                    }

                                    if (textMatched)
                                    {

                                        debug(TRACE, "tesseractFindAndClickText found text: " + stringToFind + " using comparison type: " + textComparisonType);
                                        if (!scale)
                                        {
                                            clickX = inGetX + rect.X1 + (rect.Width / 2);
                                            clickY = inGetY + rect.Y1 + (rect.Height / 2);

                                        }
                                        else
                                        {
                                            clickX = inGetX + (rect.X1 / scaleValue) + (rect.Width / scaleValue);
                                            clickY = inGetY + (rect.Y1 / scaleValue) + (rect.Height / scaleValue);

                                        }

                                        // Click
                                        if (clickType.ToLower().Equals("single"))
                                        {
                                            leftClick(clickX, clickY);
                                        }
                                        else if (clickType.ToLower().Equals("double"))
                                        {
                                            doubleClick(clickX, clickY);
                                        }
                                        else
                                        {
                                            debug(ERROR, "tesseractFindAndClickText: Unknown click type used in method : " + clickType);
                                        }
                                        imgToOCR.Dispose();
                                        return true;
                                    }
                                }

                            } while (iter.Next(myLevel));
                        }
                    }
                    imgToOCR.Dispose();
                }
            }

            return false;
        }

        // Capture a region of the screen for OCR
        public void tesseractCaptureScreen(int myGetW, int myGetH, int myGetX, int myGetY, string imageLocationFilename)
        {
            debug(TRACE, "tesseractCaptureScreen: start");

            bool pathExists = Directory.Exists(dirName + @"\cassTemplate_images\tempImages");
            if (!pathExists)
            {
                debug(ERROR, "tesseractCaptureScreen: temp image location does not exist. Check 'tempImages' folder is in 'cassTemplate_images' folder.");
            }

            using (Bitmap myScreenRegion = new Bitmap(myGetW, myGetH, PixelFormat.Format32bppArgb))
            {
                Graphics myGraphics = Graphics.FromImage(myScreenRegion as System.Drawing.Bitmap);
                myGraphics.CopyFromScreen(myGetX, myGetY, 0, 0, myScreenRegion.Size);
                myGraphics.Dispose();
                try
                {
                    myScreenRegion.Save(imageLocationFilename, System.Drawing.Imaging.ImageFormat.Tiff);
                }
                catch (Exception ex)
                {
                    debug(ERROR, "tesseractCaptureScreen: error saving temporary template file. Error : " + ex.Message);
                }
            }
            /*
                        // Capture the screen region
                        Bitmap myScreenRegion = new Bitmap(myGetW, myGetH, PixelFormat.Format32bppArgb);
                        Graphics myGraphics = Graphics.FromImage(myScreenRegion as System.Drawing.Bitmap);
                        myGraphics.CopyFromScreen(myGetX, myGetY, 0, 0, myScreenRegion.Size);

                        // save the screen out to TIFF format
                        try
                        {
                            myScreenRegion.Save(imageLocationFilename, System.Drawing.Imaging.ImageFormat.Tiff);
                        }
                        catch (Exception ex)
                        {
                            debug(ERROR, "tesseractCaptureScreen: error saving temporary template file. Error : " + ex.Message);
                        }*/
        }

        // Remove non alphanumric characters from a string
        public string removeNonAlphaChars(string inText)
        {
            //debug(TRACE, "removeNonAlphaChars: start"); // removing to make the log less chatty

            //Regex myRegEx = new Regex("[^a-zA-Z0-9 -()/]");
            Regex myRegEx = new Regex("[^a-zA-Z0-9 -]");
            string outText = myRegEx.Replace(inText, "");
            if (outText.Length > 1)
            {
                if (Char.IsLetterOrDigit(outText[0]) && Char.IsWhiteSpace(outText[1]))
                {
                    outText = outText.TrimStart(outText[0]).Trim();
                }
            }
            //debug(INFO, "removeNonAlphaChars: text after removing characters: " + outText);
            return outText;
        }

        // Clear the windows clipboard
        public void clearWindowsClipboard()
        {
            debug(TRACE, "clearWindowsClipboard: start");

            string ReturnValue = string.Empty;
            Thread STAThread = new Thread(
                delegate ()
                {
                    // Use a fully qualified name for Clipboard otherwise it
                    // will end up calling itself.
                    System.Windows.Forms.Clipboard.Clear();
                });
            STAThread.SetApartmentState(ApartmentState.STA);
            STAThread.Start();
            STAThread.Join();
            STAThread = null;

        }

        // Get what is in the current windows clipboard
        public string getTextWindowsClipboard()
        {
            debug(TRACE, "getTextWindowsClipboard: start");

            string ReturnValue = string.Empty;
            Thread STAThread = new Thread(
                delegate ()
                {
                    // Use a fully qualified name for Clipboard otherwise it
                    // will end up calling itself.
                    ReturnValue = System.Windows.Forms.Clipboard.GetText();
                });
            STAThread.SetApartmentState(ApartmentState.STA);
            STAThread.Start();
            STAThread.Join();
            STAThread = null;

            return ReturnValue;

        }

        // Get what is in the current windows clipboard
        public void setTextWindowsClipboard(string inText)
        {
            debug(TRACE, "setTextWindowsClipboard: start");

            string ReturnValue = string.Empty;
            Thread STAThread = new Thread(
                delegate ()
                {
                    // Use a fully qualified name for Clipboard otherwise it
                    // will end up calling itself.
                    System.Windows.Forms.Clipboard.SetText(inText);
                });
            STAThread.SetApartmentState(ApartmentState.STA);
            STAThread.Start();
            STAThread.Join();
            STAThread = null;

        }

        // Wait for a template image to appear in an area of the screen
        public void imageWaitArea(string templateLocation, int timeOutVal, bool moveMouse, bool checkTimeout, double imageConfidence, LocationValues areaLocation)
        {
            debug(TRACE, "imageWait: start. Template : " + templateLocation);

            if (!file_exists(templateLocation))
            {
                // Try and find the image in the images folder / sub folder
                string origtemplateLocation = templateLocation;
                templateLocation = returnImagePath(templateLocation);
                if (templateLocation == null || templateLocation == "")
                {
                    debug(ERROR, "myImageWait: File does not exist. File : " + origtemplateLocation);
                }
            }

            if (imageConfidence < 0 || imageConfidence > 1)
            {
                debug(ERROR, "myImageWait: Image confidence is < 0 or > 1. Value used was: " + imageConfidence);
            }

            Bitmap theDesktop = null;

            Mat myScreen = null;
            Mat imageToFind = null;
            Mat matchResult = null;

            double minVal;
            double maxVal = 0.0;

            Point minloc;
            Point maxloc;

            int counter = 0;

            // Area to capture
            int areaX = areaLocation.X;
            int areaY = areaLocation.Y;
            int areaH = areaLocation.H;
            int areaW = areaLocation.W;

            while (maxVal < imageConfidence && counter < timeOutVal)
            {
                // Capture desktop image and convert to MAT
                theDesktop = captureDesktopArea(areaW, areaH, areaX, areaY);
                myScreen = theDesktop.ToMat();

                // Load template from file
                imageToFind = new Mat(templateLocation);

                // Find the template on the desktop
                try
                {
                    matchResult = myScreen.MatchTemplate(imageToFind, TemplateMatchModes.CCoeffNormed);
                }
                catch (Exception e)
                {
                    //debug(1, String.Format("myImageWait: failed with error : {0}.", e.Message));
                    Debug.WriteLine("myImageWait: failed with error : {0}.", e.Message);
                }

                matchResult.MinMaxLoc(out minVal, out maxVal, out minloc, out maxloc);
                Debug.WriteLine("minval : " + minVal + " maxval : " + maxVal);

                Thread.Sleep(1000);

                if (moveMouse)
                {
                    randomMouseMove();
                }

                counter++;

                // clean up before next loop
                matchResult.Dispose();
                theDesktop.Dispose();
                myScreen.Dispose();
                imageToFind.Dispose();
            }

            Debug.WriteLine("maxval : " + maxVal);
            Debug.WriteLine("time taken : " + counter);

            theDesktop.Dispose();
            imageToFind.Dispose();
            myScreen.Dispose();
            matchResult.Dispose();

            if (checkTimeout)
            {
                if (maxVal < imageConfidence && counter >= timeOutVal)
                {
                    Debug.WriteLine("ImageWait TIMEOUT -  maxval : " + maxVal);
                    debug(ERROR, "imageWait error : image did not appear in specified time. Timeout value : " + timeOutVal + " Counter value : " + counter);
                }
            }
        }

        // Find and click a template image in a screen area
        public void imageFindClickArea(string templateLocation, int xOffSet, int yOffset, double imageConfidence, LocationValues areaLocation)
        {
            debug(TRACE, "imageFindClick: start. Template : " + templateLocation);

            if (!file_exists(templateLocation))
            {
                // Try and find the image in the images folder / sub folder
                string origtemplateLocation = templateLocation;
                templateLocation = returnImagePath(templateLocation);
                if (templateLocation == null || templateLocation == "")
                {
                    debug(ERROR, "myImageWait: File does not exist. File : " + origtemplateLocation);
                }
            }

            if (imageConfidence < 0 || imageConfidence > 1)
            {
                debug(ERROR, "myImageWait: Image confidence is < 0 or > 1. Value used was: " + imageConfidence);
            }

            Bitmap theDesktop = null;

            Mat myScreen = null;
            Mat imageToFind = null;
            Mat matchResult = null;

            double minVal;
            double maxVal;

            Point minloc;
            Point maxloc;

            // Area to capture
            int areaX = areaLocation.X;
            int areaY = areaLocation.Y;
            int areaH = areaLocation.H;
            int areaW = areaLocation.W;

            // Capture desktop image and convert to MAT
            theDesktop = captureDesktopArea(areaW, areaH, areaX, areaY);
            myScreen = theDesktop.ToMat();
            myScreen.SaveImage(dirName + @"\cassTemplate_images\tempImages\imageFindClickArea.PNG");

            // Load template from file
            imageToFind = new Mat(templateLocation);

            // Find the template on the desktop            
            matchResult = myScreen.MatchTemplate(imageToFind, TemplateMatchModes.CCoeffNormed);

            matchResult.MinMaxLoc(out minVal, out maxVal, out minloc, out maxloc);
            Debug.WriteLine("minval : " + minVal + " maxval : " + maxVal);

            // get centre of image.
            //if (maxVal >= imageConfidence)
            //{
            var toClickX = areaX + maxloc.X + imageToFind.Width / 2;
            var toClickY = areaY + maxloc.Y + imageToFind.Height / 2;

            leftClick(toClickX + xOffSet, toClickY + yOffset);
            //}
            //else
            //{
            //    debug(ERROR, "imageFindClick: Image confidence found was: "+ maxVal +". Expected confidence was: "+ imageConfidence +".");
            //}
            // clean up
            theDesktop.Dispose();
            imageToFind.Dispose();
            myScreen.Dispose();
            matchResult.Dispose();

        }

        // Capture an area of the desktop
        public Bitmap captureDesktopArea(LocationValues captureArea)
        {
            return captureDesktopArea(captureArea.W, captureArea.H, captureArea.X, captureArea.Y);
        }
        public Bitmap captureDesktopArea(int myGetW, int myGetH, int myGetX, int myGetY)
        {
            debug(TRACE, "captureDesktopArea: start");

            Rectangle rect = new Rectangle(myGetX, myGetY, myGetW, myGetH);

            Bitmap bmpScreenshot = new Bitmap(rect.Width, rect.Height, PixelFormat.Format24bppRgb);
            using (Graphics myGraphics = Graphics.FromImage(bmpScreenshot)) // ;
            {
                myGraphics.CopyFromScreen(rect.Left, rect.Top, 0, 0, bmpScreenshot.Size, CopyPixelOperation.SourceCopy);
            }

            return bmpScreenshot;
        }

        // Interact with fields
        // Searches a screen area for edit box text, allows to either:
        // - click to activate edit field
        // - search - find the search icon on an edit field
        public void interactWithFields(LocationValues areaToSearch, string fieldName, string actionToPerform)
        {
            interactWithFields(areaToSearch, fieldName, actionToPerform, 100);
        }

        public void interactWithFields(LocationValues areaToSearch, string fieldName, string actionToPerform, int interactOffset)
        {
            interactWithFields(areaToSearch, fieldName, actionToPerform, interactOffset, "contains");
        }

        public void interactWithFields(LocationValues areaToSearch, string fieldName, string actionToPerform, int interactOffset, string textComparisonType)
        {
            debug(TRACE, "interactWithFields: start");
            //260, 255, 1033, 873
            int inX = areaToSearch.X;
            int inY = areaToSearch.Y;
            int inW = areaToSearch.W;
            int inH = areaToSearch.H;

            bool textFound = false;
            tesseractCaptureScreen(inW, inH, inX, inY, dirName + @"\cassTemplate_images\tempImages\interactWithFields_capture.TIFF");
            textFound = tesseractFindAndInteractText(dirName + @"\cassTemplate_images\tempImages\interactWithFields_capture.TIFF", "textline", true, 2, inX, inY, "single", textComparisonType, fieldName, actionToPerform, false, interactOffset);
            if (!textFound)
            {
                // Try with no scaling
                textFound = tesseractFindAndInteractText(dirName + @"\cassTemplate_images\tempImages\interactWithFields_capture.TIFF", "textline", false, 2, inX, inY, "single", textComparisonType, fieldName, actionToPerform, false, interactOffset);

                if (!textFound)
                {
                    // try inverted text
                    textFound = tesseractFindAndInteractText(dirName + @"\cassTemplate_images\tempImages\interactWithFields_capture.TIFF", "textline", false, 2, inX, inY, "single", textComparisonType, fieldName, actionToPerform, true, interactOffset);

                    if (!textFound)
                    {
                        // try scaled, inverted text
                        textFound = tesseractFindAndInteractText(dirName + @"\cassTemplate_images\tempImages\interactWithFields_capture.TIFF", "textline", true, 2, inX, inY, "single", textComparisonType, fieldName, actionToPerform, true, interactOffset);
                        if (textFound)
                        {
                            debug(TRACE, "interactWithFields: text found with default scale 2, inverted image.");
                        }
                        else
                        {
                            textFound = tesseractFindAndInteractText(dirName + @"\cassTemplate_images\tempImages\interactWithFields_capture.TIFF", "textline", true, 3, inX, inY, "single", textComparisonType, fieldName, actionToPerform, true, interactOffset);
                            if (!textFound)
                            {
                                Debug.WriteLine("**** - text not found in interactwithfields!");
                                debug(WARNING, "interactWithFields: text not found after trying different image processing.");
                            }
                            else
                            {
                                debug(TRACE, "interactWithFields: text found with default scale 3, inverted image.");
                                
                            }
                        }
                    }
                    else
                    {
                        debug(TRACE, "interactWithFields: text found with no scale, inverted image.");
                    }

                }
                else
                {
                    debug(TRACE, "interactWithFields: text found with no scale.");
                }

            }
            else
            {
                debug(TRACE, "interactWithFields: text found with default scale 2.");
            }
        }

        // Utility method for interactWithFields
        // may expand in future
        public bool tesseractFindAndInteractText(string screenshotToLoad, string pageIteratorLevel, bool scale, int scaleValue, int inGetX, int inGetY, string clickType, string textComparisonType, string stringToFind, string actionType, bool invertImage)
        {
            return tesseractFindAndInteractText(screenshotToLoad, pageIteratorLevel, scale, scaleValue, inGetX, inGetY, clickType, textComparisonType, stringToFind, actionType, invertImage, 100);
        }
        public bool tesseractFindAndInteractText(string screenshotToLoad, string pageIteratorLevel, bool scale, int scaleValue, int inGetX, int inGetY, string clickType, string textComparisonType, string stringToFind, string actionType, bool invertImage, int interactOffset)
        {
            debug(TRACE, "tesseractFindAndClickText: start");

            //int clickX = 0;
            //int clickY = 0;
            bool textMatched = false;

            // set the page iterator, default to word if text unknown
            PageIteratorLevel myLevel;
            pageIteratorLevel = pageIteratorLevel.ToLower();
            switch (pageIteratorLevel)
            {
                case "word":
                    myLevel = PageIteratorLevel.Word;
                    break;
                case "textline":
                    myLevel = PageIteratorLevel.TextLine;
                    break;
                case "block":
                    myLevel = PageIteratorLevel.Block;
                    break;
                case "para":
                    myLevel = PageIteratorLevel.Para;
                    break;
                case "symbol":
                    myLevel = PageIteratorLevel.Symbol;
                    break;
                default:
                    myLevel = PageIteratorLevel.Word;
                    debug(WARNING, "tesseractFindAndClickText: Page Iterator Level not supported defaulted to WORD. Text sent to method: " + pageIteratorLevel);
                    break;
            }

            // Load the tesseract engine and try and find the text
            using (var engine = new TesseractEngine(dirName + @"\tessdata", "eng", EngineMode.Default))
            {
                using (var img = Pix.LoadFromFile(screenshotToLoad))
                {

                    var imgToOCR = img;
                    if (scale)
                    {
                        imgToOCR = imgToOCR.Scale(scaleValue, scaleValue);
                    }

                    if (invertImage)
                    {
                        imgToOCR = imgToOCR.ConvertRGBToGray();
                        imgToOCR = imgToOCR.Invert();
                    }

                    imgToOCR.Save(dirName + @"\cassTemplate_images\tempImages\interactWithFields_processed.TIFF", Tesseract.ImageFormat.Tiff); // for debug

                    // process the image for text
                    using (var page = engine.Process(imgToOCR))
                    {
                        //var text = page.GetText();
                        //debug(TRACE, "tesseractFindAndClickText: All text found : \r\n" + text + "\r\n");

                        // based on the page iterator look at the text returned and try and find the expected value
                        using (var iter = page.GetIterator())
                        {
                            iter.Begin();
                            do
                            {
                                // iterate through all the found text
                                if (iter.TryGetBoundingBox(myLevel, out var rect))
                                {
                                    // 'rect' should containt the location of the text, 'curText' contains the actual text itself
                                    var curText = iter.GetText(myLevel);
                                    Debug.WriteLine(curText);

                                    // remove non standard characters
                                    curText = removeNonAlphaChars(curText);

                                    // remove leading and trailing spaces
                                    curText = curText.TrimStart(' ');
                                    curText = curText.TrimEnd(' ');
                                    curText = curText.ToLower();

                                    switch (textComparisonType.ToLower())
                                    {
                                        case "exact":
                                            if (curText.Equals(stringToFind.ToLower()))
                                            {
                                                textMatched = true;
                                            }
                                            break;
                                        case "contains":
                                            if (curText.Contains(stringToFind.ToLower()))
                                            {
                                                textMatched = true;
                                            }
                                            break;
                                        default:
                                            if (curText.Equals(stringToFind.ToLower()))
                                            {
                                                textMatched = true;
                                            }
                                            debug(WARNING, "tesseractFindAndClickText: Text Comparison Type not supported defaulted to EXACT match. Text sent to method: " + textComparisonType);
                                            break;
                                    }

                                    if (textMatched)
                                    {
                                        LocationValues areaToSearch = new LocationValues();

                                        debug(TRACE, "tesseractFindAndClickText found text: " + stringToFind + " using comparison type: " + textComparisonType);
                                        if (!scale)
                                        {
                                            areaToSearch.X = inGetX + rect.X1;
                                            areaToSearch.Y = inGetY + rect.Y1 - 10;
                                            areaToSearch.W = rect.Width + interactOffset; // 200
                                            areaToSearch.H = rect.Height + 20;
                                        }
                                        else
                                        {
                                            areaToSearch.X = inGetX + (rect.X1 / scaleValue);
                                            areaToSearch.Y = inGetY + (rect.Y1 / scaleValue) - 10;
                                            areaToSearch.W = (rect.Width / scaleValue) + interactOffset; // 200
                                            areaToSearch.H = (rect.Height / scaleValue) + 20;
                                        }

                                        if (actionType.ToLower().Equals("search"))
                                        {
                                            imageFindClickArea(dirName + @"\cassTemplate_images\genericSearchIcon.PNG", 0, 0, 0.95, areaToSearch);
                                        }
                                        if (actionType.ToLower().Equals("edit"))
                                        {
                                            imageFindClickArea(dirName + @"\cassTemplate_images\genericEditFieldStart.PNG", 0, 0, 0.95, areaToSearch);
                                        }
                                        if (actionType.ToLower().Equals("gettext"))
                                        {
                                            imageFindClickArea(dirName + @"\cassTemplate_images\genericEditFieldStart.PNG", 0, 0, 0.95, areaToSearch);
                                            getTextFromTextBox();
                                        }
                                        // eperimental
                                        if (actionType.ToLower().Equals("ocrtext"))
                                        {
                                            imageFindClickArea(dirName + @"\cassTemplate_images\genericEditFieldStart.PNG", 0, 0, 0.95, areaToSearch);
                                            string ocrText = findandGetText(areaToSearch, true);
                                            setTextWindowsClipboard(ocrText);
                                        }
                                        imgToOCR.Dispose();
                                        return true;
                                    }
                                }

                            } while (iter.Next(myLevel));
                        }
                    }
                    imgToOCR.Dispose();
                }
            }

            return false;
        }

        public string returnImagePath(string imageName)
        {
            string[] imageFile = null;
            var foundPath = "";

            string basePath = dirName + @"\cassTemplate_images";

            // Check if full path sent to method - split out and get imagename
            if (imageName.Contains("\\"))
            {
                string[] tempArray = imageName.Split(new char[] { '\\' });
                imageName = tempArray[tempArray.Length - 1];
            }

            // Search base path
            imageFile = Directory.GetFiles(basePath, "*.png");
            foundPath = imageFile.FirstOrDefault(x => Path.GetFileName(x) == imageName);

            // Check the sub folders
            if (foundPath == null)
            {
                string[] subFolders = Directory.GetDirectories(basePath);
                foreach (string subFolder in subFolders)
                {
                    imageFile = Directory.GetFiles(subFolder, "*.png");
                    foundPath = imageFile.FirstOrDefault(x => Path.GetFileName(x) == imageName);
                    if (foundPath != null)
                    {
                        break;
                    }
                }
            }

            return foundPath;
        }

        // TESTING JUNK below here... 
        public void ocrTESTing_xpr()
        {

            var testImagePath = dirName + @"\cassTemplate_images\tempImages\ocrTESTing_xpr.TIFF";

            using (var engine = new TesseractEngine(dirName + @"\tessdata", "eng", EngineMode.Default))
            {
                using (var img = Pix.LoadFromFile(testImagePath))
                {
                    Pix img2 = img.Scale(3, 3);

                    //Pix img2 = img.ConvertRGBToGray();
                    //Pix img2 = img.Invert();
                    //img2 = img2.Scale(25,25);

                    using (var page = engine.Process(img2))
                    {
                        var text = page.GetText();
                        Debug.WriteLine("Mean confidence: {0}", page.GetMeanConfidence());

                        Debug.WriteLine("Text (GetText): \r\n{0}", text);
                        Debug.WriteLine("Text (iterator):");
                        using (var iter = page.GetIterator())
                        {
                            iter.Begin();

                            do
                            {
                                do
                                {
                                    do
                                    {
                                        do
                                        {
                                            if (iter.IsAtBeginningOf(PageIteratorLevel.Block))
                                            {
                                                Debug.WriteLine("<BLOCK>");
                                            }

                                            Debug.Write(iter.GetText(PageIteratorLevel.Word));
                                            Debug.Write(" ");

                                            if (iter.IsAtFinalOf(PageIteratorLevel.TextLine, PageIteratorLevel.Word))
                                            {
                                                Debug.WriteLine("");
                                            }
                                        } while (iter.Next(PageIteratorLevel.TextLine, PageIteratorLevel.Word));

                                        if (iter.IsAtFinalOf(PageIteratorLevel.Para, PageIteratorLevel.TextLine))
                                        {
                                            Debug.WriteLine("");
                                        }
                                    } while (iter.Next(PageIteratorLevel.Para, PageIteratorLevel.TextLine));
                                } while (iter.Next(PageIteratorLevel.Block, PageIteratorLevel.Para));
                            } while (iter.Next(PageIteratorLevel.Block));
                        }
                    }
                }
            }
        }

        public void imageFindClick_old(string templateLocation, int xOffSet, int yOffset, double imageConfidence)
        {
            debug(TRACE, "imageFindClick: start. Template : " + templateLocation);

            if (!file_exists(templateLocation))
            {
                // Try and find the image in the images folder / sub folder
                string origtemplateLocation = templateLocation;
                templateLocation = returnImagePath(templateLocation);
                if (templateLocation == null || templateLocation == "")
                {
                    debug(ERROR, "myImageWait: File does not exist. File : " + origtemplateLocation);
                }
            }

            if (imageConfidence < 0 || imageConfidence > 1)
            {
                debug(ERROR, "myImageWait: Image confidence is < 0 or > 1. Value used was: " + imageConfidence);
            }

            Bitmap theDesktop = null;

            Mat myScreen = null;
            Mat imageToFind = null;
            Mat matchResult = null;

            double minVal;
            double maxVal;

            Point minloc;
            Point maxloc;

            // Capture desktop image and convert to MAT
            theDesktop = captureDesktop();
            myScreen = theDesktop.ToMat();

            // Load template from file
            imageToFind = new Mat(templateLocation);

            // Find the template on the desktop            
            matchResult = myScreen.MatchTemplate(imageToFind, TemplateMatchModes.CCoeffNormed);

            matchResult.MinMaxLoc(out minVal, out maxVal, out minloc, out maxloc);
            Debug.WriteLine("minval : " + minVal + " maxval : " + maxVal);


            // get centre of image.
            //if (maxVal >= imageConfidence)
            //{
            var toClickX = maxloc.X + imageToFind.Width / 2;
            var toClickY = maxloc.Y + imageToFind.Height / 2;

            leftClick(toClickX + xOffSet, toClickY + yOffset);
            //}
            //else
            //{
            //    debug(ERROR, "imageFindClick: Image confidence found was: "+ maxVal +". Expected confidence was: "+ imageConfidence +".");
            //}
            // clean up
            theDesktop.Dispose();
            imageToFind.Dispose();
            imageToFind.Dispose();
            myScreen.Dispose();
            matchResult.Dispose();
        }

        public void imageFindClick_xpr(string templateLocation, int xOffSet, int yOffset, double imageConfidence)
        {
            debug(TRACE, "imageFindClick: start. Template : " + templateLocation);

            int timeoutValue = 0;

            // Check template file exists
            if (!file_exists(templateLocation))
            {
                // Try and find the image in the images folder / sub folder
                string origtemplateLocation = templateLocation;
                templateLocation = returnImagePath(templateLocation);
                if (templateLocation == null || templateLocation == "")
                {
                    debug(ERROR, "myImageWait: File does not exist. File : " + origtemplateLocation);
                }
            }

            // Check the range of the incomming image confidence
            if (imageConfidence < 0 || imageConfidence > 1)
            {
                debug(ERROR, "myImageWait: Image confidence is < 0 or > 1. Value used was: " + imageConfidence);
            }

            // Capture the primary desktop and converet to Mat format
            Bitmap theDesktop = captureDesktop();
            Mat theDesktopMat = theDesktop.ToMat();

            // Find the point to click
            using (Mat refMat = theDesktopMat)
            using (Mat tplMat = new Mat(templateLocation))
            using (Mat res = new Mat(refMat.Rows - tplMat.Rows + 1, refMat.Cols - tplMat.Cols + 1, MatType.CV_32FC1))
            {
                //Convert input images to gray
                Mat gref = refMat.CvtColor(ColorConversionCodes.BGR2GRAY);
                Mat gtpl = tplMat.CvtColor(ColorConversionCodes.BGR2GRAY);

                Cv2.MatchTemplate(gref, gtpl, res, TemplateMatchModes.CCoeffNormed);
                Cv2.Threshold(res, res, imageConfidence, 1.0, ThresholdTypes.Tozero);

                while (true && timeoutValue <= 30)
                {
                    double minval, maxval;
                    Point minloc, maxloc;

                    Cv2.MinMaxLoc(res, out minval, out maxval, out minloc, out maxloc);
                    Debug.WriteLine("minval : " + minval + " maxval : " + maxval);

                    if (maxval >= imageConfidence)
                    {

                        // Get centre of image to click
                        var toClickX = maxloc.X + tplMat.Width / 2;
                        var toClickY = maxloc.Y + tplMat.Height / 2;

                        // Add any offset
                        toClickX = toClickX + xOffSet;
                        toClickY = toClickY + yOffset;

                        // Click the image
                        leftClick(toClickX, toClickY);

                        //Setup the rectangle to draw
                        Rect r = new Rect(new Point(maxloc.X, maxloc.Y), new Size(tplMat.Width, tplMat.Height));

                        //Draw a rectangle of the matching area
                        Cv2.Rectangle(refMat, r, Scalar.LimeGreen, 2);

                        //Fill in the res Mat so you don't find the same area again in the MinMaxLoc
                        Rect outRect;
                        Cv2.FloodFill(res, maxloc, new Scalar(0), out outRect, new Scalar(0.1), new Scalar(1.0));

                    }
                    else
                    {
                        refMat.SaveImage(dirName + @"\cassTemplate_images\tempImages\bob1.png");
                        break;
                    }
                    timeoutValue++;
                }

                //refMat.SaveImage(dirName + @"\cassTemplate_images\temp7.PNG");
                theDesktop.Dispose();
                theDesktopMat.Dispose();
            }
        }

        public void imageVanishArea(string templateLocation, int timeOutVal, double imageConfidence, LocationValues areaLocation)
        {
            debug(TRACE, "imageVanish: start. Template : " + templateLocation);

            if (!file_exists(templateLocation))
            {
                // Try and find the image in the images folder / sub folder
                string origtemplateLocation = templateLocation;
                templateLocation = returnImagePath(templateLocation);
                if (templateLocation == null || templateLocation == "")
                {
                    debug(ERROR, "myImageWait: File does not exist. File : " + origtemplateLocation);
                }
            }

            if (imageConfidence < 0 || imageConfidence > 1)
            {
                debug(ERROR, "myImageWait: Image confidence is < 0 or > 1. Value used was: " + imageConfidence);
            }

            Bitmap theDesktop = null;

            Mat myScreen = null;
            Mat imageToFind = null;
            Mat matchResult = null;

            double minVal;
            double maxVal = 1.0;

            Point minloc;
            Point maxloc;

            int counter = 0;

            // Area to capture
            int areaX = areaLocation.X;
            int areaY = areaLocation.Y;
            int areaH = areaLocation.H;
            int areaW = areaLocation.W;

            while (maxVal > imageConfidence && counter < timeOutVal)
            {
                // Capture desktop image and convert to MAT
                theDesktop = captureDesktopArea(areaW, areaH, areaX, areaY);
                myScreen = theDesktop.ToMat();

                // Load template from file
                imageToFind = new Mat(templateLocation);

                // Find the template on the desktop            
                matchResult = myScreen.MatchTemplate(imageToFind, TemplateMatchModes.CCoeffNormed);

                matchResult.MinMaxLoc(out minVal, out maxVal, out minloc, out maxloc);
                Debug.WriteLine("minval : " + minVal + " maxval : " + maxVal);

                Thread.Sleep(1000);
                counter++;

                // clean up before next loop
                matchResult.Dispose();
                theDesktop.Dispose();
                myScreen.Dispose();
                imageToFind.Dispose();
            }

            Debug.WriteLine("maxval : " + maxVal);
            Debug.WriteLine("time taken : " + counter);

            theDesktop.Dispose();
            imageToFind.Dispose();
            myScreen.Dispose();
            matchResult.Dispose();

            if (maxVal < imageConfidence && counter >= timeOutVal)
            {
                debug(ERROR, "myImageVanish error : image did not vanish in specified time. Timeout value : " + timeOutVal);
            }

            debug(TRACE, "imageVanish: time for image vanish : " + counter);
        }
    }
}



