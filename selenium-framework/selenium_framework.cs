using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.Runtime.CompilerServices;
using System.Collections.ObjectModel;
using OpenQA.Selenium.Interactions;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Text;
using System.Diagnostics;
using NUnit.Framework;
using System.Reflection;
using System.Linq;

namespace selenium_framework
{
    public partial class Selenium_Framework
    {
        /*
 
        The key features of the TestPro Selenium Framework are:

        1.	Reliability - the elements and controls that make up an applications web presentation do not appear instantly.
            There is also no key flag or status that tells you when the page is complete as the presentation is being managed
            dynamically by JavaScript. This gives pages a responsive animated look and feel. Searching for, and retrying access
            to slow to download and render controls is built into the lowest level of the Selenium Frameworks management of the
            DOM and dramatically improves test reliability.

        2.	Error Handling - the use of most Selenium APIs "should" be used within a try/catch block in order to manage errors,
            trigger retries or invoke special workarounds. The framework implements this consistently.

        3.	Viewport Management - before a control such as a button can be clicked the control must be brought into the rendering
            viewport of the browser. This is a fundamental Selenium requirement and applies especially to long pages or pages rendered
            on small devices. As a human you would scroll to a control, as a test you cannot. The framework does the job of bringing
            the control into the viewport when required.

        4.	Logging - the Selenium Framework includes a comprehensive logging mechanism of all framework events that is invaluable
            for test diagnostics.

        5.	Selenium Extensions - Selenium imposes some restrictions, such as only being able to interact with controls that are in
            a displayed state. This is not always the case as modern animated application GUIs can have screen elements that are unseen
            and yet people interact with them with the mouse based on other visual cues. As an extension to Selenium these controls are
            interacted with via DOM JavaScript injection. It allows the automation to interact with elements that Selenium cannot. This
            capability is built into the fabric of the Selenium Framework.

        6.	No Positional Data - the Selenium Framework makes it easy to identify the controls and elements needed to drive the test.
            No positional information is ever used.

        7.	No Waiting - Often testers pepper their scripts with wait or sleep commands to take account of timing "issues". This is
            rarely the case with the Selenium Framework as there is automatic “sense and retry” built into its lowest layers. Tests are
            easier to write and debug when they work without needing to worry about technical timing issues.

        8.	Locators - Selenium locators are one of the more confusing aspects of using Selenium to identify DOM elements. The Selenium
            Framework mechanism does not require the use of locators at all. Selenium element identification mechanisms are handled
            internally by translating all DOM location logic into XPath. The XPath spec that is generated and used is available in the
            log and can be pasted into the Chrome Dev console for diagnostic use.

        */

        // Selenium objects
        //public IWebDriver driver = null;

        // The driver is dynamic so it can be any selenium driver that supports compatible methods

        // Dynamic version
        //public dynamic driver = null;

        // Hard class version
        public static IWebDriver driver = null;

        public static string envBrowser = "chrome"; // "chrome", "internet explorer", "firefox", "edge"
        public static string envOS = "windows"; // "android", "ios", "mac", "windows"
        public static string browserVersion;

        //private static ChromeOptions options = null;

        //private static string driverPath = AppDomain.CurrentDomain.BaseDirectory;
        private static string driverPath = "";

        private static IJavaScriptExecutor js = null;
        private static IAlert alert = null;

        private static By spec = null;    // Current By spec for current object - used in wait
        private static string frameId = "";

        private static int retry = 30;

        private string screenshotLocation = @"C:\Screenshots";

        public static string dirName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location.Replace("\\bin\\Debug\\net48", string.Empty));
        public static string dataTableLocation = dirName + @"\dataTables";

        //public static string userPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        public static string chromeDownloadPath = cGetEnvVars_WinExp.GetPath("{374DE290-123F-4565-9164-39C4925E467B}", cGetEnvVars_WinExp.KnownFolderFlags.DontVerify, false);
        //public static string chromeDownloadPath = Path.Combine(userPath, "Downloads");

        // Debug levels

        public static int NONE = 0;
        public static int ERROR = 1;
        public static int WARNING = 2;
        public static int INFO = 3;
        public static int TRACE = 4;
        public static int VERBOSE = 5;

        private static int debuglevel = TRACE;

        //
        // killProcess - to kill old browsers
        // pname = iexplore, chrome
        //

        public void killProcess(string pname)
        {
            Process[] procs = Process.GetProcessesByName(pname);
            foreach (Process proc in procs)
            {
                debug(TRACE, "killProcess: " + proc.ToString() + "[" + proc.Id.ToString() + "]");
                try
                {
                    if (!proc.HasExited)
                    {
                        proc.Kill();
                    }
                }
                catch (Exception e)
                {
                    debug(TRACE, string.Format("killProcess: Unable to kill process {0}, exception: {1}", proc.ToString(), e.Message));
                }
            }
        } // killProcess

        //
        // Browsers
        // Notes for internet explorer
        // * Set ALL FOUR "Internet Options / Security" zones to "Enable Protected Mode"
        // * Set "Local Intranet / Sites / Advanced" to a list of servers to be accessed
        //   as a workaround to network security issues
        //

        //
        // set brower environment and OS
        //

        public void setBrowser(string id)
        {
            envBrowser = id;
        }

        public void setOS(string id)
        {
            envOS = id;
        }

        //
        // setDriverPath
        //
        public void setDriverPath(string path)
        {
            // Check if running on GitHub test runner and set path to test runner selenium
            if (DetectVirtualMachine())
            {
                path = @"C:\SeleniumWebDrivers\ChromeDriver";
                debug(INFO, "setDriverPath: virtual machine detected using path: '" + path + "'");
            }
            driverPath = path;
        }

        public string getDriverPath()
        {
            return driverPath;
        }

        public void maximize()
        {
            driver.Manage().Window.Maximize();
            //var mmm = driver.GetType().GetMethod("Manage").Invoke(driver, null);
            //mmm.Window.Maximize();
        }

        //
        // setScreenShotLocation
        //
        public void setScreenShotLocation(string ssloc)
        {
            screenshotLocation = ssloc;
        }

        public string getScreenShotLocation()
        {
            return screenshotLocation;
        }

        //
        // Other support methods
        //
        public void refresh()
        {
            driver.Navigate().Refresh();
            debug(TRACE, "refresh");
        }

        public void gotoUrl(string url, string domain, string username, string password)
        {
            string server = "";
            string version;

            // From chrome 59 onwards use two step HTTP Auth
            version = browserVersion.Substring(0, 2);

            // Isolate server name
            Regex regex = new Regex("(http://.*)/");
            Match match = regex.Match(url);

            if (match.Success)
                server = match.Value;
            else
                debug(ERROR, "gotoUrl: no url match: " + url);

            // Choose browser version strategy
            debug(TRACE, "gotoUrl: browser=" + envBrowser + version);
            switch (envBrowser)
            {
                case "chrome":

                    //password = password.Replace(@"#", @"%23");
                    server = server.Replace(@"//", @"//" + domain + @"%5C" + username + ":" + password + "@");

                    // STEP 1 - dummy HTTP Auth to base server to est session
                    gotoUrl(server);

                    // STEP 2 continue without Auth
                    gotoUrl(url);
                    break;

                case "internet explorer":
                    switch (version)
                    {
                        case "11":
                            gotoUrl(url);
                            sleep(1);
                            // Handle the IE Auth popup
                            alert = driver.SwitchTo().Alert();
                            //alert.SetAuthenticationCredentials(domain + @"\" + username, password);
                            alert.Accept();
                            break;

                        default:
                            debug(ERROR, "gotoUrl: Only Internet Explorer 11 is supported");
                            break;
                    }
                    break;

                case "firefox":
                    gotoUrl(url);
                    sleep(1);
                    alert = driver.SwitchTo().Alert();
                    alert.SendKeys(domain + @"\" + username + Keys.Tab + password);
                    alert.Accept();
                    break;

                default:
                    // Default SINGLE STEP - Add domain, username and password
                    //password = password.Replace(@"#", @"%23");
                    url = url.Replace(@"//", @"//" + domain + @"%5C" + username + ":" + password + "@");
                    gotoUrl(url);
                    break;

            }
        } // gotoUrl

        public void gotoUrl(string url)
        {
            debug(TRACE, "gotoUrl: url=" + url);
            try
            {
                driver.Navigate().GoToUrl(url);

                //var nav = driver.GetType().GetMethod("Navigate").Invoke(driver, null);
                //nav.GetType().InvokeMember("GoToUrl", BindingFlags.InvokeMethod, Type.DefaultBinder,
                //    nav, new object[] {url});
            }
            catch (WebDriverException e)
            {
                debug(ERROR, "gotoUrl: WebDriverException = " + e);
            }
            catch (Exception f)
            {
                debug(ERROR, "gotoUrl: Exception = " + f);
            }

            //debug(INFO, "gotoUrl: title: '" + js.ExecuteScript("return document.title") + "'");
        } // gotoUrl

        // quit the browser and close all open windows
        public void quit()
        {
            debug(TRACE, "quit");
            driver.Quit();
            driver = null;
        } // quit

        // close the currently open window/tab
        public void close()
        {
            debug(TRACE, "close");
            driver.Close();
        } // close

        public IWebDriver getDriver()
        {
            return driver;
        } // getDriver

        public void sleep(int t)
        {
            System.Threading.Thread.Sleep(t * 1000);
            debug(TRACE, "sleep: " + t + " seconds");
        } // sleep

        //
        // Debug methods
        //

        public void setDebugLevel(int level)
        {
            debug(TRACE, "setDebugLevel: " + level);
            debuglevel = level;
        } // setDebugLevel

        public int getDebugLevel()
        {
            return debuglevel;
        } // getDebugLevel

        public void debug(int level, string text, [CallerLineNumber] int lineNumber = 0, [CallerMemberName] string caller = null)
        {
            // Call the console writeline if current filter level is
            // same or greater than message level

            string date;
            string screenshotdriverfilename;
            string screenshotfilename;

            string[] lvl = new string[] { "None", "Error", "Warning", "Info", "Trace" };

            if (NONE < level && level <= TRACE)
            {
                if (debuglevel >= level)
                {
                    if (level == ERROR || level == WARNING)
                    {
                        Console.WriteLine(lvl[level] + ":[" + caller + ":" + lineNumber + "]: " + text);

                        date = DateTime.Now.ToString("yyyyMMddHHmmss");
                        string longName = GetType().ToString();
                        string[] scriptNameSplit = longName.Split('.');
                        string scriptName = scriptNameSplit[scriptNameSplit.Length-1];

                        screenshotdriverfilename = Path.Combine(screenshotLocation, lvl[level] + "_" + scriptName + "_" + date + "_driver.png");
                        screenshotfilename = Path.Combine(screenshotLocation, lvl[level] + "_" + scriptName + "_" + date + ".png");

                        //takeScreenshot(screenshotfilename);
                        takeScreenshotdriver(screenshotdriverfilename); // webdriver screenshot - browser render area only
                        takeScreenshot(screenshotfilename); // screenshot the primary window - see any messages outside of the browser

                        if (level == ERROR)
                        {
                            if (driver != null && !Debugger.IsAttached)
                            {
                                quit(); // Collapse browser and driver
                            }
                            Assert.Fail();

                        }
                        else
                        {
                            try
                            {
                                Assert.Inconclusive();
                            }
                            catch (Exception)
                            {
                                Console.WriteLine("debug: WARNING - Assert.Inconclusive() failed");
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine(lvl[level] + ": " + text);
                    }
                }
            }
            else
            {
                Console.WriteLine("debug: bad level [" + level + "][" + text + "]");
            }

        } // debug

        //
        // getRetry
        //
        public int getRetry()
        {
            return retry;
        } // getRetry

        //
        // takeScreenshot
        // This is a third party screenshot tool. Please deploy 'screenshot.exe'
        // into the Selenium driver folder
        //

        /*
        public void xtakeScreenshot(string file)
        {
            //Console.WriteLine("debug: takeScreenshot: '" + file + "'");
            bool screenExeExists = File.Exists(Path.Combine(driverPath, "screenshot.exe"));
            if (!screenExeExists)
            {
                Console.WriteLine("debug: takeScreenshot: No 'screenshot.exe' in path '" + driverPath + "'");
            }
            else
            {
                Process process = new Process();
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.UseShellExecute = false;
                startInfo.CreateNoWindow = true;
                startInfo.FileName = Path.Combine(driverPath, "screenshot.exe");
            
                startInfo.Arguments = file;
                process.StartInfo = startInfo;
                process.Start();
                process.WaitForExit();
            }
        } // takeScreenshot
        */

        public void takeScreenshot(string file)
        {
            System.Drawing.Bitmap theDesktop = captureDesktop();
            theDesktop.Save(file, System.Drawing.Imaging.ImageFormat.Png);
        }
        
        public void takeScreenshotdriver(string file)
        {
            //Console.WriteLine("debug: takeScreenshotdriver Start");
            Boolean hasException = false;

            if (driver == null)
                Console.WriteLine("debug: takeScreenshotdriver: driver is null");
            else
            {
                ITakesScreenshot ssdriver = driver as ITakesScreenshot;
                Screenshot screenshot = null;
                try
                {
                    screenshot = ssdriver.GetScreenshot();
                }
                catch (UnhandledAlertException ex)
                {
                    hasException = true;
                    Console.WriteLine("debug: takeScreenshotdriver UnhandledAlertException Error: " + ex);
                }
                catch (NoSuchWindowException ex)
                {
                    hasException = true;
                    Console.WriteLine("debug: takeScreenshotdriver NoSuchWindowException Error: " + ex);
                }
                catch (WebDriverException ex)
                {
                    hasException = true;
                    Console.WriteLine("debug: takeScreenshotdriver WebDriverException Error: " + ex);
                }

                if (!hasException)
                {
                    //Console.WriteLine("debug: takeScreenshotdriver: " + file);
                    screenshot.SaveAsFile(file);
                }
            }
        } // takeScreenshotdriver
        

        //
        // Pop Up windows
        //

        public int getWindowCount()
        {
            int count = driver.WindowHandles.Count;

            debug(TRACE, "getWindowCount: " + count);
            return count;
        } // getWindowCount

        //
        // syncWindowCount
        //

        public void syncWindowCount(int win)
        {
            int count;

            if (!(win > 0))
                debug(ERROR, "syncWindowCount: must be greater that 0");

            count = 0;
            while (count < retry)
            {
                if (driver.WindowHandles.Count == win)
                    break;
                else
                {
                    sleep(2);
                    debug(TRACE, "syncWindowCount[" + win + "]: count=" + driver.WindowHandles.Count);
                    count++;
                }
            }

            if (count == retry)
                debug(ERROR, "syncWindowCount[" + win + "]: count not met");
            else
                debug(TRACE, "syncWindowCount[" + win + "]: count=" + driver.WindowHandles.Count);

        } // syncWindowCount

        public void popupWindow()
        {
            popupWindow(0);
        } // popupWindow

        //
        // Switch to popup window - no waiting to sync title
        //

        public void popupWindow(int index)
        {
            ReadOnlyCollection<string> handles;
            int count;

            count = 0;
            handles = driver.WindowHandles;
            while (count < retry)
            {
                if (index < handles.Count)
                {
                    driver.SwitchTo().Window(handles[index]);
                    debug(TRACE, "popupWindow: [" + index + "]: '" + driver.Title);
                    break;
                }
                else
                {
                    handles = driver.WindowHandles;
                    debug(TRACE, "popupWindow: try: " + count);
                    count++;
                    sleep(2);
                }
            }

            if (count == retry)
                debug(ERROR, "popupWindow: [" + index + "]: out of range - max: " + (handles.Count - 1));

        } // popupWindow

        //
        // Switch to pop window and wait for window title to sync
        //

        public void popupWindow(string title)
        {
            ReadOnlyCollection<string> handles;

            int count;
            bool found = false;
            string hname;

            handles = driver.WindowHandles;
            for (count = 0; count < retry && !found; count++)
            {
                // Scan the handles - try to match name
                foreach (string handle in handles)
                {
                    hname = driver.SwitchTo().Window(handle).Title;
                    if (hname.Contains(title))
                    {
                        found = true;
                        break; // look no more
                    }
                }

                if (!found)
                {
                    debug(TRACE, "popupWindow: try " + count);
                    sleep(2);
                    handles = driver.WindowHandles; // refresh handles list
                }
            }

            if (!found)
                debug(ERROR, "popupWindow: '" + title + "' not found");

            debug(TRACE, "popupWindow: switched to '" + title + "'");

        } // popupWindow

        //
        // Switch frames
        //

        public string getFrameId()
        {
            return frameId;
        }

        public void switchTo()
        {
            try
            {
                driver.SwitchTo().DefaultContent();
            }
            catch (UnhandledAlertException ex)
            {
                debug(ERROR, "switchTo: <DefaultContent> failed : " + ex.Message);
            }
            frameId = "";
            debug(TRACE, "switchTo: DefaultContent");

        } // switchTo

        public void switchTo(int frameNo)
        {
            // Frames here are addressed by #index from zero within the page

            driver.SwitchTo().Frame(frameNo);
        }

        public void switchTo(string frameid)
        {
            int i = 0;

            debug(TRACE, "switchTo (str): '" + frameid + "'");

            if (frameid.Length == 0)
            {
                switchTo();
            }
            else
            {
                while (i < retry)
                {
                    try
                    {
                        driver.SwitchTo().Frame(frameid);
                        // Set global
                        frameId = frameid;
                    }
                    catch (NoSuchFrameException e)
                    {
                        debug(TRACE, "switchTo try: " + i + " '" + e.Message + "'");
                        i++;
                        sleep(1);
                        continue;
                    }
                    catch (UnhandledAlertException e)
                    {
                        debug(TRACE, "UnhandledAlertException: " + i + " '" + e.Message + "'");
                        i++;
                        sleep(1);
                        continue;
                    }
                    break;
                }

                if (i == retry)
                    debug(ERROR, "switchTo (str): frame not found:'" + frameid + "'");
            }
        } // switchTo

        public void switchTo(IWebElement frame)
        {
            string frameId;

            frameId = frame.GetAttribute("id");
            if (frameId.Length == 0)
            {
                frameId = frame.GetAttribute("name");
            }
            switchTo(frameId); // (str)

        } // switchTo

        //
        // findElements - local version with exception handling
        //

        public ReadOnlyCollection<IWebElement> findElements(IWebElement obj, By spec)
        {
            int i = 0;
            ReadOnlyCollection<IWebElement> elts = null;

            while (i < 2)
            {
                try
                {
                    if (obj == null)
                    {
                        //debug(TRACE, "findElements: null");
                        elts = driver.FindElements(spec);
                    }
                    else
                    {
                        //debug(TRACE, "findElements: obj");
                        elts = obj.FindElements(spec);
                    }
                }
                catch (StaleElementReferenceException e)
                {
                    debug(TRACE, "findElements: StaleElementReferenceException try: " + i + " '" + e.Message + "'");
                    i++;
                    sleep(1);
                    continue;
                }
                catch (UnhandledAlertException f)
                {
                    debug(ERROR, "findElements: UnhandledAlertException try: " + i + " '" + f.Message + "'");
                    //i++;
                    //sleep(1);
                    //continue;
                }
                catch (WebDriverException g)
                {
                    debug(ERROR, "findElements: WebDriverException try: " + i + " '" + g.Message + "'");
                    //i++;
                    //sleep(1);
                    //continue;
                }

                break;
            }

            return elts;
        }

        //
        // getElements
        // This is where the transltion to XPATH occurs
        // A different retry spec is optional
        //

        public ReadOnlyCollection<IWebElement> getElements(string XPath)
        {
            return getElements(null, XPath);
        }

        public ReadOnlyCollection<IWebElement> getElements(IWebElement obj, string XPath)
        {
            return getElements(null, XPath, 0);
        }

        public ReadOnlyCollection<IWebElement> getElements(IWebElement obj, string XPath, int myRetry)
        {
            ReadOnlyCollection<IWebElement> elts = null;
            int timeout = 0;

            if (myRetry == 0)
            {
                myRetry = retry;
            }

            // Custom XPath query string
            while (timeout < myRetry)
            {
                elts = findElements(obj, By.XPath(XPath));
                if (elts.Count > 0)
                {
                    debug(TRACE, "getElements: found xpath='" + XPath + "':" + elts.Count);
                    break;
                }
                else
                    debug(TRACE, "getElements: failed xpath='" + XPath + "'");

                sleep(1);
                timeout++;
            }

            return elts;
        }

        public ReadOnlyCollection<IWebElement> getElements(IWebElement obj, string node, string attr_name, string attr_value)
        {
            return getElements(obj, node, attr_name, attr_value, retry);
        }

        public ReadOnlyCollection<IWebElement> getElements(IWebElement obj, string node, string attr_name, string attr_value, int retry)
        {
            int timeout = 0;
            string q;
            ReadOnlyCollection<IWebElement> elts = null;

            // Prepare selenium Xpath query string as a global By object
            // Handle attr_value containing a single quote
            // Only used with contains(text() ...

            string con_value = XPathLiteral(attr_value);

            if (node.Equals(""))
                node = "*";

            if (retry <= 0)
                debug(ERROR, "getElements: retry must be > 0");

            //
            // pseudo attributes to control Xpath features
            // text: to specify partial containing text
            // subtext: to specify patial containing text found in self or descendants
            //

            if (attr_name.Equals("text"))
            {
                q = @"//" + node + "[contains(translate(text(), '\u00a0', ' '), " + con_value + ")]";
            }
            else if (attr_name.Equals("subtext"))
            {
                q = @"//" + node + "/descendant-or-self::*[contains(translate(text(), '\u00a0', ' '), " + con_value + ")]";
            }
            else if (attr_name.Equals("starts"))
            {
                q = @"//" + node + "[starts-with(translate(text(), '\u00a0', ' '), " + con_value + ")]";
            }
            else if (attr_name.Equals("title"))
            {
                q = @"//" + node + "[translate(@title, '\u00a0', ' ')=" + con_value + "]";
            }
            else if (attr_name.Equals("upper"))
            {
                q = @"//" + node + "[contains(translate(text(), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), " + con_value + ")]";
            }
            else if (attr_name.Equals("lower"))
            {
                q = @"//" + node + "[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), " + con_value + ")]";
            }
            else if (!attr_name.Equals(""))
            {
                q = @"//" + node + "[@" + attr_name + "=" + con_value + "]";
            }
            else
            {
                q = @"//" + node;
            }

            // If not search context element set xpath to search from root of document using "." prefix

            if (obj != null)
                q = "." + q;

            //debug(TRACE, "getElements: spec '" + q + "'");
            spec = By.XPath(q);

            // if a search context element is given use it instead of the default driver
            while (timeout < retry)
            {
                elts = findElements(obj, spec);

                // If elements are there then we are done so drop out of loop

                if (elts != null)
                    if (elts.Count > 0)
                    {
                        debug(TRACE, "getElements: found xpath='" + q + "':" + elts.Count);
                        break;
                    }

                debug(TRACE, "getElements: failed xpath='" + q + "'");

                sleep(1);
                timeout++;
            }

            return elts;
        }

        //
        // getElement
        // Generates an ERROR is the element is not found
        //

        public IWebElement getElement(string node, string attr_name, string attr_value)
        {
            return getElement(null, node, attr_name, attr_value, 0);
        }

        public IWebElement getElement(IWebElement root, string node, string attr_name, string attr_value, int index)
        {
            ReadOnlyCollection<IWebElement> elts;

            elts = getElements(root, node, attr_name, attr_value);
            if (elts.Count > index)
            {
                return elts[index];
            }
            else
            {
                debug(ERROR, "getElement: " + node + ":" + attr_name + "=" + attr_value + ":" + index + " not found");
                return null;
            }
        }

        //
        // untilElement
        //

        public void untilElement(string node_name, string attr_name, string attr_value, string wait_attr, string wait_value)
        {
            untilElement(null, node_name, attr_name, attr_value, wait_attr, wait_value);
        } // untilElement

        public void untilElement(IWebElement obj, string node_name, string attr_name, string attr_value, string wait_attr, string wait_value)
        {
            IWebElement elt;
            string prop = null;
            int count;

            debug(TRACE, "untilElement: " + node_name + ":" + attr_name + "='" + attr_value + "'");

            elt = getElement(obj, node_name, attr_name, attr_value, 0);

            // Get element fresh each time
            if (elt != null)
            {
                // now wait for it to change state
                prop = getAttribute(obj, node_name, attr_name, attr_value, wait_attr, 0);
                for (count = 0; count < getRetry(); count++)
                {
                    if (!wait_value.Equals(prop))
                    {
                        debug(TRACE, "untilElement: Waiting ...");
                        prop = getAttribute(obj, node_name, attr_name, attr_value, wait_attr, 0);
                        sleep(1);
                    }
                    else
                    {
                        debug(TRACE, "untilElement: Found");
                        break;
                    }
                }
            }

            debug(TRACE, "untilElement: Continue");
        } // untilElement

        //
        // getParent / Child
        //

        public IWebElement getParent(IWebElement elt)
        {
            debug(TRACE, "getParent");
            return elt.FindElement(By.XPath(".."));
        }

        public IWebElement getChildLink(IWebElement elt)
        {
            IWebElement obj = null;

            try
            {
                obj = elt.FindElement(By.TagName("a"));
            }
            catch (NoSuchElementException e)
            {
                debug(ERROR, "getChildLink: no A elements found '" + e.Message + "'");
            }
            debug(TRACE, "getChildLink: '" + obj.Text + "'");
            return obj;
        }

        //
        // syncElement - get the element but just wait for it and don't bother to return it
        // Used to detect new page ready after navigation
        // The "count" param is the minimum number to exist
        //

        public void syncElement(IWebElement obj, string node, string by, string id, int count)
        {
            ReadOnlyCollection<IWebElement> elts;

            //debug(TRACE, "syncElement");
            elts = getElements(obj, node, by, id);
            if (elts.Count < count) // if too few found
            {
                debug(ERROR, "syncElement: " + node + ":" + by + "='" + id + "':" + count + " not found");
            }
        }

        //
        // Sync
        //

        public void syncWindow(string title)
        {
            string dtitle = driver.Title;
            int count = 0;

            while (!dtitle.Contains(title) && count++ < getRetry())
            {
                sleep(1);
                dtitle = driver.Title;
                debug(TRACE, "syncWindow: Try: " + count + " '" + dtitle + "'");
            }

            if (count > getRetry())
            {
                debug(ERROR, "syncWindow: '" + dtitle + "' should contain '" + title + "'");
            }

            debug(INFO, "syncWindow: '" + title + "'");
        }

        public void sync(string node, string by, string id)
        {
            sync(null, node, by, id, 1);
        }

        public void sync(IWebElement obj, string node, string by, string id, int count)
        {
            syncElement(obj, node, by, id, count);
            debug(INFO, "sync: " + node + ":" + by + "='" + id + "':" + count);
        }

        //
        // assertValue
        //

        public void assertValue(string node, string by, string id, string prop, string value)
        {
            assertValue(null, node, by, id, prop, value);
        }

        public void assertValue(IWebElement obj, string node, string by, string id, string prop, string value)
        {
            IWebElement ctrl;

            ctrl = getElement(obj, node, by, id, 0);
            if (getAttribute(ctrl, prop).Equals(value))
                debug(TRACE, "assertValue: " + node + ":" + by + ":" + id + ":" + prop + "='" + value + "'");
            else
                debug(ERROR, "assertValue: " + node + ":" + by + ":" + id + ":" + prop + "='" + value + "' failed");
        }

        public void assertValue(IWebElement ctrl, string prop, string value)
        {
            if (getAttribute(ctrl, prop).Equals(value))
                debug(TRACE, "assertValue: " + prop + "='" + value + "'");
            else
                debug(ERROR, "assertValue: " + prop + "='" + value + "' failed");
        }

        //
        // Tables
        //

        //
        // dumpTable
        // Dumps table contents into debug using std functions
        // A handy template for using tables via foreach
        //

        public void dumpTable(IWebElement tab)
        {
            ReadOnlyCollection<IWebElement> rows;
            ReadOnlyCollection<IWebElement> cols;
            int row_i;
            int col_i;

            rows = getTableRows(tab);

            row_i = 0;
            foreach (IWebElement row in rows)
            {
                debug(TRACE, "dumpTable: === Start Row: " + row_i + " ===");

                if (row_i == 0) // header row
                    cols = getTableHeader(row);
                else
                    cols = getTableData(row);

                col_i = 0;
                foreach (IWebElement col in cols)
                {
                    debug(TRACE, "dumpTable: (" + col_i + "):" + col.Text);
                    col_i++;
                }
                row_i++;
            }
        } // dumpTable

        //
        // getTableRows
        // Returns a collection of all the rows, row(0) is the header row
        //

        public ReadOnlyCollection<IWebElement> getTableRows(IWebElement tab)
        {
            ReadOnlyCollection<IWebElement> elts;

            elts = findElements(tab, By.TagName("tr"));
            debug(TRACE, "getTableRows: count:" + elts.Count);
            return elts;
        }

        //
        // getTableData
        // Returns all the <td> elements in the row
        //

        public ReadOnlyCollection<IWebElement> getTableData(IWebElement row)
        {
            ReadOnlyCollection<IWebElement> elts;

            elts = findElements(row, By.TagName("td"));
            debug(TRACE, "getTableData: count:" + elts.Count);
            return elts;
        }

        //
        // getTableHeader
        // Returns all the <th> elements in the row - use for header row
        //

        public ReadOnlyCollection<IWebElement> getTableHeader(IWebElement row)
        {
            ReadOnlyCollection<IWebElement> elts;
            elts = findElements(row, By.TagName("th"));
            debug(TRACE, "getTableHeader: count:" + elts.Count);
            return elts;
        }

        //
        // getTableCell
        // Returns a text representation of the cell contents
        // Use row=0 to access header row, headers have <th>, data rows have <td>
        // Columns start from 0 to (Count-1)
        // colname version uses column name instead of column number
        //

        public IWebElement getTableCell(IWebElement tab, int row, string colname)
        {
            ReadOnlyCollection<IWebElement> rows;
            ReadOnlyCollection<IWebElement> cols;
            int colnum = 0;

            rows = getTableRows(tab);
            cols = getTableHeader(rows[0]); // first row is header row
            foreach (IWebElement col in cols)
            {
                if (col.Text.Contains(colname))
                {
                    break;  // found it
                }
                colnum++;
            }

            debug(TRACE, "getTableCell: column:'" + colname + "'=" + colnum);
            return getTableCell(tab, row, colnum);
        } // getTableCell

        public IWebElement getTableCell(IWebElement tab, int row, int col)
        {
            ReadOnlyCollection<IWebElement> rows;
            ReadOnlyCollection<IWebElement> cols;

            rows = getTableRows(tab);

            if (row == 0) // get header row
                cols = getTableHeader(rows[0]);
            else
                cols = getTableData(rows[row]);

            debug(TRACE, "getTableCell: (" + row + "," + col + "):'" + cols[col].Text + "'");
            return cols[col];
        } // getTableCell

        //
        // Hover
        // Returns: the elt being hovered over
        // Note: There is no working implimentation of hover in selenium for Chrome due to Chrome security that will
        // move the mouse icon over the browser. It can be used to trigger hoverover events, you may see the mouse icon change
        // but the position of the mouse will not change.
        //

        public void hover(string node, string by, string id)
        {
            hover(null, node, by, id, 0);
        }

        public void hover(IWebElement obj, string node, string by, string id, int index)
        {
            IWebElement elt;
            elt = getElement(obj, node, by, id, index);
            if (elt != null)
            {
                hover(elt);
                debug(TRACE, "hover: " + node + ":" + by + "='" + id + "':" + index + " hovered");
            }
            else
            {
                debug(ERROR, "hover: " + node + ":" + by + "='" + id + "':" + index + " not found");
            }
        }

        public void hover(IWebElement obj)
        {
            // This is the std Selenium action chain to generate a simple mouseover event
            // Needs to be tested in MS Dynamics

            Actions builder = new Actions(driver);
            builder.MoveToElement(obj).Build().Perform();

            // A potential mouseover event generating workaround. Done by javascript injection.

            //string javaScript = @"if(document.createEvent){var evObj = document.createEvent('MouseEvents');evObj.initEvent('mouseover'," +
            //                    "true, false); arguments[0].dispatchEvent(evObj);} else if(document.createEventObject)" +
            //                    "{ arguments[0].fireEvent('onmouseover');}";
            //js.ExecuteScript(javaScript, obj);

            debug(INFO, "hover: obj");
        }

        //
        // Scrolling
        //

        public void scrollIntoView(IWebElement obj)
        {
            js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("arguments[0].scrollIntoView(true);", obj);
            sleep(1);
            debug(TRACE, "scrollIntoView: obj");
        }

        public void scrollIntoView_action(IWebElement obj)
        {
            Actions actions = new Actions(driver);
            actions.MoveToElement(obj).Perform();
            sleep(1);
            debug(TRACE, "scrollIntoView_action: obj");
        }

        public void scrollIntoView_center(IWebElement obj)
        {
            js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("arguments[0].scrollIntoView({block: \"center\",inline: \"center\",behavior: \"instant\"});", obj);
            sleep(1);
            debug(TRACE, "scrollIntoView: obj");
        }

        //
        // Window control
        //

        public void windowScrollBy(int x, int y)
        {
            js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("window.scrollBy(arguments[0], arguments[1]);", x, y);
            sleep(2);
            debug(TRACE, "windowScrollBy: " + x + ":" + y);
        }

        public void windowScrollTo(int x, int y)
        {
            js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("window.scrollTo(arguments[0], arguments[1]);", x, y);
            sleep(2);
            debug(TRACE, "windowScrollTo: " + x + ":" + y);
        }

        //
        // Click
        //

        public void click(IWebElement obj)
        {
            int count = 0;

            if (obj == null)
                debug(ERROR, "click: obj is null");

            if (click_displayed(obj))
            {
                debug(TRACE, "click: Displayed: true");
                scrollIntoView_center(obj);
                sleep(1);

                while (obj != null && count < 2)
                {
                    try
                    {
                        debug(INFO, "click: obj");
                        obj.Click();
                        break;
                    }
                    catch (InvalidOperationException e)
                    {
                        count++;
                        debug(TRACE, "click: InvalidOperationException: " + count + " '" + e.Message + "'");
                        sleep(1);
                        continue;
                    }
                    catch (TimeoutException f)
                    {
                        count++;
                        debug(TRACE, "click: TimeoutException: " + count + " '" + f.Message + "'");
                        sleep(1);
                        continue;
                    }
                    catch (WebDriverException g)
                    {
                        // Timed out as usual, ignore
                        debug(TRACE, "click: WebDriverException: " + count + " '" + g.Message + "'");
                        sleep(1);
                        break;
                    }
                }

                if (obj != null && count == 2)
                {
                    // Respond to reoccuring Invalid Op with JS
                    debug(TRACE, "click: exception js attempt");
                    click_js(obj);
                }

                if (obj == null)
                    debug(ERROR, "click: obj became null");

            } // Displayed
            else
            {
                click_js(obj);
                debug(TRACE, "click: non displayed obj");
            }
        } // click

        private bool click_displayed(IWebElement obj)
        {
            bool disp = false;

            try
            {
                disp = obj.Displayed && obj.Enabled;
            }
            catch (NoSuchWindowException e)
            {
                // Ignore
                debug(TRACE, "click: NoSuchWindowException: " + e.Message + "'");
            }

            return disp;
        }

        public void click_js(string node, string by, string id)
        {
            IWebElement obj;

            debug(TRACE, "click_js");
            js = (IJavaScriptExecutor)driver;
            obj = getElement(node, by, id);
            js.ExecuteScript("arguments[0].click();", obj);
        } // click_js

        public void click_js(IWebElement obj)
        {
            debug(TRACE, "click_js");
            js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("arguments[0].click();", obj);
        } // click_js

        public void click_asyncjs(string node, string by, string id)
        {
            IWebElement obj;

            // Setup async callback
            debug(TRACE, "click_asyncjs");
            js = (IJavaScriptExecutor)driver;
            obj = getElement(node, by, id);
            js.ExecuteAsyncScript(
                "var callback = arguments[arguments.length - 1];" +
                "arguments[0].click();" +
                "callback();", obj);
        } // click_js

        public void click_asyncjs(IWebElement obj)
        {
            // Setup async callback
            debug(TRACE, "click_asyncjs");
            js = (IJavaScriptExecutor)driver;
            js.ExecuteAsyncScript(
                "var callback = arguments[arguments.length - 1];" +
                "arguments[0].click();" +
                "callback();", obj);
        } // click_js

        public bool JQueryActive()
        {
            long status;

            debug(TRACE, "JQueryActive");
            js = (IJavaScriptExecutor)driver;
            if ((Boolean)js.ExecuteScript("return window.jQuery != undefined"))
            {
                status = (long)js.ExecuteScript(
                "return JQuery.active;");
                debug(TRACE, "JQueryActive: " + status);
                return status > 0;
            }
            else
            {
                return false;
            }
        }

        public bool JSActive()
        {
            string status;

            debug(TRACE, "JSActive");
            js = (IJavaScriptExecutor)driver;
            status = js.ExecuteScript(
                "return document.readyState;").ToString();
            debug(TRACE, "JQueryActive: " + status);
            return !status.Equals("complete");
        }

        public void click(string node, string by, string id)
        {
            click(null, node, by, id, 0);
        } // click

        public void click(IWebElement obj, string node, string by, string id, int index)
        {
            IWebElement elt;
            int i = 0;

            while (i < 2)
            {
                try
                {
                    elt = getElement(obj, node, by, id, index);
                    if (elt != null)
                    {
                        click(elt);
                        debug(INFO, "click: " + node + ":" + by + ":'" + id + "':" + index + " clicked");
                    }
                    else
                    {
                        debug(ERROR, "click: " + node + ":" + by + ":'" + id + "':" + index + " not found");
                    }
                }
                catch (StaleElementReferenceException e)
                {
                    debug(TRACE, "click: try: " + i + " '" + e.Message + "'");
                    i++;
                    sleep(1);
                    continue;
                }
                break;
            }

            if (i == retry)
                debug(ERROR, "click: stale element reference");

        } // click

        //
        // clear - removes value from field
        //

        public void clear(IWebElement obj)
        {
            scrollIntoView_center(obj);
            sleep(1);

            if (obj.Displayed)
            {
                debug(TRACE, "clear: Displayed: true");
                obj.Clear();
            }
            else
            {
                debug(TRACE, "click: Displayed: false");
                js = (IJavaScriptExecutor)driver;
                js.ExecuteScript("arguments[0].value = '';", obj);
            }

            debug(INFO, "click: obj");
        }

        //
        // doubleClick
        //

        public void doubleClick(IWebElement obj)
        {
            Actions builder = new Actions(driver);
            builder.MoveToElement(obj).DoubleClick(obj).Build().Perform();
            debug(TRACE, "doubleClick");
        }

        //
        // setSelect
        // Params: A ctrl of <SELECT>
        //

        public void setSelect(IWebElement ctrl, string by, string value)
        {
            int i;

            // wait for it to become displayed
            for (i = 0; !ctrl.Displayed && i < 10; i++)
                sleep(1);

            if (i == 10)
                debug(ERROR, "SELECT is not displayed");

            var selectElement = new SelectElement(ctrl);

            try
            {
                if (by.Equals("text"))
                    selectElement.SelectByText(value);

                if (by.Equals("value"))
                    selectElement.SelectByValue(value);
            }
            catch (WebDriverException g)
            {
                debug(ERROR, "setSelect: WebDriverException '" + g.Message + "'");
                sleep(1);
            }

            debug(TRACE, "setSelect:" + by + ":'" + value + "'");

        } // setSelect

        //
        // getSelectOptions
        // Returns all available options as a comma separated list
        //

        public string getSelectOptions(IWebElement ctrl)
        {
            string list = "";
            int i;
            ReadOnlyCollection<IWebElement> elts;

            debug(TRACE, "getSelectOptions");
            elts = findElements(ctrl, By.TagName("option"));
            if (elts.Count > 0)
                list = elts[0].Text;

            for (i = 1; i < elts.Count; i++)
            {
                list = list + "," + elts[i].Text;
            }

            return list;

        } // getSelectOptions

        //
        // getSelectText
        //

        public string getSelectText(IWebElement ctrl)
        {
            debug(TRACE, "getSelectText");
            var selectElement = new SelectElement(ctrl);
            return selectElement.SelectedOption.Text;

        } //getSelectText

        //
        // Type
        // Sends character to a control like an edit field
        // Use the obj version for TEXTAREA fields
        //

        public void type(string text)
        {
            new Actions(getDriver()).SendKeys(text).Perform();
        }

        public void type(IWebElement obj, string text)
        {
            type(obj, text, true);
        }
        public void type(IWebElement obj, string text, bool showText)
        {
            if (obj.Displayed)
            {
                debug(TRACE, "type: Displayed: true");
                obj.SendKeys(text);
            }
            else
            {
                debug(TRACE, "type: Displayed: false");
                js = (IJavaScriptExecutor)driver;
                js.ExecuteScript("arguments[0].value = arguments[1]", obj, text);
            }

            if (showText)
            {
                debug(INFO, "type: '" + text + "'");
            } else
            {
                text = "****";
                debug(INFO, "type: '" + text + "'");
            }
        }

        //
        // These versions of type assume INPUT onject
        //

        public void type(string by, string id, string text)
        {
            type(null, by, id, text, 0);
        }

        public void type(IWebElement obj, string by, string id, string text, int index)
        {
            IWebElement elt;
            elt = getElement(obj, "INPUT", by, id, index);
            if (elt != null)
            {
                type(elt, text);
                debug(TRACE, "type: " + by + "='" + id + "':'" + text + "':" + index + " typed");
            }
            else
            {
                debug(TRACE, "type: " + by + "='" + id + "':'" + text + "':" + index + " not found");
            }

        }

        //
        // javascript workaround by ID
        //

        public void type_js(string objectID, string newValue)
        {
            js = (IJavaScriptExecutor)driver;
            string value = (string)js.ExecuteScript("document.getElementById('" + objectID + "').value='" + newValue + "';");
        }

        //
        // backspace
        //

        public void backspace(IWebElement ctrl)
        {
            int len;
            len = getAttribute(ctrl, "value").Length;
            click(ctrl);
            ctrl.SendKeys(Keys.End);
            for (int i = 0; i < len; i++)
                ctrl.SendKeys(Keys.Backspace);
        }

        //
        // getAttribute - read a property of the DOM object
        //

        public string getAttribute(IWebElement obj, string prop)
        {
            string val = "";
            int count = 0;

            if (prop.Equals("text"))
                val = obj.Text;
            else
            {
                while (count < 2)
                {
                    try
                    {
                        val = obj.GetAttribute(prop);
                        if (val == null)
                            debug(ERROR, "getAttribute: Bad property: " + prop);

                        break;
                    }
                    catch (StaleElementReferenceException e)
                    {
                        sleep(1);
                        count++;
                        debug(TRACE, "getAttribute try: " + count + " '" + e.Message + "'");
                        continue;
                    }
                } // while

                if (count == 2)
                    debug(ERROR, "getAttribute: StaleElementReferenceException failure");
            }

            debug(TRACE, "getAttribute: " + prop + "='" + val + "'");
            return val;
        }

        public string getAttribute(string node, string by, string id, string prop)
        {
            return getAttribute(null, node, by, id, prop, 0);
        }

        public string getAttribute(IWebElement obj, string node, string by, string id, string prop, int index)
        {
            ReadOnlyCollection<IWebElement> elts;
            string val = "";

            elts = getElements(obj, node, by, id);
            if (index < elts.Count)
            {
                val = getAttribute(elts[index], prop);
                debug(TRACE, "getAttribute: " + node + ":" + id + ":" + prop + "='" + val + "'");
            }
            else
                debug(ERROR, "getAttribute: " + node + ":" + id + ":" + prop + " not found");

            return val;
        } // getAttribute

        //
        // downloadPDF - pass the By for the element <A href=""> ctrl and path to store
        // Returns filename
        //

        public string downloadPDF(By by, string path)
        {
            string downloads = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\downloads";
            string filename = "";
            string href = "";

            ReadOnlyCollection<IWebElement> elts = null;

            int i;

            // wait for and href to appear in the control

            debug(TRACE, "downloadPDF: href: '" + href + "'");
            for (i = 0; i < getRetry() && (href == "" || href.Contains("download.aspx")); i++)
            {
                sleep(2); // wait a bit, it can be slow to populate
                elts = getDriver().FindElements(by); // refresh element
                if (elts.Count > 0)
                {
                    href = getAttribute(elts[0], "href");
                }
                debug(TRACE, "downloadPDF: href: '" + href + "'");
            }

            if (href == "")
                debug(ERROR, "downloadPDF: no href to download");

            // File is in ECM - use a web client

            filename = Path.GetFileName(href);
            debug(TRACE, "ECM: filename: '" + filename + "'");

            WebClient client = new WebClient();
            // Synchronous - so it waits
            client.UseDefaultCredentials = true;
            // client.Credentials = new NetworkCredential("username", "password");
            client.DownloadFile(new Uri(href), Path.Combine(path, filename));

            return filename;

        } // downloadPDF

        //
        // XPathLiteral
        //
        // Produce an XPath literal equal to the value if possible; if not, produce
        // an XPath expression that will match the value.
        // 
        // Note that this function will produce very long XPath expressions if a value
        // contains a long run of double quotes.
        // </summary>
        // <param name="value">The value to match.</param>
        // <returns>If the value contains only single or double quotes, an XPath
        // literal equal to the value.  If it contains both, an XPath expression,
        // using concat(), that evaluates to the value.</returns>
        //
        // Attribution: https://stackoverflow.com/questions/1341847/special-character-in-xpath-query/11585487
        // Author: Robert Rossney (StackExchange MIT License)
        //

        static string XPathLiteral(string value)
        {
            // if the value contains only single or double quotes, construct
            // an XPath literal
            if (!value.Contains("\""))
            {
                return "\"" + value + "\"";
            }
            if (!value.Contains("'"))
            {
                return "'" + value + "'";
            }

            // if the value contains both single and double quotes, construct an
            // expression that concatenates all non-double-quote substrings with
            // the quotes, e.g.:
            //
            //    concat("foo", '"', "bar")

            StringBuilder sb = new StringBuilder();
            sb.Append("concat(");
            string[] substrings = value.Split('\"');
            for (int i = 0; i < substrings.Length; i++)
            {
                bool needComma = i > 0;
                if (substrings[i] != "")
                {
                    if (i > 0)
                    {
                        sb.Append(", ");
                    }
                    sb.Append("\"");
                    sb.Append(substrings[i]);
                    sb.Append("\"");
                    needComma = true;
                }
                if (i < substrings.Length - 1)
                {
                    if (needComma)
                    {
                        sb.Append(", ");
                    }
                    sb.Append("'\"'");
                }

            }
            sb.Append(")");
            return sb.ToString();
        }

        public bool waitForElement(string vNode, string vAttribute, string vAttributeValue, int timeoutValue)
        {
            IWebElement elementToFind = getElement(vNode, vAttribute, vAttributeValue);
            return waitForElement(elementToFind, timeoutValue);
        }

        public bool waitForElement(IWebElement objectToFind, int timeoutValue)
        {
            bool foundObject = false;
            int timeOut = 0;

            while (!objectToFind.Displayed && timeOut <= timeoutValue)
            {
                sleep(1);
                timeOut++;
            }

            if (objectToFind.Displayed)
            {
                foundObject = true;
            }
            return foundObject;
        }

        public void setText(string setNode, string setAttName, string setAttValue, string valueToSet)
        {
            IWebElement objectToFind = getElement(setNode, setAttName, setAttValue);
            objectToFind.Clear();
            objectToFind.SendKeys(valueToSet);
        } // setText

        //
        // Alerts
        //

        public bool waitForAlert(int t)
        {
            int i = 0;
            while (i < t)
            {
                try
                {
                    driver.SwitchTo().Alert();
                    return true;
                }
                catch (Exception e)
                {
                    debug(TRACE, "waitForAlert try: " + i + " '" + e.Message + "'");
                    i++;
                    sleep(1);
                }
            }
            return false;
        } // waitForAlert

        public bool acceptAlert(int t)
        {
            if (waitForAlert(t))
            {
                driver.SwitchTo().Alert().Accept();
                return true;
            }
            else
                return false;
        }

        public bool dismissAlert(int t)
        {
            if (waitForAlert(t))
            {
                driver.SwitchTo().Alert().Dismiss();
                return true;
            }
            else
                return false;
        }

        // Get the text from the selected edit field
        public string selectAllCopy()
        {
            clearWindowsClipboard();

            Actions a = new Actions(driver);
            a.KeyDown(Keys.Control);
            a.SendKeys("a");
            a.KeyUp(Keys.Control);
            a.Build().Perform();

            a.KeyDown(Keys.Control);
            a.SendKeys("c");
            a.KeyUp(Keys.Control);
            a.Build().Perform();
            sleep(1);

            return getTextWindowsClipboard();

        }

        public void waitForElementNotVisible(string xpathSpec, int timeoutValue)
        {
            debug(TRACE, "waitForElementNotVisible: start");

            new WebDriverWait(driver, TimeSpan.FromSeconds(timeoutValue))
                    .Until(drv => !IsElementVisible(By.XPath(xpathSpec)));
        }

        public void waitForElementVisible(string xpathSpec, int timeoutValue)
        {
            debug(TRACE, "waitForElementNotVisible: start");

            new WebDriverWait(driver, TimeSpan.FromSeconds(timeoutValue))
                    .Until(drv => IsElementVisible(By.XPath(xpathSpec)));
        }

        public bool IsElementVisible(By searchElementBy)
        {
            try
            {
                return driver.FindElement(searchElementBy).Displayed;

            }
            catch (NoSuchElementException)
            {
                return false;
            }
            catch (StaleElementReferenceException e)
            {
                debug(TRACE, "IsElementVisible: '" + e.Message + "'");
                sleep(1);
                return false;
            }
        }

        public void openNewBrowserWindow()
        {
            driver.SwitchTo().NewWindow(WindowType.Window);
        }
    }
}
