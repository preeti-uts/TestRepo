using System;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Remote;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Diagnostics;

namespace selenium_framework
{
    public partial class Selenium_Framework
    {
        // Driver initialisation
        public IWebDriver initDriver()
        {
            setTestAccountCredentials();
            return initDriver(true, "");
        }
        public IWebDriver initDriver(bool killProc)
        {
            setTestAccountCredentials();
            return initDriver(killProc, "");
        }

        public IWebDriver initDriver(bool killProc, string chrome_user)
        {
            /*
            
            DIFFERENCERS BETWEEN DRIVERS

            RemoteWebDriver: This driver class comes directly from the upstream Selenium project. This is a pretty generic driver
            where initializing the driver means making network requests to a Selenium hub to start a driver session. Since Appium
            operates on the client-server model, Appium uses this to initialize a driver session. However, directly using the
            RemoteWebDriver is not recommended since there are other drivers available that offer additional features or convenience
            functions.

            AppiumDriver: This driver class inherits from the RemoteWebDriver class, but it adds in additional functions that are useful
            in the context of a mobile automation test through the Appium server.

            AndroidDriver: This driver class inherits from AppiumDriver, but it adds in additional functions that are useful in the
            context of a mobile automation test on Android devices through Appium. Only use this driver class if you want to start a
            test on an Android device or Android emulator.

            IOSDriver: This driver class inherits from AppiumDriver, but it adds in additional functions that are useful in the context
            of a mobile automation test on iOS devices through Appium. Only use this driver class if you want to start a test on an iOS
            device or iOS emulator.
            
             */

            setTestAccountCredentials();

            if (driver == null)
            {
                // No driver - so create one
                debug(TRACE, "initDriver: new '" + envBrowser + "'");

                // Check screen resolution
                checkScreenResolution();

                //
                // Start new browser environment
                //

                switch (envBrowser)
                {
                    case "chrome":
                        switch (envOS)
                        {
                            case "windows":
                                bool driverPathFound = true;
                                if (killProc)
                                {
                                    killProcess("chrome");
                                }

                                killProcess("chromedriver");
                                ChromeDriverService cservice = null;
                                debug(TRACE, "Driver Path: " + getDriverPath());
                                string currentDriverPath = getDriverPath();
                                if (Directory.Exists(currentDriverPath))
                                {
                                    // Use the driver in the set path if it exists
                                    cservice = ChromeDriverService.CreateDefaultService(getDriverPath());
                                } 
                                else
                                {
                                    // Get selenium manager to set the path to a temp location
                                    cservice = ChromeDriverService.CreateDefaultService();
                                }
                                ChromeOptions c_options = new ChromeOptions();
                                List<string> ls = new List<string>();
                                if (DetectVirtualMachine())
                                {
                                    //ls.Add("enable-automation"); // remove again looks like change on github agent again
                                }
                                c_options.AddExcludedArguments(ls);

                                if (chrome_user.Length > 0)
                                {
                                    c_options.AddArgument(@"--user-data-dir=C:\Users\" + chrome_user + @"\AppData\Local\Google\Chrome\User Data");
                                }

                                c_options.AddArgument("--disable-gpu");
                                c_options.AddArgument("--start-maximized");
                                c_options.AddArgument("--disable-infobars");
                                c_options.Proxy = null;

                                if (getDebugLevel() == VERBOSE)
                                {
                                    cservice.LogPath = "chromedriver.log";
                                    cservice.EnableVerboseLogging = true;
                                    debug(TRACE, "initDriver: chromedriver log: " + cservice.LogPath);
                                }

                                c_options.AddUserProfilePreference("credentials_enable_service", false);
                                driver = new ChromeDriver(cservice, c_options, TimeSpan.FromMinutes(2));

                                break;

                            case "android":
                                //AppiumOptions a_options = new AppiumOptions();

                                //a_options.AddAdditionalCapability(MobileCapabilityType.PlatformName, "Android");
                                //a_options.AddAdditionalCapability(MobileCapabilityType.BrowserName, "Chrome");
                                //Uri url = new Uri("http://127.0.0.1:4723/wd/hub");
                                //driver = new AndroidDriver<AndroidElement>(url, a_options);

                                break;

                            default:
                                debug(ERROR, "initDriver: unknown OS: " + envOS);
                                break;
                        }
                        break;

                    case "internet explorer":
                        if (killProc)
                        {
                            killProcess("iexplore");
                        }

                        killProcess("IEDriverServer");
                        driver = new InternetExplorerDriver(getDriverPath());
                        break;

                    case "firefox":
                        if (killProc)
                        {
                            killProcess("firefox");
                        }

                        killProcess("geckodriver");
                        FirefoxDriverService fservice = FirefoxDriverService.CreateDefaultService(getDriverPath());
                        driver = new FirefoxDriver(fservice);
                        break;

                    default:
                        debug(ERROR, "initDriver: unknown browser: '" + envBrowser + "'");
                        break;
                }

                // Manage capabilities of this driver
                ICapabilities cap = null;
                switch (envBrowser)
                {
                    case "chrome":
                        cap = ((ChromeDriver)driver).Capabilities;
                        break;
                    case "internet explorer":
                        cap = ((InternetExplorerDriver)driver).Capabilities;
                        break;
                    case "firefox":
                        cap = ((FirefoxDriver)driver).Capabilities;
                        break;
                    default:
                        debug(ERROR, "initDriver: unknown browser: '" + envBrowser + "'");
                        break;
                }
                debug(TRACE, "BrowserName: " + cap.GetCapability("browserName"));
                debug(TRACE, "Platform: " + cap.GetCapability("platformName"));

                // Get browser version
                browserVersion = cap.GetCapability("browserVersion").ToString();
                debug(TRACE, "Version: " + browserVersion);

                // Set Javascript executor
                js = (IJavaScriptExecutor)driver;

                // Create screenshot directory if it doesn't exist
                if (!Directory.Exists(screenshotLocation)) 
                {
                    Directory.CreateDirectory(screenshotLocation);
                }

                // Maximise the browser
                maximize();

                debug(TRACE, "initDriver: '" + envBrowser + "'");

            }

            return driver;

        } // initDriver
    }
}
