using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework.Interfaces;
using NUnit.Framework;
using OpenQA.Selenium;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using AventStack.ExtentReports.Model;
using OpenQA.Selenium.Chrome;

namespace ExtentReport
{
    public class baseTest
    {
        public ExtentReports extent;
        public ExtentTest test;
        ChromeDriver driver;

        //report file
        [OneTimeSetUp]
        public void Setup()

        {
         public void SetUpExtentReport()
{
    // Use a path relative to the repository root (current working directory)
    string projectRoot = Directory.GetCurrentDirectory();
    string reportsDir = Path.Combine(projectRoot, "Reports");
    string screenshotsDir = Path.Combine(reportsDir, "Screenshots");

    // Assign to class-level variables if needed elsewhere
    reportDirectory = reportsDir;
    screenshotDirectory = screenshotsDir;

    // Ensure directories exist
    if (!Directory.Exists(reportDirectory)) Directory.CreateDirectory(reportDirectory);
    if (!Directory.Exists(screenshotDirectory)) Directory.CreateDirectory(screenshotDirectory);

    // Define the full path for the HTML test report file
    reportPath = Path.Combine(reportDirectory, "TestReport.html");
    Console.WriteLine("Extent report will be created at: " + reportPath);

    // Initialize the Spark Reporter with the defined report path
    var sparkReporter = new ExtentSparkReporter(reportPath)
    {
        Config =
        {
            ReportName = "Canvas Integration Automation Test Report",
            TimeStampFormat = "yyyy-MM-dd HH:mm:ss"
        }
    };

    // Initialize the main ExtentReports object and attach the Spark reporter
    extent = new ExtentReports();
    extent.AttachReporter(sparkReporter);

    // Add metadata/system information to the report
    extent.AddSystemInfo("Operating System", Environment.OSVersion.ToString());
    extent.AddSystemInfo("Host Name", Environment.MachineName);
}
        }

        [SetUp]
        public void startBrowser()
        {
            //start the test
            test = extent.CreateTest(TestContext.CurrentContext.Test.Name);
            //Intialise the browser
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
        }

        [Test]
        public void launchBrowser()
        {
            //launch the browser and navigate to the URL
            driver.Navigate().GoToUrl("https://google.com"); // Replace with actual URL
        }
   

        [TearDown]
        public void AfterTest()

        {
            var status = TestContext.CurrentContext.Result.Outcome.Status;
            var stackTrace = TestContext.CurrentContext.Result.StackTrace;

            DateTime time = DateTime.Now;
            String fileName = "Screenshot_" + time.ToString("h_mm_ss") + ".png";

            if (status == TestStatus.Failed)
            {

                test.Fail("Test failed", captureScreenShot(driver, fileName));
                test.Log(Status.Fail, "test failed with logtrace" + stackTrace);

            }
            else if (status == TestStatus.Passed)
            {
                test.Pass("Test passed", captureScreenShot(driver, fileName));  
            }

            extent.Flush();
            driver.Quit();
        }

        public Media captureScreenShot(IWebDriver driver, String screenShotName)
        {
            ITakesScreenshot ts = (ITakesScreenshot)driver;
            var screenshot = ts.GetScreenshot().AsBase64EncodedString;

            return MediaEntityBuilder.CreateScreenCaptureFromBase64String(screenshot, screenShotName)
                .Build();
        }
    }
}
