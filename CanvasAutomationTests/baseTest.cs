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
            string workingDirectory = Environment.CurrentDirectory;
            string projectDirectory = Path.Combine(Directory.GetParent(workingDirectory).Parent.Parent.FullName, Report);
            if (!Directory.Exists(projectDirectory)) Directory.CreateDirectory(projectDirectory);
            String reportPath = projectDirectory + "//index.html";
            var htmlReporter = new ExtentSparkReporter(reportPath);
            extent = new ExtentReports();
            extent.AttachReporter(htmlReporter);
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
