using NUnit.Framework;
using selenium_framework;

// **** change the folder for the script in the namespace
namespace CASSRegressionTests._ScriptTemplate
{
    [TestFixture]
    // **** Update the partial class name to the script name
    public partial class scriptTemplate : Selenium_Framework 
    {

        [Test]
        // **** Change the test name to the script name and amment Test to it
        public void scriptTemplateTest()
        {
            debug(TRACE, "\r\n*** Test start: " + GetType().ToString() + "\r\n");

            // Script parameters // **** and script params at the top here

            setDriverPath(@"C:\SeleniumWebDrivers\ChromeDriver"); // **** default chromedriver location on the GitHub agent
            initDriver(true);

            // start CASS Cloud and login
            loginStartStudentManagement(cassURL, userLoginName, userLoginPassword); // **** Add the right login details (in the User_library

            // wait for splash screen and click role rewards
            selectWorkplace("Rewards"); // **** select the correct role for the test

            // wait for main page to load
            imageWait(dirName + @"\cassTemplate_images\commsPortal.PNG", 60, true);

            LocationValues cassArea = new LocationValues();
            cassArea.X = 260;
            cassArea.Y = 255;
            cassArea.W = 1033 - cassArea.X;
            cassArea.H = 873 - cassArea.Y;

           // **** Your code here

            // Close student management
            exitStudentManagement(dirName);

            sleep(2);

            // wait for the second tab to close
            syncWindowCount(1);
            popupWindow(0);

            // Log off cass cloud
            logoutCASSCloud();

            // quit the browser
            quit();

            debug(TRACE, "\r\n*** Test end: " + GetType().ToString() + "\r\n");
        }
    }
}
