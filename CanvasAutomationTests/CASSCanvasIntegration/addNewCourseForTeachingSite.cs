using NUnit.Framework;
using selenium_framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CanvasAutomationTests.CASSCanvasIntegration
{
    [TestFixture]
    public partial class addNewCourseForTeachingSite : baseTest
    {
        [Test]
        public void addNewCourseForTeachingSiteTest()
        {
            debug(TRACE, "\r\n*** Test start: " + this.GetType().ToString() + "\r\n");
            // Script parameters
            String currentSession = "Autumn Session";
            DateTime now = DateTime.Now;
            string currentYr = now.ToString("yyyy");
            bool inactiveRecordfound = false;
            bool activeRecordfound = false;
            String studentId = "";
            String studyPackageCd = "";
            String subject_SISId = "";
            // **************** Login to CASS and download an excel with students having planned subjects *********************
            initDriver();
            //gotoUrl(cassURL);
            // start CASS Cloud and login
            loginStartStudentManagement(cassURL, userLoginName, userLoginPassword);
            // wait for splash screen and click role       
        }

    }
   
}
