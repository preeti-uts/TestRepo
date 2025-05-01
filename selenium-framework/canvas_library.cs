using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.DevTools.V122.Debugger;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.Common;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using Keys = OpenQA.Selenium.Keys;

namespace selenium_framework
{
    // Common library for frequently used functionality for reuse across scripts
    public partial class Selenium_Framework
    {
        // Download an excel with students having planned subjects in current yr and current session
        public void studentStudypackageExcelDownload(String currentSession, String currentYr)
        {
            // open search
            imageFindClick(dirName + @"\cassTemplate_images\searchBox\studentStudyPackageSearch.PNG", 15, 0);
            imageWait(dirName + @"\cassTemplate_images\syncImages\studentStudyPackageSYNC.PNG", 30);

            clickSearchAddCriteria("", new string[,] { { "SSP Stage", "=", "Planned" }, { "Availability Year", "=", ""+ currentYr +"" },
                { "Course Category", "=", "Subject" }, { "Session", "=", "" + currentSession + "" }, {"SSP Status Effective Date", ">=",  "01/01/"+ currentYr + "" }, { "Not On Plan", "=", "No"} });

            DataSet studentStudyPkgSearch = downloadExcelConvertToDataSet(60, true, null);
        }

        //Convert latest downloaded excel file into dataset and sort into groups based on study package cd (Subject)
        public DataTable getSortedDataTable()
        {
            //Fetch the latest downloaded file
            DataSet downloadedFileDataSet = getDataSetFromExcelFile(getFileLocationOfLatestDownloadedFile(), 4);

            //check the first row in data set
            DataRow dRow = downloadedFileDataSet.Tables[0].Rows[1];
            debug(TRACE, "################# Third row data before sort ######################" + dRow[3]);

            //************** Group datatable rows based on studyPackageCd **************
            DataTable sortedDataTable = sortDataSet(downloadedFileDataSet, "Study Package Cd");

            //check first row in sorted table
            DataRow dRowsorted = sortedDataTable.Rows[1];
            debug(TRACE, "################# Third row data after sort  ######################" + dRowsorted[3]);

            return sortedDataTable;
        }

        /// <summary>
        /// ///############moved to Pages - CAT Login page//////////////////////////
        /// </summary>
        /// <param name="subjectCode"></param>
        /// <param name="currentSession"></param>
        /// <param name="currentYr"></param>
        /// <returns></returns>

        //public void loginToCAT()
        //{
        //    debug(TRACE, "\r\n*** Test start: " + this.GetType().ToString() + "\r\n");

        //    openNewBrowserWindow();
        //    ReadOnlyCollection<IWebElement> userField;
        //    ReadOnlyCollection<IWebElement> pwField;
        //    ReadOnlyCollection<IWebElement> submitButton;
        //    ReadOnlyCollection<IWebElement> utsEmailField;

        //    //initDriver();
        //    gotoUrl(catURL);

        //    // Wait for element
        //    waitForElement("*", "text", "Sign in", 5);

        //    // Find and enter email address
        //    userField = findElements(null, By.XPath("//*[@name='loginfmt']"));
        //    type(userField[0], catUserLoginName, false);

        //    // Click next button
        //    submitButton = findElements(null, By.XPath("//*[@type='submit']"));
        //    click(submitButton[0]);

        //    // Wait for element
        //    waitForElement("*", "type", "text", 10);

        //    // Find and enter uts email address
        //    utsEmailField = findElements(null, By.XPath("//*[@type ='text']"));
        //    clear(utsEmailField[0]);
        //    sleep(2);
        //    type(utsEmailField[0], catUserLoginName, false);

        //    // Click next button
        //    submitButton = findElements(null, By.XPath("//*[@type='submit']"));
        //    click(submitButton[0]);

        //    // Wait for element
        //    waitForElement("*", "type", "password", 15);

        //    // Find and enter password
        //    pwField = findElements(null, By.XPath("//*[@type='password']"));
        //    type(pwField[0], catUserLoginPassword, false);

        //    // Find and click sign in button
        //    submitButton = findElements(null, By.XPath("//*[@type='submit']"));
        //    click(submitButton[0]);

        //    // Wait for element
        //    waitForElement("*", "text", "Manage courses", 10);
        //    }

        //public String courseSearchInCAT(String subjectCode, String currentSession, String currentYr)
        //{
        //    debug(TRACE, "\r\n*** Test start: " + this.GetType().ToString() + "\r\n");

        //    //loginToCAT();

        //    // Find and click select year dropdown
        //    ReadOnlyCollection<IWebElement> selectYrDropdown = findElements(null, By.XPath("//*[@id='multiSelectSearchCntrl_dropdownYears']"));
        //    click(selectYrDropdown[0]);

        //    //Find and enter the year
        //    ReadOnlyCollection<IWebElement> enterCurrentYr = findElements(null, By.XPath("//*[@placeholder='Please search for years']"));
        //    type(enterCurrentYr[0], currentYr, false);

        //    // Find and click current year checkbox
        //    ReadOnlyCollection<IWebElement> currentYrCheckbox = findElements(null, By.XPath("//span[contains(text(), '2024')]"));
        //    click(currentYrCheckbox[0]);

        //    //Hit escape to come out of the Select year dropdown
        //    type(Keys.Escape);

        //    // Find and click select Session dropdown
        //    ReadOnlyCollection<IWebElement> selectSessionDropdown = findElements(null, By.XPath("//*[@id='multiSelectSearchCntrl_dropdownSessions']"));
        //    click(selectSessionDropdown[0]);

        //    //Find and enter the session
        //    ReadOnlyCollection<IWebElement> enterCurrentSession = findElements(null, By.XPath("//*[@placeholder='Please search for sessions']"));
        //    type(enterCurrentSession[0], currentSession, false);

        //    // Find and click current session checkbox
        //    ReadOnlyCollection<IWebElement> currentSessionCheckbox = findElements(null, By.XPath("//span[text() = ' Autumn Session ']"));
        //    click(currentSessionCheckbox[0]);

        //    //Hit escape to come out of the Select Session dropdown
        //    type(Keys.Escape);

        //    // Find and click select status dropdown
        //    ReadOnlyCollection<IWebElement> selectStatusDropdown = findElements(null, By.XPath("//*[@id='multiSelectSearchCntrl_dropdownStatus']"));
        //    click(selectStatusDropdown[0]);
        //    sleep(2);

        //    //Find and enter the status1 - Complete
        //    ReadOnlyCollection<IWebElement> enterStatus1 = findElements(null, By.XPath("//*[@placeholder='Please search for statuses']"));
        //    type(enterStatus1[0], "Complete", true);

        //    sleep(2);

        //    // Find and click 'status1 - Complete' checkbox
        //    ReadOnlyCollection<IWebElement> status1Checkbox = findElements(null, By.XPath("//span[text() = ' Complete ']"));
        //    click(status1Checkbox[0]);

        //    sleep(2);

        //    // Find and click 'status2 - Conditionally complete' checkbox
        //    ReadOnlyCollection<IWebElement> status2Checkbox = findElements(null, By.XPath("//span[text() = ' Conditionally complete ']"));
        //    click(status2Checkbox[0]);

        //    // Hit escape to come out of the Select Status dropdown
        //    type(Keys.Escape);

        //    // Find and click Search by keyword field
        //    ReadOnlyCollection<IWebElement> searchBykeyword = findElements(null, By.XPath("//*[@id='txtCourseSearch']"));
        //    click(searchBykeyword[0]);
        //    type(searchBykeyword[0], subjectCode, true);

        //    // Find and click Search button
        //    ReadOnlyCollection<IWebElement> searchButton = findElements(null, By.XPath("//*[@id='btnCourseSearch']"));
        //    click(searchButton[0]);

        //    // Wait for element
        //    waitForElement("*", "text", "SISID", 20);

        //    sleep(2);

        //    // get course table and its rows 
        //    IWebElement courseTable = getElement("table", "id", "tblCourse");
        //    //ReadOnlyCollection<IWebElement> courseRowData = getTableRows(courseTable);

        //    //Fetch SISID of the first data row
        //    String sisID = getTableCell(courseTable, 1, 1).Text;
        //    Console.WriteLine(sisID);

        //    //Find and Click on the course name in first row
        //    IWebElement courseName = getTableCell(courseTable, 1, 2);
        //    click(courseName);

        //    //wait for the progress bar  
        //    //bool progressBarFound = waitForElement("", "role", "progressbar", 5);

        //    waitForElementVisible("//*[@role='progressbar']", 10);

        //    waitForElementNotVisible("//*[@role='progressbar']", 40);

        //    //Get CASS subjects or courses table
        //    IWebElement subjectCourseTable = getElement("table", "id", "selectedSubjectTable");
        //    ReadOnlyCollection<IWebElement> subjectCourseRowData = getTableRows(subjectCourseTable);

        //    int numberOfSubjectRecords = subjectCourseRowData.Count();

        //    Console.WriteLine("No of rows in subject table: " + numberOfSubjectRecords);

        //    //Check if subjects or courses table is having atleast one row 

        //    if (numberOfSubjectRecords >= 1)
        //    {
        //        //subjectCode = getTableCell(subjectCourseTable, 1, 2).Text;
        //        debug(TRACE, "############ Subject present in CAT  ################" + sisID);
        //        return sisID;
        //    }
        //    else
        //    {
        //        return "";
        //    }

        //}


        //Canvas Beta login
        public void canvasTestLogin()
        {
            loginWithOkta(canvasURL, canvasUserLoginName, canvasUserLoginPassword, canvasUserLoginAnswerToSecurityQuestionAnswer);

            //Wait for the home page to load completely
            waitForElement("*", "id", "dashboard", 15);

        }

        //Find a student from sorted datatable which is not enrolled in Canvas to the corresponding subject
        public List<String> checkSubjectEnrolmentInCanvas(String studyPackageCd, String subject_SISId, bool inactiveRecordfound)
        {
            //Find and Click on Admin section under menu
            ReadOnlyCollection<IWebElement> AdminMenu = findElements(null, By.XPath("//ul[@id = 'menu']/li[2]/a[@id = 'global_nav_accounts_link']"));
            click(AdminMenu[0]);

            //Wait for Admin menu to come on screen before clicking on root link
            waitForElement("*", "text", "Admin", 5);

            //Find and Click UTS link and wait for the courses to load
            ReadOnlyCollection<IWebElement> utsLink = findElements(null, By.XPath("//a[contains(text(), 'University of Technology Sydney')]"));
            click_js(utsLink[0]);
            waitForElement("*", "data-automation", "courses list", 5);

            //Find 'search courses' searchbox and enter the course under test
            ReadOnlyCollection<IWebElement> courseSearch = findElements(null, By.XPath("//input[@type = 'search']"));
            type(courseSearch[0], subject_SISId, true);

            //wait for the search result 
            bool textFound = waitForElement("", "text", subject_SISId, 15);

            //Find and click on the course link if text found in search result
            if (textFound)
            {

                ReadOnlyCollection<IWebElement> courseLink = findElements(null, By.XPath("//span[starts-with(text()," + "'" + studyPackageCd + "'" + ")]"));
                click_js(courseLink[0]);
            }

            //Wait for the course detail page to load
            waitForElement("*", "text", "Recent announcements", 5);

            //Find and click on People link in the left panel
            ReadOnlyCollection<IWebElement> peopleLink = findElements(null, By.XPath(".//a[contains(text(), 'People')]"));
            click(peopleLink[0]);

            //Wait for the table with students list to load conpletely
            waitForElement("th", "text", "Name", 15);

            //Select Student from the drop down
            ReadOnlyCollection<IWebElement> studentFromDropDown = findElements(null, By.XPath("//*[@name = 'enrollment_role_id']"));
            setSelect(studentFromDropDown[0], "value", "3");

            sleep(8);

            // get people table and its rows 
            IWebElement peopleTable = getElement("tbody", "class", "collectionViewItems");
            ReadOnlyCollection<IWebElement> peopleTableRows = getTableRows(peopleTable);
            List<String> inactiveStudentsList = new List<String>();
            //int loopCount = 1;
            foreach (IWebElement studentRow in peopleTableRows)
            {
                //ReadOnlyCollection<IWebElement> inactiveStudentRow = findElements(null, By.XPath("//span[@class = 'label' and contains(text(), 'inactive')]"));
                if (studentRow.Text.Contains("inactive"))
                {
                    String studentRowText = studentRow.Text;
                    debug(TRACE, "################# student row text #################" + studentRowText);
                    var studentSSIDMatch = Regex.Match(studentRowText, "\\binactive\\s\\b\\d+");
                    String studentSSID = (studentSSIDMatch.Groups[0].Value).Substring(9);
                    Console.WriteLine("student id is: " + studentSSID);
                    inactiveStudentsList.Add(studentSSID);
                }

            }
            return inactiveStudentsList;
        }

        //Perform subject enrolment in CASS
        public String subjectEnrolmentInCASS(List<String> inactiveStudents, String studentId, String studyPackageCd, String currentSession, String currentYr)
        {
            // open search
            imageFindClick(dirName + @"\cassTemplate_images\searchBox\studentStudyPackageSearch.PNG", 15, 0);
            imageWait(dirName + @"\cassTemplate_images\syncImages\studentStudyPackageSYNC.PNG", 30);

            foreach (String id in inactiveStudents)
            {

                if (id == inactiveStudents.First())
                {
                    clickSearchAddCriteria("", new string[,] { { "Student Id", "=", "" + id + "" }, { "Study Package Cd", "=", "" + studyPackageCd + "" }, { "SSP Stage", "=", "Planned" }, { "Availability Year", "=", "" + currentYr + "" },
                    { "Session", "=", "" + currentSession + "" }, { "SSP Status Effective Date", ">=", "01/01/" + currentYr + "" }, { "Not On Plan", "=", "No" }});
                }
                else
                {   
                    imageFindClick(dirName + @"\cassTemplate_images\editFields\studentIdEdit.PNG");
                    sleep(1);
                    SendKeys.SendWait(@"{TAB}");
                    sleep(1);
                    SendKeys.SendWait(@"{TAB}");
                    sleep(1);
                    clearTextInField();
                    typeTextInField(id);

                    // click on retrieve
                    imageWait(dirName + @"\cassTemplate_images\retrieveSearch.PNG", 30);
                    imageFindClick(dirName + @"\cassTemplate_images\retrieveSearch.PNG");
                }

                //Accept any warnings
                checkForWarningsClickOk(15);

                bool imageFound = imageExists(dirName + @"\cassTemplate_images\sspEmptySearch.PNG", 0.98);
                debug(TRACE, "image found is - "+ imageFound);
                if (!imageFound)
                {
                    debug(TRACE, "Student Found");
                    studentId = id;
                    break;
                }
                else
                {
                    debug(TRACE, "Student search returned no results");
                }
            }

            // click the first result
            imageFindClick(dirName + @"\cassTemplate_images\windowsImages\okButtonBlackBorder.PNG");
            imageVanish(dirName + @"\cassTemplate_images\syncImages\studentStudyPackageSYNC.PNG", 15);
            imageWait(dirName + @"\cassTemplate_images\linkImages\saveAction.PNG", 10);

            // Change SSP Stage to Enrolled
            leftClick(437, 427);
            typeTextInField("Enrolled" + @"{TAB}");

            // Change SSP Status to enrolled
            leftClick(437, 452);
            typeTextInField("Enrolled" + @"{TAB}");

            // save
            saveAction(120);
            sleep(1);

            return studentId;
        }

        public bool confirmSubjectEnrolmentInCanvas(string studentId)
        {

            bool activeRecordfound = false;

            for (int count = 0; count < 10; count++)
            {
                //Refresh canvas application
                refresh();

                //Wait for the table with students list to load conpletely
                waitForElement("th", "text", "Name", 15);

                //Find 'search people' searchbox and enter the studentId
                ReadOnlyCollection<IWebElement> peopleSearch = findElements(null, By.XPath("//input[@type = 'text' and @placeholder = 'Search people']"));
                type(peopleSearch[0], studentId, true);

                //Wait for the search result table to load
                try
                {
                    waitForElement("*", "text", studentId, 15);

                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                    debug(TRACE, "SIS Id did not appear.");

                }

                //Check if the search result gives record NOT found or an Active record 
                if (IsElementVisible(By.XPath("//h2[contains(text(), 'No people found')]")))
                {

                    debug(TRACE, "################# Student NOT found  ######################");

                }
                else
                {
                    sleep(3);
                    //Find if the record found is inactive
                    ReadOnlyCollection<IWebElement> inactiveRecord = findElements(null, By.XPath("//span[@class = 'label' and contains(text(), 'inactive')]"));
                    Console.WriteLine("inactive count:" + inactiveRecord.Count);
                    if (inactiveRecord.Count == 0)
                    {
                        activeRecordfound = true;
                        Console.WriteLine("active found, Enrolment flown to canvas");
                        break;
                    }

                }
                sleep(20);
            }

            return activeRecordfound;
        }


    }
}