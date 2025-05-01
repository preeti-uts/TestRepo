using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace selenium_framework
{
    public partial class Selenium_Framework
    {
        //page elements - Manage Course Page
        private By courseSearch = By.XPath("//*[@id='btnCourseSearch']");
        private By selectYrDropdown = By.XPath("//*[@id='multiSelectSearchCntrl_dropdownYears']");
        private By enterCurrentYr = By.XPath("//*[@placeholder='Please search for years']");
        private By currentYrCheckbox = By.XPath("//span[contains(text(), '2024')]");
        private By selectSessionDropdown = By.XPath("//*[@id='multiSelectSearchCntrl_dropdownSessions']");
        private By enterCurrentSession = By.XPath("//*[@placeholder='Please search for sessions']");
        private By currentSessionCheckbox = By.XPath("//span[text() = ' Autumn Session ']");
        private By selectStatusDropdown = By.XPath("//*[@id='multiSelectSearchCntrl_dropdownStatus']");
        private By enterStatus1 = By.XPath("//*[@placeholder='Please search for statuses']");
        private By status1Checkbox = By.XPath("//span[text() = ' Complete ']");
        private By status2Checkbox = By.XPath("//span[text() = ' Conditionally complete ']");
        private By searchByKeyword = By.XPath("//*[@id='txtCourseSearch']");
        private By searchButton = By.XPath("//*[@id='btnCourseSearch']");
        private string progressBar = "//*[@role='progressbar']";

        public String courseSearchInCAT(String subjectCode, String currentSession, String currentYr)
        {
            debug(TRACE, "\r\n*** Test start: " + this.GetType().ToString() + "\r\n");

            // Find and click select year dropdown
            ReadOnlyCollection<IWebElement> selectYrDropdown = findElements(null, this.selectYrDropdown);
            click(selectYrDropdown[0]);

            //Find and enter the year
            ReadOnlyCollection<IWebElement> enterCurrentYr = findElements(null, this.enterCurrentYr);
            type(enterCurrentYr[0], currentYr, false);

            // Find and click current year checkbox
            ReadOnlyCollection<IWebElement> currentYrCheckbox = findElements(null, this.currentYrCheckbox);
            click(currentYrCheckbox[0]);

            //Hit escape to come out of the Select year dropdown
            type(Keys.Escape);

            // Find and click select Session dropdown
            ReadOnlyCollection<IWebElement> selectSessionDropdown = findElements(null, this.selectSessionDropdown );
            click(selectSessionDropdown[0]);

            //Find and enter the session
            ReadOnlyCollection<IWebElement> enterCurrentSession = findElements(null, this.enterCurrentSession );
            type(enterCurrentSession[0], currentSession, false);

            // Find and click current session checkbox
            ReadOnlyCollection<IWebElement> currentSessionCheckbox = findElements(null, this.currentSessionCheckbox);
            click(currentSessionCheckbox[0]);

            //Hit escape to come out of the Select Session dropdown
            type(Keys.Escape);

            // Find and click select status dropdown
            ReadOnlyCollection<IWebElement> selectStatusDropdown = findElements(null, this.selectStatusDropdown);
            click(selectStatusDropdown[0]);
            sleep(2);

            //Find and enter the status1 - Complete
            ReadOnlyCollection<IWebElement> enterStatus1 = findElements(null, this.enterStatus1 );
            type(enterStatus1[0], "Complete", true);

            sleep(2);

            // Find and click 'status1 - Complete' checkbox
            ReadOnlyCollection<IWebElement> status1Checkbox = findElements(null, this.status1Checkbox );
            click(status1Checkbox[0]);

            sleep(2);

            // Find and click 'status2 - Conditionally complete' checkbox
            ReadOnlyCollection<IWebElement> status2Checkbox = findElements(null,  this.status2Checkbox);
            click(status2Checkbox[0]);

            // Hit escape to come out of the Select Status dropdown
            type(Keys.Escape);

            // Find and click Search by keyword field
            ReadOnlyCollection<IWebElement> searchBykeyword = findElements(null, this.searchByKeyword);
            click(searchBykeyword[0]);
            type(searchBykeyword[0], subjectCode, true);

            // Find and click Search button
            ReadOnlyCollection<IWebElement> searchButton = findElements(null, this.searchButton);
            click(searchButton[0]);

            // Wait for element
            waitForElement("*", "text", "SISID", 20);

            sleep(2);

            // get course table and its rows 
            IWebElement courseTable = getElement("table", "id", "tblCourse");
            //ReadOnlyCollection<IWebElement> courseRowData = getTableRows(courseTable);

            //Fetch SISID of the first data row
            String sisID = getTableCell(courseTable, 1, 1).Text;
            Console.WriteLine(sisID);

            //Find and Click on the course name in first row
            IWebElement courseName = getTableCell(courseTable, 1, 2);
            click(courseName);

            //wait for the progress bar  
            //bool progressBarFound = waitForElement("", "role", "progressbar", 5);

            waitForElementVisible( progressBar, 10);

            waitForElementNotVisible(progressBar, 40);

            //Get CASS subjects or courses table
            IWebElement subjectCourseTable = getElement("table", "id", "selectedSubjectTable");
            ReadOnlyCollection<IWebElement> subjectCourseRowData = getTableRows(subjectCourseTable);

            int numberOfSubjectRecords = subjectCourseRowData.Count();

            Console.WriteLine("No of rows in subject table: " + numberOfSubjectRecords);

            //Check if subjects or courses table is having atleast one row 

            if (numberOfSubjectRecords >= 1)
            {
                //subjectCode = getTableCell(subjectCourseTable, 1, 2).Text;
                debug(TRACE, "############ Subject present in CAT  ################" + sisID);
                return sisID;
            }
            else
            {
                return "";
            }

        }
    }
}
