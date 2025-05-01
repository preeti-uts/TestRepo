using ExcelDataReader;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.UI.WebControls;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media.TextFormatting;

namespace selenium_framework
{
    // Common library for frequently used functionality for reuse across scripts
    public partial class Selenium_Framework
    {
        // CASS user login details
        public static string userLoginName = "";
        public static string userLoginPassword = "";
        public static string userLoginAnswerToSecurityQuestionAnswer = "";

        public static string userLoginNameTwo = "";
        public static string userLoginPasswordTwo = "";
        public static string userLoginAnswerToSecurityQuestionAnswerTwo = "";

        public static string userLoginNameThree = "";
        public static string userLoginPasswordThree = "";
        public static string userLoginAnswerToSecurityQuestionAnswerThree = "";

        // Canvas user login details
        public static string canvasUserLoginName = "";
        public static string canvasUserLoginPassword = "";
        public static string canvasUserLoginAnswerToSecurityQuestionAnswer = "";

        //CAT application login details
        public static string catUserLoginName = "";
        public static string catUserLoginPassword = "";
        
        // url
        //public static string cassURL = "https://uts-test.t1cloud.com/"; // CCTEST
        public static string cassURL = "https://uts-uat.t1cloud.com/"; //CCUAT
        //public static string cassURL = "https://uts.t1cloud.com/"; //CCPROD

        //CAT UAT URL
        public static string catURL = "https://cat-uat.uts.edu.au/";

        //Canvas Test URL
        public static string canvasURL = "https://uts.test.instructure.com/";
        
        private static DataSet excelSheetTestData;

        // Overloaded function that determines the userLoginSecurityQuestionAnswer based on loginName
        public void loginCASSCloud(string cassURL, string loginName, string loginPassword)
        {

            string loginAnswerToSecretQuestion = "";

            if (loginName == userLoginName)
            {
                loginAnswerToSecretQuestion = userLoginAnswerToSecurityQuestionAnswer;
            }
            else if (loginName == userLoginNameTwo)
            {
                loginAnswerToSecretQuestion = userLoginAnswerToSecurityQuestionAnswerTwo;
            }
            else if (loginName == userLoginNameThree)
            {
                loginAnswerToSecretQuestion = userLoginAnswerToSecurityQuestionAnswerThree;
            }
            else if (loginName == canvasUserLoginName)
            {
                loginAnswerToSecretQuestion = canvasUserLoginAnswerToSecurityQuestionAnswer;
            }
            else
            {
                debug(ERROR, "loginCASSCloud: test account email address was not found");
            }
            loginCASSCloud(cassURL, loginName, loginPassword, loginAnswerToSecretQuestion);
        }

        // ------------- New Method ---------------
        // Login to CASS Cloud (Okta)
        public void loginCASSCloud(string cassURL, string loginName, string loginPassword, string loginAnswerToSecretQuestion)
        {
            debug(TRACE, "loginCASSCloud: start");

            loginWithOkta(cassURL, loginName, loginPassword, loginAnswerToSecretQuestion);

            sleep(2);
            JSActive();

            // Wait for page to load
            waitForElement("*", "text", "Ci in the Cloud", 15);
            sleep(2);
        }

        // ------------- Old Method ---------------
        //Login to CASS Cloud (Standard Authentication)
        //public void loginCASSCloud(string cassURL, string userLoginName, string userLoginPassword)
        //{
        //    debug(TRACE, "loginStartStudentManagement: start");

        //    gotoUrl(cassURL);

        //    try
        //    {

        //        ReadOnlyCollection<IWebElement> userField = findElements(null, By.Id("LogonName"));
        //        type(userField[0], userLoginName);

        //        ReadOnlyCollection<IWebElement> pwField = findElements(null, By.Name("Password"));
        //        type(pwField[0], userLoginPassword, false);

        //        ReadOnlyCollection<IWebElement> loginButton = findElements(null, By.Name("logOnButton"));
        //        click(loginButton[0]);

        //    }
        //    catch (Exception e)
        //    {
        //        debug(ERROR, "loginStartStudentManagement: failed to login. Exception: \r\n" + e.Message.ToString() + "\r\n");
        //    }
        //    sleep(2);
        //    JSActive();

        //    // wait for page to load
        //    waitForElement("*", "text", "Ci in the Cloud", 15);
        //    sleep(2);
        //}

        // Log into the CASS Cloud website and start up student management
        public void loginStartStudentManagement(string cassURL, string userLoginName, string userLoginPassword)
        {
            debug(TRACE, "loginStartStudentManagement: start");

            loginCASSCloud(cassURL, userLoginName, userLoginPassword);

            ReadOnlyCollection<IWebElement> studentManagement = findElements(null, By.XPath("//span [text()='Student Management']"));
            if (studentManagement.Count != 0)
            {
                click(studentManagement[0]);
            }
            else
            {
                // If didn't find student management link, find and expand the Ci in the cloud tile
                ReadOnlyCollection<IWebElement> ciInTheCloud = findElements(null, By.XPath("//a [@aria-description='Ci in the Cloud']"));
                click(ciInTheCloud[0]);
                //refresh();
                sleep(1);
                studentManagement = findElements(null, By.XPath("//span [text()='Student Management']"));
                click(studentManagement[0]);
            }
            syncWindowCount(2);
            popupWindow(1);
        }

        // Logout of CASS Cloud
        public void logoutCASSCloud()
        {
            debug(TRACE, "logoutCASSCloud: start");

            bool loginPageFound = false;
            bool homePageFound = false;
            int elapsedSeconds = 0;

            // Click on the user profile button
            ReadOnlyCollection<IWebElement> userIcon = findElements(null, By.ClassName("userDetails"));
            click(userIcon[0]);

            // Select Log Off
            ReadOnlyCollection<IWebElement> logoutLink = findElements(null, By.XPath("//a [text()='Log Off']"));
            click(logoutLink[0]);

            ReadOnlyCollection<IWebElement> loginPage;
            ReadOnlyCollection<IWebElement> HomePage;

            // Wait for the login screen to appear
            while (elapsedSeconds < 15)
            {
                loginPage = findElements(null, By.XPath("//*[@id ='LogonName']"));
                if (loginPage.Count > 0)
                {
                    loginPageFound = true;
                    break;
                }
                sleep(1);
                elapsedSeconds++;
            }

            // Wait for the home page screen to appear. Since SSO has been implemented, when you logout, it does not redirect to the login page anymore, instead it redirects you to the home page
            if (!loginPageFound)
            {
                elapsedSeconds = 0;
                while (elapsedSeconds < 15)
                {
                    HomePage = findElements(null, By.XPath("//*[text() ='Ci in the Cloud']"));
                    if (HomePage.Count > 0)
                    {
                        homePageFound = true;
                        break;
                    }
                    sleep(1);
                    elapsedSeconds++;
                }
            }

            if (loginPageFound)
            {
                debug(TRACE, "logoutCASSCloud: Logout successful. Login page found");
            }
            if (homePageFound)
            {
                debug(TRACE, "logoutCASSCloud: Logout successful. Home page found.");
            }
            else
            {
                debug(ERROR, "logoutCASSCloud: Logout not successful. Neither login nor CiAnywhere Home Page found.");
            }
        }

        // Exit student management - no checks for pop ups as yet, add as they occur
        public void exitStudentManagement(string imagePath)
        {
            debug(TRACE, "exitSrudentManagement: start");

            imageFindClick(imagePath + @"\cassTemplate_images\menuFile.PNG");
            imageWait(imagePath + @"\cassTemplate_images\MenuExit.PNG", 60);
            imageFindClick(imagePath + @"\cassTemplate_images\MenuExit.PNG");
            sleep(2);

            // check to see if save button appears
            int timeout = 0;
            bool saveDialogPopup = imageExists(dirName + @"\cassTemplate_images\windowsImages\noButton.PNG");

            // if pop up already displayed, close it.
            if (saveDialogPopup)
            {
                imageWait(dirName + @"\cassTemplate_images\windowsImages\noButton.PNG", 30);
                imageFindClick(dirName + @"\cassTemplate_images\windowsImages\noButton.PNG");
            }

            // If pop up not found, wait and see if it does pop up
            while (!saveDialogPopup && timeout < 10)
            {
                saveDialogPopup = imageExists(dirName + @"\cassTemplate_images\windowsImages\noButton.PNG");
                if (saveDialogPopup)
                {
                    imageWait(dirName + @"\cassTemplate_images\windowsImages\noButton.PNG", 30);
                    imageFindClick(dirName + @"\cassTemplate_images\windowsImages\noButton.PNG");
                }
                sleep(1);
                timeout++;
            }

        }

        // Wait for the student management splash screen and select the workplace / role
        public void selectWorkplace(string workplaceName)
        {
            debug(TRACE, "selectWorkplace: start");

            bool searchField = false;
            bool textFound = false;
            LocationValues searchArea = new LocationValues();
            searchArea.X = 660;
            searchArea.Y = 370;
            searchArea.W = 150;
            searchArea.H = 120;

            // wait for the splash screen to appear
            imageWait(dirName + @"\cassTemplate_images\\workplaceIcons\pleaseChooseaWorkplace.PNG", 120);

            // Check if search button exists
            searchField = imageExists(dirName + @"\cassTemplate_images\\workplaceIcons\workplaceSearch.PNG");
            if (searchField)
            {
                // type in the name to show the workplace
                leftClick(716, 617);
                sleep(1);
                typeTextInField(workplaceName);
                sleep(2);
            }
            workplaceName = workplaceName.Replace("{", "");
            workplaceName = workplaceName.Replace("}", "");

            int rowOneCount = 0;
            int rowTwoCount = 0;
            int loopCount = 0;
            while (!textFound && loopCount <= 10)
            {
                tesseractCaptureScreen(searchArea.W, searchArea.H, searchArea.X, searchArea.Y, dirName + @"\cassTemplate_images\tempImages\chooseWorkplace.TIFF");
                textFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\chooseWorkplace.TIFF", "block", true, 2, searchArea.X, searchArea.Y, "single", "contains", workplaceName, false);
                if (!textFound)
                {
                    debug(INFO, "chooseWorkplace: tesseract - scale 2 - invert");
                    textFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\chooseWorkplace.TIFF", "block", true, 2, searchArea.X, searchArea.Y, "single", "contains", workplaceName, true);

                    if (!textFound)
                    {
                        debug(INFO, "chooseWorkplace: tesseract - no scale - invert");
                        textFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\chooseWorkplace.TIFF", "block", false, 2, searchArea.X, searchArea.Y, "single", "contains", workplaceName, true);

                        if (!textFound)
                        {
                            debug(INFO, "chooseWorkplace: tesseract - no scale - no invert");
                            textFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\chooseWorkplace.TIFF", "block", false, 2, searchArea.X, searchArea.Y, "single", "contains", workplaceName, false);

                            if (!textFound)
                            {
                                debug(INFO, "chooseWorkplace: tesseract - scale 3 - no invert");
                                textFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\chooseWorkplace.TIFF", "block", true, 3, searchArea.X, searchArea.Y, "single", "contains", workplaceName, false);

                                if (!textFound)
                                {
                                    debug(INFO, "chooseWorkplace: tesseract - scale 3 - invert");
                                    textFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\tempImages\chooseWorkplace.TIFF", "block", true, 3, searchArea.X, searchArea.Y, "single", "contains", workplaceName, true);
                                }
                            }
                        }
                    }
                }
                if (!textFound && (rowOneCount < 4))
                {
                    searchArea.X = searchArea.X + 155;
                    searchArea.Y = 370;
                    searchArea.W = 150;
                    searchArea.H = 120;

                    rowOneCount++;
                }

                if (!textFound && (rowTwoCount < 4) && (rowOneCount == 4))
                {
                    if (rowTwoCount == 0)
                    {
                        searchArea.X = 660;
                    }
                    else
                    {
                        searchArea.X = searchArea.X + 155;
                    }
                    searchArea.Y = 505;
                    searchArea.W = 150;
                    searchArea.H = 120;

                    rowTwoCount++;
                }
                loopCount++;
            }
            // Check if text found and clicked
            if (textFound)
            {
                debug(TRACE, "chooseWorkplace: workplace found: " + workplaceName);
            }
            else
            {
                debug(ERROR, "chooseWorkplace: workplace not found: " + workplaceName);
            }

        }

        // Click save action to perform save and check for warning pop up
        public void saveAction()
        {
            saveAction(60);
        }

        public void saveAction(int timeoutValue)
        {
            debug(TRACE, "saveAction: start");

            int loopCount = 0;
            bool saveComplete = false;
            bool warningMessage = false;
            bool errorMessage = false;
            bool authPopUp = false;

            imageWait(dirName + @"\cassTemplate_images\saveAction.PNG", 5);
            imageFindClick(dirName + @"\cassTemplate_images\saveAction.PNG");

            saveComplete = imageExists(dirName + @"\cassTemplate_images\saveCompleteIcon.PNG");

            // if save completion icon doesnt exist, keep checking for warning messages
            while (!saveComplete && loopCount <= timeoutValue)
            {
                // Check the current state
                saveComplete = imageExists(dirName + @"\cassTemplate_images\saveCompleteIcon.PNG");
                warningMessage = imageExists(dirName + @"\cassTemplate_images\warningList.PNG");
                errorMessage = imageExists(dirName + @"\cassTemplate_images\syncImages\errorIcon.PNG");
                authPopUp = imageExists(dirName + @"\cassTemplate_images\syncImages\revisionAuthDetailSYNC.PNG");


                if (warningMessage)
                {
                    // Check for Yes radio button
                    if (imageExists(dirName + @"\cassTemplate_images\checkBox\warningYES.PNG"))
                    {
                        imageFindClick(dirName + @"\cassTemplate_images\checkBox\warningYES.PNG");
                        sleep(1);
                    }
                    typeTextInField(@"{ENTER}");
                    imageVanish(dirName + @"\cassTemplate_images\warningList.PNG", 5);
                }

                // Deal with warnings
                if (errorMessage)
                {
                    debug(ERROR, "saveAction: error icon on screen, see screenshot.");
                }

                // Deal with revsion auth message
                if (authPopUp)
                {
                    debug(TRACE, "saveAction: Revision authorisation detail pop up, break out of save.");
                    break;
                }

                sleep(1);

                loopCount++;
            }

            if (!saveComplete && loopCount >= timeoutValue)
            {
                debug(ERROR, "saveAction: save took longer then timeout value: " + timeoutValue + " seconds.");
            }

            debug(TRACE, "saveAction: save took : " + loopCount + " seconds.");
        }

        // Waits for a manual save to complete (ie doesn't click the save link) for when other pop up actions are required before save completes
        public void waitForSaveComplete(int timeoutValue)
        {
            debug(TRACE, "waitForSaveComplete: start");

            int loopCount = 0;
            bool saveComplete = false;
            bool warningMessage = false;
            bool errorMessage = false;
            bool authPopUp = false;

            //imageWait(dirName + @"\cassTemplate_images\saveAction.PNG", 5);
            //imageFindClick(dirName + @"\cassTemplate_images\saveAction.PNG");

            saveComplete = imageExists(dirName + @"\cassTemplate_images\saveCompleteIcon.PNG");

            // if save completion icon doesnt exist, keep checking for warning messages
            while (!saveComplete && loopCount <= timeoutValue)
            {
                // Check the current state
                saveComplete = imageExists(dirName + @"\cassTemplate_images\saveCompleteIcon.PNG");
                warningMessage = imageExists(dirName + @"\cassTemplate_images\warningList.PNG");
                errorMessage = imageExists(dirName + @"\cassTemplate_images\syncImages\errorIcon.PNG");
                authPopUp = imageExists(dirName + @"\cassTemplate_images\syncImages\revisionAuthDetailSYNC.PNG");


                if (warningMessage)
                {
                    // Check for Yes radio button
                    if (imageExists(dirName + @"\cassTemplate_images\checkBox\warningYES.PNG"))
                    {
                        imageFindClick(dirName + @"\cassTemplate_images\checkBox\warningYES.PNG");
                        sleep(1);
                    }
                    typeTextInField(@"{ENTER}");
                    imageVanish(dirName + @"\cassTemplate_images\warningList.PNG", 5);
                }

                // Deal with warnings
                if (errorMessage)
                {
                    debug(ERROR, "saveAction: error icon on screen, see screenshot.");
                }

                // Deal with revsion auth message
                if (authPopUp)
                {
                    debug(TRACE, "saveAction: Revision authorisation detail pop up, break out of save.");
                    break;
                }

                sleep(1);

                loopCount++;
            }

            if (!saveComplete && loopCount >= timeoutValue)
            {
                debug(ERROR, "waitForSaveComplete: save took longer then timeout value: " + timeoutValue + " seconds.");
            }

            debug(TRACE, "waitForSaveComplete: save took : " + loopCount + " seconds.");
        }
        // // Use save and clear action
        public void saveClearAction(int timeoutValue)
        {
            saveClearAction(timeoutValue, dirName + @"\cassTemplate_images\saveCompleteIcon.PNG");
        }

        public void saveClearAction(int timeoutValue, string syncImage)
        {
            debug(TRACE, "saveAction: start");

            int loopCount = 0;
            bool saveComplete = false;
            bool warningMessage = false;
            bool errorMessage = false;

            imageWait(dirName + @"\cassTemplate_images\saveClearLink.PNG", 5);
            imageFindClick(dirName + @"\cassTemplate_images\saveClearLink.PNG");

            saveComplete = imageExists(syncImage);

            // if save completion icon doesnt exist, keep checking for warning messages
            while (!saveComplete && loopCount <= timeoutValue)
            {
                // Check the current state
                saveComplete = imageExists(syncImage);
                warningMessage = imageExists(dirName + @"\cassTemplate_images\warningList.PNG");
                errorMessage = imageExists(dirName + @"\cassTemplate_images\syncImages\errorIcon.PNG");

                if (warningMessage)
                {
                    typeTextInField(@"{ENTER}");
                    imageVanish(dirName + @"\cassTemplate_images\warningList.PNG", 5);
                }
                if (errorMessage)
                {
                    debug(ERROR, "saveAction: error icon on screen, see screenshot.");
                }
                sleep(1);

                loopCount++;
            }

            if (!saveComplete && loopCount >= timeoutValue)
            {
                debug(ERROR, "saveAction: save took longer then timeout value: " + timeoutValue + " seconds.");
            }

            debug(TRACE, "saveAction: save took : " + loopCount + " seconds.");
        }

        // Click load search from the standard student search screen, with 3 search criteriasni
        public void clickLoadSearch()
        {
            debug(TRACE, "clickLoadSearch: start");

            leftClick(1020, 481);
            imageWait(dirName + @"\cassTemplate_images\genericTemplatePicklistSYNC.PNG", 20);
            sleep(1);
        }

        // Click save search from the standard student search screen, with 3 search criteria
        public void clickSaveSearch()
        {
            debug(TRACE, "clickSaveSearch: start");

            leftClick(942, 461);
            // to do add sync
        }

        // Click clear values from the standard student search screen, with 3 search criteria
        public void clickClearValues()
        {
            debug(TRACE, "clickClearValues: start");

            leftClick(775, 461);
            // to do add sync
        }

        // Search for a CASS function using the sidebar search
        public void sideBarSearchFunction(string functionToSearch)
        {
            sideBarSearchFunction(functionToSearch, "contains");
        }

        public void sideBarSearchFunction(string functionToSearch, string textComparisonType)
        {
            sideBarSearchFunction(functionToSearch, textComparisonType, true);
        }

        public void sideBarSearchFunction(string inFunctionToSearch, string textComparisonType, bool scaleResults)
        {
            debug(TRACE, "sideBarSearchFunction: start");

            bool textWasFound;
            // Check if empty table row visable
            LocationValues searchBarArea = new LocationValues();
            searchBarArea.X = 72;
            searchBarArea.Y = 281;
            searchBarArea.W = 237 - searchBarArea.X;
            searchBarArea.H = 1000 - searchBarArea.Y;

            bool reachedEndOfPage = false;
            bool searchResultsScrollBarExists;

            string functionToSearch = inFunctionToSearch;

            // Move mouse to make sure it isn't over the side bar search already
            MouseMove(200, 100);

            // Click the search side button
            imageWait(dirName + @"\cassTemplate_images\linkImages\sideBarSearchLink.PNG", 15);
            imageFindClick(dirName + @"\cassTemplate_images\linkImages\sideBarSearchLink.PNG");
            imageWait(dirName + @"\cassTemplate_images\searchFunctionSYNC.PNG", 15);

            // type in the text to search
            typeTextInField(@functionToSearch + "{ENTER}");
            sleep(1);

            // Remove and curly brackets ({, }) from text so we can ocr the search results to click
            if (inFunctionToSearch.Contains("{"))
            {
                functionToSearch = removeNonAlphaChars(inFunctionToSearch);
                //functionToSearch = inFunctionToSearch.Replace("{", "");
                //functionToSearch = functionToSearch.Replace("}", "");
            }
            else
            {
                functionToSearch = inFunctionToSearch;
            }

            // wait for search area to have results
            imageVanish(dirName + @"\cassTemplate_images\syncImages\sidebarSearchSYNC.PNG", 200);
            sleep(1);

            // Initial check if the specified function is found using the sidebar search.
            textWasFound = findAndClickTextInSideBar(functionToSearch, textComparisonType, scaleResults, searchBarArea);

            // Check if a scrollbar exists in the search results area
            searchResultsScrollBarExists = imageExistsArea(dirName + @"\cassTemplate_images\searchResultsScrollDown.PNG", 0.95, searchBarArea);

            // If a scrollbar exists, attempt to scroll through the search results until the text is found or until the end
            if (searchResultsScrollBarExists)
            {
                while (!textWasFound && !reachedEndOfPage)
                {
                    debug(TRACE, "sideBarSearchFunction_new: Text not found. Attempting to scroll down and find the function");

                    // Perform a scroll in the sidebar
                    if (!textWasFound)
                    {
                        leftClick(83, 287);
                        scrollSearchResultsScreen(searchBarArea);
                        //leftClick(83, 287);
                    }
                    // Check if the specified function is found using the sidebar search
                    textWasFound = findAndClickTextInSideBar(functionToSearch, textComparisonType, scaleResults, searchBarArea);

                    // Check if the end of the search results page has been reached
                    reachedEndOfPage = imageExistsArea(dirName + @"\cassTemplate_images\searchResultsScrollDownEndOfList.PNG", 0.95, searchBarArea);
                }
            }

            // If the text is still not found after all attempts, log an error
            if (!textWasFound)
            {
                debug(ERROR, "sideBarSearchFunction_new: text not found after scrolling through the search results.");
            }

            // wait for search to vanish (assume page is loaded....)
            imageVanish(dirName + @"\cassTemplate_images\searchFunctionSYNC.PNG", 60);
        }

        // findAndClickTextInSideBar
        public bool findAndClickTextInSideBar(string functionToSearch, string textComparisonType, bool scaleResults, LocationValues searchBarArea)
        {
            bool textWasFound;

            // Screen coords of the search area for sidebar search
            int myGetX = searchBarArea.X;
            int myGetY = searchBarArea.Y;
            int myGetW = searchBarArea.W;
            int myGetH = searchBarArea.H;

            // find and click text
            tesseractCaptureScreen(myGetW, myGetH, myGetX, myGetY, dirName + @"\cassTemplate_images\resultsScreenshot.TIFF");
            if (scaleResults)
            {
                debug(INFO, "sideBarSearchFunction_new: Scale value for call was '" + scaleResults + "'");
                textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\resultsScreenshot.TIFF", "TextLine", scaleResults, 2, myGetX, myGetY, "double", textComparisonType, functionToSearch, false);

                if (!textWasFound)
                {
                    debug(INFO, "sideBarSearchFunction_new: text not found. Scale value for call was '" + scaleResults + "', trying scale result false");
                    textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\resultsScreenshot.TIFF", "TextLine", false, 2, myGetX, myGetY, "double", textComparisonType, functionToSearch, false);

                    if (!textWasFound)
                    {
                        debug(INFO, "sideBarSearchFunction_new: text not found. Scale value for call was '" + scaleResults + "', trying scale result false, inverse");
                        textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\resultsScreenshot.TIFF", "TextLine", false, 2, myGetX, myGetY, "double", textComparisonType, functionToSearch, true);

                        if (!textWasFound)
                        {
                            debug(INFO, "sideBarSearchFunction_new: text not found. Scale value for call was '" + scaleResults + "', trying scale result true, inverse ");
                            textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\resultsScreenshot.TIFF", "TextLine", true, 2, myGetX, myGetY, "double", textComparisonType, functionToSearch, true);

                            if (!textWasFound)
                            {
                                debug(INFO, "sideBarSearchFunction_new: text not found. Scale value for call was '" + scaleResults + "', trying scale result true, inverse ");
                                textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\resultsScreenshot.TIFF", "TextLine", true, 3, myGetX, myGetY, "double", textComparisonType, functionToSearch, true);
                                if (!textWasFound)
                                {
                                    debug(INFO, "sideBarSearchFunction_new: text not found. Scale value for call was '" + scaleResults + "', trying scale result true, inverse ");
                                    textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\resultsScreenshot.TIFF", "TextLine", true, 3, myGetX, myGetY, "double", textComparisonType, functionToSearch, false);
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                debug(INFO, "sideBarSearchFunction_new: Scale value for call was '" + scaleResults + "'");
                textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\resultsScreenshot.TIFF", "TextLine", false, 2, myGetX, myGetY, "double", textComparisonType, functionToSearch, false);

                if (!textWasFound)
                {
                    debug(INFO, "sideBarSearchFunction_new: text not found. Scale value for call was '" + scaleResults + "', trying scale result true");
                    textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\resultsScreenshot.TIFF", "TextLine", true, 2, myGetX, myGetY, "double", textComparisonType, functionToSearch, false);

                    if (!textWasFound)
                    {
                        debug(INFO, "sideBarSearchFunction_new: text not found. Scale value for call was '" + scaleResults + "', trying scale result true");
                        textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\resultsScreenshot.TIFF", "TextLine", true, 2, myGetX, myGetY, "double", textComparisonType, functionToSearch, true);
                        if (!textWasFound)
                        {
                            debug(INFO, "sideBarSearchFunction_new: text not found. Scale value for call was '" + scaleResults + "', trying scale result true, inverse ");
                            textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\resultsScreenshot.TIFF", "TextLine", true, 3, myGetX, myGetY, "double", textComparisonType, functionToSearch, true);
                        }
                    }
                }
            }
            return textWasFound;
        }

        // Expand a study plan
        public void expandStudyPlanCourse(String studyPlanID, LocationValues locationToExtractText)
        {
            debug(TRACE, "expandStudyPlanCourse: start");

            bool expanded = false;
            int waitTimeout = 0;

            int myGetX = locationToExtractText.X;
            int myGetY = locationToExtractText.Y;
            int myGetH = locationToExtractText.H;
            int myGetW = locationToExtractText.W;

            //findandClickTextLine(myGetW, myGetH, myGetX, myGetY, studyPlanID, false, "single"); // This is not working for some spk cd since always same position use a left click
            findandClickText(myGetW, myGetH, myGetX, myGetY, studyPlanID, true);

            // Expand the course                                                                                  
            imageFindClick(dirName + @"\cassTemplate_images\linkImages\expandLink.PNG");

            // Check for any warning pop ups
            checkForWarningsClickOk(15);


            checkIfStudyPlanTemplateSearchExistsClickOk(35);

            // Wait for the expand link to change
            //imageVanish(dirName + @"\cassTemplate_images\linkImages\expandLink.PNG", 120);

            while (expanded && waitTimeout <= 120)
            {
                expanded = imageExists(dirName + @"\cassTemplate_images\linkImages\expandLink.PNG");
                checkIfWarningExistsClickOk();
                sleep(1);
                waitTimeout++;
            }
        }

        // Find and enrol in a subject
        public void enrolInSubject(string subjectID, LocationValues locationToExtractText)
        {
            debug(TRACE, "enrolInSubject: start");

            bool isSubjectAvailableBlank;
            string textExtractedInTable;
            bool subjectFoundInTable = false;
            bool tableScrolledDown = false;
            bool reachedEndOfTable = false;

            int myGetX = locationToExtractText.X;
            int myGetY = locationToExtractText.Y;
            int myGetH = locationToExtractText.H;
            int myGetW = locationToExtractText.W;

            while (!subjectFoundInTable && !reachedEndOfTable)
            {
                // Attempt to extract text and find the subject exists in the extracted text without scrolling through multiple methods
                textExtractedInTable = findandGetText(locationToExtractText);

                if (!isKeywordFoundInText(textExtractedInTable, subjectID))
                {
                    // Attempt to extract text with scaling
                    textExtractedInTable = findandGetText(locationToExtractText, true);
                    debug(TRACE, "enrolInSubject: findandGetText without scaling - subject not found");

                    if (!isKeywordFoundInText(textExtractedInTable, subjectID))
                    {
                        // Attempt to extract text with different scaling parameters
                        textExtractedInTable = findandGetText(myGetW, myGetH, myGetX, myGetY, true, 4, false);
                        subjectFoundInTable = isKeywordFoundInText(textExtractedInTable, subjectID);
                        debug(TRACE, "enrolInSubject: findandGetText with scaling x4 - " + (subjectFoundInTable ? "subject found" : "subject not found"));
                    }
                    else
                    {
                        subjectFoundInTable = true;
                        debug(TRACE, "enrolInSubject: findandGetText with scaling - subject found");
                    }
                }
                else
                {
                    subjectFoundInTable = true;
                    debug(TRACE, "enrolInSubject: findandGetText without scaling - subject found");
                }

                // Check if the end of the table has been reached
                reachedEndOfTable = imageExistsArea(dirName + @"\cassTemplate_images\windowsImages\studyPlanScrollDownEndOfTable.PNG", 0.95, locationToExtractText);

                // Scroll down once if the subject is not found and end of the table is not reached
                if (!subjectFoundInTable && !reachedEndOfTable)
                {
                    scrollScreenArea("down", locationToExtractText);
                    debug(TRACE, "enrolInSubject: scrolled occured");
                }
            }

            // If subject is found, click it
            if (subjectFoundInTable)
            {
                findandClickText(myGetW, myGetH, myGetX, myGetY, subjectID, true, 3);
                debug(TRACE, "enrolInSubject: subject code " + subjectID + " was found in the table");
            }
            else
            {
                // subject code was not found after scrolling until the end of the table
                debug(ERROR, "enrolInSubject: subject code " + subjectID + " was not found in the table after scrolling down");
            }

            // Find and click the enrol link
            imageWait(dirName + @"\cassTemplate_images\linkImages\enrolLink.PNG", 5);
            imageFindClick(dirName + @"\cassTemplate_images\linkImages\enrolLink.PNG");

            // Wait for pop up
            imageWait(dirName + @"\cassTemplate_images\syncImages\studyPlanEnrolmentSYNC.PNG", 30);

            // Add a check if available
            isSubjectAvailableBlank = imageExists(dirName + @"\cassTemplate_images\syncImages\studyPlanEnrolmentCheckAvailability.PNG");

            if (isSubjectAvailableBlank)
            {
                imageFindClick(dirName + @"\cassTemplate_images\linkImages\cancelButtonLink.PNG");
                debug(WARNING, "Subject Code: " + subjectID + " was not available for enrolment");
            }
            else
            {
                imageFindClick(dirName + @"\cassTemplate_images\linkImages\okButtonLink.PNG");
                try
                {
                    imageWait(dirName + @"\cassTemplate_images\warningList.PNG", 5, false, false);
                    SendKeys.SendWait(@"{ENTER}");
                    imageVanish(dirName + @"\cassTemplate_images\warningList.PNG", 5);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                    debug(TRACE, "saveAction: warning list dialog did not appear.");
                }
            }

            imageVanish(dirName + @"\cassTemplate_images\syncImages\studyPlanEnrolmentSYNC.PNG", 30);
            sleep(1);

            // Don't think this is needed currently..
            //if (isSubjectAvailableBlank)
            //{
                //findandClickText(myGetW, myGetH, myGetX, myGetY, subjectID, true, 3);
                //findandClickTextLine(myGetW, myGetH, myGetX, myGetY, subjectID, true, 3, "single");
                //findandClickTextLine(108, 297, 455, 415, subjectID, true, "single");
            //}

            // Reset the scroll bar to the top
            if (subjectFoundInTable)
            {
                scrollStudyPlanTableToStart(locationToExtractText);
            }
        }

        // Check if the keyword is found in text
        private bool isKeywordFoundInText(string text, string keyword)
        {
            return text.ToLower().Contains(keyword.ToLower());
        }

        // Scroll the study plan table to the start (top)
        public bool scrollStudyPlanTableToStart(LocationValues studyPlanTableLocation)
        {
            debug(TRACE, "scrollStudyPlanTableToStart: start");

            bool imageFound = false;
            int attempts = 0;
            int clickOffset = 0;
            string scrollBarToUse = "";

            scrollBarToUse = dirName + @"\cassTemplate_images\\windowsImages\studyPlanScrollUp.PNG";
            clickOffset = 10;

            // Scroll down to Selection Status
            while (!imageFound && attempts <= 30)
            {
                imageWaitArea(scrollBarToUse, 1, false, false, 0.95, studyPlanTableLocation);
                imageFindClickArea(scrollBarToUse, 0, clickOffset, 0.95, studyPlanTableLocation);
                //sleep(1);

                MouseMove(1881, 675); // move mousde so not over scroll bar
                imageFound = imageExists(dirName + @"\cassTemplate_images\windowsImages\studyPlanTableStart.PNG");
                attempts++;
            }

            return imageFound;
        }

        // Scroll the study plan table to the end (bottom)
        public bool scrollStudyPlanTableToEnd(LocationValues studyPlanTableLocation)
        {
            debug(TRACE, "scrollStudyPlanTableToEnd: start");

            bool imageFound = false;
            int attempts = 0;
            int clickOffset = 0;
            string scrollBarToUse = "";

            scrollBarToUse = dirName + @"\cassTemplate_images\windowsImages\studyPlanDownArrow.PNG";
            clickOffset = -10;

            if (imageExists(dirName + @"\cassTemplate_images\windowsImages\studyPlanDownArrow.PNG"))
            {
                // Scroll down to Selection Status
                while (!imageFound && attempts <= 30)
                {
                    imageWaitArea(scrollBarToUse, 1, false, false, 0.95, studyPlanTableLocation);
                    imageFindClickArea(scrollBarToUse, 0, clickOffset, 0.95, studyPlanTableLocation);
                    sleep(1);

                    //MouseMove(1881, 675); // move mouse so not over scroll bar
                    leftClick(1881, 675); // Change to a mouse click, app not registering click after move
                    imageFound = imageExists(dirName + @"\cassTemplate_images\windowsImages\studyPlanTableEnd.PNG");
                    attempts++;
                }
            }
            debug(TRACE, "scrollStudyPlanTableToEnd: imageFound value: " + imageFound);
            return imageFound;
        }

        // Scroll any specified screen area
        public void scrollScreenArea(string direction, LocationValues screenArea)
        {
            scrollScreenArea(direction, 1, screenArea);
        }

        public void scrollScreenArea(string direction, int numberOfScrolls, LocationValues screenArea)
        {
            debug(TRACE, "scrollScreenArea: start");

            int clickOffset = 0;
            string scrollBarToUse = "";

            switch (direction.ToLower())
            {
                case "up":
                    scrollBarToUse = dirName + @"\cassTemplate_images\studyPlanScrollUp.PNG";
                    clickOffset = 20;
                    break;

                case "down":
                    scrollBarToUse = dirName + @"\cassTemplate_images\studyPlanScrollDown_new.PNG";
                    clickOffset = -20;
                    break;

                default:
                    debug(ERROR, "scrollScreenArea: invalid scroll direction selected. Sent value:" + direction);
                    break;
            }

            for (int i = 0; i < numberOfScrolls; i++)
            {
                imageFindClickArea(scrollBarToUse, 0, clickOffset, 0.95, screenArea);
                MouseMoveRelative(20, 20);
            }
        }

        // Scroll the study plan table to find a string value
        public bool scrollStudyPlanTable(string direction, string textToFind, LocationValues studyPlanTableLocation)
        {
            debug(TRACE, "scrollStudyPlanTable: start");

            bool textFound = false;
            int attempts = 0;
            int clickOffset = 0;
            string scrollBarToUse = "";
            string foundText = "";

            switch (direction.ToLower())
            {
                case "up":
                    scrollBarToUse = dirName + @"\cassTemplate_images\studyPlanScrollUp.PNG";
                    clickOffset = 20;
                    break;

                case "down":
                    scrollBarToUse = dirName + @"\cassTemplate_images\studyPlanScrollDown.PNG";
                    clickOffset = -20;
                    break;

                default:
                    debug(ERROR, "scrollStudyPlanTable: invalid scroll dirstion selected. Sent value:" + direction);
                    break;
            }

            foundText = findandGetText(studyPlanTableLocation);
            Console.WriteLine("foundText after scrolling attempt#"+ attempts);

            if (foundText.ToLower().Contains(textToFind.ToLower()))
            {
                textFound = true;
            }
            // Scroll down to Selection Status
            while (!textFound && attempts <= 8)
            {
                imageWaitArea(scrollBarToUse, 1, false, false, 0.95, studyPlanTableLocation);
                //imageFindClickArea(scrollBarToUse, 0, clickOffset, 0.95, studyPlanTableLocation);
                leftClick(1859, 711);
                sleep(2);
                leftClick(1859, 778);
                sleep(2);

                foundText = findandGetText(studyPlanTableLocation);
                if (foundText.ToLower().Contains(textToFind.ToLower()))
                {
                    textFound = true;
                }
                attempts++;
            }

            return textFound;
        }

        // Enter values for a field on the Student screen for creating / updating student general details (update fields as required)
        public void enterStudentField(string fieldName, string textValue)
        {
            debug(TRACE, "enterStudentField: start");

            // Check on student page
            imageWait(dirName + @"\cassTemplate_images\StudentSYNC.PNG", 5);

            switch (fieldName)
            {
                case "familyName":
                    leftClick(427, 386);
                    typeTextInField(textValue);
                    break;

                case "givenName":
                    leftClick(427, 408);
                    typeTextInField(textValue);
                    break;

                case "dateOfBirth":
                    leftClick(706, 361);
                    typeTextInField(textValue);
                    break;

                case "gender":
                    leftClick(706, 385);
                    typeTextInField(textValue);
                    break;

                default:
                    debug(ERROR, "Field name used not supported. Field name used : " + fieldName);
                    break;
            }
        }

        // Enter values for a field on the Application screen for creating Application general details (update fields as required)
        public void enterApplicationField(string fieldName, string textValue)
        {
            debug(TRACE, "enterApplicationField: start");

            // Check on application page
            imageWait(dirName + @"\cassTemplate_images\syncImages\applicationSYNC.PNG", 5);

            switch (fieldName)
            {
                case "liabilityCategory":
                    leftClick(411, 450);
                    typeTextInField(textValue);
                    break;

                case "loadCategory":
                    leftClick(411, 474);
                    typeTextInField(textValue);
                    break;

                case "attendanceMode":
                    leftClick(411, 497);
                    typeTextInField(textValue);
                    break;

                case "studyMode":
                    leftClick(411, 523);
                    typeTextInField(textValue);
                    break;

                default:
                    debug(ERROR, "Field name used not supported. Field name used : " + fieldName);
                    break;
            }
        }

        // Select clear form from the drop down menu
        public void clearForm()
        {
            debug(TRACE, "clearForm: start");

            imageWait(dirName + @"\cassTemplate_images\menuActions.PNG", 5);
            imageFindClick(dirName + @"\cassTemplate_images\menuActions.PNG");
            imageWait(dirName + @"\cassTemplate_images\menuClear.PNG", 5);
            imageFindClick(dirName + @"\cassTemplate_images\menuClear.PNG");
        }

        // Select close form from the drop down menu
        public void closeForm()
        {
            debug(TRACE, "closeForm: start");

            imageWait(dirName + @"\cassTemplate_images\menuActions.PNG", 5);
            imageFindClick(dirName + @"\cassTemplate_images\menuActions.PNG");
            imageWait(dirName + @"\cassTemplate_images\menuClose.PNG", 5);
            imageFindClick(dirName + @"\cassTemplate_images\menuClose.PNG");

        }

        // Select close all forms from the drop down menu
        public void closeAllForms()
        {
            debug(TRACE, "closeAllForms: start");

            imageFindClick(dirName + @"\cassTemplate_images\menuWindow.PNG");
            imageWait(dirName + @"\cassTemplate_images\menuCloseAllForms.PNG", 5);
            imageFindClick(dirName + @"\cassTemplate_images\menuCloseAllForms.PNG");
            imageWait(dirName + @"\cassTemplate_images\syncImages\closeAllFormsSYNC.PNG", 10);
        }

        // Scroll the main screen up or down until a template image is displayed on the screen - has a set click offset
        public bool scrollMainScreen(string direction, string imageToFindLocation)
        {
            debug(TRACE, "scrollMainScreen: start");

            bool imageFound = false;
            int attempts = 0;
            int clickOffset = 0;
            string scrollBarToUse = "";

            switch (direction.ToLower())
            {
                case "up":
                    scrollBarToUse = dirName + @"\cassTemplate_images\windowScrollUp.PNG";
                    clickOffset = 20;
                    break;

                case "down":
                    scrollBarToUse = dirName + @"\cassTemplate_images\windowScrollDown.PNG";
                    clickOffset = -20;
                    break;

                default:
                    debug(ERROR, "scrollMainScreen: invalid scroll dirstion selected. Sent value:" + direction);
                    break;
            }
            // Scroll down to Selection Status
            while (!imageFound && attempts <= 60)
            {
                imageWait(scrollBarToUse, 10);
                imageFindClick(scrollBarToUse, 0, clickOffset);
                sleep(1);

                imageFound = imageExists(imageToFindLocation);
                attempts++;
            }

            return imageFound;
        }

        // Scroll the main screen up or down until a template image is displayed on the screen - can specify a click offset
        public bool scrollMainScreen(string direction, string imageToFindLocation, int clickOffset)
        {
            debug(TRACE, "scrollMainScreen: start");

            bool imageFound = false;
            int attempts = 0;
            string scrollBarToUse = "";
            LocationValues areaToSearch = new LocationValues();
            areaToSearch.X = 1890;
            areaToSearch.Y = 270;
            areaToSearch.H = 1020 - areaToSearch.Y;
            areaToSearch.W = 1920 - areaToSearch.X;

            switch (direction.ToLower())
            {
                case "up":
                    scrollBarToUse = dirName + @"\cassTemplate_images\windowScrollUp.PNG";
                    break;

                case "down":
                    scrollBarToUse = dirName + @"\cassTemplate_images\windowScrollDownNew.PNG";
                    break;

                default:
                    debug(ERROR, "scrollMainScreen: invalid scroll dirstion selected. Sent value:" + direction);
                    break;
            }
            // Scroll down to Selection Status
            while (!imageFound && attempts <= 60)
            {
                imageWait(scrollBarToUse, 10);
                imageFindClickArea(scrollBarToUse, 0, clickOffset, 0.99, areaToSearch);
                sleep(1);

                imageFound = imageExists(imageToFindLocation, 0.98);
                attempts++;
            }

            return imageFound;
        }

        public bool scrollTableArea(string direction, string imageToFindLocation, LocationValues areaToSearch, int clickOffset)
        {
            debug(TRACE, "scrollMainScreen: start");

            bool imageFound = false;
            int attempts = 0;
            string scrollBarToUse = "";

            switch (direction.ToLower())
            {
                case "up":
                    scrollBarToUse = dirName + @"\cassTemplate_images\windowScrollUp.PNG";
                    break;

                case "down":
                    scrollBarToUse = dirName + @"\cassTemplate_images\windowScrollDown.PNG";
                    break;

                default:
                    debug(ERROR, "scrollMainScreen: invalid scroll dirstion selected. Sent value:" + direction);
                    break;
            }
            // Scroll down to Selection Status
            while (!imageFound && attempts <= 120)
            {
                imageWait(scrollBarToUse, 10);
                imageFindClickArea(scrollBarToUse, 0, clickOffset, 0.99, areaToSearch);
                sleep(1);

                imageFound = imageExists(imageToFindLocation, 0.99);
                attempts++;
            }

            return imageFound;
        }

        // Scroll the search results panel until a template image is displayed
        public void scrollSearchResultsScreen(LocationValues searchBarArea)
        {
            debug(TRACE, "scrollSearchResultsScreen: start");

            int clickOffset = -40;
            string scrollBarToUse = dirName + @"\cassTemplate_images\searchBox\searchResultsScrollDown.PNG";

            imageWait(scrollBarToUse, 10);
            imageFindClickArea(scrollBarToUse, 0, clickOffset, 0.95, searchBarArea);

            // Move mouse away from the search area
            MouseMove(500, 700);

            sleep(1);
        }

        // Scroll the action panel until a template image is displayed
        public bool scrollActionsScreen(string direction, string imageToFindLocation)
        {
            debug(TRACE, "scrollActionsScreen: start");

            bool imageFound = false;
            int attempts = 0;
            int clickOffset = 0;
            string scrollBarToUse = "";

            switch (direction.ToLower())
            {
                case "up":
                    scrollBarToUse = dirName + @"\cassTemplate_images\actionWindowScrollUp.PNG";
                    clickOffset = 20;
                    break;

                case "down":
                    scrollBarToUse = dirName + @"\cassTemplate_images\actionWindowScrollDown.PNG";
                    clickOffset = -20;
                    break;

                default:
                    debug(ERROR, "scrollActionsScreen: invalid scroll dirstion selected. Sent value:" + direction);
                    break;
            }
            // Scroll down to Selection Status
            while (!imageFound && attempts <= 60)
            {
                imageWait(scrollBarToUse, 10);
                imageFindClick(scrollBarToUse, 0, clickOffset, 0.99);
                sleep(1);

                imageFound = imageExists(imageToFindLocation);
                attempts++;
            }

            return imageFound;
        }

        // scroll up on Server Folder Viewer - Import Receipt
        public bool scrollServerFolderScreen(string direction, string imageToFindLocation)
        {
            debug(TRACE, "scrollMainScreen: start");

            bool imageFound = false;
            int attempts = 0;
            int clickOffset = 0;
            string scrollBarToUse = "";

            switch (direction.ToLower())
            {
                case "up":
                    scrollBarToUse = dirName + @"\cassTemplate_images\serverFolderWindowScrollUp.PNG";
                    clickOffset = 20;
                    break;

                case "down":
                    scrollBarToUse = dirName + @"\cassTemplate_images\serverFolderWindowScrollDown.PNG";
                    clickOffset = -20;
                    break;

                default:
                    debug(ERROR, "scrollMainScreen: invalid scroll dirstion selected. Sent value:" + direction);
                    break;
            }
            // Scroll down to Selection Status
            while (!imageFound && attempts <= 60)
            {
                imageWait(scrollBarToUse, 10);
                imageFindClick(scrollBarToUse, 0, clickOffset);

                //randomMouseMove();
                leftClick(1650, 300);
                sleep(1);
                imageFound = imageExists(imageToFindLocation);
                attempts++;
            }

            return imageFound;
        }

        // get file location of latest downloaded file - Import Receipts
        public string getFileLocationOfLatestDownloadedFile()
        {
            debug(TRACE, "getFileLocationOfLatestDownloadedFile: start");

            string fileDirectory = "";
            string[] files = Directory.GetFiles(chromeDownloadPath);
            string latestFile = files.OrderByDescending(f => new FileInfo(f).CreationTime).FirstOrDefault();
            if (latestFile != null)
            {
                string fileName = System.IO.Path.GetFileName(latestFile);
                string fileLocation = System.IO.Path.GetDirectoryName(latestFile);
                fileDirectory = fileLocation + "\\" + fileName;
                Console.WriteLine("fileDirectory: " + fileDirectory);
                return fileDirectory;
            }
            else
            {
                Console.WriteLine("No files downloaded in the given directory");
                return fileDirectory;
            }
        }

        // click on search button and add criteria
        //  to run the function: clickSearchAddCriteria("Study Plan", new string[,]{ { "Availability Year", "=", "2022" }, { "Study Period", "=", "Autumn Session" } });
        public void clickSearchAddCriteria(string screenName, string[,] searchCriteriaDetails)
        {
            debug(TRACE, "clickSearchAddCriteria: start");

            switch (screenName)
            {
                case "Student":
                    // wait for Student screen
                    imageWait(dirName + @"\cassTemplate_images\syncImages\studentSYNC.PNG", 30);

                    // click on search button
                    imageWait(dirName + @"\cassTemplate_images\searchStudent.PNG", 30);
                    imageFindClick(dirName + @"\cassTemplate_images\searchStudent.PNG", 5, 0);

                    // wait for the search window
                    imageWait(dirName + @"\cassTemplate_images\syncImages\studentSearchSYNC.PNG", 30);

                    break;

                case "Study Package":
                    // wait for Study Package screen
                    imageWait(dirName + @"\cassTemplate_images\syncImages\studyPackageSYNC.PNG", 30);

                    // remove cursor
                    Thread.Sleep(500);
                    SendKeys.SendWait(@"{TAB}");
                    Thread.Sleep(500);

                    // click on search button
                    imageWait(dirName + @"\cassTemplate_images\studyPackageSearch.PNG", 30);
                    imageFindClick(dirName + @"\cassTemplate_images\studyPackageSearch.PNG", 90, 0);

                    // wait for the search window
                    imageWait(dirName + @"\cassTemplate_images\studyPackageSearch.PNG", 30);

                    break;

                case "Study Package Availability":
                    // wait for Study Package Availability screen
                    imageWait(dirName + @"\cassTemplate_images\syncImages\studyPackageAvailabilitySYNC.PNG", 30);

                    // click on search button
                    imageWait(dirName + @"\cassTemplate_images\studyPackageAvailabilitySearchButton.PNG", 30);
                    imageFindClick(dirName + @"\cassTemplate_images\studyPackageAvailabilitySearchButton.PNG", 20, 0);

                    // wait for the search window
                    imageWait(dirName + @"\cassTemplate_images\studyPackageAvailabilitySearchSYNC.PNG", 30);

                    break;

                case "Student Study Package":
                    // wait for Student Study Package screen
                    imageWait(dirName + @"\cassTemplate_images\syncImages\studentStudyPackageSYNC.PNG", 30);

                    // click on search button
                    imageWait(dirName + @"\cassTemplate_images\studentStudyPackageSearch.PNG", 30);
                    imageFindClick(dirName + @"\cassTemplate_images\studentStudyPackageSearch.PNG", 20, 0);

                    // wait for the search window
                    imageWait(dirName + @"\cassTemplate_images\syncImages\studentStudyPackageSearchSYNC.PNG", 30);

                    break;

                case "Study Plan":
                    // wait for Study Plan screen
                    imageWait(dirName + @"\cassTemplate_images\syncImages\studyPlanSYNC.PNG", 30);

                    // click on search button
                    imageWait(dirName + @"\cassTemplate_images\studyPlanSearch.PNG", 30);
                    imageFindClick(dirName + @"\cassTemplate_images\studyPlanSearch.PNG", 15, 0);

                    // wait for the search windows
                    imageWait(dirName + @"\cassTemplate_images\syncImages\studentStudyPackageSearchSYNC.PNG", 30);

                    break;

                default:
                    debug(TRACE, "clickSearchAddCriteria: No screen name added, adding criteria.");
                    break;

            }

            sleep(2);

            // clear any existing search results
            //imageFindClick(dirName + @"\cassTemplate_images\buttonImages\searchClearIcon.PNG");
            leftClick(820, 288); // using coords clear is not unique image
            sleep(2);

            // click on Clear Criteria
            imageWait(dirName + @"\cassTemplate_images\clearCriteria.PNG", 30);
            imageFindClick(dirName + @"\cassTemplate_images\clearCriteria.PNG");

            // loop through all search criterias
            for (int i = 0; i < searchCriteriaDetails.GetLength(0); i++)
            {
                for (int j = 0; j < searchCriteriaDetails.GetLength(1); j++)
                {
                    if (j == 0)
                    {
                        // click on blank criteria textfeild
                        imageWait(dirName + @"\cassTemplate_images\searchBox\blankCriteria.PNG", 30);
                        imageFindClick(dirName + @"\cassTemplate_images\searchBox\blankCriteria.PNG", -200, 0);

                        // enter search Criteria
                        Thread.Sleep(500);
                        SendKeys.SendWait(searchCriteriaDetails[i, j]);
                        Thread.Sleep(500);
                        SendKeys.SendWait(@"{TAB}");
                        Thread.Sleep(500);
                    }
                    else if (j == 1)
                    {
                        // enter search Criteria operator
                        clearTextInField();
                        Thread.Sleep(500);
                        SendKeys.SendWait(searchCriteriaDetails[i, j]);
                        Thread.Sleep(500);
                        SendKeys.SendWait(@"{TAB}");
                    }
                    else if (j == 2)
                    {
                        // enter search Criteria value
                        Thread.Sleep(500);
                        SendKeys.SendWait(searchCriteriaDetails[i, j]);
                        Thread.Sleep(500);

                        // click on Add Criteria
                        imageWait(dirName + @"\cassTemplate_images\searchBox\addCriteria.PNG", 30);
                        imageFindClick(dirName + @"\cassTemplate_images\searchBox\addCriteria.PNG");
                    }
                }
            }

            // click on retrieve
            imageWait(dirName + @"\cassTemplate_images\retrieveSearch.PNG", 30);
            imageFindClick(dirName + @"\cassTemplate_images\retrieveSearch.PNG");

            LocationValues searchDialogArea = new LocationValues();
            searchDialogArea.X = 640;
            searchDialogArea.Y = 239;
            searchDialogArea.W = 1307 - searchDialogArea.X;
            searchDialogArea.H = 861 - searchDialogArea.Y;

            // Wait for search results
            //imageVanish(dirName + @"\cassTemplate_images\syncImages\popupSearchCountSYNC.PNG", 30, 0.99);
            imageVanishArea(dirName + @"\cassTemplate_images\syncImages\popupSearchCountSYNC.PNG", 60, 0.99, searchDialogArea);

            return;
        }

        // Add criteria for searches in the main screen and not a search pop up in CASS
        //  to run the function: clickSearchAddCriteria("Study Plan", new string[,]{ { "Availability Year", "=", "2022" }, { "Study Period", "=", "Autumn Session" } });
        public void clickMainAddCriteria(string screenName, string[,] searchCriteriaDetails)
        {
            clickMainAddCriteria(screenName, searchCriteriaDetails, false);
        }
        public void clickMainAddCriteria(string screenName, string[,] searchCriteriaDetails, bool isSacntionScreen)
        {
            string screenToSearch = "";

            if (!isSacntionScreen)
            {
                screenToSearch = dirName + @"\cassTemplate_images\searchBox\blankCriteriaFullScreen.PNG";
            }
            else
            {
                screenToSearch = dirName + @"\cassTemplate_images\searchBox\blankCriteriaFullScreenSanctions.PNG";
            }

            // click on Clear Criteria
            imageWait(dirName + @"\cassTemplate_images\clearCriteria.PNG", 30);
            imageFindClick(dirName + @"\cassTemplate_images\clearCriteria.PNG");

            // loop through all search criterias
            for (int i = 0; i < searchCriteriaDetails.GetLength(0); i++)
            {
                for (int j = 0; j < searchCriteriaDetails.GetLength(1); j++)
                {
                    if (j == 0)
                    {
                        // click on blank criteria textfeild
                        imageWait(screenToSearch, 30, false, false, 0.83);
                        imageFindClick(screenToSearch, -200, 0, 0.83);

                        // enter search Criteria
                        Thread.Sleep(500);
                        SendKeys.SendWait(searchCriteriaDetails[i, j]);
                        Thread.Sleep(500);
                        SendKeys.SendWait(@"{TAB}");
                        Thread.Sleep(500);
                    }
                    else if (j == 1)
                    {
                        // enter search Criteria operator
                        clearTextInField();
                        Thread.Sleep(500);
                        SendKeys.SendWait(searchCriteriaDetails[i, j]);
                        Thread.Sleep(500);
                        SendKeys.SendWait(@"{TAB}");
                    }
                    else if (j == 2)
                    {
                        // enter search Criteria value
                        Thread.Sleep(500);
                        SendKeys.SendWait(searchCriteriaDetails[i, j]);
                        Thread.Sleep(500);

                        // click on Add Criteria
                        imageWait(dirName + @"\cassTemplate_images\searchBox\addCriteria.PNG", 30);
                        imageFindClick(dirName + @"\cassTemplate_images\searchBox\addCriteria.PNG");
                    }
                }
            }

            // click on retrieve
            imageWait(dirName + @"\cassTemplate_images\retrieveButtonGeneric.PNG", 30);
            imageFindClick(dirName + @"\cassTemplate_images\retrieveButtonGeneric.PNG");

            // wait for results
            imageVanish(dirName + @"\cassTemplate_images\syncImages\popupSearchCountSYNC.PNG", 30, 0.99);

            return;
        }

        // Alternative to the sidebar search
        // Opens available functions pop up and searches for the specified function name to open
        public void availableFunctionsSelect(string functionToSearch, string textComparisonType)
        {
            debug(TRACE, "availableFunctionsSelect: start");

            // Click the small icon on the bottom right of the screen for Available Functions
            imageWait(dirName + @"\cassTemplate_images\linkimages\availableFunctionsButton.PNG", 15, false);
            imageFindClick(dirName + @"\cassTemplate_images\linkimages\availableFunctionsButton.PNG");

            // Wait for the window to appear
            imageWait(dirName + @"\cassTemplate_images\syncimages\availableFunctionsSYNC.PNG", 60, false);

            imageExists(dirName + @"\cassTemplate_images\availableFunctionsStartPosition.PNG", 0.99);

            // Check if at top of list - if not close and restart - this resets the starting position
            if (!imageExists(dirName + @"\cassTemplate_images\availableFunctionsStartPosition.PNG", 0.99))
            {
                // Close window
                imageFindClick(dirName + @"\cassTemplate_images\availableFunctionsCloseButton.PNG");
                imageVanish(dirName + @"\cassTemplate_images\syncimages\availableFunctionsSYNC.PNG", 60);

                // Reopen available functions
                imageWait(dirName + @"\cassTemplate_images\linkimages\availableFunctionsButton.PNG", 15, false);
                imageFindClick(dirName + @"\cassTemplate_images\linkimages\availableFunctionsButton.PNG");
                imageWait(dirName + @"\cassTemplate_images\syncimages\availableFunctionsSYNC.PNG", 60, false);
            }

            // Calculate size(based on hard coded values and the sync value as this may move around and content will change
            // depending on the role
            LocationValues locations = getImageLocation(dirName + @"\cassTemplate_images\syncimages\availableFunctionsSYNC.PNG");

            // Calculate the Height and Width
            locations.H = locations.Y + (670 - locations.Y);
            locations.W = locations.X + (400 - locations.X);
            Debug.WriteLine("The locations: " + locations.X + " " + locations.Y + " " + locations.H + " " + locations.W);

            // Click the tiles icon to change the view and makes the text more readable to OCR
            imageWait(dirName + @"\cassTemplate_images\linkimages\availableFunctionsTilesIcon.PNG", 15, false);
            imageFindClick(dirName + @"\cassTemplate_images\linkimages\availableFunctionsTilesIcon.PNG");
            sleep(1);

            // Values for the scroll bar click
            string scrollBarToUse = dirName + @"\cassTemplate_images\availableFunctionsScrollDown.PNG";
            int clickOffset = -10;

            // Vars for tesseract capture
            bool scaleResults = true; // leave this as true for now, may add similiar to sidebar search where tries scaled, then not scaled.
            bool textWasFound = false;
            bool endOfScroll = false;

            // Loop through the
            while (!textWasFound && !endOfScroll)
            {
                tesseractCaptureScreen(locations.W, locations.H, locations.X, locations.Y, dirName + @"\cassTemplate_images\resultsScreenshot.TIFF");
                textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\resultsScreenshot.TIFF", "TextLine", scaleResults, 2, locations.X, locations.Y, "double", textComparisonType, functionToSearch, false);
                if (!textWasFound)
                {
                    textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\resultsScreenshot.TIFF", "TextLine", !scaleResults, 2, locations.X, locations.Y, "double", textComparisonType, functionToSearch, false);
                    if (!textWasFound)
                    {
                        textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\resultsScreenshot.TIFF", "TextLine", scaleResults, 2, locations.X, locations.Y, "double", textComparisonType, functionToSearch, true);
                        if (!textWasFound)
                        {
                            textWasFound = tesseractFindAndClickText(dirName + @"\cassTemplate_images\resultsScreenshot.TIFF", "TextLine", !scaleResults, 2, locations.X, locations.Y, "double", textComparisonType, functionToSearch, true);
                        }
                    }
                }

                if (!textWasFound)
                {
                    imageWaitArea(scrollBarToUse, 10, false, false, 0.95, locations);
                    imageFindClickArea(scrollBarToUse, 0, clickOffset, 0.95, locations);
                }
                imageFindClick(dirName + @"\cassTemplate_images\syncimages\availableFunctionsSYNC.PNG");
                sleep(1);

                // Check if at end of the scrollbar
                endOfScroll = imageExists(dirName + @"\cassTemplate_images\availableFunctionsScrollDownEnd.PNG");
            }

            // Wait for the pop up window to close
            if (endOfScroll)
            {
                imageFindClick(dirName + @"\cassTemplate_images\availableFunctionsCloseButton.PNG");
            }
            imageVanish(dirName + @"\cassTemplate_images\syncimages\availableFunctionsSYNC.PNG", 60);
        }

        // reads the study plan excel sheet and check whether the subject's SSP staus is Enrolled
        public bool checkIfStudentWasEnrolled(string studyPackageCode, string excelFilePath)
        {
            debug(TRACE, "checkIfStudentWasEnrolled: start");

            bool studentEnrolled = false;
            using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    excelSheetTestData = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true,
                            FilterRow = rowReader => rowReader.Depth > 2,
                            ReadHeaderRow = (rowReader) =>
                            {
                                rowReader.Read();
                                rowReader.Read();
                                rowReader.Read();
                            }
                        },
                    });

                    DataTable dataTable = excelSheetTestData.Tables[0];

                    foreach (DataRow dataRow in dataTable.Rows)
                    {
                        if (dataRow["Spk Cd"].ToString() == studyPackageCode)
                        {
                            string sspStatus = dataRow["SSP Status"].ToString();
                            Console.WriteLine("sspStatus: " + sspStatus);
                            if (sspStatus.Trim() == "Enrolled")
                            {
                                studentEnrolled = true;
                            }
                        }
                    }
                }
            }
            return studentEnrolled;
        }

        // reads excel file and returns the extracted DataSet
        // param: coloumnNameStartLine --> row value of the table where the coloumn names start
        public DataSet getDataSetFromExcelFile(string excelFilePath, int coloumnNameStartLine)
        {
            debug(TRACE, "getDataSetFromExcelFile: start");

            using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    excelSheetTestData = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true,
                            FilterRow = rowReader => rowReader.Depth > coloumnNameStartLine - 2,
                            ReadHeaderRow = (rowReader) =>
                            {
                                for (int i = 0; i < coloumnNameStartLine - 1; i++)
                                {
                                    rowReader.Read();
                                }
                            }
                        },
                    });
                }
            }
            return excelSheetTestData;
        }

        //Read data from CSV file and returns dataset
        public DataSet ReadCsv(string filePath, int headerRowIndex)
        {
            var config = new ExcelReaderConfiguration()
            {
                FallbackEncoding = Encoding.GetEncoding(1252),
                AutodetectSeparators = new char[] { ',', ';', '\t', '|', '#' },
            };

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateCsvReader(stream, config))
                {
                    var dataSetConfig = new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true,
                            ReadHeaderRow = (rowReader) =>
                            {
                                // Read rows until we get to the header row
                                for (int i = 0; i < headerRowIndex; i++)
                                {
                                    rowReader.Read();
                                }
                            },
                        }
                    };

                    var resultCsvDataset = reader.AsDataSet(dataSetConfig);
                    return resultCsvDataset;
                }
            }
        }

        // get no of BAND in excel file 
        public int getNoOfBandInExcelFile(string excelFilePath)
        {
            DataSet excelData = getDataSetFromExcelFile(excelFilePath, 4);
            int occurences = 0;
            string pattern = "BAND";

            foreach (DataTable table in excelData.Tables)
            {
                foreach (DataRow row in table.Rows)
                {
                    string cellValue = row[0].ToString();
                    occurences += Regex.Matches(cellValue, pattern, RegexOptions.IgnoreCase).Count;
                }
            }
            return occurences;
        }


        // check the download folder every sec whether a file was downloaded
        // if true return true and file path
        // if false return false
        public (bool, string) WaitForFileToBeDownloaded(int timeoutInSeconds)
        {
            debug(TRACE, "WaitForFileToBeDownloaded: start");

            var elapsedTime = 0;
            var initialFiles = Directory.GetFiles(chromeDownloadPath).ToList();

            while (elapsedTime < timeoutInSeconds * 1000)
            {
                var currentFiles = Directory.GetFiles(chromeDownloadPath).ToList();
                var newFiles = currentFiles.Except(initialFiles).ToList();

                if (newFiles.Count > 0 && !newFiles[0].Contains("crdownload") //&& !newFiles[0].Contains("tmp") - excel downloads have tmp in the filename
)
                {
                    string[] checkExtension = newFiles[0].Split('.'); // split to check file extension isn't temp
                    if (!checkExtension[checkExtension.Length - 1].Equals("tmp"))
                    {
                        debug(TRACE, "WaitForFileToBeDownloaded: file: " + newFiles[0]);
                        return (true, newFiles[0]);
                    }
                }

                Thread.Sleep(1000);
                elapsedTime += 1000;
            }
            debug(TRACE, "WaitForFileToBeDownloaded: no files found");
            return (false, null);
        }

        // Nihal?
        public void checkIfWarningExistsClickOk()
        {
            debug(TRACE, "checkIfWarningExistsClickOk: start");

            if (imageExists(dirName + @"\cassTemplate_images\warningList.PNG"))
            {
                SendKeys.SendWait(@"{ENTER}");
                imageVanish(dirName + @"\cassTemplate_images\warningList.PNG", 5);
            }
        }

        public void checkIfStudyPlanTemplateSearchExistsClickOk(int timeoutValue)
        {
            debug(TRACE, "checkIfStudyPlanTemplateSearchExistsClickOk: start");

            int counter = 0;
            while (counter < timeoutValue)
            {
                if (imageExists(dirName + @"\cassTemplate_images\studyPlanTemplateSearch.PNG"))
                {
                    typeTextInField(@"{ENTER}");
                    imageVanish(dirName + @"\cassTemplate_images\studyPlanTemplateSearch.PNG", 5);
                }
                sleep(1);
                counter++;
            }
            return;
        }

        // check if current directory is empty
        public bool isCurrentDirectoryEmpty(DataSet excelDataSet)
        {
            debug(TRACE, "isCurrentDirectoryEmpty: start");

            DataTable dataTable = excelDataSet.Tables[0];

            foreach (DataRow dataRow in dataTable.Rows)
            {
                if (dataRow["Type"].ToString() != "Folder")
                {
                    return false;
                }
            }
            return true;
        }

        // Check for warnings, if true click Ok
        public void checkForWarningsClickOk(int timeoutValue)
        {
            debug(TRACE, "checkForWarningsClickOk: start");

            int counter = 0;
            while (counter < timeoutValue)
            {
                if (imageExists(dirName + @"\cassTemplate_images\warningList.PNG"))
                {
                    typeTextInField(@"{ENTER}");
                    imageVanish(dirName + @"\cassTemplate_images\warningList.PNG", 5);
                }
                sleep(1);
                counter++;
            }
            return;
        }

        // upload file to CASS using the chervon on top on CASS application screen
        public void uploadFileToCASS(string filePath)
        {
            debug(TRACE, "uploadFileToCASS: start");
            int windowCountBefore = 0;
            int windowCountAfter = 0;
            int timeoutInSeconds = 5;
            int elapsedTimeInSeconds = 0;
            int scenario = 0;

            // click on hidden menu on the top
            imageWait(dirName + @"\cassTemplate_images\topHiddenBar.PNG", 20);
            imageFindClick(dirName + @"\cassTemplate_images\topHiddenBar.PNG");

            // get current window count
            windowCountBefore = getWindowCount();

            // click on upload button
            imageWait(dirName + @"\cassTemplate_images\uploadFromPC.PNG", 20);
            imageFindClick(dirName + @"\cassTemplate_images\uploadFromPC.PNG");

            // check if a new browser has spawned
            while (elapsedTimeInSeconds < timeoutInSeconds)
            {
                windowCountAfter = getWindowCount();
                if (windowCountAfter == windowCountBefore + 1)
                {
                    scenario = 1;
                    break;
                }
                else
                {
                    elapsedTimeInSeconds += 1;
                    sleep(1);
                }
            }

            switch (scenario)
            {
                case 1:
                    // Scenario #1 - New Browser Pop Window with a button

                    // Focus on new broswer window
                    syncWindowCount(3);
                    popupWindow(2);

                    // click on the button
                    ReadOnlyCollection<IWebElement> selectFileButton = findElements(null, By.XPath("//a [text()='Click here to select your file!']"));
                    click(selectFileButton[0]);

                    // wait for the new browser window to pop up
                    imageWait(dirName + @"\cassTemplate_images\uploadWindowHeadingNew.PNG", 20, false, true, 0.90);
                    break;
                default:
                    // Scenario #2 - New Pop Window to select file from local machine
                    // wait for the new file window to pop up
                    imageWait(dirName + @"\cassTemplate_images\uploadWindowHeadingNew.PNG", 20, false, true, 0.90);
                    break;
            }

            // enter file name
            imageFindClick(dirName + @"\cassTemplate_images\uploadWindowFileName.PNG", 50, 0);
            Thread.Sleep(500);
            SendKeys.SendWait(filePath);
            Thread.Sleep(500);

            // click on Open button
            imageWait(dirName + @"\cassTemplate_images\openButtonBlueBorder.PNG", 20);
            imageFindClick(dirName + @"\cassTemplate_images\openButtonBlueBorder.PNG");
            sleep(5);

            // if image exists, overwrite the file
            if (imageExists(dirName + @"\cassTemplate_images\importFolderFileAlreadyExistsError.PNG"))
            {
                imageFindClick(dirName + @"\cassTemplate_images\yesButtonWithBlueBorder.PNG");
                sleep(5);
            }
            return;
        }

        // Check Fee Roll Over Log File
        public void checkFeeRollOverLogFile(string logFileLocation)
        {
            return;
        }

        // Selects Job Option in the actions panel and unchecks the restart job option
        public void checkJobOptionsUnchecked()
        {

            debug(TRACE, "checkJobOptionsUnchecked: start");
            bool jobOptionsVisible = imageExists(dirName + @"\cassTemplate_images\linkImages\jobOptions.PNG");
            if (!jobOptionsVisible)
            {
                bool scrollImageFound = scrollActionsScreen("Up", dirName + @"\cassTemplate_images\linkImages\jobOptions.PNG");
                if (!scrollImageFound)
                {
                    debug(ERROR, "scrollActionsScreen did not find the image to scroll to");

                }
            }

            imageWait(dirName + @"\cassTemplate_images\linkImages\jobOptions.PNG", 60);
            imageFindClick(dirName + @"\cassTemplate_images\linkImages\jobOptions.PNG", 20, 0);
            imageWait(dirName + @"\cassTemplate_images\syncImages\jobOptionsSYNC.PNG", 10);

            // Untick Restart job option
            imageFindClick(dirName + @"\cassTemplate_images\checkBox\restartJobOption.PNG", -5, 0);
            typeTextInField(@"{TAB}");

            // Check the box is unchecked
            imageWait(dirName + @"\cassTemplate_images\syncImages\restartJobUnChecked.PNG", 10);

            // Close the job options pop up
            imageWait(dirName + @"\cassTemplate_images\linkImages\okButtonSelectedLink.PNG", 10);
            imageFindClick(dirName + @"\cassTemplate_images\linkImages\okButtonSelectedLink.PNG", 20, 0);
            imageVanish(dirName + @"\cassTemplate_images\syncImages\jobOptionsSYNC.PNG", 10);
        }

        public void checkReportOptionsUnchecked()
        {

            debug(TRACE, "checkJobOptionsUnchecked: start");

            imageWait(dirName + @"\cassTemplate_images\linkImages\reportOptions.PNG", 60);
            imageFindClick(dirName + @"\cassTemplate_images\linkImages\reportOptions.PNG", 20, 0);
            imageWait(dirName + @"\cassTemplate_images\syncImages\jobOptionsSYNC.PNG", 10);

            // Untick Restart job option
            imageFindClick(dirName + @"\cassTemplate_images\checkBox\restartJobOption.PNG", -5, 0);
            typeTextInField(@"{TAB}");

            // Check the box is unchecked
            imageWait(dirName + @"\cassTemplate_images\syncImages\restartJobUnChecked.PNG", 10);

            // Set the queue to default "$DEFAULT"
            imageFindClick(dirName + @"\cassTemplate_images\editFields\queueText.PNG", 40, 0);
            clearTextInField();
            typeTextInField("$DEFAULT{TAB}");

            // Close the job options pop up
            imageWait(dirName + @"\cassTemplate_images\linkImages\okButtonSelectedLink.PNG", 10);
            imageFindClick(dirName + @"\cassTemplate_images\linkImages\okButtonSelectedLink.PNG", 20, 0);
            imageVanish(dirName + @"\cassTemplate_images\syncImages\jobOptionsSYNC.PNG", 10);
        }
        // Scrolls the Actions panel and checks that test run is checked
        public void checkTestRunChecked()
        {

            debug(TRACE, "checkTestRunChecked: start");

            bool checkTestRun = imageExists(dirName + @"\cassTemplate_images\syncImages\testRun.PNG", 0.99);
            if (!checkTestRun)
            {
                bool scrollImageFound = scrollActionsScreen("Down", dirName + @"\cassTemplate_images\syncImages\testRun.PNG");
                if (!scrollImageFound)
                {
                    debug(ERROR, "scrollActionsScreen did not find the image to scroll to");

                }
            }

            // Check if Test run is checked, if not, click it
            checkTestRun = imageExists(dirName + @"\cassTemplate_images\syncImages\testRunChecked.PNG", 0.99);
            if (!checkTestRun)
            {
                imageFindClick(dirName + @"\cassTemplate_images\syncImages\testRunUnchecked.PNG", -25, 0);
                typeTextInField(@"{TAB}"); // Tab off the field
                sleep(1);
                imageWait(dirName + @"\cassTemplate_images\syncImages\testRunChecked.PNG", 10, false, true, 0.93);
            }
        }

        // Scrolls the Actions panel and checks that test run is unchecked
        public void checkTestRunUnchecked()
        {

            debug(TRACE, "checkTestRunUnchecked: start");

            bool checkTestRun = imageExists(dirName + @"\cassTemplate_images\syncImages\testRun.PNG", 0.99);
            if (!checkTestRun)
            {
                bool scrollImageFound = scrollActionsScreen("Down", dirName + @"\cassTemplate_images\syncImages\testRun.PNG");
                if (!scrollImageFound)
                {
                    debug(ERROR, "scrollActionsScreen did not find the image to scroll to");

                }
            }

            // Check if Test run is checked, if not, click it
            bool isTestRunChecked = imageExists(dirName + @"\cassTemplate_images\syncImages\testRunChecked.PNG", 0.99);
            if (isTestRunChecked)
            {
                imageFindClick(dirName + @"\cassTemplate_images\syncImages\testRunChecked.PNG", -25, 0);
                typeTextInField(@"{TAB}{TAB}"); // Tab off the field
                sleep(2);
                imageWait(dirName + @"\cassTemplate_images\syncImages\testRunUnchecked.PNG", 10, false, true, 0.93);
            }
        }

        // Clicks the submit to server link to send a DP job to the server, handles and warning messages, checks for error icon
        public void submitRequestToServer()
        {
            submitRequestToServer(190);
        }
        public void submitRequestToServer(int announcementTimeout)
        {
            debug(TRACE, "submitRequestToServer: start");

            bool submitRequestVisable = imageExists(dirName + @"\cassTemplate_images\linkImages\submitRequestToServer.PNG");
            if (!submitRequestVisable)
            {
                bool scrollImageFound = scrollActionsScreen("Up", dirName + @"\cassTemplate_images\linkImages\submitRequestToServer.PNG");
                if (!scrollImageFound)
                {
                    debug(ERROR, "scrollActionsScreen did not find the image to scroll to");

                }
            }
            sleep(1);

            imageWait(dirName + @"\cassTemplate_images\linkImages\submitRequestToServer.PNG", 10);

            imageFindClick(dirName + @"\cassTemplate_images\linkImages\submitRequestToServer.PNG");

            // Check for warning message pop ups
            int timeout = 0;
            bool warningPopUp = false; // imageExists(dirName + @"\cassTemplate_images\warningList.PNG");
            bool windowsPopuUp = imageExists(dirName + @"\cassTemplate_images\yesButtonWithBlueBorder.PNG");
            if (windowsPopuUp)
            {
                imageFindClick(dirName + @"\cassTemplate_images\linkImages\yesButtonWithBlueBorder.PNG");
                imageWait(dirName + @"\cassTemplate_images\linkImages\yesButtonWithBlueBorder.PNG", 10);
                windowsPopuUp = false;
            }

            bool announcementShowing = false;
            while (!warningPopUp && !windowsPopuUp && timeout <= 30)
            {
                // Deal with warning if exists
                warningPopUp = imageExists(dirName + @"\cassTemplate_images\warningList.PNG");
                if (warningPopUp)
                {

                    typeTextInField(@"{ENTER}");

                    imageVanish(dirName + @"\cassTemplate_images\warningList.PNG", 5);

                    break;

                }
                windowsPopuUp = imageExists(dirName + @"\cassTemplate_images\yesButtonWithBlueBorder.PNG");
                if (windowsPopuUp)
                {
                    imageFindClick(dirName + @"\cassTemplate_images\linkImages\yesButtonWithBlueBorder.PNG");
                    imageVanish(dirName + @"\cassTemplate_images\linkImages\yesButtonWithBlueBorder.PNG", 10);
                    windowsPopuUp = false;
                }

                // jump out of warning check is announcement pop up is showing
                announcementShowing = imageExists(dirName + @"\cassTemplate_imagessyncImages\popupAnnuncementsSYNC.PNG");
                if (announcementShowing)
                {
                    break;
                }

                sleep(1);
                timeout++;

            }

            // Wait for pop up announcement to appear and wait for it to close
            imageWait(dirName + @"\cassTemplate_images\syncImages\popupAnnuncementsSYNC.PNG", 60);

            // Wait while the job is still waiting to be processed
            timeout = 0;
            bool processStartWaiting = imageExists(dirName + @"\cassTemplate_images\syncImages\jobWillBeProcessedNextByTheServer.PNG");
            while (processStartWaiting && timeout <= announcementTimeout)
            {
                processStartWaiting = imageExists(dirName + @"\cassTemplate_images\syncImages\jobWillBeProcessedNextByTheServer.PNG");
                sleep(1);
                timeout++;
            }
            debug(TRACE, "submitRequestToServer: is process still waiting to process: " + processStartWaiting);
            // Wait for the job to process and the announcement pop up to close

            timeout = 0;
            bool stillAnnouncement = imageExists(dirName + @"\cassTemplate_images\syncImages\popupAnnuncementsSYNC.PNG");
            while (stillAnnouncement && timeout <= announcementTimeout)
            {
                stillAnnouncement = imageExists(dirName + @"\cassTemplate_images\syncImages\popupAnnuncementsSYNC.PNG");
                sleep(1);
                timeout++;
            }
            debug(TRACE, "submitRequestToServer: is announcement still visible: " + stillAnnouncement);
        }

        // Where save spawns an announcement dialog on save, this method will wait for the announcment dialog to disappear
        public void saveRequestWithAnnouncement(int announcementTimeout)
        {
            debug(TRACE, "submitRequestToServer: start");

            bool submitRequestVisable = imageExists(dirName + @"\cassTemplate_images\saveAction.PNG");
            if (!submitRequestVisable)
            {
                bool scrollImageFound = scrollActionsScreen("Up", dirName + @"\cassTemplate_images\saveAction.PNG");
                if (!scrollImageFound)
                {
                    debug(ERROR, "scrollActionsScreen did not find the image to scroll to");

                }
            }
            sleep(1);

            imageWait(dirName + @"\cassTemplate_images\saveAction.PNG", 10);

            imageFindClick(dirName + @"\cassTemplate_images\saveAction.PNG");

            // Check for warning message pop ups
            int timeout = 0;
            bool warningPopUp = imageExists(dirName + @"\cassTemplate_images\warningList.PNG");
            bool windowsPopuUp = imageExists(dirName + @"\cassTemplate_images\yesButtonWithBlueBorder.PNG");
            if (windowsPopuUp)
            {
                imageFindClick(dirName + @"\cassTemplate_images\linkImages\yesButtonWithBlueBorder.PNG");
                imageWait(dirName + @"\cassTemplate_images\linkImages\yesButtonWithBlueBorder.PNG", 10);
                windowsPopuUp = false;
            }

            bool announcementShowing = false;
            while (!warningPopUp && !windowsPopuUp && timeout <= 20)

            {
                // Deal with warning if exists
                warningPopUp = imageExists(dirName + @"\cassTemplate_images\warningList.PNG");
                if (warningPopUp)
                {

                    typeTextInField(@"{ENTER}");

                    imageVanish(dirName + @"\cassTemplate_images\warningList.PNG", 5);

                    break;

                }
                windowsPopuUp = imageExists(dirName + @"\cassTemplate_images\yesButtonWithBlueBorder.PNG");
                if (windowsPopuUp)
                {
                    imageFindClick(dirName + @"\cassTemplate_images\linkImages\yesButtonWithBlueBorder.PNG");
                    imageWait(dirName + @"\cassTemplate_images\linkImages\yesButtonWithBlueBorder.PNG", 10);
                }

                // jump out of warning check is announcement pop up is showing
                announcementShowing = imageExists(dirName + @"\cassTemplate_imagessyncImages\popupAnnuncementsSYNC.PNG");
                if (announcementShowing)
                {
                    break;
                }

                sleep(1);
                timeout++;

            }

            // Wait for pop up announcement to appear and wait for it to close
            imageWait(dirName + @"\cassTemplate_images\syncImages\popupAnnuncementsSYNC.PNG", 30);
            //imageVanish(dirName + @"\cassTemplate_images\syncImages\popupAnnuncementsSYNC.PNG", 190);

            // Wait while the job is still waiting to be processed
            timeout = 0;
            bool processStartWaiting = imageExists(dirName + @"\cassTemplate_images\syncImages\jobWillBeProcessedNextByTheServer.PNG");
            while (processStartWaiting && timeout <= announcementTimeout)
            {
                processStartWaiting = imageExists(dirName + @"\cassTemplate_images\syncImages\jobWillBeProcessedNextByTheServer.PNG");
                sleep(1);
                timeout++;
            }

            timeout = 0;
            bool stillAnnouncement = imageExists(dirName + @"\cassTemplate_images\syncImages\popupAnnuncementsSYNC.PNG");
            while (stillAnnouncement && timeout <= announcementTimeout)
            {
                stillAnnouncement = imageExists(dirName + @"\cassTemplate_images\syncImages\popupAnnuncementsSYNC.PNG");
                sleep(1);
                timeout++;
            }
        }

        // Clicks the select worksheet link and waits for it to be selected
        public void selectWorksheet(int announcementTimeout)
        {
            debug(TRACE, "submitRequestToServer: start");

            bool submitRequestVisable = imageExists(dirName + @"\cassTemplate_images\linkImages\selectWorksheet.PNG");
            if (!submitRequestVisable)
            {
                bool scrollImageFound = scrollActionsScreen("Up", dirName + @"\cassTemplate_images\linkImages\selectWorksheet.PNG");
                if (!scrollImageFound)
                {
                    debug(ERROR, "scrollActionsScreen did not find the image to scroll to");

                }
            }
            sleep(1);

            imageWait(dirName + @"\cassTemplate_images\linkImages\selectWorksheet.PNG", 10);
            imageFindClick(dirName + @"\cassTemplate_images\linkImages\selectWorksheet.PNG");

            // Check for warning message pop ups
            int timeout = 0;
            bool warningPopUp = imageExists(dirName + @"\cassTemplate_images\warningList.PNG");
            bool announcementShowing = false;
            while (!warningPopUp && timeout <= 20)

            {
                // Deal with warning if exists
                warningPopUp = imageExists(dirName + @"\cassTemplate_images\warningList.PNG");
                if (warningPopUp)
                {

                    typeTextInField(@"{ENTER}");

                    imageVanish(dirName + @"\cassTemplate_images\warningList.PNG", 5);

                    break;

                }

                // jump out of warning check is announcement pop up is showing
                announcementShowing = imageExists(dirName + @"\cassTemplate_imagessyncImages\popupAnnuncementsSYNC.PNG");
                if (announcementShowing)
                {
                    break;
                }

                sleep(1);
                timeout++;

            }

            // Wait for pop up announcement to appear and wait for it to close
            imageWait(dirName + @"\cassTemplate_images\syncImages\popupAnnuncementsSYNC.PNG", 30);
            //imageVanish(dirName + @"\cassTemplate_images\syncImages\popupAnnuncementsSYNC.PNG", 190);

            // Wait for this job will be processed next by the server text to change
            timeout = 0;
            bool processStartWaiting = imageExists(dirName + @"\cassTemplate_images\syncImages\jobWillBeProcessedNextByTheServer.PNG");
            while (processStartWaiting && timeout <= announcementTimeout)
            {
                processStartWaiting = imageExists(dirName + @"\cassTemplate_images\syncImages\jobWillBeProcessedNextByTheServer.PNG");
                sleep(1);
                timeout++;
            }

            // Wait for the announcement pop up to disappear
            timeout = 0;
            bool stillAnnouncement = imageExists(dirName + @"\cassTemplate_images\syncImages\popupAnnuncementsSYNC.PNG");
            while (stillAnnouncement && timeout <= announcementTimeout)
            {
                stillAnnouncement = imageExists(dirName + @"\cassTemplate_images\syncImages\popupAnnuncementsSYNC.PNG");
                sleep(1);
                timeout++;
            }
        }

        // Check Formula Constant Updated Successfully
        public static (bool, string, double, double) checkWhetherValuesInFormulaConstantDataSetsAreEqual(DataSet dataSet1, DataSet dataSet2)
        {

            DataTable table1 = dataSet1.Tables[0];
            DataTable table2 = dataSet2.Tables[0];
            string bandNumber = "";
            double valueFromDataSet1 = 0.0;
            double valueFromDataSet2 = 0.0;

            // Calculate new values for table1
            for (int i = 0; i < table1.Rows.Count; i++)
            {
                DataRow row = table1.Rows[i];
                double currentValue = (double)row[1];
                double amountIncrease = (double)row[2];
                double percentIncrease = (double)row[3];
                double newValue = currentValue + amountIncrease + (currentValue * percentIncrease) / 100;
                row[1] = newValue;
            }

            // Compare the rows of table1 and table2
            for (int i = 0; i < table1.Rows.Count; i++)
            {
                DataRow row1 = table1.Rows[i];
                DataRow row2 = table2.Rows[i];

                double roundedValue1 = Math.Round((double)row1[1], 3);
                double roundedValue2 = Math.Round((double)row2[1], 3);

                if (roundedValue1 != roundedValue2)
                {
                    bandNumber = (string)row1[0];
                    valueFromDataSet1 = (double)row1[1];
                    valueFromDataSet2 = (double)row2[1];
                    return (false, bandNumber, valueFromDataSet1, valueFromDataSet2);
                }
            }
            return (true, bandNumber, valueFromDataSet1, valueFromDataSet2);
        }


        // Temp method for Sheetal's script.
        public void sideBarSearchFunction_OCRMismatch(string functionToSearch, string textToClick, string textComparisonType, bool scaleResults)
        {
            debug(TRACE, "sideBarSearchFunction: start");

            // Click the search side button
            //leftClick(13, 336);
            imageWait(dirName + @"\cassTemplate_images\linkImages\sideBarSearchLink.PNG", 15);
            imageFindClick(dirName + @"\cassTemplate_images\linkImages\sideBarSearchLink.PNG");
            imageWait(dirName + @"\cassTemplate_images\searchFunctionSYNC.PNG", 15);

            // type in the text to search
            typeTextInField(@functionToSearch + "{ENTER}");
            sleep(1);
            // sidebarSearchLoading
            // wait for search area to have results
            imageVanish(dirName + @"\cassTemplate_images\syncImages\sidebarSearchSYNC.PNG", 200);
            sleep(1);

            // find and click text
            findandClickTextLine_OCRMismatch(204, 740, 33, 257, textComparisonType, textToClick, scaleResults, 3, "double");

            // wait for search to vanish (assume page is loaded....)
            imageVanish(dirName + @"\cassTemplate_images\searchFunctionSYNC.PNG", 60);
        }

        // Enter values for a field on the external org screen for creating new record
        public void enterExtOrgField(string fieldName, string textValue)
        {
            debug(TRACE, "enterExternalOrgField: start");

            // Check on ext org page
            imageWait(dirName + @"\cassTemplate_images\externalOrganisationSYNC.PNG", 5);

            switch (fieldName)
            {
                case "organisationName":
                    leftClick(461, 359);
                    typeTextInField(textValue);
                    break;

                case "organisationType":
                    leftClick(431, 407);
                    typeTextInField(textValue);
                    break;

                case "status":
                    leftClick(437, 436);
                    typeTextInField(textValue);
                    break;

                default:
                    debug(ERROR, "Field name used not supported. Field name used : " + fieldName);
                    break;
            }
        }

        // Creates a basic student (will extend options over time)
        public string createBasicStudent(string familyName, string givenName, string gender, string dateOfBirth, string nationality)
        {
            debug(TRACE, "createBasicStudent: start");

            // open student details
            sideBarSearchFunction("Student", "exact");

            // open the student window
            imageWait(dirName + @"\cassTemplate_images\StudentSYNC.PNG", 60);
            imageWait(dirName + @"\cassTemplate_images\studentHistoricDataExistsSYNC.PNG", 60);

            // Clear the form
            clearForm();
            sleep(2);

            // Enter student details
            leftClick(427, 416);
            //interactWithFields();
            typeTextInField(familyName);

            leftClick(427, 438);
            typeTextInField(givenName);

            leftClick(427, 466);
            typeTextInField("AutomationTest"); // hard coded so can be used in searches.

            leftClick(706, 391);
            clearTextInField();
            typeTextInField(dateOfBirth + @"{TAB}");

            leftClick(706, 415);
            typeTextInField(gender + @"{TAB}");

            // save
            saveAction();
            sleep(1);

            imageFindClick(dirName + @"\cassTemplate_images\linkImages\citizenShipResidencyLink.PNG");
            imageWait(dirName + @"\cassTemplate_images\citizenShipResidencySYNC.PNG", 60);
            //sleep(3);
            imageWait(dirName + @"\cassTemplate_images\StudentCitzAboriginalSYNC.PNG", 60);

            // Select the nationality
            // Define the area to search for text 
            LocationValues screenArea = new LocationValues();
            screenArea.X = 260;
            screenArea.Y = 348;
            screenArea.W = 1033 - screenArea.X;
            screenArea.H = 392 - screenArea.Y;

            // Search for a test user
            interactWithFields(screenArea, "Citizenship", "search", 500);
            imageWait(dirName + @"\cassTemplate_images\syncImages\citizenshipCodeSearchSYNC.PNG", 30);
            clickSearchAddCriteria("", new string[,] { { "Code Description", "like", "" + nationality + "" }, { "Code Type", "like", "CITIZENSHIP" } });
            sleep(5);
            screenArea.X = 655;
            screenArea.Y = 451;
            screenArea.W = 593;
            screenArea.H = 400;
            findandClickTextLine(screenArea.W, screenArea.H, screenArea.X, screenArea.Y, "contains", nationality, true, 3, false, "double");

            // save
            saveAction();
            sleep(1);

            /*LocationValues location = getImageLocation(dirName + @"\cassTemplate_images\studentBlankIDField.PNG");
            int myGetX = location.X;
            int myGetY = location.Y;
            int myGetH = location.H;
            int myGetW = location.W;

            var tempText = findandGetText(myGetW, myGetH, myGetX, myGetY, true);
            var studentIDText = findTextInString(@"[0-9]{8}", tempText);*/
            // Change to copy from field
            LocationValues cassArea = new LocationValues();
            cassArea.X = 260;
            cassArea.Y = 255;
            cassArea.W = 1033 - cassArea.X;
            cassArea.H = 873 - cassArea.Y;

            interactWithFields(cassArea, "Student Id", "gettext");
            string studentIDText = getTextWindowsClipboard();
            debug(TRACE, "createBasicStudent: student id created : " + studentIDText);
            return studentIDText;
        }

        public string createBasicStudentWithID(string title, string familyName, string givenName, string gender, string dateOfBirth, string nationality)
        {
            debug(TRACE, "createBasicStudentWithID: start");

            // open student details
            sideBarSearchFunction("Student", "exact");

            // open the student window
            imageWait(dirName + @"\cassTemplate_images\StudentSYNC.PNG", 60);
            imageWait(dirName + @"\cassTemplate_images\studentHistoricDataExistsSYNC.PNG", 60);

            // Clear the form
            clearForm();
            sleep(2);

            // Enter student ID
            LocationValues screenArea = new LocationValues();
            screenArea.X = 260;
            screenArea.Y = 214;
            screenArea.W = 800 - screenArea.X;
            screenArea.H = 600 - screenArea.Y;

            interactWithFields(screenArea, "Title", "edit", 500);
            clearTextInField();
            typeTextInField(title + @"{TAB}" + @"{TAB}");

            // Enter student details
            leftClick(427, 386);
            clearTextInField();
            typeTextInField(familyName);

            leftClick(427, 408);
            clearTextInField();
            typeTextInField(givenName);

            leftClick(427, 436);
            clearTextInField();
            typeTextInField("AutomationTest"); // hard coded so can be used in searches.

            leftClick(706, 361);
            clearTextInField();
            typeTextInField(dateOfBirth + @"{TAB}");

            leftClick(706, 385);
            clearTextInField();
            typeTextInField(gender + @"{TAB}");

            // save
            saveAction();
            sleep(1);

            imageFindClick(dirName + @"\cassTemplate_images\linkImages\citizenShipResidencyLink.PNG");
            imageWait(dirName + @"\cassTemplate_images\citizenShipResidencySYNC.PNG", 60);
            //sleep(3);
            imageWait(dirName + @"\cassTemplate_images\StudentCitzAboriginalSYNC.PNG", 60);

            // Select the nationality
            // Define the area to search for text 
            screenArea = new LocationValues();
            screenArea.X = 260;
            screenArea.Y = 348;
            screenArea.W = 1033 - screenArea.X;
            screenArea.H = 379 - screenArea.Y;

            // Search for a test user
            interactWithFields(screenArea, "Citizenship", "search", 500);
            imageWait(dirName + @"\cassTemplate_images\syncImages\citizenshipCodeSearchSYNC.PNG", 30);
            clickSearchAddCriteria("", new string[,] { { "Code Description", "like", "" + nationality + "" }, { "Code Type", "like", "CITIZENSHIP" } });
            sleep(5);
            screenArea.X = 655;
            screenArea.Y = 451;
            screenArea.W = 593;
            screenArea.H = 400;
            findandClickTextLine(screenArea.W, screenArea.H, screenArea.X, screenArea.Y, "contains", nationality, true, 3, false, "double");

            // save
            saveAction();
            sleep(1);

            /*LocationValues location = getImageLocation(dirName + @"\cassTemplate_images\studentBlankIDField.PNG");
            int myGetX = location.X;
            int myGetY = location.Y;
            int myGetH = location.H;
            int myGetW = location.W;

            var tempText = findandGetText(myGetW, myGetH, myGetX, myGetY, true);
            var studentIDText = findTextInString(@"[0-9]{8}", tempText);*/
            // Change to copy from field
            LocationValues cassArea = new LocationValues();
            cassArea.X = 260;
            cassArea.Y = 255;
            cassArea.W = 1033 - cassArea.X;
            cassArea.H = 873 - cassArea.Y;

            interactWithFields(cassArea, "Student Id", "gettext");
            string studentIDText = getTextWindowsClipboard();
            debug(TRACE, "createBasicStudent: student id created : " + studentIDText);
            return studentIDText;
        }

        public string updateBasicStudentWithID(string userId, string title, string familyName,
               string givenName, string gender, string dateOfBirth, string nationality)
        {
            debug(TRACE, "updateBasicStudentWithID: start");

            // open student details
            sideBarSearchFunction("Student", "exact");

            // open the student window
            imageWait(dirName + @"\cassTemplate_images\StudentSYNC.PNG", 60);
            imageWait(dirName + @"\cassTemplate_images\studentHistoricDataExistsSYNC.PNG", 60);

            // Clear the form
            clearForm();
            sleep(2);

            LocationValues screenArea = new LocationValues();
            screenArea.X = 260;
            screenArea.Y = 214;
            screenArea.W = 800 - screenArea.X;
            screenArea.H = 1000 - screenArea.Y;

            // Enter student ID
            interactWithFields(screenArea, "Student Id", "edit", 500);
            clearTextInField();
            typeTextInField(userId + @"{TAB}");

            imageFindClick(dirName + @"\cassTemplate_images\retrieveButtonGeneric.PNG");
            imageWait(dirName + @"\cassTemplate_images\linkImages\addFormerName.PNG", 120);

            interactWithFields(screenArea, "Title", "edit", 500);
            clearTextInField();
            typeTextInField(title + @"{TAB}" + @"{TAB}");

            // Enter student details
            leftClick(427, 386);
            clearTextInField();
            typeTextInField(familyName);

            leftClick(427, 408);
            clearTextInField();
            typeTextInField(givenName);

            leftClick(427, 436);
            clearTextInField();
            typeTextInField("AutomationTest"); // hard coded so can be used in searches.

            leftClick(706, 361);
            clearTextInField();
            typeTextInField(dateOfBirth + @"{TAB}");

            leftClick(706, 385);
            clearTextInField();
            typeTextInField(gender + @"{TAB}");

            //interactWithFields(screenArea, "Name Effective Date", "edit", 500);
            //clearTextInField();
            //typeTextInField(nameEffectiveDate + @"{TAB}");
            sleep(2);

            // save
            saveAction();
            sleep(1);

            imageFindClick(dirName + @"\cassTemplate_images\linkImages\citizenShipResidencyLink.PNG");
            imageWait(dirName + @"\cassTemplate_images\citizenShipResidencySYNC.PNG", 60);
            //sleep(3);
            imageWait(dirName + @"\cassTemplate_images\StudentCitzAboriginalSYNC.PNG", 60);

            // Select the nationality
            // Define the area to search for text 
            screenArea = new LocationValues();
            screenArea.X = 260;
            screenArea.Y = 348;
            screenArea.W = 1033 - screenArea.X;
            screenArea.H = 379 - screenArea.Y;

            // Search for a test user
            interactWithFields(screenArea, "Citizenship", "search", 500);
            imageWait(dirName + @"\cassTemplate_images\syncImages\citizenshipCodeSearchSYNC.PNG", 30);
            clickSearchAddCriteria("", new string[,] { { "Code Description", "like", "" + nationality + "" }, { "Code Type", "like", "CITIZENSHIP" } });
            sleep(5);
            screenArea.X = 655;
            screenArea.Y = 451;
            screenArea.W = 593;
            screenArea.H = 400;
            findandClickTextLine(screenArea.W, screenArea.H, screenArea.X, screenArea.Y, "contains", nationality, true, 3, false, "double");

            // click on checkbox 'Student Norminated Citizenship
            imageFindClick(dirName + @"\cassTemplate_images\checkBox\studentNominatedCitizenship.PNG");
            sleep(5);

            // save
            saveAction();
            sleep(1);

            /*LocationValues location = getImageLocation(dirName + @"\cassTemplate_images\studentBlankIDField.PNG");
            int myGetX = location.X;
            int myGetY = location.Y;
            int myGetH = location.H;
            int myGetW = location.W;

            var tempText = findandGetText(myGetW, myGetH, myGetX, myGetY, true);
            var studentIDText = findTextInString(@"[0-9]{8}", tempText);*/
            // Change to copy from field
            LocationValues cassArea = new LocationValues();
            cassArea.X = 260;
            cassArea.Y = 255;
            cassArea.W = 1033 - cassArea.X;
            cassArea.H = 873 - cassArea.Y;

            interactWithFields(cassArea, "Student Id", "gettext");
            string studentIDText = getTextWindowsClipboard();
            debug(TRACE, "createBasicStudent: student id created : " + studentIDText);
            return studentIDText;
        }


        // Creats an application for a student
        public void createIndividualApplication(string studyPackageCD, string liabilityCategory, string loadCategory, string attendanceMode, string studyMode, string assessmentType, string assessment)
        {
            debug(TRACE, "createIndividualApplication: start");

            sideBarSearchFunction("Application", "exact");
            //imageFindClick(dirName + @"\cassTemplate_images\linkImages\applicationLink.PNG");
            //imageWait(dirName + @"\cassTemplate_images\syncImages\applicationSYNC.PNG", 60);

            // get coords for extracting Application No
            LocationValues location = getImageLocation(dirName + @"\cassTemplate_images\locationImages\applicationNoLOC.PNG");
            int myGetX = location.X;
            int myGetY = location.Y;
            int myGetH = location.H;
            int myGetW = location.W;

            // Open search on Study Package Code
            // Define the area to search for text 
            LocationValues screenArea = new LocationValues();
            screenArea.X = 264;
            screenArea.Y = 333;
            screenArea.W = 513 - screenArea.X;
            screenArea.H = 358 - screenArea.Y;

            // Search for Study Package Code
            interactWithFields(screenArea, "Study Package Code", "search");
            imageWait(dirName + @"\cassTemplate_images\syncImages\studyPackageSearch.PNG", 30);
            clickSearchAddCriteria("", new string[,] { { "Stage", "like", "Active" }, { "Study Package Category", "like", "Course" }, { "Study Package Cd", "like", studyPackageCD } });
            //sleep(5);

            // select the frist row (Study Package Code) from the search result table by clicking on the OK button
            // Click ok
            imageFindClick(dirName + @"\cassTemplate_images\okButtonBlackBorder.PNG");
            imageVanish(dirName + @"\cassTemplate_images\okButtonBlackBorder.PNG", 30);

            // Enter the study package code
            // delete on 10/11/2023, Study Package Code is enter by searching on Study Package Code
            //leftClick(411, 322);
            //clearTextInField();
            //typeTextInField(@"" + studyPackageCD + "{TAB}");
            //imageVanish(dirName + @"\cassTemplate_images\syncImages\applicationTypeSYNC.PNG", 15);

            // fill in mandatory fields

            // liability
            enterApplicationField("liabilityCategory", liabilityCategory + @"{TAB}");

            // load
            enterApplicationField("loadCategory", loadCategory + @"{TAB}");

            // attendance
            enterApplicationField("attendanceMode", attendanceMode + @"{TAB}");

            // study mode
            enterApplicationField("studyMode", studyMode + @"{TAB}");

            // scroll down
            imageFindClick(dirName + @"\cassTemplate_images\windowScrollDown.PNG", 0, -20);
            sleep(1);
            imageWait(dirName + @"\cassTemplate_images\syncImages\assessmentSYNC.PNG", 15);
            sleep(1);

            // click the assessment checkbox
            imageWait(dirName + @"\cassTemplate_images\\applicationAssessmentCheckbox.PNG", 15);
            imageFindClick(dirName + @"\cassTemplate_images\applicationAssessmentCheckbox.PNG");
            sleep(2);

            // wait for data to be populated
            imageVanish(dirName + @"\cassTemplate_images\syncImages\applicationAssessmentSYNC.PNG", 10);

            //leftClick(345, 516);  
            imageFindClick(dirName + @"\cassTemplate_images\editFields\typeHeader.PNG", 0, 20);
            typeTextInField(assessmentType + @"{TAB}");

            //leftClick(444, 516); 
            imageFindClick(dirName + @"\cassTemplate_images\editFields\asseementHeader.PNG", 0, 20);
            typeTextInField(assessment + @"{TAB}");

            // save
            saveAction(120);
            sleep(1);

            // Wait for the application number to change after save.
            imageVanish(dirName + @"\cassTemplate_images\locationImages\applicationNoLOC.PNG", 10);

            // sort out OCR for application # at some point.. ocr not picking up 1 from image.. 
            //var tempText = findandGetText(myGetW, myGetH, myGetX, myGetY, true, true);
            //var studentIDText = findTextInString(@"[0-9]{1}", tempText);
            //debug(TRACE, "createBasicStudent: student id created : " + studentIDText);

        }

        // Updates the offer status for a student
        public void setIndividualOffer(string offerStatus)
        {
            debug(TRACE, "setIndividualOffer: start");

            sideBarSearchFunction("Single Offer");

            // Search for the student on single offer screen
            imageWait(dirName + @"\cassTemplate_images\syncImages\singleOfferSYNC.PNG", 60);

            // change offer status to Offered
            leftClick(462, 463);
            typeTextInField(offerStatus + @"{TAB}");

            // save
            saveAction();
            sleep(1);

        }

        // Opens SSP and changes the stage and status
        public void setSSPStageStatus(string sspStage, string sspStatus)
        {
            debug(TRACE, "setSSPStageStatus: start");

            // Goto Student Study Package
            sideBarSearchFunction("Student Study Package");

            // Change SSP Stage to Admitted
            leftClick(409, 427);
            typeTextInField(sspStage + @"{TAB}");

            // Change SSP Status to Admitted
            leftClick(409, 451);
            typeTextInField(sspStatus + @"{TAB}");

            // save
            saveAction(120);
            sleep(1);
        }

        // Enrolls in a list of subjects in the Study plan table
        public void studyPlanEnrolSubjects(string studyPackageCD, List<string> subjectList)
        {
            debug(TRACE, "studyPlanEnrolSubjects: start");

            LocationValues location = new LocationValues();

            sideBarSearchFunction("Study Plan", "exact");

            // syn on multiple images as the save seems to have multiple parts to it
            imageWait(dirName + @"\cassTemplate_images\saveAction.PNG", 30);
            imageWait(dirName + @"\cassTemplate_images\syncImages\studyPlanApprovalBlankSYNC.PNG", 90);
            imageVanish(dirName + @"\cassTemplate_images\locationImages\studyPlanLOC.PNG", 15);

            // Get location of results grid
            // location = getImageLocation(dirName + @"\cassTemplate_images\locationImages\studyPlanLOC.PNG");

            Console.WriteLine("location in studyPlanEnrolSubjects", location);

            location.X = 262;
            location.Y = 383;
            location.W = 1870 - location.X;
            location.H = 753 - location.Y;

            // Expand the course to show subjects
            expandStudyPlanCourse(studyPackageCD, location);

            // Click and reenter the approval field so  it registers with the app. (known functionality)
            studyPlanSetApproval();
            //imageFindClick(dirName + @"\cassTemplate_images\linkImages\studyPlanEnrolmentApprovedField.PNG");
            //clearTextInField();
            typeTextInField(@"Enrolment Approved{TAB}");

            // Enrol in subjects in the subject list subjects
            foreach (var item in subjectList)
            {
                if (!item.Equals("null"))
                {
                    enrolInSubject(item, location);
                }
            }

            // save
            saveAction(120);
            sleep(1);
            imageWait(dirName + @"\cassTemplate_images\syncImages\studyPlanApprovalBlankSYNC.PNG", 90);

        }

        // Add an elective in the study plan screen
        public void studyPlanAddElective(string electiveCode)
        {
            debug(TRACE, "studyPlanAddElective: start");

            // Check to see if there is an elective to choose from
            bool isElectiveSelectionAvailable = false;

            isElectiveSelectionAvailable = imageExists(dirName + @"\cassTemplate_images\linkImages\studyPlanCoose.PNG");

            if (isElectiveSelectionAvailable)
            {
                // Open the electives pop up
                imageFindClick(dirName + @"\cassTemplate_images\linkImages\studyPlanCoose.PNG");
                imageWait(dirName + @"\cassTemplate_images\syncImages\studyPackageChoose.PNG", 30);
                imageVanish(dirName + @"\cassTemplate_images\syncImages\studyPlanLoadingRules.PNG", 30);

                // Define the area to search for text
                LocationValues chooseArea = new LocationValues();
                chooseArea.X = 612;
                chooseArea.Y = 544;
                chooseArea.W = 1340 - chooseArea.X;
                chooseArea.H = 882 - chooseArea.Y;
                //612 544 1308 882

                // first line coordinates
                int xCheck = 622;
                int yCheck = 601;
                int offset = 20;

                int loopCount = 0;

                string tempText = "";
                bool electiveFound = false;

                // Download the data table
                DataSet studyPlanSearch = downloadExcelConvertToDataSet(60, true, chooseArea);
                int rowsReturned = studyPlanSearch.Tables["Sheet1"].Rows.Count;
                if (rowsReturned != 0)
                {
                    foreach (DataRow row in studyPlanSearch.Tables["Sheet1"].Rows)
                    {
                        // Get the elective code and seelct if found
                        tempText = row["Spk Cd"].ToString();
                        if (tempText == electiveCode)
                        {
                            leftClick(xCheck, yCheck + (offset * loopCount));
                            electiveFound = true;
                        }
                    }

                    loopCount++;
                }
                else
                {
                    debug(TRACE, "studyPlanAddElective: no electives to pick");
                }

                if (!electiveFound)
                {
                    debug(TRACE, "studyPlanAddElective: elective not found, expected: " + electiveCode);
                }

                // find and click the elective
                //LocationValues textLocation = findTextLocation(chooseArea, electiveCode, true, 2);
                //leftClick(textLocation.X - 30, textLocation.Y + 10);
                sleep(2); // needs a few secs to register

                // Click OK
                imageWait(dirName + @"\cassTemplate_images\studyPackageOkButton.PNG", 10);
                imageFindClick(dirName + @"\cassTemplate_images\studyPackageOkButton.PNG");
                imageVanish(dirName + @"\cassTemplate_images\syncImages\studyPackageChoose.PNG", 30);
                imageWait(dirName + @"\cassTemplate_images\syncImages\studyPlanRedX.PNG", 30);
            }
            else
            {
                debug(TRACE, "studyPlanAddElective: no option in study plan to choose electives, skipping");
            }
        }

        // Set the approval ro Enrolment Approved on study plan screen
        public void studyPlanSetApproval()
        {
            debug(TRACE, "studyPlanSetApproval: start");

            //imageFindClick(dirName + @"\cassTemplate_images\linkImages\studyPlanEnrolmentApprovedFieldEmpty.PNG");
            LocationValues screenArea = new LocationValues();
            screenArea.X = 10;
            screenArea.Y = 20;
            screenArea.W = 500 - screenArea.X;
            screenArea.H = 500 - screenArea.Y;
            interactWithFields(screenArea, "Approval", "edit", 100);
            clearTextInField();
            typeTextInField(@"Enrolment Approved{TAB}");
            sleep(1);
        }

        // does something eh Nihal?? ;)
        // check if the data DataSet contains the name and value of the newly created Fee Band
        public bool checkIfDataSetContainsNewlyCreatedBand(DataSet excelFileDataSet, string newName, double newValue)
        {
            debug(TRACE, "checkIfDataSetContainsNewlyCreatedBand: start");

            bool found = false;
            DataTable table1 = excelFileDataSet.Tables[0];

            // check if dataset contains newly created band
            for (int i = 0; i < table1.Rows.Count; i++)
            {
                DataRow row = table1.Rows[i];
                if ((string)row[0] == newName && (double)row[1] == newValue)
                {
                    found = true;
                }
            }
            return found;
        }

        // Scrolls to the end of the inner window in CASS 
        public bool scrollToEndInnerWindow(string screenName, LocationValues screenLocation)
        {
            debug(TRACE, "scrollToEndInnerWindow: start");

            bool imageFound = false;
            int attempts = 0;
            int clickOffset = 0;
            string scrollBarToUse = "";

            scrollBarToUse = dirName + @"\cassTemplate_images\studyPlanScrollDown.PNG";
            clickOffset = -10;

            // Scroll down to Selection Status
            while (!imageFound && attempts <= 30)
            {
                imageWaitArea(scrollBarToUse, 1, false, false, 0.95, screenLocation);
                imageFindClickArea(scrollBarToUse, 0, clickOffset, 0.90, screenLocation);

                leftClick(1881, 280); // move mouse so not over scroll bar

                if (screenName == "Formula Constants")
                {
                    imageFound = imageExists(dirName + @"\cassTemplate_images\formuaConstantTableEnd.PNG");
                }
                if (screenName == "Sanctions")
                {
                    imageFound = imageExistsArea(dirName + @"\cassTemplate_images\emptyTableRowStar.PNG", 0.95, screenLocation);
                }
                attempts++;
            }

            return imageFound;
        }

        // Copy text from the currently selected text field to the clipboard
        public string getTextFromTextBox()
        {
            debug(TRACE, "getTextFromTextBox: start");

            string extractedString = "";

            Thread.Sleep(500);
            SendKeys.SendWait("^a");
            Thread.Sleep(500);
            SendKeys.SendWait("^c");
            Thread.Sleep(500);

            extractedString = getTextWindowsClipboard();
            return extractedString;
        }

        // Read a text file into a string array
        public string[] readTextFileToStringArray(string fileToLoad)
        {
            debug(TRACE, "readTextFileToStringArray: start");

            string[] outText = null;

            outText = File.ReadAllLines(fileToLoad);

            return outText;
        }

        // create a laeave of absense record for a sttudent
        public string createStudentLOA(string startDate, string returnDate)
        {
            debug(TRACE, "createStudentLOA: start");

            // Search for a student
            imageFindClick(dirName + @"\cassTemplate_images\searchBox\loaSearch.PNG", 20, 0);
            imageWait(dirName + @"\cassTemplate_images\syncImages\studentStudyPackageSearchSYNC.PNG", 60, true);

            clickSearchAddCriteria("", new string[,] { { "Other Given Names", "like", "automation" }, { "SSP Status", "=", "Admitted" },
                { "Leave of Absence Exists", "=", "No" }});

            // Click ok
            imageFindClick(dirName + @"\cassTemplate_images\okButtonBlackBorder.PNG");
            imageVanish(dirName + @"\cassTemplate_images\okButtonBlackBorder.PNG", 30);

            // wait for record to load
            imageVanish(dirName + @"\cassTemplate_images\editFields\loaStudentIdField.PNG", 30);

            // Screen Area
            LocationValues screenArea = new LocationValues();
            screenArea.X = 260;
            screenArea.Y = 255;
            screenArea.W = 1033 - screenArea.X;
            screenArea.H = 873 - screenArea.Y;

            // Enter start and return dates
            //interactWithFields(screenArea, "Start Date", "edit", 600);
            imageFindClick(dirName + @"\cassTemplate_images\loaStartDate.PNG", 20, 0);
            typeTextInField(startDate + "{TAB}");

            //interactWithFields(screenArea, "Return Date", "edit", 400);
            imageFindClick(dirName + @"\cassTemplate_images\loaReturnDate.PNG", 20, 0);
            typeTextInField(returnDate + "{TAB}");

            // enter reason
            interactWithFields(screenArea, "Reason", "edit", 200);
            typeTextInField("Other Reasons{TAB}");

            // Enter reason
            leftClick(411, 483);
            typeTextInField("Created wwith automation test{TAB}");

            // save
            saveAction();
            sleep(1);

            // Get the student number used
            imageFindClick(dirName + @"\cassTemplate_images\editFields\studentIdTEXT.PNG", 100, 0);
            string studentId = getTextFromTextBox();
            debug(TRACE, "createStudentLOA: student used: " + studentId);

            return studentId;
        }

        // retrieve a laeave of absense record for a sttudent
        public void retrieveStudentLOA(string studentId)
        {
            debug(TRACE, "retrieveStudentLOA: start");

            // Search for a student
            imageFindClick(dirName + @"\cassTemplate_images\searchBox\loaSearchAll.PNG", 20, 0);
            imageWait(dirName + @"\cassTemplate_images\syncImages\leaveOfAbsenceSearchSYNC.PNG", 60, true);

            clickSearchAddCriteria("", new string[,] { { "Student Id", "=", "" + studentId + "" } });

            // Click ok
            imageFindClick(dirName + @"\cassTemplate_images\okButtonBlackBorder.PNG");
            imageVanish(dirName + @"\cassTemplate_images\okButtonBlackBorder.PNG", 30);

            // wait for record to load
            imageVanish(dirName + @"\cassTemplate_images\editFields\loaLeaveNo.PNG", 30);

        }


        // Returns a randomised datatable from a given dataset
        public DataTable randomiseDataSet(DataSet inDataSet)
        {
            // randomise the data
            DataTable rndTable = inDataSet.Tables[0];
            rndTable.Columns.Add("SortBy", typeof(Int32));
            foreach (DataColumn col in rndTable.Columns)
                col.ReadOnly = false;

            Random rnd = new Random();
            foreach (DataRow row in rndTable.Rows)
            {
                row["SortBy"] = rnd.Next(1, 100);
            }
            DataView dv = rndTable.DefaultView;
            dv.Sort = "SortBy";
            DataTable sortedDT = dv.ToTable();

            return sortedDT;
        }

        // Sheethk: 28/05 Added this mentod to sort the datatable based on a column

        public DataTable sortDataSet(DataSet inDataSet, string columnToSort)
        {
            // create the table
            DataTable myTable = inDataSet.Tables[0];
            // Create data view
            DataView dv = myTable.DefaultView;
            dv.Sort = columnToSort;
            DataTable sortedDT = dv.ToTable();

            return sortedDT;
        }

        // Export search results from CASS and select a random row of data from a specified column
        public string selectRandomDataFromExportedSearch(string columnName)
        {
            return selectRandomDataFromExportedSearch(columnName, false);
        }
        public string selectRandomDataFromExportedSearch(string columnName, bool searchPopup)
        {
            bool excelButtonExists = false;

            LocationValues searchDialogArea = new LocationValues();
            searchDialogArea.X = 640;
            searchDialogArea.Y = 239;
            searchDialogArea.W = 1307 - searchDialogArea.X;
            searchDialogArea.H = 861 - searchDialogArea.Y;

            // Check excel export exists
            excelButtonExists = imageExists(dirName + @"\cassTemplate_images\linkImages\excelButton.PNG");
            if (!excelButtonExists)
            {
                debug(ERROR, "selectRandomStudentFromExportedSearch: Export to excel button not found.");
            }

            // export the search results
            if (searchPopup)
            {
                imageFindClickArea(dirName + @"\cassTemplate_images\linkImages\excelButton.PNG", 0, 0, 0.95, searchDialogArea);
            }
            else
            {
                imageFindClick(dirName + @"\cassTemplate_images\linkImages\excelButton.PNG");
            }
            // wait for the download
            (bool, string) fileDownloaded = WaitForFileToBeDownloaded(60);
            if (fileDownloaded.Item1)
            {
                debug(TRACE, "selectRandomStudentFromExportedSearch - log file downloaded within 60 seconds");
            }
            else
            {
                debug(ERROR, "selectRandomStudentFromExportedSearch - log file not downloaded within 60 seconds");
            }
            sleep(1);

            // load the excel file into a dataset
            DataSet studentData = getDataSetFromExcelFile(fileDownloaded.Item2, 4);

            // Enter students
            int loopCount = 1;
            string tempText = "";

            int rowsReturned = studentData.Tables["Sheet1"].Rows.Count;
            Random random = new Random();
            int randStudent = random.Next(0, rowsReturned);

            foreach (DataRow row in studentData.Tables["Sheet1"].Rows)
            {
                tempText = row[columnName].ToString();

                // return the random data
                if (loopCount == randStudent)
                {
                    return tempText;
                }

                loopCount++;
            }

            return tempText;
        }

        // To Run: checkIfARowExistsInDataSet(excelFileDataSet, new string[,]{ { "Student ID", "0123456789" }, { "First Name", "John" }}) 
        // check whether all the specified values in the array are present in a row of a DataSet 
        public bool checkIfARowExistsInDataSet(DataSet excelFileDataSet, string[,] coloumnNameAndRowValuesPair)
        {
            debug(TRACE, "checkIfARowExistsInDataSet: start");

            string columnNameToSearch = "";
            string rowValueToSearch = "";
            DataTable table1 = excelFileDataSet.Tables[0];

            // traverse through each row of the table
            for (int i = 0; i < table1.Rows.Count; i++)
            {
                DataRow row = table1.Rows[i];

                // traverse through each element of coloumnNameAndRowValuesPair array
                for (int j = 0; j < coloumnNameAndRowValuesPair.GetLength(0); j++)
                {
                    bool found = false;
                    columnNameToSearch = coloumnNameAndRowValuesPair[j, 0];
                    rowValueToSearch = coloumnNameAndRowValuesPair[j, 1];

                    // convert column name to column number
                    int columnIndexToSearch = table1.Columns.IndexOf(columnNameToSearch);

                    // compare row value from the dataset against row value in the array
                    if (row[columnIndexToSearch].ToString().Trim() == rowValueToSearch)
                    {
                        found = true; // if a match is found, set found to true
                        debug(TRACE, "checkIfARowExistsInDataSet: Match found for a row value");
                        debug(TRACE, "checkIfARowExistsInDataSet: row value from excel DataSet = " + row[columnIndexToSearch].ToString().Trim());
                        debug(TRACE, "checkIfARowExistsInDataSet: row value from array (user input) = " + rowValueToSearch);
                    }

                    // if a match was not found in this coloumnNameAndRowValuesPair pair, break and move to the next row
                    if (!found)
                    {
                        break;
                    }

                    // if the current coloumnNameAndRowValuesPair pair matches the current row value of the table and 
                    // we're at the last column-value pair, then return true
                    if (found && j == coloumnNameAndRowValuesPair.GetLength(0) - 1)
                    {
                        return true;
                    }
                }
            }
            // if no match was found after checking all rows, return false
            return false;
        }

        // click on the execel button and convert the downloaded excel file into a dataset.
        // bool flag to specify using the search dialog excel export
        public DataSet downloadExcelConvertToDataSet(int timeoutInSecs)
        {
            return downloadExcelConvertToDataSet(timeoutInSecs, false, null);
        }
        public DataSet downloadExcelConvertToDataSet(int timeoutInSecs, bool useAreaSearch, LocationValues areaLocation)
        {

            debug(TRACE, "downloadExcelConvertToDataSet: start");

            bool excelButtonExists = false;
            LocationValues searchPopUpLocation = new LocationValues();

            if (areaLocation == null)
            {
                searchPopUpLocation.X = 640;
                searchPopUpLocation.Y = 239;
                searchPopUpLocation.W = 1307 - searchPopUpLocation.X;
                searchPopUpLocation.H = 861 - searchPopUpLocation.Y;
            }
            else
            {
                searchPopUpLocation = areaLocation;
            }

            // Check excel export exists
            if (useAreaSearch)
            {
                excelButtonExists = imageExistsArea(dirName + @"\cassTemplate_images\linkImages\excelButton.PNG", 0.95, searchPopUpLocation);
            }
            else
            {
                excelButtonExists = imageExists(dirName + @"\cassTemplate_images\linkImages\excelButton.PNG");
            }
            if (!excelButtonExists)
            {
                debug(ERROR, "downloadExcelConvertToDataSet: Export to excel button not found.");
            }

            // click on the excel button
            if (useAreaSearch)
            {
                imageFindClickArea(dirName + @"\cassTemplate_images\excelButton.PNG", 0, 0, 0.95, searchPopUpLocation);
            }
            else
            {
                imageFindClick(dirName + @"\cassTemplate_images\excelButton.PNG");
            }
            // wait for the file to be downloaded
            (bool excelFileDownloaded, string downloadedExcelFilePath) = WaitForFileToBeDownloaded(timeoutInSecs);

            // show error if downloaded file not found
            if (excelFileDownloaded == false)
            {
                debug(ERROR, "Downloaded excel file not found in the specified time");
            }

            // extract dataset from downloaded excel file
            DataSet downloadedFileDataSet = getDataSetFromExcelFile(downloadedExcelFilePath, 4);

            return downloadedFileDataSet;
        }

        // selects a random row of data from a column specified in the dataset
        public string selectRandomDataFromDataset(DataSet inDataset, string columnName)
        {
            // Enter students
            int loopCount = 1;
            string tempText = "";

            int rowsReturned = inDataset.Tables["Sheet1"].Rows.Count;
            Random random = new Random(Guid.NewGuid().GetHashCode());
            int randStudent = random.Next(0, rowsReturned);

            foreach (DataRow row in inDataset.Tables["Sheet1"].Rows)
            {
                tempText = row[columnName].ToString();

                // return the random data
                if (loopCount == randStudent)
                {
                    return tempText;
                }

                loopCount++;
            }

            return tempText;
        }

        // click clear screen button and wait it save button dissapears
        public void clearScreen()
        {

            debug(TRACE, "clearScreen: start");

            bool saveButtonExists = false;
            int elapsedTime = 0;
            int timeoutInSeconds = 30;

            // Check if save button exists
            saveButtonExists = imageExists(dirName + @"\cassTemplate_images\linkImages\saveAction.PNG");
            if (!saveButtonExists)
            {
                debug(ERROR, "clearScreen: save button not found.");
            }

            // click on the clear button
            imageFindClick(dirName + @"\cassTemplate_images\buttonImages\searchClearIcon.PNG");

            // check whether the save button has disappeared
            while (elapsedTime < timeoutInSeconds * 1000 && saveButtonExists)
            {
                saveButtonExists = imageExists(dirName + @"\cassTemplate_images\linkImages\saveAction.PNG");
                Thread.Sleep(1000);
                elapsedTime += 1000;
            }
            debug(TRACE, "clearScreen: save button did not disappear in the specified time");
        }

        // Finds text in a row in a table. uses the find icon from the table side icons
        public void tableFindRow(string valueToSearch)
        {
            // Click on the table search icon
            imageFindClick(dirName + @"\cassTemplate_images\searchBox\tableSearchIcon.PNG");
            imageWait(dirName + @"\cassTemplate_images\syncImages\findDialogTextSYNC.PNG", 60);

            // search for EAP in the table
            imageFindClick(dirName + @"\cassTemplate_images\windowsImages\findWhatText.PNG", 30, 0);
            typeTextInField(valueToSearch);
            imageFindClick(dirName + @"\cassTemplate_images\windowsImages\findNextButton.PNG");
            sleep(3);
            imageFindClick(dirName + @"\cassTemplate_images\windowsImages\findCancel.PNG");
        }

        // Paste special - uses the chevron to use clipboard to input text from a string array. Selects paste special from the edit menu.
        // Need to have the field selected for the paste special to work 
        // ***** May need more work, but seems to be ok at the moment. - seems to work different for different screens, added screen specific versions below
        public void pasteSpecial(string[] inStringArray, string imagePathTest)
        {

            // click on hidden menu on the top
            imageWait(dirName + @"\cassTemplate_images\topHiddenBar.PNG", 20);
            imageFindClick(dirName + @"\cassTemplate_images\topHiddenBar.PNG");
            MouseMove(1083, 139);
            // wait and click the clipboard
            imageWait(dirName + @"\cassTemplate_images\buttonImages\clipboardIcon.PNG", 20);
            imageFindClick(dirName + @"\cassTemplate_images\buttonImages\clipboardIcon.PNG");

            // wait for the popup
            imageWait(dirName + @"\cassTemplate_images\buttonImages\toClip.PNG", 20);

            typeTextInField("^v");
            //for (int i = 0; i < inStringArray.Length; i++)
            //{
            //    typeTextInField(inStringArray[i] + "\r\n");
            //type(inStringArray[i] + OpenQA.Selenium.Keys.Enter);
            //}

            //imageWait(dirName + @"\cassTemplate_images\\toClip.PNG", 20);
            //imageFindClick(dirName + @"\cassTemplate_images\\toClip.PNG");

            // set the field
            if (imagePathTest != null)
            {
                imageFindClick(imagePathTest, 30, 0);
            }

            // paste special
            imageFindClick(dirName + @"\cassTemplate_images\editMenu.PNG");
            imageWait(dirName + @"\cassTemplate_images\menuPasteSpecial.PNG", 20);
            imageFindClick(dirName + @"\cassTemplate_images\menuPasteSpecial.PNG");

            // close the clipboard pop up
            imageWait(dirName + @"\cassTemplate_images\buttonImages\closeClipboard.PNG", 20);
            imageFindClick(dirName + @"\cassTemplate_images\buttonImages\closeClipboard.PNG");

            imageWait(dirName + @"\cassTemplate_images\okButtonBlueSelected.PNG", 40);
            imageFindClick(dirName + @"\cassTemplate_images\okButtonBlueSelected.PNG");
        }

        // Randomly selects the specified number of students from a dataset
        public string[] selectRandomStudentsFromDataset(DataSet inDataSet, string columnName, int numStudents)
        {
            // select some random students
            int loopCount = 0;
            int studentCount = 0;
            string[] studentIds = new string[numStudents];
            Random rnd = new Random();
            int randNum = 0;

            while (studentCount < numStudents)
            {
                // get a random number
                randNum = rnd.Next(1, inDataSet.Tables["Sheet1"].Rows.Count);
                loopCount = 0;
                foreach (DataRow row in inDataSet.Tables["Sheet1"].Rows)
                {
                    // return the random data
                    if (loopCount == randNum)
                    {
                        studentIds[studentCount] = row[columnName].ToString();
                    }
                    loopCount++;
                }
                studentCount++;
            }

            return studentIds;
        }

        // paste special for bulk offer test on student reward screen
        public void pasteSpecialStudentReward()
        {

            // click on hidden menu on the top
            imageWait(dirName + @"\cassTemplate_images\topHiddenBar.PNG", 20);
            imageFindClick(dirName + @"\cassTemplate_images\topHiddenBar.PNG");
            MouseMove(1083, 139);
            // wait and click the clipboard
            imageWait(dirName + @"\cassTemplate_images\buttonImages\clipboardIcon.PNG", 20);
            imageFindClick(dirName + @"\cassTemplate_images\buttonImages\clipboardIcon.PNG");

            // wait for the popup
            imageWait(dirName + @"\cassTemplate_images\buttonImages\toClip.PNG", 20);

            typeTextInField("^v");

            imageFindClick(dirName + @"\cassTemplate_images\editFields\studentIdTEXT.PNG", 200, 0);

            // paste special
            imageFindClick(dirName + @"\cassTemplate_images\editMenu.PNG");
            imageWait(dirName + @"\cassTemplate_images\menuPasteSpecial.PNG", 20);
            imageFindClick(dirName + @"\cassTemplate_images\menuPasteSpecial.PNG");

            // close the clipboard pop up
            imageWait(dirName + @"\cassTemplate_images\buttonImages\closeClipboard.PNG", 20);
            imageFindClick(dirName + @"\cassTemplate_images\buttonImages\closeClipboard.PNG");

            imageWait(dirName + @"\cassTemplate_images\okButtonBlueSelected.PNG", 300);
            imageFindClick(dirName + @"\cassTemplate_images\okButtonBlueSelected.PNG");
        }

        // paste special for student group comment screen
        public void pasteSpecialStudentGroupComments()
        {

            // click on hidden menu on the top
            imageWait(dirName + @"\cassTemplate_images\topHiddenBar.PNG", 20);
            imageFindClick(dirName + @"\cassTemplate_images\topHiddenBar.PNG");
            MouseMove(1083, 139);
            // wait and click the clipboard
            imageWait(dirName + @"\cassTemplate_images\buttonImages\clipboardIcon.PNG", 20);
            imageFindClick(dirName + @"\cassTemplate_images\buttonImages\clipboardIcon.PNG");

            // wait for the popup
            imageWait(dirName + @"\cassTemplate_images\buttonImages\toClip.PNG", 20);

            typeTextInField("^v");
            sleep(1);

            if (imageExists(dirName + @"\cassTemplate_images\tableSelectedArrow.PNG"))
            {
                imageFindClick(dirName + @"\cassTemplate_images\tableSelectedArrow.PNG", 20, 0);
                clearTextInField();
            }
            else
            {
                imageFindClick(dirName + @"\cassTemplate_images\emptyTableRowStar.PNG", 20, 0);
                clearTextInField();
            }
            // paste special
            imageFindClick(dirName + @"\cassTemplate_images\editMenu.PNG");
            imageWait(dirName + @"\cassTemplate_images\menuPasteSpecial.PNG", 20);
            imageFindClick(dirName + @"\cassTemplate_images\menuPasteSpecial.PNG");

            // close the clipboard pop up
            imageWait(dirName + @"\cassTemplate_images\buttonImages\closeClipboard.PNG", 20);
            imageFindClick(dirName + @"\cassTemplate_images\buttonImages\closeClipboard.PNG");

            imageWait(dirName + @"\cassTemplate_images\okButtonBlueSelected.PNG", 40);
            imageFindClick(dirName + @"\cassTemplate_images\okButtonBlueSelected.PNG");
        }

        // remove columns from a table
        public void removeTableColumn(string template)
        {
            imageFindClick(template, 0, 0, 0.95, "rightClick");
            imageWait(dirName + @"\cassTemplate_images\windowsImages\removeThiscolumn.PNG", 30);
            imageFindClick(dirName + @"\cassTemplate_images\windowsImages\removeThiscolumn.PNG");
            imageVanish(template, 30);
        }

        // change date format from dd-mmm-yyyy to dd/MM/yyyy
        public string changeTextDateFormat(string inTextDate)
        {
            string month = "0";
            string[] inTextDateSplit = inTextDate.Split(new char[] { '-' });

            switch (inTextDateSplit[1])
            {
                case "Jan":
                    month = "01";
                    break;
                case "Feb":
                    month = "02";
                    break;
                case "Mar":
                    month = "03";
                    break;
                case "Apr":
                    month = "04";
                    break;
                case "May":
                    month = "05";
                    break;
                case "Jun":
                    month = "06";
                    break;
                case "Jul":
                    month = "07";
                    break;
                case "Aug":
                    month = "08";
                    break;
                case "Sep":
                    month = "09";
                    break;
                case "Oct":
                    month = "10";
                    break;
                case "Nov":
                    month = "11";
                    break;
                case "Dec":
                    month = "12";
                    break;
                default:
                    break;
            }

            return inTextDateSplit[0] + "/" + month + "/" + inTextDateSplit[2];
        }

        // Configure an alert in CASS using the alrt text in the alert portlet pop up window
        public bool configureCASSAlert(string alertToFind)
        {
            // Try to configure alerts
            imageWait(dirName + @"\cassTemplate_images\linkImages\configureAlertsDisplayed.PNG", 30);
            imageFindClick(dirName + @"\cassTemplate_images\linkImages\configureAlertsDisplayed.PNG", -20, 0);
            imageWait(dirName + @"\cassTemplate_images\syncImages\alertPortletConfigSYNC.PNG", 30);

            // define the location area for the alert popup
            LocationValues alertPortletConfig = new LocationValues();
            alertPortletConfig.X = 713;
            alertPortletConfig.Y = 317;
            alertPortletConfig.W = 1227 - alertPortletConfig.X;
            alertPortletConfig.H = 880 - alertPortletConfig.Y;

            //string alertToFind = "Unprocessed GRS LOA";
            bool endOfScroll = false;
            bool foundText = false;
            string scrollBarToUse = dirName + @"\cassTemplate_images\availableFunctionsScrollDown.PNG";
            int clickOffset = -10;

            LocationValues foundAlertLoc = null;

            while (!foundText && !endOfScroll)
            {


                foundAlertLoc = findTextLocation(alertPortletConfig, alertToFind, "textline", true, 2);
                if (foundAlertLoc.X != 0)
                {
                    foundText = true;
                }

                if (!foundText)
                {
                    imageWaitArea(scrollBarToUse, 10, false, false, 0.90, alertPortletConfig);
                    imageFindClickArea(scrollBarToUse, 0, clickOffset, 0.95, alertPortletConfig);
                }

                //MouseMove(1245, 850);
                leftClick(1218, 844);
                endOfScroll = imageExists(dirName + @"\cassTemplate_images\availableFunctionsScrollDownEnd.PNG");
            }

            // expand the locations to be able to capture the 'Activate' text
            foundAlertLoc.X = foundAlertLoc.X - 20;
            foundAlertLoc.Y = foundAlertLoc.Y - 10;
            foundAlertLoc.W = foundAlertLoc.W + 400;
            foundAlertLoc.H = foundAlertLoc.H + 20;

            imageFindClickArea(dirName + @"\cassTemplate_images\linkImages\activate.PNG", 0, 0, 0.95, foundAlertLoc);

            imageFindClick(dirName + @"\cassTemplate_images\windowsImages\alertsPopupOK.PNG");

            return foundText;
        }

        public void createStudentManagementUser(string userId, string title, string firstName, string lastName,
                                                string password, string emailAddr, string role, string effectiveDate, string staff, string org)
        {
            debug(TRACE, "create Student Management User: start");

            // open student details
            sideBarSearchFunction("Student Management User", "exact");

            // open the student window
            imageWait(dirName + @"\cassTemplate_images\studentManagementUserSYNC.PNG", 60);
            imageWait(dirName + @"\cassTemplate_images\emailAddressSYNC.PNG", 60);

            // Clear the form
            clearForm();
            sleep(2);

            // Enter user id
            leftClick(392, 253);
            typeTextInField(userId + @"{TAB}");

            // Enter title
            leftClick(433, 319);
            typeTextInField(title + @"{TAB}");

            // Enter firstName
            LocationValues screenArea = new LocationValues();
            screenArea.X = 260;
            screenArea.Y = 350;
            screenArea.W = 1033 - screenArea.X;
            screenArea.H = 1200 - screenArea.Y;

            // Search for a test user
            interactWithFields(screenArea, "First Name", "edit", 500);
            //leftClick(433, 342);
            typeTextInField(firstName + @"{TAB}"); // hard coded so can be used in searches.

            // Enter lastName
            //leftClick(433, 366);
            interactWithFields(screenArea, "Last Name", "edit", 500);
            //clearTextInField();
            typeTextInField(lastName + @"{TAB}");

            //leftClick(433, 529);
            screenArea = new LocationValues();
            screenArea.X = 260;
            screenArea.Y = 565;
            screenArea.W = 1033 - screenArea.X;
            screenArea.H = 1200 - screenArea.Y;
            interactWithFields(screenArea, "Password", "edit", 500);
            typeTextInField(password + @"{TAB}" + @"{TAB}");

            //leftClick(433, 765);
            interactWithFields(screenArea, "Email Address", "edit", 500);
            typeTextInField(emailAddr + @"{TAB}");

            // scroll to 'Assigned Access' section
            bool scrollImageFound = scrollMainScreen("down", dirName + @"\cassTemplate_images\linkImages\AssignedAccessSection.PNG", -20);
            if (!scrollImageFound)
            {
                debug(ERROR, "scrollMainScreen did not find section 'Assigned Access' to scroll to");
            }
            // Enter Access Method Code
            imageFindClick(dirName + @"\cassTemplate_images\emptyTableRowStar.PNG", 20, 0);
            sleep(1);
            typeTextInField("Directly Assigned Roles" + @"{TAB}" + @"{TAB}");

            // Enter role 
            typeTextInField(role + @"{TAB}");
            //enter Effective Date
            typeTextInField(effectiveDate + @"{TAB}" + @"{TAB}");

            // collapse 'Workflow' section show that 'Student Management' section will come into view.
            imageFindClick(dirName + @"\cassTemplate_images\linkImage\workflow.png");
            sleep(1);

            interactWithFields(screenArea, "Staff Id", "edit", 500);
            typeTextInField(staff + @"{TAB}");

            interactWithFields(screenArea, "Internal Org Unit", "edit", 500);
            typeTextInField(org + @"{TAB}");
            sleep(10);


            // save
            saveAction();
            sleep(1);
        }
        public void updateStudentManagementUser(string userId, string title, string firstName, string lastName,
                                                string password, string emailAddr, string role, string effectiveDate, string staff, string org)
        {
            debug(TRACE, "update Student Management User: start");

            // open student details
            sideBarSearchFunction("Student Management User", "exact");

            // open the student window
            imageWait(dirName + @"\cassTemplate_images\studentManagementUserSYNC.PNG", 60);
            imageWait(dirName + @"\cassTemplate_images\emailAddressSYNC.PNG", 60);

            // Clear the form
            clearForm();
            sleep(2);

            // Enter user id
            leftClick(392, 253);
            typeTextInField(userId + @"{TAB}");

            // retrieve user for update
            imageFindClick(dirName + @"\cassTemplate_images\retrieveButtonGeneric.PNG");
            imageWait(dirName + @"\cassTemplate_images\linkImages\delete.PNG", 120);

            // Enter title
            leftClick(433, 319);
            clearTextInField();
            typeTextInField(title + @"{TAB}");

            // Enter firstName
            LocationValues screenArea = new LocationValues();
            screenArea.X = 260;
            screenArea.Y = 350;
            screenArea.W = 1033 - screenArea.X;
            screenArea.H = 1200 - screenArea.Y;

            // update first name
            interactWithFields(screenArea, "First Name", "edit", 500);
            clearTextInField();
            typeTextInField(firstName + @"{TAB}"); // hard coded so can be used in searches.

            // update lastName            
            interactWithFields(screenArea, "Last Name", "edit", 500);
            clearTextInField();
            typeTextInField(lastName + @"{TAB}");

            screenArea = new LocationValues();
            screenArea.X = 260;
            screenArea.Y = 565;
            screenArea.W = 1033 - screenArea.X;
            screenArea.H = 1200 - screenArea.Y;
            // not upating password for now...cannot use same password
            //interactWithFields(screenArea, "Password", "edit", 500);
            //clearTextInField();
            //typeTextInField(password + @"{TAB}" + @"{TAB}");

            //leftClick(433, 765);
            interactWithFields(screenArea, "Email Address", "edit", 500);
            clearTextInField();
            typeTextInField(emailAddr + @"{TAB}");

            // scroll to 'Assigned Access' section
            bool scrollImageFound = scrollMainScreen("down", dirName + @"\cassTemplate_images\linkImages\AssignedAccessSection.PNG", -20);
            if (!scrollImageFound)
            {
                debug(ERROR, "scrollMainScreen did not find section 'Assigned Access' to scroll to");
            }

            // remove the existing role first
            imageFindClick(dirName + @"\cassTemplate_images\linkImages\clickHere.PNG");
            sleep(5);
            imageWait(dirName + @"\cassTemplate_images\syncImages\associateRoles.PNG", 120);

            imageFindClick(dirName + @"\cassTemplate_images\buttonImages\goingLeftButton.PNG");
            imageFindClick(dirName + @"\cassTemplate_images\windowsImages\alertsPopupOK.PNG");

            // Enter Access Method Code
            imageFindClick(dirName + @"\cassTemplate_images\emptyTableRowStar.PNG", 20, 0);
            sleep(1);
            clearTextInField();
            typeTextInField("Directly Assigned Roles" + @"{TAB}");

            // Enter role 
            clearTextInField();
            typeTextInField(role + @"{TAB}");
            //enter Effective Date
            clearTextInField();
            typeTextInField(effectiveDate + @"{TAB}" + @"{TAB}");

            // collapse 'Workflow' section show that 'Student Management' section will come into view.
            imageFindClick(dirName + @"\cassTemplate_images\linkImage\workflow.png");
            sleep(1);

            interactWithFields(screenArea, "Staff Id", "edit", 500);
            clearTextInField();
            typeTextInField(staff + @"{TAB}");

            interactWithFields(screenArea, "Internal Org Unit", "edit", 500);
            clearTextInField();
            typeTextInField(org + @"{TAB}");
            sleep(10);


            // save
            saveAction();
            sleep(1);
        }

    }
}