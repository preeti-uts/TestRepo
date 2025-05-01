using System.IO;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System;
using System.Text;
using OpenQA.Selenium;
using System.Collections.ObjectModel;

namespace selenium_framework
{
    public partial class Selenium_Framework
    {
        // Retreieve test account credentails from either a local json file or github env variable (remote)
        public JToken getUserDetails(string envVarName)
        {
            var fileContent = "";
            var myResult = (JToken)null;
            string pathToJson = "C:\\GitHubRepoConfig\\CanvasAutomationTests\\localTestAccountCredentials.json";

            // Check if running local
            if (Debugger.IsAttached)
            {
                if (file_exists(pathToJson))
                {
                    fileContent = File.ReadAllText(pathToJson);
                    myResult = JsonConvert.DeserializeObject<JToken>(fileContent);
                }
                else
                {
                    debug(ERROR, "Please specify a path for the user details file.");
                }
            }
            else
            {
                string gitSecretText = System.Environment.GetEnvironmentVariable(envVarName);

                if (gitSecretText != null)
                {
                    myResult = JsonConvert.DeserializeObject<JToken>(gitSecretText);
                }
                else
                {
                    debug(ERROR, "Environment variable not found. Check the workflow.");
                }
            }

            return myResult;
        }

        // Set test account credentails from either a local json file or github env variable (remote)
        public void setTestAccountCredentials()
        {

            debug(TRACE, "setTestAccounts: start");

            // github environment variable name in which test account credentials are stored
            string gitHubEnvName = "TEST_ACCOUNTS_CREDENTIALS";

            // retrieve test account creds 
            JToken testAccountCredentials = getUserDetails(gitHubEnvName);

            // set the test account credential
            userLoginName = (string)testAccountCredentials["user1_email"];
            userLoginPassword = DecodeFrom64((string)testAccountCredentials["user1_password"]);
            userLoginAnswerToSecurityQuestionAnswer = (string)testAccountCredentials["user1_answer_to_security_question"];
            userLoginNameTwo = (string)testAccountCredentials["user2_email"];
            userLoginPasswordTwo = DecodeFrom64((string)testAccountCredentials["user2_password"]);
            userLoginAnswerToSecurityQuestionAnswerTwo = (string)testAccountCredentials["user2_answer_to_security_question"];
            userLoginNameThree = (string)testAccountCredentials["user3_email"];
            userLoginPasswordThree = DecodeFrom64((string)testAccountCredentials["user3_password"]);
            userLoginAnswerToSecurityQuestionAnswerThree = (string)testAccountCredentials["user3_answer_to_security_question"];
            catUserLoginName = (string)testAccountCredentials["cat_user1_email"];
            catUserLoginPassword = DecodeFrom64((string)testAccountCredentials["cat_user1_password"]);
            canvasUserLoginName = (string)testAccountCredentials["canvas_user1_email"];
            canvasUserLoginPassword = DecodeFrom64((string)testAccountCredentials["canvas_user1_password"]);
            canvasUserLoginAnswerToSecurityQuestionAnswer = (string)testAccountCredentials["canvas_user1_answer_to_security_question"];

            // check if running local
            if (Debugger.IsAttached)
            {
                debug(TRACE, "setTestAccounts: testAccountCredentials are " + testAccountCredentials);
            }

            Console.WriteLine("I am alive");
        }

        // Navigates to the specified URL and logs into the application using the provided username, password and answer to the security question
        public void loginWithOkta(string applicationURL, string loginName, string loginPassword, string loginAnswerToSecretQuestion)
        {

            debug(TRACE, "loginWithOkta: start");

            ReadOnlyCollection<IWebElement> userField;
            ReadOnlyCollection<IWebElement> pwField;
            ReadOnlyCollection<IWebElement> answerField;
            ReadOnlyCollection<IWebElement> submitButton;

            if (string.IsNullOrEmpty(loginName) || string.IsNullOrEmpty(loginPassword) || string.IsNullOrEmpty(loginAnswerToSecretQuestion))
            {
                debug(ERROR, "loginWithOkta: test account email or password or answer to the security question is empty");
            }

            gotoUrl(applicationURL);

            // Attempting Okta authentication
            debug(TRACE, "loginWithOkta: Attempting Okta authentication.");

            // Wait for element
            waitForElement("*", "type", "submit", 15);

            // Find and enter email address
            userField = findElements(null, By.XPath("//*[@name='identifier']"));
            type(userField[0], loginName, false);

            // Click next button
            submitButton = findElements(null, By.XPath("//*[@type='submit']"));
            click(submitButton[0]);

            // Wait for element
            waitForElement("*", "type", "password", 15);

            // Find and enter password
            pwField = findElements(null, By.XPath("//*[@type='password']"));
            type(pwField[0], loginPassword, false);

            // Find and click sign in button
            submitButton = findElements(null, By.XPath("//*[@type='submit']"));
            click(submitButton[0]);

            // Wait for the security question
            bool textFound = waitForElement("", "text", "Verify with your Security Question", 30);

            if (textFound)
            {
                // Find and enter answer to security question
                answerField = findElements(null, By.XPath("//*[@type='password']"));
                type(answerField[0], loginAnswerToSecretQuestion);

                // Click sign in button
                submitButton = findElements(null, By.XPath("//*[@type='submit']"));
                click(submitButton[0]);
            }
        }


        // Method to get the download path from windows. As it may be different to the default C:/user/<userprofile> path
        // Other guids for special windows folders that can be used with this method:
        //  Contacts =Guid("{56784854-C6CB-462B-8169-88E350ACB882}");
        //  Desktop = new Guid("{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}");
        //  Documents = new Guid("{FDD39AD0-238F-46AF-ADB4-6C85480369C7}");
        //  Downloads = new Guid("{374DE290-123F-4565-9164-39C4925E467B}");
        //  Favorites = new Guid("{1777F761-68AD-4D8A-87BD-30B759FA33DD}");
        //  Links = new Guid("{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}");
        //  Music = new Guid("{4BD8D571-6D19-48D3-BE97-422220080E43}");
        //  Pictures = new Guid("{33E28130-4E1E-4676-835A-98395C3BC3BB}");
        //  SavedGames = new Guid("{4C5C32FF-BB9D-43B0-B5B4-2D72E54EAAA4}");
        //  SavedSearches = new Guid("{7D1D3A04-DEBB-4115-95CF-2F29DA2920DA}");
        //  Videos = new Guid("{18989B1D-99B5-455B-841C-AB7C74E4DDFC}");
        //
        // usage:
        // cGetEnvVars_WinExp.GetPath("{<guid here>}", cGetEnvVars_WinExp.KnownFolderFlags.DontVerify, false);
        //
        static class cGetEnvVars_WinExp
        {
            [DllImport("Shell32.dll")]
            private static extern int SHGetKnownFolderPath(
                [MarshalAs(UnmanagedType.LPStruct)] Guid rfid, uint dwFlags, IntPtr hToken,
                out IntPtr ppszPath);

            [Flags]
            public enum KnownFolderFlags : uint
            {
                SimpleIDList = 0x00000100
                , NotParentRelative = 0x00000200, DefaultPath = 0x00000400, Init = 0x00000800
                , NoAlias = 0x00001000, DontUnexpand = 0x00002000, DontVerify = 0x00004000
                , Create = 0x00008000, NoAppcontainerRedirection = 0x00010000, AliasOnly = 0x80000000
            }
            public static string GetPath(string RegStrName, KnownFolderFlags flags, bool defaultUser)
            {
                IntPtr outPath;
                int result =
                    SHGetKnownFolderPath(
                        new Guid(RegStrName), (uint)flags, new IntPtr(defaultUser ? -1 : 0), out outPath
                    );
                if (result >= 0)
                {
                    return Marshal.PtrToStringUni(outPath);
                }
                else
                {
                    throw new ExternalException("Unable to retrieve the known folder path. It may not "
                        + "be available on this system.", result);
                }
            }

        }

        //
        // Random data generation
        //

        private static Random random = new Random();

        public static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        public static string RandomStringNumber(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        public static string RandomNumber(int length)
        {
            const string chars = "0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        public static string EncodeTo64(string toEncode)
        {
            byte[] toEncodeAsBytes = Encoding.ASCII.GetBytes(toEncode);
            return Convert.ToBase64String(toEncodeAsBytes);
        }

        public static string DecodeFrom64(string encodedData)
        {
            byte[] encodedDataAsBytes = Convert.FromBase64String(encodedData);
            return Encoding.ASCII.GetString(encodedDataAsBytes);
        }
    }
    //
    // Reverse - reverse a string
    //

    public static class StringExtensions
    {
        public static string Reverse(this string input)
        {
            return new string(input.ToCharArray().Reverse().ToArray());
        }
    } // Reverse

}
