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
        //page elements
        private By userField = By.XPath("//*[@name='loginfmt']");
        private By pwField = By.XPath("//*[@type='password']"); 
        private By submitButton = By.XPath("//*[@type='submit']");
        private By utsEmailField = By.XPath("//*[@type='text']");

        public void loginToCAT()
        {
            debug(TRACE, "\r\n*** Test start: " + this.GetType().ToString() + "\r\n");

            openNewBrowserWindow();
            ReadOnlyCollection<IWebElement> userField;
            ReadOnlyCollection<IWebElement> pwField;
            ReadOnlyCollection<IWebElement> submitButton;
            ReadOnlyCollection<IWebElement> utsEmailField;

            //initDriver();
            gotoUrl(catURL);

            // Wait for element
            waitForElement("*", "text", "Sign in", 5);

            // Find and enter email address
            userField = findElements(null, this.userField); 
            type(userField[0], catUserLoginName, false);

            // Click next button
            submitButton = findElements(null, this.submitButton);
            click(submitButton[0]);

            // Wait for element
            waitForElement("*", "type", "text", 10);

            // Find and enter uts email address
            utsEmailField = findElements(null, this.utsEmailField );
            clear(utsEmailField[0]);
            sleep(2);
            type(utsEmailField[0], catUserLoginName, false);

            // Click next button
            submitButton = findElements(null, this.submitButton);
            click(submitButton[0]);

            // Wait for element
            waitForElement("*", "type", "password", 15);

            // Find and enter password
            pwField = findElements(null, this.pwField);
            type(pwField[0], catUserLoginPassword, false);

            // Find and click sign in button
            submitButton = findElements(null, this.submitButton);
            click(submitButton[0]);

            // Wait for element
            waitForElement("*", "text", "Manage courses", 10);
        }
    }
}
