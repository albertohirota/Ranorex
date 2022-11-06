/*
 * Created by Ranorex
 * User: alberto.hirota
 * Date: 11/6/2022
 * Time: 1:58 PM
 * 
 * To change this template use Tools > Options > Coding > Edit standard headers.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Threading;
using WinForms = System.Windows.Forms;

using Ranorex;
using Ranorex.Core;
using Ranorex.Core.Testing;

namespace Ranorex_POC.Outlook
{
    /// <summary>
    /// Description of OutlookMethods.
    /// </summary>
    [TestModule("A94F4800-85F5-411D-8A08-CCD2D63DFDF7", ModuleType.UserCode, 1)]
    public class CreateAndSendEmails : ITestModule
    {
        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public CreateAndSendEmails()
        {
            // Do not delete - a parameterless constructor is required!
        }

        /// <summary>
        /// Performs the playback of actions in this module.
        /// </summary>
        /// <remarks>You should not call this method directly, instead pass the module
        /// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
        /// that will in turn invoke this method.</remarks>
        void ITestModule.Run()
        {
            Mouse.DefaultMoveTime = 300;
            Keyboard.DefaultKeyPressTime = 100;
            Delay.SpeedFactor = 1.0;
            
            // Email can be changed here
            string email = Environment.UserName + "@" +Environment.UserDomainName + ".local"; 
            
            for(int i = 1; i < 5; i++){
            	Outlook.OutlookMethods.CreateNewEmail();
            	Outlook.MessageMethods.PopulateNewEmail(email,"Test Subject"+i.ToString(),"Email body information "+i.ToString());
            	Outlook.MessageMethods.SendEmail();
            }
        }
    }
}
