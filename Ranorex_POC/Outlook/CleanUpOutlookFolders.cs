/*
 * Created by Ranorex
 * User: alberto.hirota
 * Date: 11/6/2022
 * Time: 3:29 PM
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
    /// Description of CleanUpOutlookFolders.
    /// </summary>
    [TestModule("A7211D91-A3E8-40DD-8FB9-F3412A79F7EB", ModuleType.UserCode, 1)]
    public class CleanUpOutlookFolders : ITestModule
    {
        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public CleanUpOutlookFolders()
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
            Outlook.OutlookMethods.CleanOutlookFolders();
        }
    }
}
