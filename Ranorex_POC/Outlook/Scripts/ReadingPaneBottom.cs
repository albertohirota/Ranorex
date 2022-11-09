/*
 * Created by Ranorex
 * User: alberto.hirota
 * Date: 11/7/2022
 * Time: 10:48 AM
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

namespace Ranorex_POC.Outlook.Scripts
{
    /// <summary>
    /// Description of ReadingPaneBottom.
    /// </summary>
    [TestModule("F3A9AEDF-361A-41DB-8C55-A0E2A839A9F7", ModuleType.UserCode, 1)]
    public class ReadingPaneBottom : ITestModule
    {
        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public ReadingPaneBottom()
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
            Outlook.OutlookMethods.ReadPaneBottom();
        }
    }
}
