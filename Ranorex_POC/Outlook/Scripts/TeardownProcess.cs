/*
 * Created by Ranorex
 * User: alberto.hirota
 * Date: 11/9/2022
 * Time: 7:52 AM
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
    /// Description of TeardownProcess.
    /// </summary>
    [TestModule("F9F74096-AEA6-4CD0-9EA5-77A9930A1233", ModuleType.UserCode, 1)]
    public class TeardownProcess : ITestModule
    {
        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public TeardownProcess()
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
            Outlook.MessageMethods.ClickCloseMessageAndDoNotSave();
        }
    }
}
