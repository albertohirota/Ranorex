/*
 * Created by Ranorex
 * User: alberto.hirota
 * Date: 11/9/2022
 * Time: 2:22 PM
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

namespace Ranorex_POC.Word.Scripts
{
    /// <summary>
    /// Description of ClenupFiles.
    /// </summary>
    [TestModule("D1B116FB-F6B5-4F3A-B83C-9EBDC5AECA70", ModuleType.UserCode, 1)]
    public class ClenupFiles : ITestModule
    {
        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public ClenupFiles()
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
            Common.CommonMethods.DeleteAllFilesInsideFolder(@"C:\Temp");
        }
    }
}
