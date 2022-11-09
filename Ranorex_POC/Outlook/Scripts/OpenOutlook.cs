/*
 * Created by Ranorex
 * User: alberto.hirota
 * Date: 11/6/2022
 * Time: 1:46 PM
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
    /// Description of OpenOutlook.
    /// </summary>
    [TestModule("234650DE-F730-4EDF-BDC5-A2301E3852FB", ModuleType.UserCode, 1)]
    public class OpenOutlook : ITestModule
    {
    	public static Ranorex_POCRepository repo = Ranorex_POCRepository.Instance;
    	
        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public OpenOutlook()
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
            Common.OpeningApps.OpenApplication("Outlook");
            Common.CommonMethods.WaitUntilExist(repo.Outlook.SelfInfo,60);
            Common.OpeningApps.MaximizeOutlook();
        }
    }
}
