/*
 * Created by Ranorex
 * User: alberto.hirota
 * Date: 11/9/2022
 * Time: 9:52 AM
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
    /// Description of OpenWord.
    /// </summary>
    [TestModule("E633668A-4E75-413E-9419-9F74702D8D21", ModuleType.UserCode, 1)]
    public class OpenWord : ITestModule
    {
    	public static Ranorex_POCRepository repo = Ranorex_POCRepository.Instance;
    	
        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public OpenWord()
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
            Common.OpeningApps.OpenApplication("Winword");
            Common.CommonMethods.WaitUntilExist(repo.Word.SelfInfo,60);
            Common.OpeningApps.MaximizeWord();
            Word.WordMethods.Click_NewDocument();
        }
    }
}
