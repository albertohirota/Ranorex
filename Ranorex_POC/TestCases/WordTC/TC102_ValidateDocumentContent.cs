﻿///////////////////////////////////////////////////////////////////////////////
//
// This file was automatically generated by RANOREX.
// DO NOT MODIFY THIS FILE! It is regenerated by the designer.
// All your modifications will be lost!
// http://www.ranorex.com
//
///////////////////////////////////////////////////////////////////////////////

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
using Ranorex.Core.Repository;

namespace Ranorex_POC.TestCases.WordTC
{
#pragma warning disable 0436 //(CS0436) The type 'type' in 'assembly' conflicts with the imported type 'type2' in 'assembly'. Using the type defined in 'assembly'.
    /// <summary>
    ///The TC102_ValidateDocumentContent recording.
    /// </summary>
    [TestModule("cb33fa61-fb00-4eb9-8c6a-f2b054fcbd09", ModuleType.Recording, 1)]
    public partial class TC102_ValidateDocumentContent : ITestModule
    {
        /// <summary>
        /// Holds an instance of the global::Ranorex_POC.Ranorex_POCRepository repository.
        /// </summary>
        public static global::Ranorex_POC.Ranorex_POCRepository repo = global::Ranorex_POC.Ranorex_POCRepository.Instance;

        static TC102_ValidateDocumentContent instance = new TC102_ValidateDocumentContent();

        /// <summary>
        /// Constructs a new instance.
        /// </summary>
        public TC102_ValidateDocumentContent()
        {
        }

        /// <summary>
        /// Gets a static instance of this recording.
        /// </summary>
        public static TC102_ValidateDocumentContent Instance
        {
            get { return instance; }
        }

#region Variables

#endregion

        /// <summary>
        /// Starts the replay of the static recording <see cref="Instance"/>.
        /// </summary>
        [System.CodeDom.Compiler.GeneratedCode("Ranorex", global::Ranorex.Core.Constants.CodeGenVersion)]
        public static void Start()
        {
            TestModuleRunner.Run(Instance);
        }

        /// <summary>
        /// Performs the playback of actions in this recording.
        /// </summary>
        /// <remarks>You should not call this method directly, instead pass the module
        /// instance to the <see cref="TestModuleRunner.Run(ITestModule)"/> method
        /// that will in turn invoke this method.</remarks>
        [System.CodeDom.Compiler.GeneratedCode("Ranorex", global::Ranorex.Core.Constants.CodeGenVersion)]
        void ITestModule.Run()
        {
            Mouse.DefaultMoveTime = 300;
            Keyboard.DefaultKeyPressTime = 20;
            Delay.SpeedFactor = 1.00;

            Init();

            OpenDoc();
            Delay.Milliseconds(0);
            
            ValidateDocContent();
            Delay.Milliseconds(0);
            
            CloseDocument();
            Delay.Milliseconds(0);
            
        }

#region Image Feature Data
#endregion
    }
#pragma warning restore 0436
}
