/*
 * Created by Ranorex
 * User: alberto.hirota
 * Date: 11/6/2022
 * Time: 1:26 PM
 * 
 * To change this template use Tools > Options > Coding > Edit standard headers.
 */
using System;
using Ranorex;
using System.Diagnostics;
using System.IO;

namespace Ranorex_POC.Common
{
	/// <summary>
	/// Description of OpeningApps.
	/// </summary>
	public class OpeningApps
	{
		public static Ranorex_POCRepository repo = Ranorex_POCRepository.Instance;
		        
        /// <summary>
        /// Description: StartsExplorer in case it did not automatically open it.
        /// </summary>
        public static void StartExplorer()
        {
        	try
        	{
	        	Process.Start(Path.Combine(Environment.GetEnvironmentVariable("windir"), "explorer.exe"));
				Delay.Seconds(10);	        	
        	} catch (Exception e) {
        		Report.Error(e.Message);
        	}
        }
        
        
        /// <summary>
        /// Description: Open an application
        /// </summary>
        /// <param name="application">Need the Application name</param>
        public static void OpenApplication(string application)
        {
        	Report.Log(ReportLevel.Info, "Keyboard", "Key sequence '{LWin down}r{LWin up}'. Opening Run.");
            Keyboard.Press("{LWin down}r{LWin up}");
            Delay.Milliseconds(500);
            
            Report.Log(ReportLevel.Info, "Keyboard", "Key sequence: '"+ application+"' with focus on 'Run.OpenText'.", repo.Run.OpenTextInfo);
            repo.Run.OpenText.PressKeys(application);
            Delay.Milliseconds(200);

			Report.Log(ReportLevel.Info, "Mouse", "Mouse click OK Button in RUN window.", repo.Run.ButtonOkInfo);
            repo.Run.ButtonOk.Click();
            Delay.Milliseconds(200);        
        }
        
        /// <summary>
        /// Method to maximize Outlook if it is not maximized
        /// </summary>
        public static void MaximizeOutlook()
        {
        	Report.Info("Checking if Outlook is maximized...");
        	if(repo.Outlook.ButtonMaximizeInfo.Exists())
        	{
        		Report.Info("Maximizing Outlook");
        		Report.Log(ReportLevel.Info, "Click", "Mouse Click item 'Outlook.ButtonMaximize' at Center.", repo.Outlook.ButtonMaximizeInfo);
        		repo.Outlook.ButtonMaximize.Click();
            	Delay.Milliseconds(200);
        	}
        }
        
        /// <summary>
        /// Method to maximize Word if it is not maximized
        /// </summary>
        public static void MaximizeWord()
        {
        	Report.Info("Checking if Word is maximized...");
        	if(repo.Word.ButtonMaximizeInfo.Exists())
        	{
        		Report.Info("Maximizing Outlook");
        		Report.Log(ReportLevel.Info, "Click", "Mouse Click item 'Word.ButtonMaximize' at Center.", repo.Word.ButtonMaximizeInfo);
        		repo.Word.ButtonMaximize.Click();
            	Delay.Milliseconds(200);
        	}
        }
	}
}
