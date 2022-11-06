/*
 * Created by Ranorex
 * User: alberto.hirota
 * Date: 11/6/2022
 * Time: 2:08 PM
 * 
 * To change this template use Tools > Options > Coding > Edit standard headers.
 */
using System;
using Ranorex;

namespace Ranorex_POC.Outlook
{
	/// <summary>
	/// Description of OutlookMethods.
	/// </summary>
	public class OutlookMethods
	{
		public static Ranorex_POCRepository repo = Ranorex_POCRepository.Instance;
		
		public OutlookMethods()
		{
		}
		
		/// <summary>
		/// Click New Email
		/// </summary>
		public static void CreateNewEmail()
		{
			Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'Outlook.NewEmail' at Center.", repo.Outlook.NewEmailInfo);
            repo.Outlook.NewEmail.Click();
            Delay.Milliseconds(200);
            Common.CommonMethods.WaitUntilExist(repo.OutlookMessage.SelfInfo, 20);
		}
		
		/// <summary>
		/// Clean Inbox and Draft folders in Outlook
		/// </summary>
		public static void CleanOutlookFolders()
		{
			Report.Log(ReportLevel.Info, "Mouse", "Mouse Right Click item 'Outlook.MailFolders.InboxFolder' at Center.", repo.Outlook.MailFolders.InboxFolderInfo);
            repo.Outlook.MailFolders.InboxFolder.Click(System.Windows.Forms.MouseButtons.Right);
            Delay.Milliseconds(200);
            DeleteAll("Inbox");
            Report.Log(ReportLevel.Info, "Mouse", "Mouse Right Click item 'Outlook.MailFolders.DraftFolder' at Center.", repo.Outlook.MailFolders.DraftFolderInfo);
            repo.Outlook.MailFolders.DraftFolder.Click(System.Windows.Forms.MouseButtons.Right);
            Delay.Milliseconds(200);
            DeleteAll("Draft");
		}
		
		/// <summary>
		/// Delete all emails in a specific folder.
		/// </summary>
		/// <param name="folderName"></param>
		private static void DeleteAll(string folderName)
		{
			Report.Info("Deleting all emails...");
			try
        	{
	        	if (repo.Outlook.DeleteAll.Enabled)
	        	{
	        		Report.Info("Click Delete All in "+folderName+" folder.");
	        		repo.Outlook.DeleteAll.Click();
	        		Delay.Milliseconds(300);
	        		Report.Info("Click Yes.");
	        		repo.Outlook.ButtonYes.Click();
	        	}
	        	else
	        		Report.Info("No emails to delete");
        	}
        	catch {
        		Report.Info("Could not find context menu.");
        	}
		}
	}
}
