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
			Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'Outlook.NewEmail' at Center.", repo.Outlook.Ribbon.NewEmailInfo);
            repo.Outlook.Ribbon.NewEmail.Click();
            Delay.Milliseconds(200);
            Common.CommonMethods.WaitUntilExist(repo.OutlookMessage.SelfInfo, 20);
		}
		
		/// <summary>
		/// Clean Inbox and Draft folders in Outlook
		/// </summary>
		public static void CleanOutlookFolders()
		{
			Report.Log(ReportLevel.Info, "Mouse", "Mouse Right Click item 'Outlook.MailFolders.InboxFolder' at Center.", repo.Outlook.MailFolders.InboxFolderInfo);
            repo.Outlook.MailFolders.InboxFolder.Click(System.Windows.Forms.MouseButtons.Right,"43;10");
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
		
		/// <summary>
		/// Method to click InBox folder
		/// </summary>
		public static void Click_InBoxFolder()
		{
			Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'Outlook.MailFolders.InboxFolder' at Center.", repo.Outlook.MailFolders.InboxFolderInfo);
			repo.Outlook.MailFolders.InboxFolder.Click("43;10");
            Delay.Milliseconds(200);
		}
		
		/// <summary>
		/// Turn the Reading Pane to Bottom
		/// </summary>
		public static void ReadPaneBottom()
		{
			Click_MenuButtonView();
			Click_ButtonReadingPane();
			Click_MenuBottom();
			Click_MenuButtonHome();
		}
		
		/// <summary>
		/// Turn the Reading Pane to Off
		/// </summary>
		public static void ReadPaneOff()
		{
			Click_MenuButtonView();
			Click_ButtonReadingPane();
			Click_MenuOff();
			Click_MenuButtonHome();
		}
		
		/// <summary>
		/// Click View Menu Button in Outlook menu bar
		/// </summary>
		public static void Click_MenuButtonView()
		{
			Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'Outlook.MenuButtonBar.Tab_View' at Center.", repo.Outlook.MenuButtonBar.Tab_ViewInfo);
            repo.Outlook.MenuButtonBar.Tab_View.Click();
            Delay.Milliseconds(200);
		}
		
		/// <summary>
		/// Click Button Reading Panein Outlook Ribbon
		/// </summary>
		public static void Click_ButtonReadingPane()
		{
			Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'Outlook.ButtonReadingPane' at Center.", repo.Outlook.Ribbon.ButtonReadingPaneInfo);
            repo.Outlook.Ribbon.ButtonReadingPane.Click();
            Delay.Milliseconds(200);
		}
		
		/// <summary>
		/// Click Menu Bottom item
		/// </summary>
		public static void Click_MenuBottom()
		{
			Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'Outlook.MenuBottom' at Center.", repo.Outlook.Ribbon.MenuBottomInfo);
            repo.Outlook.Ribbon.MenuBottom.Click();
            Delay.Milliseconds(200);
		}
		
		/// <summary>
		/// Click Menu Bottom item
		/// </summary>
		public static void Click_MenuOff()
		{
			Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'Outlook.MenuOff' at Center.", repo.Outlook.Ribbon.MenuOffInfo);
            repo.Outlook.Ribbon.MenuOff.Click();
            Delay.Milliseconds(200);
		}
		
		/// <summary>
		/// Click Home Menu Button in Outlook menu bar
		/// </summary>
		public static void Click_MenuButtonHome()
		{
			Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'Outlook.MenuButtonBar.Tab_Home' at Center.", repo.Outlook.MenuButtonBar.Tab_HomeInfo);
            repo.Outlook.MenuButtonBar.Tab_Home.Click();
            Delay.Milliseconds(200);
		}
		
		/// <summary>
		/// Click Email List element
		/// </summary>
		public static void Click_EmailList()
		{
			Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'Outlook.EmailList.GroupBy' at Center.", repo.Outlook.EmailList.GroupByInfo);
            repo.Outlook.EmailList.GroupBy.Click();
            Delay.Milliseconds(200);
		}
		
		/// <summary>
		/// Description: method created to select the last email created.
		/// </summary>
		public static void SelectLastEmailReceived()
		{
			Click_EmailList();
			Delay.Seconds(40);
			Ranorex.Keyboard.Press("{Down}");
		}
		
		/// <summary>
		/// Description: select the email based on the email subject
		/// </summary>
		/// <param name="emailSubject">Email subject to be selected</param>
		public static void ClickEmailReceived(string emailSubject)
		{
			repo.EmailSubject = emailSubject;
			Delay.Seconds(2);
			Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'Outlook.EmailList.EmailSubject' at Center.", repo.Outlook.EmailList.EmailSubjectInfo);
			repo.Outlook.EmailList.EmailSubject.Click();
			Delay.Milliseconds(200);
		}
		
		/// <summary>
		/// Description: Open the email in a new window, based on the email subject
		/// </summary>
		/// <param name="emailSubject">Email subject to be opened</param>
		public static void OpenEmailReceived(string emailSubject)
		{
			repo.EmailSubject = emailSubject;
			Delay.Seconds(2);
			Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'Outlook.EmailList.EmailSubject' at Center.", repo.Outlook.EmailList.EmailSubjectInfo);
			repo.Outlook.EmailList.EmailSubject.DoubleClick();
			Delay.Milliseconds(200);
		}
	}
}
