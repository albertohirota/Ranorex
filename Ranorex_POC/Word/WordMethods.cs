/*
 * Created by Ranorex
 * User: alberto.hirota
 * Date: 11/9/2022
 * Time: 10:03 AM
 * 
 * To change this template use Tools > Options > Coding > Edit standard headers.
 */
using System;
using Ranorex;

namespace Ranorex_POC.Word
{
	/// <summary>
	/// Description of WordMethods.
	/// </summary>
	public class WordMethods
	{
		public static Ranorex_POCRepository repo = Ranorex_POCRepository.Instance;
		
		public WordMethods()
		{
		}
		
		/// <summary>
		/// Click New Document, if word opens in Home mode... 
		/// </summary>
		public static void Click_NewDocument()
		{
			Report.Info("Checking if new Document object exists to be clicked...");
			if(repo.Word.NewDocumentInfo.Exists())
			{
				Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'Word.NewDocument' at Center.", repo.Word.NewDocumentInfo);
            	repo.Word.NewDocument.Click();
            	Delay.Milliseconds(200);
			}
		}
		
		/// <summary>
		/// If document word is not closed, it will close it.
		/// </summary>
		public static void Click_CloseDocumentAndDoNotSave()
		{
			Report.Info("Checking if Word is open...");
			if(repo.Word.SelfInfo.Exists())
			{
				Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'Word.ButtonClose' at Center.", repo.Word.ButtonCloseInfo);
            	repo.Word.ButtonClose.Click();
            	Delay.Milliseconds(200);
			}
            	
			Report.Info("Checking if new Don't save window is showing up...");
			if(repo.OfficeWarn.ButtonDoNotSaveInfo.Exists())
			{
				Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'OfficeWarn.ButtonDoNotSave' at Center.", repo.OfficeWarn.ButtonDoNotSaveInfo);
            	repo.OfficeWarn.ButtonDoNotSave.Click();
            	Delay.Milliseconds(200);
			}
		}
		
		/// <summary>
		/// Adding Document body information
		/// </summary>
		/// <param name="text">Add the text here</param>
		public static void AddDocBodyText(string text)
		{
			Report.Log(ReportLevel.Info, "Keyboard", "Text: '"+ text+"' with focus on 'Run.OpenText'.", repo.Word.DocumentBodyInfo);
            repo.Word.DocumentBody.PressKeys(text);
            Delay.Milliseconds(200);
		}
		
		/// <summary>
		/// Description: method to save as the file
		/// </summary>
		/// <param name="fileNamePath">It requires filePath and Name to be saved</param>
		public static void SaveAsDocument(string fileNamePath)
		{
			Click_FileMenuButton();
			Click_SaveAsMenuButton();
			Click_BrowseMenuButton();
			Delay.Seconds(5);
			FileSavePathTextBox(fileNamePath);
			Click_ButtonSave();
			VerifyAndClickReplaceIfNeeded();
		}
		
		/// <summary>
		/// Click File Menu Button in Word bar
		/// </summary>
		public static void Click_FileMenuButton()
		{
			Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'Word.MenuButtonBar.ButtonFile' at Center.", repo.Word.MenuButtonBar.ButtonFileInfo);
            repo.Word.MenuButtonBar.ButtonFile.Click();
            Delay.Milliseconds(200);
		}
		
		/// <summary>
		/// Method to click Save As button
		/// </summary>
		public static void Click_SaveAsMenuButton()
		{
			Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'Word.InfoMenuBar.MenuSaveAs' at Center.", repo.Word.InfoMenuBar.MenuSaveAsInfo);
            repo.Word.InfoMenuBar.MenuSaveAs.Click();
            Delay.Milliseconds(200);
		}
		
		/// <summary>
		/// Method to click Browse Button
		/// </summary>
		public static void Click_BrowseMenuButton()
		{
			Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'Word.Information.BrowseMenuButton' at Center.", repo.Word.Information.BrowseMenuButtonInfo);
            repo.Word.Information.BrowseMenuButton.Click();
            Delay.Milliseconds(200);
		}
		
		/// <summary>
		/// Method to fill the File name Text box, where it will be saved and file name
		/// </summary>
		/// <param name="filePath">File name and path to be saved</param>
		public static void FileSavePathTextBox(string filePath)
		{
			Report.Log(ReportLevel.Info, "Keyboard", "Filepath: '"+ filePath+".", repo.Word.SaveAsWindow.FileNameTextBoxInfo);
            repo.Word.SaveAsWindow.FileNameTextBox.PressKeys(filePath);
            Delay.Milliseconds(200);
		}
		
		/// <summary>
		/// Method to click Save Button, in the Save As Windows
		/// </summary>
		public static void Click_ButtonSave()
		{
			Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'Word.SaveAsWindow.SaveButton' at Center.", repo.Word.SaveAsWindow.SaveButtonInfo);
            repo.Word.SaveAsWindow.SaveButton.Click();
            Delay.Milliseconds(200);
		}
		
		/// <summary>
		/// Description: verify if Replace window exists. If exists, it should click OK button
		/// </summary>
		public static void VerifyAndClickReplaceIfNeeded()
		{
			Report.Info("Verifing if Replace Windows exists...");
			if(repo.Word.ReplaceWindow.OkButtonInfo.Exists())
			{
				Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'Word.ReplaceWindow.OkButton' at Center.", repo.Word.ReplaceWindow.OkButtonInfo);
            	repo.Word.ReplaceWindow.OkButton.Click();
            	Delay.Milliseconds(200);
			}
		}
	}
}
