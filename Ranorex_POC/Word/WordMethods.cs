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
		
		public static void Click_CloseDocumentAndDoNotSave()
		{
			Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'Word.ButtonClose' at Center.", repo.Word.ButtonCloseInfo);
            repo.Word.ButtonClose.Click();
            Delay.Milliseconds(200);
            	
			Report.Info("Checking if new Don't save window is showing up...");
			if(repo.OfficeWarn.ButtonDoNotSaveInfo.Exists())
			{
				Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'OfficeWarn.ButtonDoNotSave' at Center.", repo.OfficeWarn.ButtonDoNotSaveInfo);
            	repo.OfficeWarn.ButtonDoNotSave.Click();
            	Delay.Milliseconds(200);
			}
		}
	}
}
