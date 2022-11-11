/*
 * Created by Ranorex
 * User: alberto.hirota
 * Date: 11/9/2022
 * Time: 2:33 PM
 * 
 * To change this template use Tools > Options > Coding > Edit standard headers.
 */
using System;
using System.IO;
using Ranorex;
using System.Drawing;

namespace Ranorex_POC.Word
{
	/// <summary>
	/// Description of WordValidation.
	/// </summary>
	public class WordValidation
	{
		public static Ranorex_POCRepository repo = Ranorex_POCRepository.Instance;
		
		public WordValidation()
		{
		}
		
		/// <summary>
		/// Ranorex validation if File exists
		/// </summary>
		/// <param name="filePath">File name and path</param>
		public static void ValidateFileExists(string filePath)
		{
			bool fileExists = false;
			try
        	{
				if (File.Exists(filePath))
					fileExists = true;
				Validate.IsTrue(fileExists);
        	}
			catch (Exception ex){
				Report.Screenshot();
				Report.Error(ex.Message);
			}
		}
		
		/// <summary>
		/// Validate specific icon.
		/// If we want to make the can pass the repository (repo.Word... as argument).
		/// </summary>
		public static void ValidateInsertPictureIconExists()
		{
			try
        	{
				CompressedImage InsertPictureIcon_Screenshot1 = repo.Word.Ribbon.InsertPictureIconInfo.GetScreenshot1(new Rectangle(0, 0, 47, 69));
            	Imaging.FindOptions InsertPictureIcon_Screenshot1_Options = Imaging.FindOptions.Default;
            	Report.Log(ReportLevel.Info, "Validation", "Validating ContainsImage (Screenshot: 'Screenshot1' with region {X=0,Y=0,Width=47,Height=69}) on item 'menuitemInfo'.", repo.Word.Ribbon.InsertPictureIconInfo);
            	Validate.ContainsImage(repo.Word.Ribbon.InsertPictureIconInfo, InsertPictureIcon_Screenshot1, InsertPictureIcon_Screenshot1_Options);
        	}
			catch (Exception ex){
				Report.Screenshot();
				Report.Error(ex.Message);
			}
		}
		
		/// <summary>
		/// Validate if the text contains in the body of the document 
		/// </summary>
		/// <param name="bodyText"></param>
		public static void ValidateDocumentBodyTextExists(string bodyText)
		{
			try
        	{
				Report.Log(ReportLevel.Info, "Validation", "Validating appearance of item 'repo.Word.DocumentBody': "+bodyText);
      			Validate.AttributeContains(repo.Word.DocumentBodyInfo,"Text",bodyText);
      			Delay.Milliseconds(100);
        	}
			catch (Exception ex){
				Report.Screenshot();
				Report.Error(ex.Message);
			}
		}
		
		
		public static void ValidateTagExists(string tag)
		{
			try
        	{
				Report.Log(ReportLevel.Info, "Validation", "Validating appearance of item 'repo.Word.Information.Tag': "+tag);
      			Validate.AttributeContains(repo.Word.Information.TagInfo,"Text",tag);
      			Delay.Milliseconds(100);
        	}
			catch (Exception ex){
				Report.Screenshot();
				Report.Error(ex.Message);
			}
		}
	}
}
