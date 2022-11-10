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

namespace Ranorex_POC.Word
{
	/// <summary>
	/// Description of WordValidation.
	/// </summary>
	public class WordValidation
	{
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
	}
}
