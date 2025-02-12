﻿/*
 * Created by Ranorex
 * User: alberto.hirota
 * Date: 11/6/2022
 * Time: 2:57 PM
 * 
 * To change this template use Tools > Options > Coding > Edit standard headers.
 */
using System;
using System.IO;
using Ranorex;
using Ranorex.Core.Repository;

namespace Ranorex_POC.Common
{
	/// <summary>
	/// Description of CommonMethods.
	/// </summary>
	public class CommonMethods
	{
		public CommonMethods()
		{
		}
		
		/// <summary>
		/// Description: Method to wait the repository element until exist. It requires two parameters
		/// </summary>
		/// <param name="elementInfo">it requires the object repository element.</param>
		/// <param name="timeOutInSeconds">How many seconds, it will wait for the element to exist.</param>
		public static void WaitUntilExist(RepoItemInfo elementInfo, int timeOutInSeconds)
		{
			System.DateTime start = System.DateTime.Now;
        	TimeSpan duration = new TimeSpan(-1);
        	bool open = false;
        	Report.Info("Absolute Path is " + elementInfo.AbsolutePath + ".");
        	
        	do
        	{
        		Report.Info("Waiting for " + elementInfo.FullName + "...");
        		try
        		{
        			open = elementInfo.Exists();
        		}
        		catch (Exception e)
        		{
        			open = false;
        			Report.Error("Exception trying to find item:\r\n" + e.Message);
        		}
        	} while(start.AddSeconds(timeOutInSeconds) > System.DateTime.Now && !open);
        	
        	duration = System.DateTime.Now - start;
        	if (open)
        		Report.Info("Opened in " +duration.TotalSeconds + " seconds.");
        	else
        		Report.Failure (elementInfo.Name + " failed to open within the specified timeout.Total time of waiting is " +duration.TotalSeconds + " seconds.");
        	Report.Screenshot();
		}
		
		/// <summary>
		/// Create folder, if the folder doesn't exist
		/// </summary>
		/// <param name="folderName">Folder path</param>
		public static void CreateFolderIfDoesNotExist(string folderName)
		{
			bool exists = System.IO.Directory.Exists(folderName);
			if(!exists)
    			System.IO.Directory.CreateDirectory(folderName);
		}
		
		/// <summary>
		/// Description: delete all files inside a specific folder
		/// </summary>
		/// <param name="folderName">Folder path</param>
		public static void DeleteAllFilesInsideFolder(string folderName)
		{
			try
			{
				System.IO.DirectoryInfo directory = new DirectoryInfo(folderName);
				foreach (FileInfo file in directory.GetFiles())
				{
				    file.Delete(); 
				}
			}
			catch(Exception ex){
				Report.Screenshot();
				Report.Error(ex.Message);
			}	
		}
		
		/// <summary>
		/// Return string path of the testing files.
		/// </summary>
		/// <param name="localPath"></param>
		/// <returns>Test file name. Ex: \\Files\\Word.docx</returns>
		public static string ReturnStringPath(string localPath)
		{
		    string currentDir = Environment.CurrentDirectory;
		    DirectoryInfo directory = new DirectoryInfo(
		        Path.GetFullPath(Path.Combine(currentDir, @"..\..\" + localPath)));
		    return directory.ToString();
		}
	}
}
