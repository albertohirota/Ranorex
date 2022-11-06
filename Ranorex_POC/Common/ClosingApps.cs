/*
 * Created by Ranorex
 * User: alberto.hirota
 * Date: 11/6/2022
 * Time: 10:42 AM
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
	/// Description of CloseAllApps.
	/// </summary>
	public class ClosingApps
	{
		public static Ranorex_POCRepository repo = Ranorex_POCRepository.Instance;
		
		/// <summary>
    	/// This is kill all applications that could be running during scripts.
    	/// </summary>
        public static void CloseAllApplications()
        {
        	Report.Info("Checking for open processes.");
        	string[] apps = {"Outlook","regedit","WinWord","Excel","PowerPnt","OneNote"};
        	foreach(string app in apps)
        	{
        		TaskkillProcess(app);
        	}
        	RestartExplorer();        
        }
        
        /// <summary>
        /// Kill all open process before starting to run Test Cases
        /// </summary>
        /// <param name="process">Process name to be killed required</param>
        public static void TaskkillProcess(string process)
        {
            Delay.Milliseconds(250);
            Process[] appList = ReturnProcessListByName(process);
            Report.Info("Checking for " + process + " processes.");

            if (appList != null && appList.Length > 0)
            {
                Report.Info("Found " + appList.Length + " process(es) running.");
                foreach (Process p in appList)
                    using (p)
                    {
                        try
                        {
                            Report.Info("Ensuring the process is responding.");
                            if (p.Responding)
                            	TaskKill(process, p);
                            else if (!p.Responding && !p.HasExited)
	        				{
	        					Report.Warn("Process is not responding.");
	        					Report.Info("Killing process and continuing task");
	        					TaskkillIfNotResponding(process);
	        				}
                            else
                                Report.Info("Main window was closed...");
                        }
                        catch (Exception ex)
                        {
                            Report.Error("Exception while killing process:\r\n" + ex.Message);
                        }
                    }
            }
        }
        
        /// <summary>
		///		Description: Checks if a process is is open and is not responding, and if 
		/// 				it is not responding, it TaskKills the process. It also kills the '___ has stopped working' windows
		/// 				that popup when a program breaks
		/// 				Have added an extra measure to make sure that it cannot kill explorer
        /// </summary>
        /// <param name="process">Process name that is not responding</param>
        public static void TaskkillIfNotResponding(string process)
        {
            Delay.Milliseconds(250);
            Process[] appList = ReturnProcessListByName(process);
            Report.Info("Checking for " + process + " processes.");

            if (appList != null && appList.Length > 0)
            {
                Report.Info("Found " + appList.Length + " process(es) running.");
                foreach (Process p in appList)
                    using (p)
                    {
                        try
                        {
                            Report.Info("Checking if the process exists, is not responding, and is not explorer.");
                            if (!p.Responding && !p.ProcessName.Equals("explorer") && !p.HasExited)
                            {
                            	TaskKill(process, p);
                                Report.Info("Killing '____ Has stopped working' Popups");
                                TaskkillProcess("WerFault");
                            }
                            else if (p.Responding && !p.HasExited)
	        					Report.Info("Process is responding.");
                            else
                                Report.Info("Main window was closed...");
                        }
                        catch (Exception ex)
                        {
                            Report.Error("Exception while killing process:\r\n" + ex.Message);
                        }
                    }
            }
        }
        
        private static void TaskKill(string process, Process p)
        {
        	int timeout = 60000;
        	bool hasExited = false;
        	Report.Info("Sending task kill F command..");
            Process.Start("taskkill", "/F /IM " + process + ".exe");
            Report.Info("Waiting " + timeout + " milliseconds.");                  
            hasExited = p.WaitForExit(timeout);
            if (hasExited)
                Report.Info("killed process  " + p.ToString());
            else
                Report.Info("problem killing process   " + p.ToString());
        }
        
        /// <summary>
        /// Return the Process List by name
        /// </summary>
        /// <param name="process">String of the process name</param>
        /// <returns>Process list in array</returns>
        private static Process[] ReturnProcessListByName(string process)
        {
        	Process[] appList = null;
        	try
            {
                appList = System.Diagnostics.Process.GetProcessesByName(process);
            }
            catch (Exception ex)
            {
                Report.Error("Error when retrieving list of process.\r\n" + ex.Message);
            }
            return appList;
        }
        
        /// <summary>
		/// Clean the temporary files and directories under SubPath generated by recordings.
		/// The root directory, SubPath, will be kept.
		/// </summary>
		/// <param name="folder">Folder path to delete all files</param>
		public static void CleanFolderAndSubFolder(string folder)
		{
			try
			{
				CleanAllFiles(folder);
				foreach(string subdirectory in Directory.GetDirectories(folder))
				{
					CleanAllFiles(subdirectory);
					Directory.Delete(subdirectory);
				}
			}
			catch(Exception e)
			{
				Report.Error("User", e.Message);
			}
		}
		
		/// <summary>
		/// Clean the temporary files.
		/// </summary>
		/// <param name="folder">Folder path to delete all files</param>
		public static void CleanAllFiles(string folder)
		{
			if(Directory.Exists(folder))
			{
				foreach(string FileName in Directory.GetFiles(folder))
				{
					try{
						File.Delete(FileName);
					}catch(Exception e){
						Report.Error("User", e.Message);
					}										
				}
			} else
				throw new Exception("The directory" + folder + " is not exist.");
		}
		
        /// <summary>
        /// Description: It should restart windows explorer
        /// </summary>
        public static void RestartExplorer()
        {
        	try
        	{
        		Report.Info("Killing explorer process.");
        		TaskkillProcess("explorer");
        		repo.Explorer.TaskBarInfo.WaitForExists(10000);
        		Report.Info("Explorer has restarted");  		
        	} catch (Exception e) {
        		Report.Warn("Explorer has not restarted.");
        		OpeningApps.StartExplorer();
	        	if (repo.Explorer.TaskBarInfo.Exists()) 
	        		Report.Info("Explorer has restarted");
	        	else 
	        		Report.Error("Could not start explorer");Report.Error(e.Message);		
        	}
        }
        
        /// <summary>
        /// Close specific Windows Application
        /// </summary>
        /// <param name="application">Application name is required</param>
        public static void CloseApplication(string application)
        {
        	TaskkillProcess(application);
        }
	}
}
