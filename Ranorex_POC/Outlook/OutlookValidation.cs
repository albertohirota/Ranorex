/*
 * Created by Ranorex
 * User: alberto.hirota
 * Date: 11/7/2022
 * Time: 10:14 AM
 * 
 * To change this template use Tools > Options > Coding > Edit standard headers.
 */
using System;
using Ranorex;

namespace Ranorex_POC.Outlook
{
	/// <summary>
	/// Description of OutlookValidation.
	/// </summary>
	public class OutlookValidation
	{
		public static Ranorex_POCRepository repo = Ranorex_POCRepository.Instance;
		
		public OutlookValidation()
		{
		}
		
		/// <summary>
		/// Validating Email subject Exists
		/// </summary>
		/// <param name="emailSubject">Required Subject text to be validated</param>
		public static void EmailSubjectExists(string emailSubject)
		{
			repo.EmailSubject = emailSubject;
			try
        	{
        		Report.Log(ReportLevel.Info, "Validation", "Validating appearance of item 'repo.Outlook.EmailList.EmailSubject': "+emailSubject);
      			Validate.Exists(repo.Outlook.EmailList.EmailSubjectInfo);
      			Delay.Milliseconds(100);
        	}
			catch (Exception ex)
			{
				Report.Screenshot();
				Report.Error(ex.Message);
			}
		}
		
		/// <summary>
		/// Validating Email body Exists
		/// </summary>
		/// <param name="emailBody">Required EmailBody text to be validated</param>
		public static void EmailBodyExists(string emailBody)
		{
			try
        	{
        		Report.Log(ReportLevel.Info, "Validation", "Validating appearance of item 'repo.Outlook.EmailBottomView.EmailBody': "+emailBody);
      			Validate.AttributeContains(repo.Outlook.EmailBottomView.EmailBodyInfo,"Text",emailBody);
      			Delay.Milliseconds(100);
        	}
			catch (Exception ex)
			{
				Report.Screenshot();
				Report.Error(ex.Message);
			}
		}
		
		
		public static void EmailRecipientExists(string emailRecipient)
		{
			try
        	{
        		Report.Log(ReportLevel.Info, "Validation", "Validating appearance of item 'repo.Outlook.EmailBottomView.EmailRecipient': "+emailRecipient);
      			Validate.AttributeContains(repo.Outlook.EmailBottomView.EmailRecipientInfo,"Text",emailRecipient);
      			Delay.Milliseconds(100);
        	}
			catch (Exception ex)
			{
				Report.Screenshot();
				Report.Error(ex.Message);
			}
		}
	}
}
