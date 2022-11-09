/*
 * Created by Ranorex
 * User: alberto.hirota
 * Date: 11/6/2022
 * Time: 5:52 PM
 * 
 * To change this template use Tools > Options > Coding > Edit standard headers.
 */
using System;
using Ranorex;

namespace Ranorex_POC.Outlook
{
	/// <summary>
	/// Description of MessageMethods.
	/// </summary>
	public class MessageMethods
	{
		public static Ranorex_POCRepository repo = Ranorex_POCRepository.Instance;
		
		public MessageMethods()
		{
		}
		
		/// <summary>
		/// Populate all email
		/// </summary>
		/// <param name="email">Recipient email</param>
		/// <param name="subject">Email subject</param>
		/// <param name="emailBody">Email body info</param>
		public static void PopulateNewEmail(string email, string subject, string emailBody)
		{
			PopulateEmail(email);
			PopulateSubject(subject);
			PopulateEmailBody(emailBody);
		}
		
		/// <summary>
		/// Funtion to click Send Email Button
		/// </summary>
		public static void ClickSendEmail()
		{
			Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'OutlookMessage.ButtonSend' at Center.", repo.OutlookMessage.ButtonSendInfo);
            repo.OutlookMessage.ButtonSend.Click();
            Delay.Milliseconds(200);
		}
		
		/// <summary>
		/// Description: Populate Email Body
		/// </summary>
		/// <param name="emailBody">Email body description</param>
		public static void PopulateEmailBody(string emailBody)
		{
			Report.Log(ReportLevel.Info, "Keyboard", "Typing email body: "+emailBody);
			repo.OutlookMessage.EmailBody.PressKeys(emailBody);
			Delay.Milliseconds(200);
			repo.OutlookMessage.EmailBody.PressKeys("{Return}");
            Delay.Milliseconds(200);
		}
		
		/// <summary>
		/// Description: Populate Email Subject
		/// </summary>
		/// <param name="subject">Require subject string</param>
		public static void PopulateSubject(string subject)
		{
			Report.Log(ReportLevel.Info, "Keyboard", "Typing subject: "+subject);
			repo.OutlookMessage.Subject.PressKeys("{Back}");
			Delay.Milliseconds(200);
			repo.OutlookMessage.Subject.PressKeys(subject);
			Delay.Milliseconds(200);
			repo.OutlookMessage.Subject.PressKeys("{Return}");
            Delay.Milliseconds(200);
		}
		
		/// <summary>
		/// Populate Email 'To'
		/// </summary>
		/// <param name="email">Require recipient email</param>
		public static void PopulateEmail(string email)
		{
			Report.Log(ReportLevel.Info, "Keyboard", "Typing email: "+email);
			repo.OutlookMessage.To.PressKeys("{Back}");
			Delay.Milliseconds(200);
			repo.OutlookMessage.To.PressKeys(email);
			Delay.Milliseconds(200);
			repo.OutlookMessage.To.PressKeys("{Return}");
            Delay.Milliseconds(200);
		}
		
		/// <summary>
		/// Method to send a standard new email. It require 3 arguments
		/// </summary>
		/// <param name="subject">This argument should be the email subject</param>
		/// <param name="body">This argument should be the email body</param>
		/// <param name="email">Optional argument to include the recipient email, if you do not add this argument, them it will pick the computer user and default local domain</param>
		public static void SendNewEmail(string subject, string body, string email=null)
		{
			if(email == null)
				email = Environment.UserName + "@" +Environment.UserDomainName + ".local"; 
			Outlook.OutlookMethods.CreateNewEmail();
            Outlook.MessageMethods.PopulateNewEmail(email,subject,body);
            Outlook.MessageMethods.ClickSendEmail();
		}
		
		/// <summary>
		/// Click CloseMessage button
		/// </summary>
		public static void ClickCloseMessageAndDoNotSave()
		{
			Report.Info("Checking if Message window is open...");
			if(repo.OutlookMessage.SelfInfo.Exists())
			{
				Report.Info("Message is open...");
				Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'OutlookMessage.ButtonCloseMessage' at Center.", repo.OutlookMessage.ButtonCloseMessageInfo);
            	repo.OutlookMessage.ButtonCloseMessage.Click();
            	Delay.Milliseconds(200);
			}
            
            Report.Info("Checking if 'You want to save your message' warns exists");
            if(repo.OfficeWarn.SelfInfo.Exists())
            {
            	Report.Info("Warn exists...");
            	Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'OutlookWarn.ButtonNo' at Center.", repo.OfficeWarn.ButtonNoInfo);
            	repo.OfficeWarn.ButtonNo.Click();
            	Delay.Milliseconds(200);
            }
		}
		
		/// <summary>
		/// Click Reply Message button
		/// </summary>
		public static void ClickReplyMessage()
		{
			Report.Log(ReportLevel.Info, "Mouse", "Mouse Click item 'OutlookMessage.ButtonReply' at Center.", repo.OutlookMessage.ButtonReplyInfo);
            repo.OutlookMessage.ButtonReply.Click();
            Delay.Milliseconds(200);
		}
		
	}
}
