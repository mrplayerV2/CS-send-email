using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using System.Net.Mail;
public class EmailUtil
{
	public EmailUtil(){	
	}

    public static bool SendMail(string strFrom, string strTo, string strSubject, string strBody){
        bool bRet = false;
        try       {
            string[] Tos = strTo.Split(',');
            string[] BCC = ConfigurationManager.AppSettings["mailFromCC"].ToString().Split(',');
            MailMessage message = new MailMessage();
            message.From = new MailAddress(strFrom);
            foreach (string toemail in Tos) message.To.Add(new MailAddress(toemail));
            foreach (string bbcemail in BCC) message.Bcc.Add(new MailAddress(bbcemail));
            message.Subject = strSubject;
            message.Body = strBody;
            message.BodyEncoding = System.Text.Encoding.UTF8;
            message.IsBodyHtml = true;
            SmtpClient emailHost = new SmtpClient();
            emailHost.Host = ConfigurationManager.AppSettings["mailServer"].ToString();
            string UserName = ConfigurationManager.AppSettings["UserName"].ToString();
            string Pw = ConfigurationManager.AppSettings["Password"].ToString();
            emailHost.Credentials = new System.Net.NetworkCredential(UserName, Pw);
            emailHost.Send(message);
            bRet = true;
        }
        catch (Exception ex){
            throw ex;
            bRet = false;
        }
        return bRet;
    }

    public static bool SendMail(string strFrom, string strTo, string strSubject, string strBody, string[] strFile){
        bool bRet = false;
        try{
            string[] Tos = strTo.Split(',');
            string[] BCC = ConfigurationManager.AppSettings["mailFromCC"].ToString().Split(',');
            MailMessage message = new MailMessage();
            message.From = new MailAddress(strFrom);
            foreach (string toemail in Tos) message.To.Add(new MailAddress(toemail));
            foreach (string bbcemail in BCC) message.Bcc.Add(new MailAddress(bbcemail));
            message.Subject = strSubject;
            message.Body = strBody;
            message.BodyEncoding = System.Text.Encoding.UTF8;
            message.IsBodyHtml = true;
            for (int i = 0; i < strFile.Length; i++){
                if (System.IO.File.Exists(strFile[i]))
                    message.Attachments.Add(new Attachment(strFile[i]));
            }
            SmtpClient emailHost = new SmtpClient();
            emailHost.Host = ConfigurationManager.AppSettings["mailServer"].ToString();
            string UserName = ConfigurationManager.AppSettings["UserName"].ToString();
            string Pw = ConfigurationManager.AppSettings["Password"].ToString();
            emailHost.Credentials = new System.Net.NetworkCredential(UserName, Pw);
            emailHost.Send(message);
            bRet = true;
        }
         catch (Exception ex)
        {
            throw ex;
            bRet = false;
        }
        return bRet;
    }
    
    public static bool SendMailSSL(string strTo, string strSubject, string strBody){
        bool bRet = false;
        try{
            const string SMTP_SERVER = "http://schemas.microsoft.com/cdo/configuration/smtpserver";
            const string SMTP_SERVER_PORT = "http://schemas.microsoft.com/cdo/configuration/smtpserverport";
            const string SEND_USING = "http://schemas.microsoft.com/cdo/configuration/sendusing";
            const string SMTP_USE_SSL = "http://schemas.microsoft.com/cdo/configuration/smtpusessl";
            const string SMTP_AUTHENTICATE = "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate";
            const string SEND_USERNAME = "http://schemas.microsoft.com/cdo/configuration/sendusername";
            const string SEND_PASSWORD = "http://schemas.microsoft.com/cdo/configuration/sendpassword";
            string mailServer = ConfigurationManager.AppSettings["mailServer"].ToString();
            string UserName = ConfigurationManager.AppSettings["UserName"].ToString();
            string Pw = ConfigurationManager.AppSettings["Password"].ToString();
            string mailFrom = ConfigurationManager.AppSettings["mailFrom"].ToString();
            string mailBCC = ConfigurationManager.AppSettings["mailFromCC"].ToString();
            System.Web.Mail.MailMessage mail = new System.Web.Mail.MailMessage();
            mail.Fields[SMTP_SERVER] = mailServer;
            mail.Fields[SMTP_SERVER_PORT] = 465;
            mail.Fields[SEND_USING] = 2;
            mail.Fields[SMTP_USE_SSL] = true;
            mail.Fields[SMTP_AUTHENTICATE] = 1;
            mail.Fields[SEND_USERNAME] = UserName;
            mail.Fields[SEND_PASSWORD] = Pw;
            mail.Subject = strSubject;
            mail.Body = strBody;
            mail.BodyFormat = System.Web.Mail.MailFormat.Html;
            mail.BodyEncoding = System.Text.Encoding.UTF8;
            mail.Bcc = mailBCC;
            mail.From = mailFrom;
            mail.To = strTo;
            System.Web.Mail.SmtpMail.Send(mail);

            bRet = true;
        }
         catch (Exception ex)
        {
            throw ex;
            bRet = false;
        }
        return bRet;
    }
  
    public static bool SendMailSSL(string strTo, string strSubject, string strBody, string[] strFile){
        bool bRet = false;
        try{
            const string SMTP_SERVER = "http://schemas.microsoft.com/cdo/configuration/smtpserver";
            const string SMTP_SERVER_PORT = "http://schemas.microsoft.com/cdo/configuration/smtpserverport";
            const string SEND_USING = "http://schemas.microsoft.com/cdo/configuration/sendusing";
            const string SMTP_USE_SSL = "http://schemas.microsoft.com/cdo/configuration/smtpusessl";
            const string SMTP_AUTHENTICATE = "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate";
            const string SEND_USERNAME = "http://schemas.microsoft.com/cdo/configuration/sendusername";
            const string SEND_PASSWORD = "http://schemas.microsoft.com/cdo/configuration/sendpassword";
            string mailServer = ConfigurationManager.AppSettings["mailServer"].ToString();
            string UserName = ConfigurationManager.AppSettings["UserName"].ToString();
            string Pw = ConfigurationManager.AppSettings["Password"].ToString();
            string mailFrom = ConfigurationManager.AppSettings["mailFrom"].ToString();
            string mailBCC = ConfigurationManager.AppSettings["mailFromCC"].ToString();
            System.Web.Mail.MailMessage mail = new System.Web.Mail.MailMessage();
            mail.Fields[SMTP_SERVER] = mailServer;
            mail.Fields[SMTP_SERVER_PORT] = 465;
            mail.Fields[SEND_USING] = 2;
            mail.Fields[SMTP_USE_SSL] = true;
            mail.Fields[SMTP_AUTHENTICATE] = 1;
            mail.Fields[SEND_USERNAME] = UserName;
            mail.Fields[SEND_PASSWORD] = Pw;
            mail.Subject = strSubject;
            mail.Body = strBody;
            for (int i = 0; i < strFile.Length; i++){
                if (System.IO.File.Exists(strFile[i])) 
                 mail.Attachments.Add(new System.Web.Mail.MailAttachment(strFile[i]));
            }
            mail.BodyFormat = System.Web.Mail.MailFormat.Html;
            mail.BodyEncoding = System.Text.Encoding.UTF8;
            mail.Bcc = mailBCC;
            mail.From = mailFrom;
            mail.To = strTo;
            System.Web.Mail.SmtpMail.Send(mail);
            bRet = true;
        }
        catch (Exception ex){
            bRet = false;
        }
        return bRet;
    }
    public static bool SendMail(string strTo, string strSubject, string strBody)
    {
        bool bRet = false;
        try
        {
            string[] Tos = strTo.Split(',');
            string[] BCC = ConfigurationManager.AppSettings["mailFromCC"].ToString().Split(',');
            MailMessage message = new MailMessage();
            message.From = new MailAddress(ConfigurationManager.AppSettings["mailFrom"].ToString(), ConfigurationManager.AppSettings["DisplayName"].ToString());
            foreach (string toemail in Tos) message.To.Add(new MailAddress(toemail));
            foreach (string bbcemail in BCC) message.Bcc.Add(new MailAddress(bbcemail));
            message.Subject = strSubject;
            message.Body = strBody;
            message.BodyEncoding = System.Text.Encoding.UTF8;
            message.IsBodyHtml = true;
            SmtpClient emailHost = new SmtpClient();
            emailHost.Host = ConfigurationManager.AppSettings["mailServer"].ToString();
            //----------------------
            string UserName = ConfigurationManager.AppSettings["UserName"].ToString();
            string Pw = ConfigurationManager.AppSettings["Password"].ToString();
            emailHost.Credentials = new System.Net.NetworkCredential(UserName, Pw);
            //----------------------
            emailHost.Send(message);
            bRet = true;
        }
        catch (Exception ex)
        { 
            bRet = false;
         return SendMailSSL(strTo, strSubject, strBody);

        }
        return bRet;
    }
    public static bool SendMail(string strTo, string strSubject, string strBody, string[] strFile)
    {
        bool bRet = false;
        try
        {
            string[] Tos = strTo.Split(',');
            string[] BCC = ConfigurationManager.AppSettings["mailFromCC"].ToString().Split(',');
            MailMessage message = new MailMessage();
            message.From = new MailAddress(ConfigurationManager.AppSettings["mailFrom"].ToString(), ConfigurationManager.AppSettings["DisplayName"].ToString());

            foreach (string toemail in Tos) message.To.Add(new MailAddress(toemail));
            foreach (string bbcemail in BCC) message.Bcc.Add(new MailAddress(bbcemail));

			 for (int i = 0; i < strFile.Length; i++)
            {
                if (System.IO.File.Exists(strFile[i])) message.Attachments.Add(new Attachment(strFile[i]));
            }
            message.Subject = strSubject;
            message.Body = strBody;
            message.BodyEncoding = System.Text.Encoding.UTF8;
            message.IsBodyHtml = true;
            SmtpClient emailHost = new SmtpClient();
            emailHost.Host = ConfigurationManager.AppSettings["mailServer"].ToString();
            string UserName = ConfigurationManager.AppSettings["UserName"].ToString();
            string Pw = ConfigurationManager.AppSettings["Password"].ToString();
            emailHost.Credentials = new System.Net.NetworkCredential(UserName, Pw);
            emailHost.Send(message);
            bRet = true;
        } 
        catch (Exception ex){
         return SendMailSSL(strTo, strSubject, strBody ,strFile);
            bRet = false;
        }
        return bRet;
    }
    public static bool SendMail(string FromName, string strFrom, string ToName, string strTo, string strSubject, string strBody, string[] strFile)
    {
        bool bRet = false;
        try{
            string[] BCC = ConfigurationManager.AppSettings["mailFromCC"].ToString().Split(',');
            MailMessage message = new MailMessage();
            message.From = new MailAddress(strFrom, FromName);
            message.To.Add(new MailAddress(strTo, ToName));
            foreach (string bbcemail in BCC) message.Bcc.Add(new MailAddress(bbcemail));
            message.Subject = strSubject;
            message.Body = strBody;
            message.BodyEncoding = System.Text.Encoding.UTF8;
            message.IsBodyHtml = true;
            for (int i = 0; i < strFile.Length; i++)
            {
                if (System.IO.File.Exists(strFile[i]))
                    message.Attachments.Add(new Attachment(strFile[i]));
            }
            SmtpClient emailHost = new SmtpClient();
            emailHost.Host = ConfigurationManager.AppSettings["mailServer"].ToString();
            string UserName = ConfigurationManager.AppSettings["UserName"].ToString();
            string Pw = ConfigurationManager.AppSettings["Password"].ToString();
            emailHost.Credentials = new System.Net.NetworkCredential(UserName, Pw);
            emailHost.Send(message);
            bRet = true;
        }
         catch (Exception ex)
        {
            throw ex;
            bRet = false;
        }
        return bRet;
    }
    public static bool SendMail(string FromName, string strFrom, string ToName, string strTo, string strSubject, string strBody)
    {
        bool bRet = false;
        try
        {
            string[] BCC = ConfigurationManager.AppSettings["mailFromCC"].ToString().Split(',');
            MailMessage message = new MailMessage();
            message.From = new MailAddress(strFrom, FromName);
            message.To.Add(new MailAddress(strTo, ToName));
            foreach (string bbcemail in BCC) message.Bcc.Add(new MailAddress(bbcemail));
            message.Subject = strSubject;
            message.Body = strBody;
            message.BodyEncoding = System.Text.Encoding.UTF8;
            message.IsBodyHtml = true;
            SmtpClient emailHost = new SmtpClient();
            emailHost.Host = ConfigurationManager.AppSettings["mailServer"].ToString();
            //----------------------
            string UserName = ConfigurationManager.AppSettings["UserName"].ToString();
            string Pw = ConfigurationManager.AppSettings["Password"].ToString();
            emailHost.Credentials = new System.Net.NetworkCredential(UserName, Pw);
            emailHost.Send(message);
            bRet = true;
        }
        catch (Exception ex)
        {
            HttpContext.Current.Response.Write(ex.Message.ToString());
            bRet = false;
        }
        return bRet;
    }
    public static bool SendMail(string FromName, string strFrom, string ToName, string strTo, string strCC, string strSubject, string strBody){
        bool bRet = false;
        try
        {
            string[] BCC = ConfigurationManager.AppSettings["mailFromCC"].ToString().Split(',');
            MailMessage message = new MailMessage();
            message.From = new MailAddress(strFrom, FromName);
            message.To.Add(new MailAddress(strTo, ToName));
            foreach (string bbcemail in BCC) message.Bcc.Add(new MailAddress(bbcemail));
            message.CC.Add(strCC);
            message.Subject = strSubject;
            message.Body = strBody;
            message.BodyEncoding = System.Text.Encoding.UTF8;
            message.IsBodyHtml = true;
            SmtpClient emailHost = new SmtpClient();
            emailHost.Host = ConfigurationManager.AppSettings["mailServer"].ToString();
            string UserName = ConfigurationManager.AppSettings["UserName"].ToString();
            string Pw = ConfigurationManager.AppSettings["Password"].ToString();
            emailHost.Credentials = new System.Net.NetworkCredential(UserName, Pw);
            emailHost.Send(message);
            bRet = true;
        }
        catch (Exception ex)
        {
            HttpContext.Current.Response.Write(ex.Message.ToString());
            bRet = false;
        }
        return bRet;
    }   
}