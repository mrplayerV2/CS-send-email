public static bool SendMail(string strFrom, string strTo, string strSubject, string strBody)
	{
		bool flag = false;
		try
		{
			MailAddress from = new MailAddress(strFrom);
			MailAddress to = new MailAddress(strTo);
			MailMessage mailMessage = new MailMessage(from, to);
			mailMessage.Bcc.Add(ConfigurationManager.AppSettings["bccMail"].ToString());
			mailMessage.Subject = strSubject;
			mailMessage.Body = strBody;
			mailMessage.BodyEncoding = Encoding.UTF8;
			mailMessage.IsBodyHtml = true;
			SmtpClient smtpClient = new SmtpClient();
			smtpClient.Host = ConfigurationManager.AppSettings["MailServer"].ToString();
			smtpClient.Send(mailMessage);
			return true;
		}
		catch
		{
			return false;
		}
	}

	public static bool SendMail(string strFrom, string strTo, string strSubject, string strBody, string[] strFile)
	{
		bool flag = false;
		try
		{
			MailAddress from = new MailAddress(strFrom);
			MailAddress to = new MailAddress(strTo);
			MailMessage mailMessage = new MailMessage(from, to);
			mailMessage.Bcc.Add(ConfigurationManager.AppSettings["bccMail"].ToString());
			mailMessage.Subject = strSubject;
			mailMessage.Body = strBody;
			mailMessage.BodyEncoding = Encoding.UTF8;
			mailMessage.IsBodyHtml = true;
			for (int i = 0; i < strFile.Length; i++)
			{
				if (File.Exists(strFile[i]))
				{
					mailMessage.Attachments.Add(new Attachment(strFile[i]));
				}
			}
			SmtpClient smtpClient = new SmtpClient();
			smtpClient.Host = ConfigurationManager.AppSettings["MailServer"].ToString();
			smtpClient.Send(mailMessage);
			return true;
		}
		catch
		{
			return false;
		}
	}

	public static bool SendMail(string strTo, string strSubject, string strBody)
	{
		bool flag = false;
		try
		{
			MailAddress from = new MailAddress(ConfigurationManager.AppSettings["SenderAddr"].ToString(), ConfigurationManager.AppSettings["DisplayName"].ToString());
			MailAddress to = new MailAddress(strTo);
			MailMessage mailMessage = new MailMessage(from, to);
			mailMessage.Subject = strSubject;
			mailMessage.Body = strBody;
			mailMessage.BodyEncoding = Encoding.UTF8;
			mailMessage.IsBodyHtml = true;
			SmtpClient smtpClient = new SmtpClient();
			smtpClient.Host = ConfigurationManager.AppSettings["MailServer"].ToString();
			smtpClient.Credentials = new NetworkCredential(ConfigurationManager.AppSettings["SenderID"].ToString(), ConfigurationManager.AppSettings["SenderPWD"].ToString());
			smtpClient.Send(mailMessage);
			return true;
		}
		catch
		{
			return false;
		}
	}

	public static bool SendMail(string strTo, string strSubject, string strBody, string[] strFile)
	{
		bool flag = false;
		try
		{
			MailAddress from = new MailAddress(ConfigurationManager.AppSettings["SenderAddr"].ToString(), ConfigurationManager.AppSettings["DisplayName"].ToString());
			MailAddress to = new MailAddress(strTo);
			MailMessage mailMessage = new MailMessage(from, to);
			mailMessage.Bcc.Add(ConfigurationManager.AppSettings["bccMail"].ToString());
			mailMessage.Subject = strSubject;
			mailMessage.Body = strBody;
			mailMessage.BodyEncoding = Encoding.UTF8;
			mailMessage.IsBodyHtml = true;
			for (int i = 0; i < strFile.Length; i++)
			{
				if (File.Exists(strFile[i]))
				{
					mailMessage.Attachments.Add(new Attachment(strFile[i]));
				}
			}
			SmtpClient smtpClient = new SmtpClient();
			smtpClient.Host = ConfigurationManager.AppSettings["MailServer"].ToString();
			smtpClient.Credentials = new NetworkCredential(ConfigurationManager.AppSettings["SenderID"].ToString(), ConfigurationManager.AppSettings["SenderPWD"].ToString());
			smtpClient.Send(mailMessage);
			return true;
		}
		catch
		{
			return false;
		}
	}