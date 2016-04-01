using System;
using System.Collections.Generic;
using System.Text;

namespace MonitorService
{
    public class MailSettings
    {
        
        public string SMTPServer { get; set; }
        public string FromAddress { get; set; }
        public string FromName { get; set; }
        public string ToAddress { get; set; }
        public string CCAddress { get; set; }
        public string Subject { get; set; }
        public string MailBody { get; set; }

       
        public MailSettings(string _smtpServer, string _fromAddress, string _fromName, string _toAddress, string _ccAddress, string _subject, string _mailBody)
        {
            SMTPServer = _smtpServer;
            FromAddress = _fromAddress;
            FromName = _fromName;
            ToAddress = _toAddress;
            CCAddress = _ccAddress;
            Subject = _subject;
            MailBody = _mailBody;
        }

    }
}
