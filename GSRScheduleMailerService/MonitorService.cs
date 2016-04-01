using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;
using System.Text;
using System.Data.SqlClient;
using System.Net.Mail;
using System.IO;
using System.Reflection;
using System.Timers;
using INI;
using LogUtility;

namespace MonitorService
{   

    public partial class MonitorService : ServiceBase
    {
        public bool refreshCacheEvents = true;
        protected Timer timer;
        string AppPath = AppDomain.CurrentDomain.BaseDirectory;
        public DataTable CloseEvents = null;
        public MonitorService()
        {  
            this.ServiceName = "MonitorService";
            this.CanPauseAndContinue = false;
            
            timer = new Timer();
            timer.Interval = 2000;
            timer.Enabled = true;
            timer.Elapsed += new ElapsedEventHandler(OnTimer);
			
         }
    
        private static bool SendMailMessage(MailSettings ms, bool Enabled)
        {
            if (!Enabled)
                return true;

            Log LocalLog = new Log(AppDomain.CurrentDomain.BaseDirectory + "Log.xml");

            SmtpClient client = new SmtpClient(ms.SMTPServer);
            MailMessage message = new MailMessage() { From = new MailAddress(ms.FromAddress, ms.FromName) };


            if (ms.ToAddress.Trim() != "")
            {
                ms.ToAddress = ms.ToAddress.Replace(";", ",");
                string[] TO = ms.ToAddress.Split(',');
                for (int i = 0; i < TO.Length; i++)
                    message.To.Add(new MailAddress(TO[i], ""));
            }
            if (ms.CCAddress.Trim() != "")
            {
                ms.CCAddress = ms.CCAddress.Replace(";", ",");
                string[] CCC = ms.CCAddress.Split(',');
                for (int i = 0; i < CCC.Length; i++)
                    message.CC.Add(new MailAddress(CCC[i], ""));
            }


            message.Subject = ms.Subject;
            message.Body = ms.MailBody;
            message.IsBodyHtml = true;
            try
            {
                client.Send(message);
                LocalLog.AddMessage(Log.LogMessageType.Info, string.Format("Inviata mail: {0}", ms.Subject));
            }
            catch (Exception ex)
            {
                LocalLog.AddMessage(Log.LogMessageType.Error, string.Format("Impossibile inviare mail. Si e' verificata la seguente eccezione: {0}", ex.ToString()));
            }

            return true;
        }

        protected void MainCycle()
        {
            Log LocalLog = new Log(AppPath + "Log.xml");
            IniFile settings = new IniFile(AppPath + "ServiceMonitor.ini");
            

            int PollingMinutes = Convert.ToInt16(settings.IniReadValue("MAIN", "PollingPeriodMin"));
            bool InvioAbilitato = (Convert.ToInt16(settings.IniReadValue("MAIN", "invioMail")) == 1);
            string ServiceToMonitor = settings.IniReadValue("MAIN", "ServiceToMonitor");
            string EventViewerSource = settings.IniReadValue("EVENTVIEWER", "Source");
            int timeoutMilliseconds = Convert.ToInt32(settings.IniReadValue("MAIN", "msTimeoutPeriod"));
            TimeSpan timeout = TimeSpan.FromMilliseconds(timeoutMilliseconds);

            int ErrorCodeToTrap = Convert.ToInt32(settings.IniReadValue("EVENTVIEWER", "ErrorToCatch"));
            string EventViewerLog = settings.IniReadValue("EVENTVIEWER", "LogName");

            MailSettings ms= new MailSettings(
                settings.IniReadValue("SMTP", "server"),
                settings.IniReadValue("SMTP", "From"),
                settings.IniReadValue("SMTP", "FromName"),
                settings.IniReadValue("SMTP", "To"),
                settings.IniReadValue("SMTP", "CC"),
                settings.IniReadValue("SMTP", "Subject"),
                ""
                );

            EventLog el = new EventLog(EventViewerLog);

            
            foreach (EventLogEntry logEntry in el.Entries)
            {
                if (logEntry.TimeGenerated > DateTime.Now.AddMinutes(-1 * PollingMinutes).
                    AddMilliseconds(-2000d))  //we check the events occurred since last execution of MainCycle(). Added an overlapping 2 second interval to avoid missing any event
                { 
                    if (logEntry.InstanceId == ErrorCodeToTrap && logEntry.EntryType==EventLogEntryType.Error && logEntry.Source==EventViewerSource)
                    {
                        LocalLog.AddMessage(Log.LogMessageType.Error, string.Format("Individuato nel registro eventi {0}, l'evento di ID:{1}. Si procede al restart del servizio {2}", 
                            EventViewerLog, ErrorCodeToTrap.ToString(), ServiceToMonitor));
                        try
                        {
                            ServiceController service = new ServiceController(ServiceToMonitor);
                            LocalLog.AddMessage(Log.LogMessageType.Warning, string.Format("Il servizio {0} si trova correntemente in fase {1}. Tuttavia è occorso l'errore {2} e si procederà al riavvio.", service.ServiceName, service.Status, ErrorCodeToTrap));

                            service.Stop();
                            service.WaitForStatus(ServiceControllerStatus.Stopped, timeout);

                            if(service.Status != ServiceControllerStatus.Stopped)
                            {
                                LocalLog.AddMessage(Log.LogMessageType.Error, string.Format("Impossibile arrestare il servizio {0} a seguito di un evento {1}", ServiceToMonitor, ErrorCodeToTrap));
                                ms.Subject = ms.Subject.Replace("%NOME%", string.Format("Impossibile arrestare il servizio {0}", service.ServiceName));
                                ms.MailBody = string.Format("Attenzione, il servizio {0} a seguito di un evento {1} richiede un restart che non è stato possibile effettuare. Si prega di procedere manualmente.",
                                    service.ServiceName, ErrorCodeToTrap);
                                SendMailMessage(ms, InvioAbilitato);
                                return;
                            }

                            service.Start();
                            service.WaitForStatus(ServiceControllerStatus.Running, timeout);


                            if (service.Status == ServiceControllerStatus.Running)
                            {
                                LocalLog.AddMessage(Log.LogMessageType.Info, string.Format("Il servizio {0} e' stato arrestato ed e' stato correttamente riavviato nel tempo di timeout ({1}ms). Tutto Ok", ServiceToMonitor, timeoutMilliseconds.ToString()));
                                ms.Subject = ms.Subject.Replace("%NOME%", string.Format("Servizio {0} riavviato correttamente", service.ServiceName));
                                ms.MailBody = string.Format("E' stato individuato nel registro eventi {0} l'errore {1}. Il servizio {2} e' stato arrestato e successivamente riavviato da MonitorService",
                                    EventViewerLog, ErrorCodeToTrap.ToString(), service.ServiceName);
                                SendMailMessage(ms, InvioAbilitato);

                            }
                            else
                            {
                                LocalLog.AddMessage(Log.LogMessageType.Error, string.Format("Il Servizio {0} non si e' riavviato entro il periodo di timeout. Verrà effettuato un nuovo tentativo al prossimo polling", ServiceToMonitor));
                                ms.Subject = ms.Subject.Replace("%NOME%", string.Format("Impossibile riavviare il servizio {0}", service.ServiceName));
                                ms.MailBody = string.Format("Il servizio {0} in seguito all'evento di id {4} si trova in fase {1} e non si e' riavviato nel timeout period ({2}ms). Un nuovo tentativo verra' efettuato tra {3} minuti",
                                    service.ServiceName, service.Status.ToString(), timeoutMilliseconds.ToString(), settings.IniReadValue("MAIN", "PollingPeriodMin"), ErrorCodeToTrap);
                                SendMailMessage(ms, InvioAbilitato);
                            }



                        }
                        catch (Exception ex)
                        {
                            LocalLog.AddMessage(Log.LogMessageType.Error, string.Format("Attenzione, durante la fase di gestione dell'errore {0} Si è verificata la seguente eccezione durante l'esecuzione di MonitorService: {1}", ErrorCodeToTrap, ex.ToString()));
                            ms.Subject = ms.Subject.Replace("%NOME%", "Eccezione durante l'esecuzione di MonitorService");
                            ms.MailBody = string.Format("Attenzione, durante la fase di gestione dell'errore Event Log {0} si e' verificata la seguente eccezione {1}", ErrorCodeToTrap, ex.ToString());
                            SendMailMessage(ms, InvioAbilitato);
                        }
                        return;  //do not check if service is running.
                    }
                }
            }




            try
            {
                 ServiceController service = new ServiceController(ServiceToMonitor);
            
                if (service.Status == ServiceControllerStatus.Running)
                {
#if DEBUG
                    LocalLog.AddMessage(Log.LogMessageType.Info, "Servizio " + service.ServiceName + " up and running. Tutto Ok");
#endif
                }
                else
                {
                    

                    LocalLog.AddMessage(Log.LogMessageType.Warning, string.Format("Il servizio {0} si trova correntemente in fase {1}. Verrà ora eseguito un tentativo di restart", service.ServiceName, service.Status));

                    service.Start();
                    service.WaitForStatus(ServiceControllerStatus.Running, timeout);

                    if (service.Status == ServiceControllerStatus.Running)
                    {
                        LocalLog.AddMessage(Log.LogMessageType.Info, string.Format("Servizio {0} ripartito nel tempo di timeout ({1}ms). Tutto Ok", service.ServiceName, timeoutMilliseconds));
                        ms.Subject = ms.Subject.Replace("%NOME%", string.Format("Servizio {0} riavviato correttamente", service.ServiceName));
                        ms.MailBody = string.Format("Il servizio {0} e' stato appena riavviato da MonitorService", service.ServiceName);
                        SendMailMessage(ms, InvioAbilitato);
                       
                    }
                    else
                    {
                        LocalLog.AddMessage(Log.LogMessageType.Error, string.Format("Servizio {0} non riavviato entro {1}ms. Nuovo tentativo al prossimo polling", service.ServiceName, timeoutMilliseconds));
                        ms.Subject = ms.Subject.Replace("%NOME%", string.Format("Impossibile riavviare il servizio {0}", service.ServiceName));
                        ms.MailBody = string.Format("Il servizio {0} si trova in fase {1} e non si e' riavviato nel timeout period ({2}ms). Un nuovo tentativo verra' efettuato tra {3} minuti", 
                            service.ServiceName, service.Status.ToString(), timeoutMilliseconds.ToString(), settings.IniReadValue("MAIN", "PollingPeriodMin"));
                        SendMailMessage(ms, InvioAbilitato);
                    }
                }
            }
            catch(Exception ex)
            {
                LocalLog.AddMessage(Log.LogMessageType.Error, string.Format("Si è verificata la seguente eccezione durante l'esecuzione di MonitorService: {0}", ex.ToString()));
                ms.Subject = ms.Subject.Replace("%NOME%", "Eccezione durante l'esecuzione di MonitorService");
                ms.MailBody = string.Format("Attenzione, durante l'esecuzione del servizio MonitorService si e' verificata la seguente eccezione {0}", ex.ToString());
                SendMailMessage(ms, InvioAbilitato);
            }
        }

        protected override void OnStart(string[] args)
         {
             Log LocalLog = new Log(AppPath + "Log.xml");
             LocalLog.AddMessage(Log.LogMessageType.Info, "Servizio avviato");
         }

        protected override void OnStop()
        {
            Log LocalLog = new Log(AppPath + "Log.xml");
            timer.Enabled = false;
            LocalLog.AddMessage(Log.LogMessageType.Info, "Servizio arrestato");
        }


      

        private void OnTimer(object sender, EventArgs e)
        {
            IniFile settings = new IniFile(AppPath + "ServiceMonitor.ini");
            int PollingMinutes = Convert.ToInt16(settings.IniReadValue("MAIN", "PollingPeriodMin"));

            timer.Interval = 60000 * PollingMinutes;
#if DEBUG
            timer.Interval = 60000;
#endif
            MainCycle();
           
        }
    }
}
