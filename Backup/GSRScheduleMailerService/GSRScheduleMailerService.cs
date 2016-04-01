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

namespace GSRScheduleMailerService
{
    public partial class GSRScheduleMailerService : ServiceBase
    {
        protected Timer timer;
        string AppPath = AppDomain.CurrentDomain.BaseDirectory;
        public GSRScheduleMailerService()
        {

            this.ServiceName = "GSRSchedMail";
            this.CanPauseAndContinue = false;
            
            timer = new Timer();
            timer.Interval = 2000;
            timer.Enabled = true;
            timer.Elapsed += new ElapsedEventHandler(OnTimer);
			
         }
        private static string connectionString
        {
            get
            {
                string AppPath = AppDomain.CurrentDomain.BaseDirectory;
                IniFile settings = new IniFile(AppPath +"GSRSched.ini");
                string SQLC = "Data Source=@@SRV;Initial Catalog=@@DBC;User Id=@@UID;Password=@@PWD;";
                SQLC = SQLC.Replace("@@SRV", settings.IniReadValue("SQL", "Server"));
                SQLC = SQLC.Replace("@@DBC", settings.IniReadValue("SQL", "Database"));
                SQLC = SQLC.Replace("@@UID", settings.IniReadValue("SQL", "Username"));
                SQLC = SQLC.Replace("@@PWD", settings.IniReadValue("SQL", "Password"));
                return SQLC;
            }
        }
        private static DataTable cacheEvents(DateTime Ora)
        {
            string SQLConnectionString = connectionString;
            DataTable events = new DataTable();
            using (SqlConnection Conn = new SqlConnection(SQLConnectionString))
            {
                Conn.Open();
                string SQLStatement = "SELECT * FROM GSR_TaskTable WHERE Stato=0 AND DataOraInizio >= '" + Ora.AddMinutes(-31).ToString("s") + "' AND DataOraInizio <='" + Ora.AddMinutes(31).ToString("s") + "'";
                SqlCommand cmd = new SqlCommand(SQLStatement, Conn);
                cmd.CommandText = SQLStatement;
                DataSet DS = new DataSet();
                DS.Tables.Add(events);
                DS.Load(cmd.ExecuteReader(), 0, events);

            }
            return events;
        }
        private static string Dlookup(string Field, string Table, string Condition)
        {
            string SQLStatement = "SELECT " + Field + " FROM  " + Table + " WHERE " + Condition;
            using (SqlConnection Conn = new SqlConnection(connectionString))
            {
                Conn.Open();
                SqlCommand cmd = new SqlCommand(SQLStatement, Conn);
                cmd.CommandText = SQLStatement;
                object ObjDato = cmd.ExecuteScalar();
                string Dato = "";
                if (ObjDato == null)
                    return "";
                if (ObjDato.GetType() == DateTime.Now.GetType())
                    Dato = String.Format("{0:d}", ObjDato);
                else
                    Dato = Convert.ToString(ObjDato);
                Conn.Close();
                return Dato;
            }

        }
        private static bool SendMailMessage(string SMTPServer, string fromAddress, string fromName, string toAddress, string CCAddress, string msgSubject, string msgBody)
        {
            SmtpClient client = new SmtpClient(SMTPServer);
            MailMessage message = new MailMessage();
            message.From = new MailAddress(fromAddress, fromName);


            if (toAddress.Trim() != "")
            {
                toAddress = toAddress.Replace(";", ",");
                string[] TO = toAddress.Split(',');
                for (int i = 0; i < TO.Length; i++)
                    message.To.Add(new MailAddress(TO[i], ""));
            }
            if (CCAddress.Trim() != "")
            {
                CCAddress = CCAddress.Replace(";", ",");
                string[] CCC = CCAddress.Split(',');
                for (int i = 0; i < CCC.Length; i++)
                    message.CC.Add(new MailAddress(CCC[i], ""));
            }



            message.Subject = msgSubject;
            message.Body = msgBody;
            message.IsBodyHtml = true;
            client.Send(message);

            return true;
        }
        protected void MainCycle()
        {

            Log LocalLog = new Log(AppPath + "Log.xml");
           // LocalLog.AddMessage(Log.LogMessageType.Info, "Polling");
            IniFile settings = new IniFile(AppPath + "GSRSched.ini");
            DataTable CloseEvents = null;


            //main cycle
                DateTime Adesso = DateTime.Now;
                CloseEvents = cacheEvents(Adesso);
                int MinutiAvviso = 0;
                for (int i = 0; i < CloseEvents.Rows.Count; i++)
                {
                    DateTime TaskTime = (DateTime)CloseEvents.Rows[i].ItemArray[3];
                    TimeSpan TS = new TimeSpan(Adesso.Ticks - TaskTime.Ticks);

                    string TaskID = CloseEvents.Rows[i].ItemArray[1].ToString();
                    MinutiAvviso = Convert.ToInt16(Dlookup("EmailAlert", "GSR_TaskList", "ID=" + TaskID).Trim());

                   // LocalLog.AddMessage(Log.LogMessageType.Info, "TaskID:" + TaskID.ToString() +" - EventID:" +  CloseEvents.Rows[i].ItemArray[0].ToString() + " - CurrRange:" +  TS.Minutes.ToString() + " - Alert:" + MinutiAvviso.ToString() );

                    if (TS.Minutes == MinutiAvviso && MinutiAvviso != 0)
                    {
                        string MessageBody = "";
                        StreamReader StreamFile;
                        string Line = "";
                        StreamFile = File.OpenText(AppPath + settings.IniReadValue("SMTP", "MailBodyFile"));
                        Line = StreamFile.ReadLine();
                        while (Line != null)
                        {
                            MessageBody += Line;
                            Line = StreamFile.ReadLine();
                        }
                        StreamFile.Close();

                        //serie di replace;
                        MessageBody = MessageBody.Replace("%NOME%", Dlookup("Nome", "GSR_TaskList", "ID=" + TaskID).Trim());
                        MessageBody = MessageBody.Replace("%DESCR%", Dlookup("Descrizione", "GSR_TaskList", "ID=" + TaskID).Trim());
                        MessageBody = MessageBody.Replace("%LINK%", Dlookup("Link", "GSR_TaskList", "ID=" + TaskID).Trim());
                        MessageBody = MessageBody.Replace("%DATAINIZIO%", ((DateTime)CloseEvents.Rows[i].ItemArray[3]).ToShortDateString().Trim());
                        MessageBody = MessageBody.Replace("%ORAINIZIO%", ((DateTime)CloseEvents.Rows[i].ItemArray[3]).ToShortTimeString().Trim());

                        string subject = settings.IniReadValue("SMTP", "Subject");

                        subject = subject.Replace("%NOME%", Dlookup("Nome", "GSR_TaskList", "ID=" + TaskID).Trim());
                        subject = subject.Replace("%DESCR%", Dlookup("Descrizione", "GSR_TaskList", "ID=" + TaskID).Trim());
                        subject = subject.Replace("%LINK%", Dlookup("Link", "GSR_TaskList", "ID=" + TaskID).Trim());
                        subject = subject.Replace("%DATAINIZIO%", ((DateTime)CloseEvents.Rows[i].ItemArray[3]).ToShortDateString().Trim());
                        subject = subject.Replace("%ORAINIZIO%", ((DateTime)CloseEvents.Rows[i].ItemArray[3]).ToShortTimeString().Trim());



                        try
                        {
                            SendMailMessage(
                                settings.IniReadValue("SMTP", "server"),
                                settings.IniReadValue("SMTP", "From"),
                                settings.IniReadValue("SMTP", "FromName"),
                                settings.IniReadValue("SMTP", "To"),
                                settings.IniReadValue("SMTP", "CC"),
                                subject,
                                MessageBody);

                            LocalLog.AddMessage(Log.LogMessageType.Info, "Invio avviso per " +
                                Dlookup("Nome", "GSR_TaskList", "ID=" + TaskID).Trim() + " del " +
                                ((DateTime)CloseEvents.Rows[i].ItemArray[3]).ToShortDateString().Trim() + " alle " +
                                ((DateTime)CloseEvents.Rows[i].ItemArray[3]).ToShortTimeString().Trim());

                            LocalLog.AddMessage(Log.LogMessageType.Info, "TaskID:" + TaskID + " - Alert! (" + TS.Minutes + ")");
                        }
                        catch (SmtpException e)
                        {
                            LocalLog.AddMessage(Log.LogMessageType.Error, e.Message);
                        }

                    }
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
            timer.Interval = 60000;
           
            MainCycle();
        }
    }
}
