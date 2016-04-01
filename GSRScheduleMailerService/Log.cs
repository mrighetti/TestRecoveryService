using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;

namespace LogUtility
{
    public class Log
    {
        #region Constructors
        public Log()
        {
            LogCreateFile();
        }
        public Log(string filename)
        {
            _filename = filename;
            LogCreateFile();
        }
        #endregion
        #region Private
        string _filename = ""; //Server.MapPath("App_Data") + "\\ApplicationLog.xml";
        private void LogCreateFile()
        {
            if (!File.Exists(_filename))
            {
                XmlTextWriter writer = new XmlTextWriter(_filename, Encoding.UTF8) { Formatting = Formatting.Indented };
                writer.WriteStartDocument(true);
                writer.WriteStartElement("Log");
                writer.WriteStartElement("FileInfo");
                writer.WriteEndElement();
                writer.WriteStartElement("Messages");
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
                writer.Close();
                AddLogInfo("Created");
            }
        }
        #endregion
        #region Properties
        public enum LogMessageType { Info, Warning, Error, Debug }
        public string LogFileName
        {
            get
            {
                return _filename;
            }
        }
        public int MessageCount
        {
            get
            {
                XmlDocument XD = new XmlDocument();
                XD.Load(_filename);
                XmlNode NL = XD.SelectSingleNode("/Log/Messages");
                return NL.ChildNodes.Count;
            }
        }
        #endregion
        #region Methods
        public void AddLogInfo(string Message)
        {

            XmlDocument XD = new XmlDocument();
            XD.Load(_filename);
            XmlNode NL = XD.SelectSingleNode("/Log/FileInfo");
            StringWriter sw = new StringWriter();
            XmlTextWriter XmlW = new XmlTextWriter(sw);
            XmlW.WriteStartElement("Info");
            XmlW.WriteAttributeString("TimeStamp", DateTime.Now.ToString("s"));
            XmlW.WriteString(Message);
            XmlW.WriteEndElement();
            NL.InnerXml += sw.ToString();
            XD.Save(_filename);
        }
        public void AddMessage(LogMessageType MessageType, string Message)
        {
            XmlDocument XD = new XmlDocument();
            XD.Load(_filename);
            XmlNode NL = XD.SelectSingleNode("/Log/Messages");
            StringWriter sw = new StringWriter();
            XmlTextWriter XmlW = new XmlTextWriter(sw);
            XmlW.WriteStartElement("Message");
            XmlW.WriteAttributeString("TimeStamp", DateTime.Now.ToString("s"));
            XmlW.WriteAttributeString("Type", MessageType.ToString());
            XmlW.WriteString(Message);
            XmlW.WriteEndElement();
            NL.InnerXml += sw.ToString();
            XD.Save(_filename);
        }
        public void AddMessage(LogMessageType MessageType, string Message, string User)
        {
            AddMessage(MessageType, string.Format("{0}: {1}", User, Message));
        }
        public void AddMessage(LogMessageType MessageType, string Message, string User, string NetAddress)
        {
            AddMessage(MessageType, string.Format("{0} ({1}): {2}", User.Trim(), NetAddress.Trim(), Message));
        }
        public void Backup(string NewFileName)
        {
            XmlDocument XD = new XmlDocument();
            XD.Load(_filename);
            XD.Save(NewFileName);
            AddLogInfo("Backup - " + NewFileName);
        }
        public void Clear(bool doBackup)
        {
            if (doBackup == true)
            {
                string newfilename = DateTime.Now.ToString("s").Replace("-", "").Replace(":", "");
                newfilename = string.Format("{0}.{1}.xml", this._filename, newfilename);
                this.Backup(newfilename);
            }
            XmlDocument XD = new XmlDocument();
            XD.Load(_filename);
            XmlNode NL = XD.SelectSingleNode("/Log/Messages");
            int nodecount = NL.ChildNodes.Count;
            for (int i = 0; i < nodecount; i++)
                NL.RemoveChild(NL.ChildNodes[0]);
            XD.Save(_filename);
            AddLogInfo("Cleared");
        }
        public void Prune(int LeaveLastMessages, bool doBackup)
        {
            if (doBackup == true)
            {
                string newfilename = DateTime.Now.ToString("s").Replace("-", "").Replace(":", "");
                newfilename = string.Format("{0}.{1}.xml", this._filename, newfilename);
                this.Backup(newfilename);
            }
            XmlDocument XD = new XmlDocument();
            XD.Load(_filename);
            XmlNode NL = XD.SelectSingleNode("/Log/Messages");
            int nodecount = NL.ChildNodes.Count - LeaveLastMessages;
            for (int i = 0; i < nodecount; i++)
                NL.RemoveChild(NL.ChildNodes[0]);
            XD.Save(_filename);
            AddLogInfo(string.Format("Pruned - Last {0} messages", LeaveLastMessages));

        }

        #endregion
    }
    
    
}
