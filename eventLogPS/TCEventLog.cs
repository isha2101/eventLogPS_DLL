using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace eventLogPS
{
    class TCEventLog
    {
        eventLog objLog = new eventLog();
        public void TestFun()
        {
            //getByCmd();
            //GetLogEntries();
            //getLogEntries();
            getLogById();
            //getLogByDateTime();
            //getLogBySource();
            //createLog();
            //clrLog();
            //getLogByType();
            //getGPOList();
        }
        //public DataTable getGPOList()
        //{
        //    DataTable dt = new DataTable();
        //    dt = objLog.GetAppliedGroupPolicies();
        //    return dt;
        //}
        public DataTable GetLogEntries()
        {
            DataTable dt = new DataTable();
            DateTime startDateTime = new DateTime(2024, 4, 10, 21, 00, 00);
            DateTime endDateTime = new DateTime(2024, 4, 14, 21, 00, 00);
            dt = objLog.GetEventLog("Setup", "WIN-SI4O25LI52M", "lastexecute.txt",4,startDateTime, endDateTime);
            return dt;
        }
        public DataTable getLogEntries()
        {
            DataTable dt = new DataTable();
            dt = objLog.getEventLogEntries("Security", "WIN-SI4O25LI52M", "lastExecute.txt"); //Application
            return dt;
        }
        public DataTable getLogById()
        {
            DataTable dt = new DataTable();
            dt = objLog.getEventLogEntriesID("Security", 4768, "192.168.100.115", "lastExecutAdmin.txt","tectonas\\Administrator","Tectona#123");  //2
            return dt;
        }
        public DataTable getLogByDateTime()
        {
            DataTable dt = new DataTable();
            DateTime startDateTime = new DateTime(2024, 3, 14, 11,59,03);//14-03-2024 11:59:03
            DateTime endDateTime = new DateTime(2024,4,12,10,16,40);//12-04-2024 10:16:40
            dt = objLog.GetEventLogEntriesByDateTime("Setup",startDateTime,endDateTime, "WIN-SI4O25LI52M", "lastExecute.txt");//startDateTime  endDateTime
            return dt;
        }
        public DataTable getLogBySource()
        {
            DataTable dt = new DataTable();
            string SourceName = "Service Control Manager";
            dt = objLog.GetEventLogEntriesBySource("System", SourceName, "WIN-SI4O25LI52M", "lastExecute.txt");
            return dt;
        }

        public bool createLog()
        {
            bool bl = objLog.writeLog("MyCustomSource", "Security");
            return bl;
        }
        public DataTable getByCmd()
        {
            DataTable dt = new DataTable();
            dt = objLog.ExecutePSCmd("System", "lastExecute.txt");
            return dt;
        }

        public bool clrLog()
        {
            bool bl = objLog.clearLog("Security", "WIN-SI4O25LI52M");
            return bl;
        }
        public DataTable getLogByType()
        {
            DataTable dt = new DataTable();
            string eventType = "Information";
            dt = objLog.GetEventLogByEventType("Security", eventType, "WIN-SI4O25LI52M", "lastExecute.txt");
            return dt;
        }
    }
}
