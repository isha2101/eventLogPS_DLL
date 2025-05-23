using System;
using System.Data;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Collections.ObjectModel;
using System.Collections;
using System.Linq;
using System.Globalization;
using log4net.Repository.Hierarchy;
using System.IO;
using System.Diagnostics;
using System.Xml;
using System.Collections.Generic;
using System.Security;

namespace eventLogPS
{
    public class eventLog
    {
        /// <summary>
        /// Retrieves event log entries from a specified log and remote computer by Event ID,
        /// using given user credentials. It only fetches new logs since the last execution time,
        /// and updates the timestamp file upon successful retrieval.
        /// </summary>
        public DataTable getEventLogEntriesID(string logName, int eventId, string computerName, string fileName, string userName, string password)
        {
            DataTable logDataTable = CreateLogDataTable();
            // Initialize last execution time to avoid reprocessing old logs
            string lastExecutionTimeFile = fileName;
            DateTime lastExecutionTime = DateTime.MinValue;

            WriteLog("logEntries", "log", "eventLogPS", "Starting event log retrieval...", true);

            if (File.Exists(lastExecutionTimeFile))
            {
                string lastExecutionString = File.ReadAllText(lastExecutionTimeFile).Trim();
                if (!string.IsNullOrEmpty(lastExecutionString))
                {
                    lastExecutionTime = DateTime.Parse(lastExecutionString).ToUniversalTime();
                    WriteLog("logEntries", "log", "eventLogPS", $"Last execution time found: {lastExecutionTime}", true);
                }
            }
            else
            {
                lastExecutionTime = DateTime.UtcNow.AddDays(-30);
                WriteLog("logEntries", "log", "eventLogPS", $"Effective lastExecutionTime: {lastExecutionTime}", true);

                WriteLog("logEntries", "log", "eventLogPS", "No previous execution time found. Retrieving all logs.", true);
            }
            // Prepare credentials for remote connection
            string username = $"{userName}";
            string Password = $"{password}";
            SecureString securePass = new SecureString();
            foreach (char c in Password)
            {
                securePass.AppendChar(c);
            }
            PSCredential credential = new PSCredential(username, securePass);
            using (PowerShell ps = PowerShell.Create(RunspaceMode.NewRunspace))
            {
                if (computerName.Equals(Environment.MachineName, StringComparison.OrdinalIgnoreCase))
                {
                    // Local call with FilterHashtable
                    ps.AddCommand("Get-WinEvent");
                    ps.AddParameter("FilterHashtable", new Dictionary<string, object>
        {
            { "LogName", logName },
            { "ID", eventId },
            { "StartTime", lastExecutionTime }
        });
                }
                else
                {
                    // Remote call using Invoke-Command and ScriptBlock
                    ps.AddCommand("Invoke-Command");
                    ps.AddParameter("ComputerName", computerName);
                    ps.AddParameter("Credential", credential);
                    WriteLog("logEntries", "log", "eventLogPS", $"Using lastExecutionTime: {lastExecutionTime}", true);

                                ps.AddParameter("ScriptBlock", ScriptBlock.Create($@"
                        Get-WinEvent -FilterHashtable @{{ LogName='{logName}'; ID={eventId}; StartTime=([datetime]::Parse('{lastExecutionTime:yyyy-MM-dd HH:mm:ss}')) }} |
                        Select-Object Message, TimeCreated, Id, LogName, MachineName, ProviderName, Task, LevelDisplayName, OpcodeDisplayName, Keywords
                    "));

    //                ps.AddParameter("ScriptBlock", ScriptBlock.Create($@"
    //Get-WinEvent -FilterHashtable @{{ LogName='{logName}'; ID={eventId}; StartTime=([datetime]::Parse('{lastExecutionTime:yyyy-MM-dd HH:mm:ss}')) }}"));

                }


                List<DateTime> logTimestamps = ExecutePowerShellCommand(ps, logDataTable, fileName);

                // Update the last execution time file if any new logs were retrieved
                if (logTimestamps.Count > 0)
                {
                    DateTime latestLogTime = logTimestamps.Max();
                    File.WriteAllText(lastExecutionTimeFile, latestLogTime.ToString("yyyy-MM-dd HH:mm:ss"));
                    WriteLog("logEntries", "log", "eventLogPS", $"Updated last execution time: {latestLogTime}", true);
                }
                else
                {
                    WriteLog("logEntries", "log", "eventLogPS", "No new event logs found.", true);
                }
            }

            WriteLog("logEntries", "log", "eventLogPS", "Event log retrieval completed.", true);
            return logDataTable;
        }

        /// <summary>
        /// Creates and returns a DataTable structure to store event log data.
        /// </summary>
        public DataTable CreateLogDataTable()
        {
            DataTable logDataTable = new DataTable();
            // Add columns to the DataTable with appropriate data types
            logDataTable.Columns.Add("Details", typeof(string));
            logDataTable.Columns.Add("Time", typeof(DateTime));
            logDataTable.Columns.Add("EventID", typeof(int));
            logDataTable.Columns.Add("LogName", typeof(string));
            logDataTable.Columns.Add("Computer", typeof(string));
            logDataTable.Columns.Add("Source", typeof(string));
            logDataTable.Columns.Add("Taskcategory", typeof(string));
            logDataTable.Columns.Add("Level", typeof(string));
            logDataTable.Columns.Add("OpCode", typeof(string));
            logDataTable.Columns.Add("Keywords", typeof(string));
            
            return logDataTable;
        }

        /// <summary>
        /// Populates a DataTable with Windows Event Log entries from a PowerShell result collection.
        /// It filters out logs that were already processed based on a saved timestamp in a file.
        /// </summary>
        public void PopulateDataTable(Collection<PSObject> results, DataTable logDataTable, string fileName)
        {
            // Read last execution time again (to prevent duplicate logs)
            string lastExecutionTimeFile = fileName;
            DateTime lastExecutionTime = DateTime.MinValue;

            if (File.Exists(lastExecutionTimeFile))
            {
                string lastExecutionString = File.ReadAllText(lastExecutionTimeFile).Trim();
                if (!string.IsNullOrEmpty(lastExecutionString))
                {
                    lastExecutionTime = DateTime.Parse(lastExecutionString);
                }
            }
            // Iterate over each PowerShell result (Event Log entry)
            foreach (PSObject result in results)
            {
                DateTime logTime = (DateTime)result.Properties["TimeCreated"].Value;

                // Skip logs that were already processed
                if (logTime <= lastExecutionTime)
                    continue;

                DataRow dataRow = logDataTable.NewRow();
                dataRow["Details"] = result.Properties["Message"].Value;
                dataRow["Time"] = logTime;
                dataRow["EventID"] = result.Properties["Id"].Value;
                dataRow["LogName"] = result.Properties["LogName"].Value;
                dataRow["Computer"] = result.Properties["MachineName"].Value;
                dataRow["Source"] = result.Properties["ProviderName"].Value;
                dataRow["TaskCategory"] = result.Properties["Task"].Value;
                dataRow["Level"] = result.Properties["LevelDisplayName"]?.Value ?? "N/A";
                dataRow["OpCode"] = result.Properties["OpcodeDisplayName"]?.Value ?? "N/A";
                dataRow["Keywords"] = result.Properties["Keywords"]?.Value ?? "N/A";

                logDataTable.Rows.Add(dataRow);
            }
        }

        /// <summary>
        /// Retrieves Windows Event Logs from a specified computer and log name,
        /// filtered by the event type (e.g., Information, Warning, Error),
        /// and populates the results into a DataTable.
        /// </summary>
        public DataTable GetEventLogByEventType(string logName, string eventType, string computerName, string fileName)
        {
            DataTable logDataTable = CreateLogDataTable();
            // Initialize a new PowerShell instance to execute commands
            using (PowerShell ps = PowerShell.Create())
            {
                ps.AddCommand("Get-WinEvent");
                ps.AddParameter("LogName", logName);
                ps.AddParameter("ComputerName", computerName);
                // Add an XPath filter to only include logs matching the specified event level
                ps.AddParameter("FilterXPath", $"*[System[Level={GetEventLevel(eventType)}]]");
                // Execute the PowerShell command and populate the DataTable with results
                ExecutePowerShellCommand(ps, logDataTable,fileName);
            }
            return logDataTable; // Return the populated DataTable
        }
        public int GetEventLevel(string eventType)
        {
            // Map event type names to corresponding event levels
            switch (eventType.ToLower())
            {
                case "critical":
                    return 1; // Critical
                case "error":
                    return 2; // Error
                case "warning":
                    return 3; // Warning
                case "information":
                    return 4; // Information
                case "verbose":
                    return 5; // Verbose
                default:
                    throw new ArgumentException("Invalid event type specified.");
            }
        }

        /// <summary>
        /// Retrieves all Windows Event Log entries from the specified log on a remote or local computer,
        /// and populates them into a DataTable. It avoids reprocessing logs based on the last execution timestamp.
        /// </summary>
        public DataTable getEventLogEntries(string logName, string computerName, string fileName)
        {
            DataTable logDataTable = CreateLogDataTable();
            using (PowerShell ps = PowerShell.Create())
            {
                ps.AddCommand("Get-WinEvent");
                ps.AddParameter("LogName", logName);
                ps.AddParameter("ComputerName", computerName);
                // Execute the command and populate the DataTable with the filtered log entries
                ExecutePowerShellCommand(ps, logDataTable,fileName);
            }
            return logDataTable;
        }

        /// <summary>
        /// Retrieves event log entries from a specified log and remote computer by Event ID,
        /// using given user credentials. It only fetches new logs since the last execution time,
        /// and updates the timestamp file upon successful retrieval.
        /// </summary>
        public DataTable getEventLogEntriesByID(string logName, int eventId, string computerName, string fileName, string userName, string password)
        {
            DataTable logDataTable = CreateLogDataTable();
            // Initialize last execution time to avoid reprocessing old logs
            string lastExecutionTimeFile = fileName;
            DateTime lastExecutionTime = DateTime.MinValue;

            WriteLog("logEntries", "log", "eventLogPS", "Starting event log retrieval...", true);

            if (File.Exists(lastExecutionTimeFile))
            {
                string lastExecutionString = File.ReadAllText(lastExecutionTimeFile).Trim();
                if (!string.IsNullOrEmpty(lastExecutionString))
                {
                    lastExecutionTime = DateTime.Parse(lastExecutionString).ToUniversalTime();
                    WriteLog("logEntries", "log", "eventLogPS", $"Last execution time found: {lastExecutionTime}", true);
                }
            }
            else
            {
                WriteLog("logEntries", "log", "eventLogPS", "No previous execution time found. Retrieving all logs.", true);
            }
            // Prepare credentials for remote connection
            string username = $"{userName}";
            string Password = $"{password}";
            SecureString securePass = new SecureString();
            foreach (char c in Password)
            {
                securePass.AppendChar(c);
            }
            PSCredential credential = new PSCredential(username, securePass);
            using (PowerShell ps = PowerShell.Create(RunspaceMode.NewRunspace))
            {
                ps.AddCommand("Get-WinEvent");
                ps.AddParameter("LogName", logName);
                ps.AddParameter("ComputerName", computerName);
                ps.AddParameter("Credential", credential);
                long timeDiffMs = GetTimeDifferenceInMilliseconds(lastExecutionTime);
                WriteLog("logEntries", "log", "eventLogPS", $"Calculated time difference: {timeDiffMs} ms", true);

                // Filter: match by EventID and only logs created within the time difference
                string filter = $"*[System[(EventID={eventId}) and TimeCreated[timediff(@SystemTime) <= {timeDiffMs}]]]";
                ps.AddParameter("FilterXPath", filter);
                WriteLog("logEntries", "log", "eventLogPS", $"Executing PowerShell command with filter: {filter}", true);

                List<DateTime> logTimestamps = ExecutePowerShellCommand(ps, logDataTable, fileName);

                // Update the last execution time file if any new logs were retrieved
                if (logTimestamps.Count > 0)
                {
                    DateTime latestLogTime = logTimestamps.Max();
                    File.WriteAllText(lastExecutionTimeFile, latestLogTime.ToString("yyyy-MM-dd HH:mm:ss"));
                    WriteLog("logEntries", "log", "eventLogPS", $"Updated last execution time: {latestLogTime}", true);
                }
                else
                {
                    WriteLog("logEntries", "log", "eventLogPS", "No new event logs found.", true);
                }
            }

            WriteLog("logEntries", "log", "eventLogPS", "Event log retrieval completed.", true);
            return logDataTable;
        }

        /// <summary>
        /// Executes a PowerShell command to retrieve event log entries,
        /// populates the results into a provided DataTable, and collects timestamps of each log entry.
        /// It logs both execution status and any errors encountered during the process.
        /// </summary>
        public List<DateTime> ExecutePowerShellCommand(PowerShell ps, DataTable logDataTable, string fileName)
        {
            List<DateTime> logTimestamps = new List<DateTime>();
            WriteLog("logEntries", "log", "eventLogPS", "Executing PowerShell command...", true);
            // Execute the PowerShell command and retrieve the results

            Collection<PSObject> results = ps.Invoke();

            if (ps.HadErrors)
            {
                foreach (ErrorRecord error in ps.Streams.Error)
                {
                    // Log each PowerShell error message
                    string errorMessage = $"PowerShell error: {error.Exception.Message}";
                    WriteLog("logEntries", "log", "eventLogPS", errorMessage, true);
                }
            }
            else
            {
                WriteLog("logEntries", "log", "eventLogPS", $"PowerShell command executed successfully. Retrieved {results.Count} events.", true);
                // Populate the DataTable with event log results
                PopulateDataTable(results, logDataTable, fileName);

                // Extract timestamps from the results
                foreach (PSObject obj in results)
                {
                    if (obj.Properties["TimeCreated"]?.Value != null)
                    {
                        DateTime eventTime = DateTime.Parse(obj.Properties["TimeCreated"].Value.ToString());
                        logTimestamps.Add(eventTime);
                        WriteLog("logEntries", "log", "eventLogPS", $"Processed event timestamp: {eventTime}", true);
                    }
                }
            }

            WriteLog("logEntries", "log", "eventLogPS", $"Total events processed: {logTimestamps.Count}", true);
            return logTimestamps;
        }

        /// <summary>
        /// Calculates the time difference in milliseconds between the current UTC time and the last execution time.
        /// If this is the first execution (no previous timestamp), it returns int.MaxValue to fetch all logs.
        /// Ensures a minimum threshold of 100 milliseconds to avoid issues with very small time spans.
        /// </summary>
        private int GetTimeDifferenceInMilliseconds(DateTime lastExecutionTime)
        {
            // If this is the first run (no previous execution time), fetch all logs
            if (lastExecutionTime == DateTime.MinValue)
                return int.MaxValue;

            // Calculate the time difference between now and the last execution
            TimeSpan diff = DateTime.UtcNow - lastExecutionTime;
            return Math.Max((int)diff.TotalMilliseconds, 100);
        }

        /// <summary>
        /// Retrieves event log entries from a specified event log on a remote computer between a given date-time range.
        /// Uses a hardcoded administrator credential to authenticate and applies a time-based XPath filter for log selection.
        /// </summary>
        public DataTable GetEventLogEntriesByDateTime(string logName,DateTime startTime, DateTime endTime, string computerName,string fileName)
        {
            string username = "DEMO\\Administrator";
            // Convert plain password to SecureString for use in PowerShell
            string Password = "RESEt21";
            SecureString securePass=new SecureString();
            foreach (char c in Password)
            {
                securePass.AppendChar(c);
            }
            PSCredential credential = new PSCredential(username, securePass);
            DataTable logDataTable = CreateLogDataTable();
            using (PowerShell ps = PowerShell.Create())
            {
                ps.AddCommand("Get-WinEvent");
                ps.AddParameter("LogName", logName);
                ps.AddParameter("ComputerName", computerName);
                ps.AddParameter("Credential", credential);
                // Build XPath filter to select events between startTime and endTime
                string filter = "*[System[TimeCreated[@SystemTime >= '" + startTime.ToString("yyyy-MM-ddTHH:mm:ss") + "' and @SystemTime <= '" + endTime.ToString("yyyy-MM-ddTHH:mm:ss") + "']]]";
                ps.AddParameter("FilterXPath", filter);

                // Execute the PowerShell command and populate the DataTable
                ExecutePowerShellCommand(ps, logDataTable,fileName);
            }
            return logDataTable;
        }

        /// <summary>
        /// Retrieves event log entries from a specific Windows Event Log where the source matches the given provider name (source name).
        /// Filters the logs on a remote machine using XPath based on the event source (provider).
        /// </summary>
        public DataTable GetEventLogEntriesBySource(string logName, string sourceName, string computerName, string fileName)
        {
            DataTable logDataTable = CreateLogDataTable();
            using (PowerShell ps = PowerShell.Create())
            {
                ps.AddCommand("Get-WinEvent");
                ps.AddParameter("LogName", logName);
                ps.AddParameter("ComputerName", computerName);
                // Filter events by the specified source/provider using XPath
                ps.AddParameter("FilterXPath", "*[System[Provider[@Name='" + sourceName + "']]]");
                // Execute PowerShell command and populate the DataTable
                ExecutePowerShellCommand(ps, logDataTable,fileName);
            }
            return logDataTable;
        }

        /// <summary>
        /// Retrieves Windows Event Log entries from a specified log on a remote computer, with optional filtering
        /// by event ID and/or a date-time range. Constructs a dynamic XPath filter based on the provided parameters.
        /// </summary>
        public DataTable GetEventLog(string logName, string computerName,string fileName,int? eventId = null, DateTime? startTime = null, DateTime? endTime = null)
        {
            DataTable logDataTable = CreateLogDataTable();
            using (PowerShell ps = PowerShell.Create())
            {
                ps.AddCommand("Get-WinEvent");
                ps.AddParameter("LogName", logName);
                ps.AddParameter("ComputerName", computerName);
                // Build the XPath filter dynamically based on which optional parameters are provided

                // Case 1: Filter by event ID AND a date-time range
                if (eventId.HasValue && startTime.HasValue && endTime.HasValue)
                {
                    string filter = "*[System/EventID=" + eventId + " and System[TimeCreated[@SystemTime >= '" + startTime.Value.ToString("yyyy-MM-ddTHH:mm:ss") + "' and @SystemTime <= '" + endTime.Value.ToString("yyyy-MM-ddTHH:mm:ss") + "']]]";
                    ps.AddParameter("FilterXPath", filter);
                }
                // Case 2: Filter by only event ID
                else if (eventId.HasValue)
                {
                    string filter = "*[System/EventID=" + eventId + "]";
                    ps.AddParameter("FilterXPath", filter);
                }
                // Case 3: Filter by only date-time range
                else if (startTime.HasValue && endTime.HasValue)
                {
                    string filter = "*[System[TimeCreated[@SystemTime >= '" + startTime.Value.ToString("yyyy-MM-ddTHH:mm:ss") + "' and @SystemTime <= '" + endTime.Value.ToString("yyyy-MM-ddTHH:mm:ss") + "']]]";
                    ps.AddParameter("FilterXPath", filter);
                }
                // Execute the command and populate the DataTable
                ExecutePowerShellCommand(ps, logDataTable,fileName);
            }
            return logDataTable;
        }

        /// <summary>
        /// Writes a test informational entry to the Windows Event Log under the specified source and log name.
        /// If the event source does not exist, it creates it first.
        /// </summary>
        public bool writeLog(string sourceName,string logName)
        {
            bool bl = false;
            try
            {
                // Check if the event source already exists on the machine
                if (!EventLog.SourceExists(sourceName))
                { 
                    // Create a new event source associated with the specified log name
                    EventLog.CreateEventSource(sourceName, logName);
                    bl = true;
                }
                // Write a test informational log entry to the event log using the specified source
                EventLog.WriteEntry(sourceName, "this is test log entry", EventLogEntryType.Information);
                
            }
            catch(Exception ex)
            { 
            }
            return bl;
        }

        /// <summary>
        /// Clears the specified Windows Event Log on a remote or local computer using PowerShell's Clear-EventLog cmdlet.
        /// </summary>
        public bool clearLog(string logName, string computerName)
        {
            bool bl = false; 
            try
            {
                // Add the Clear-EventLog PowerShell command with parameters for log name and target computer
                using (PowerShell ps = PowerShell.Create())
                {
                    ps.AddCommand("Clear-EventLog");
                    ps.AddParameter("LogName", logName);
                    ps.AddParameter("ComputerName", computerName);
                    ps.Invoke();
                    // Check if any errors occurred during command execution
                    if (ps.HadErrors)
                    {
                        foreach (ErrorRecord error in ps.Streams.Error)
                        {
                            string errorMessage = $"PowerShell error: {error.Exception.Message}";
                            WriteLog("logEntries", "log", "eventLogPS", errorMessage, true);
                        }
                    }
                    else return bl = true;
                }
                
            }
            catch( Exception ex )
            {
                bl = false;
            }
            return bl;
        }

        /// <summary>
        /// Executes the PowerShell command to retrieve event logs from the specified log name,
        /// and returns the results as a DataTable.
        /// </summary>
        public DataTable ExecutePSCmd(string logName,string fileName)
        {
            DataTable logDataTable = CreateLogDataTable();
            using(PowerShell  ps = PowerShell.Create())
            {
                // Create and open a new Runspace for the PowerShell pipeline execution
                using (Runspace rs = RunspaceFactory.CreateRunspace())
                {
                    rs.Open();
                    ps.Runspace = rs;
                    // Build the PowerShell command string to get event logs for the specified log name
                    string command = $"Get-EventLog -LogName {logName}";
                    ps.AddScript(command);
                    // Execute the PowerShell command and populate the DataTable with results
                    ExecutePowerShellCommand(ps, logDataTable,fileName);
                    rs.Close();
                }
            }
            return logDataTable;
        }

        // Path to the log directory
        static string LogPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\AssertYIT\\";

        // Writes a log entry to a file
            public static void WriteLog(string FileName, string Extension, string ProductName, string Message, bool HourWise)
            {
                try
                {
                    if (HourWise)
                        FileName = FileName + "_" + System.DateTime.Now.ToString("yyyy-MM-dd HH") + "." + Extension;
                    else
                        FileName = FileName + "." + Extension;
                    FileName = LogPath + ProductName + "\\" + FileName;
                    if (!Directory.Exists(LogPath + ProductName))
                        Directory.CreateDirectory(LogPath + ProductName);
                    FileStream fs;
                    fs = new FileStream(FileName, FileMode.Append);
                    StreamWriter s = new StreamWriter(fs);
                    s.WriteLine(System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " : " + Message);
                    s.Close();
                    fs.Close();
                }
                catch (Exception)
                {
                }
            }

        
    }
}
