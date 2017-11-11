using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Telegraph.LogModule.Loggers.TraceLogger.Writers;
using Telegraph.LogModule.LogTypes;

namespace Telegraph.LogModule.Loggers.TraceLogger
{
    public class TraceLogger : AbstractLogger
    {
        private readonly ErrorLogType errorLog;
        private readonly DebugLogType debugLog;
        private readonly InfoLogType infoLog;

        private readonly Dictionary<LogType, AbstractLogType> logTypeTable;

        public TraceLogger()
        {
            ConsoleTraceListener consoleTraceListener = new ConsoleTraceListener();
            Trace.Listeners.Add(consoleTraceListener);

            errorLog = new ErrorLogType(new ErrorTraceWriter());
            debugLog = new DebugLogType(new DebugTraceWriter());
            infoLog = new InfoLogType(new InfoTraceWriter());

            logTypeTable = new Dictionary<LogType, AbstractLogType>
            {
                {LogType.Error, errorLog},
                {LogType.Debug, debugLog},
                {LogType.Info, infoLog}
            };
        }

        public override void Error(object message)
        {
            Log(errorLog, message);
            LogMessageToFile(message, LogType.Error);
        }

        public override void Debug(object message)
        {
            Log(debugLog, message);
            LogMessageToFile(message, LogType.Debug);
        }

        public override void Info(object message)
        {
            Log(infoLog, message);
            LogMessageToFile(message, LogType.Info);
        }

        public override void AddActionBefore(LogType logType, Action handler)
        {
            foreach (var typePair in logTypeTable)
            {
                if ((logType & typePair.Key) == typePair.Key)
                {
                    typePair.Value.BeginExecute += handler;
                }
            }
        }

        public override void RemoveActionBefore(LogType logType, Action handler)
        {
            foreach (var typePair in logTypeTable)
            {
                if ((logType & typePair.Key) == typePair.Key)
                {
                    typePair.Value.BeginExecute -= handler;
                }
            }
        }

        public override void AddActionAfter(LogType logType, Action handler)
        {
            foreach (var typePair in logTypeTable)
            {
                if ((logType & typePair.Key) == typePair.Key)
                {
                    typePair.Value.EndExecute += handler;
                }
            }
        }

        public override void RemoveActionAfter(LogType logType, Action handler)
        {
            foreach (var typePair in logTypeTable)
            {
                if ((logType & typePair.Key) == typePair.Key)
                {
                    typePair.Value.EndExecute -= handler;
                }
            }
        }

        public override void ClearAllActions(LogType logType)
        {
            foreach (var typePair in logTypeTable)
            {
                if ((logType & typePair.Key) == typePair.Key)
                {
                    typePair.Value.BeginExecute = null;
                    typePair.Value.EndExecute = null;
                }
            }
        }

        private void Log(AbstractLogType logType, object message)
        {
            logType.Execute($"{GetCallerMethodName()} => {message}");
        }

        private void LogMessageToFile(object msg, LogType type)
        {
            string filename = "info.log";
            switch (type)
            {
                case LogType.Error:
                    filename = "errors.log";
                    break;
                case LogType.Debug:
                    filename = "debug.log";
                    break;
                default:
                    filename = "info.log";
                    break;
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + filename;
            if (File.Exists(filepath))
            {
                if ((int)new FileInfo(filepath).Length > 330000)
                {
                    File.Delete(filepath);
                }
            }

            StreamWriter sw = System.IO.File.AppendText(filename);

            try
            {
                string logLine = System.String.Format(
                    "{0:G}: {1}.", System.DateTime.Now, msg);
                sw.WriteLine(logLine);
            }
            finally
            {
                sw.Close();
            }
        }
    }
}