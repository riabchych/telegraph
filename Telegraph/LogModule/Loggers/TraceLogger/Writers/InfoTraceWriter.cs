using System;
using System.Diagnostics;

namespace Telegraph.LogModule.Loggers.TraceLogger.Writers
{
    public class InfoTraceWriter : ILogWriter
    {
        public string LogType { get; } = "INFO";

        public void Write(object message)
        {
            Trace.WriteLine(message, LogType);
        }
    }
}