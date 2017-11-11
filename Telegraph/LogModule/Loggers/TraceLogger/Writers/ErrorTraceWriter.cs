using System;
using System.Diagnostics;

namespace Telegraph.LogModule.Loggers.TraceLogger.Writers
{
    public class ErrorTraceWriter : ILogWriter
    {
        public string LogType { get; } = "ERROR";

        public void Write(object message)
        {
            Trace.WriteLine(message, LogType);
        }
    }
}