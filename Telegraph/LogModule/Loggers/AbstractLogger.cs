using System;
using System.Diagnostics;
using System.Reflection;

namespace Telegraph.LogModule.Loggers
{
    public abstract class AbstractLogger : ILogger
    {
        protected string GetCallerMethodName()
        {

            MethodBase callerMethodBase = new StackTrace().GetFrame(3).GetMethod();
            string callerMethod = callerMethodBase.Name;
            string callerClass = callerMethodBase.ReflectedType?.FullName;
            return callerClass + "." + callerMethod;
        }

        public abstract void Error(object message);
        public abstract void Debug(object message);
        public abstract void Info(object message);

        public abstract void AddActionBefore(LogType logType, Action handler);
        public abstract void RemoveActionBefore(LogType logType, Action handler);

        public abstract void AddActionAfter(LogType logType, Action handler);
        public abstract void RemoveActionAfter(LogType logType, Action handler);

        public abstract void ClearAllActions(LogType logType);
    }
}