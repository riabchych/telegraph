using System;

namespace Telegraph.LogModule.LogTypes
{
    public abstract class AbstractLogType
    {
        public Action BeginExecute { get; set; }
        public Action EndExecute { get; set; }

        protected ILogWriter LogWriter;

        protected AbstractLogType(ILogWriter logWriter)
        {
            LogWriter = logWriter;
        }

        public virtual void Execute(object message)
        {
            BeginExecute?.Invoke();
            LogWriter.Write(message);
            EndExecute?.Invoke();
        }
    }
}