namespace Telegraph.LogModule.LogTypes
{
    public class ErrorLogType : AbstractLogType
    {
        public ErrorLogType(ILogWriter logWriter) : base(logWriter)
        {
        }
    }
}