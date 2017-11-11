namespace Telegraph.LogModule
{
    public interface ILogWriter
    {
        string LogType { get; }
        void Write(object message);
    }
}