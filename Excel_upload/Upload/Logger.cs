using log4net;

namespace Excel_upload
{
    public static class Logger
    {
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public static void Write(MessageType messageType, string message)
        {
            switch (messageType)
            {
                case MessageType.Info:
                    Log.Info(message);
                    break;
                case MessageType.Warn:
                    Log.Warn(message);
                    break;
                case MessageType.Error:
                    Log.Error(message);
                    break;
                case MessageType.Fatal:
                    Log.Fatal(message);
                    break;
                default:
                    Log.Info(message);
                    break;
            }
        }
    }
}
