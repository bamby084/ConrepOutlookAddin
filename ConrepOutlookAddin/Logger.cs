using log4net;

namespace ConrepOutlookAddin
{
    public static class Logger
    {
        private static readonly ILog _logger = LogManager.GetLogger("ConrepLogger");

        public static void Error(System.Exception ex)
        {
            _logger.Error(ex.Message, ex);
        }

        public static void Info(string info)
        {
            _logger.Info(info);
        }
    }
}
