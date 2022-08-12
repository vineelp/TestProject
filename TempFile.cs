public static class Logger
    {
        private static readonly ILogger log;
        private static readonly ILogger logDB;
        static Logger()
        {
            log = new LoggerConfiguration().Enrich.WithProperty("ApplicationName", ConfigurationManager.AppSettings["ApplicationName"])
               .Enrich.WithProperty("ApplicationVersion", Assembly.GetExecutingAssembly().GetName().Version)
               //.Enrich.WithProperty("UserName", HttpContext.Current.User.Identity.Name)
               .Enrich.FromLogContext()
               .Enrich.WithProperty("EnvironmentName", ConfigurationManager.AppSettings["Environment"])
               .WriteTo.File(new JsonFormatter(), ConfigurationManager.AppSettings["LogFilePath"],
                rollingInterval: RollingInterval.Month,
                fileSizeLimitBytes: int.Parse(ConfigurationManager.AppSettings["LogFileSize"].ToString()),
                rollOnFileSizeLimit: true, shared: true)
                .CreateLogger();
            logDB = new LoggerConfiguration()
               .WriteTo.File(new JsonFormatter(), ConfigurationManager.AppSettings["LogFilePath_DB"],
                rollingInterval: RollingInterval.Month,
                fileSizeLimitBytes: int.Parse(ConfigurationManager.AppSettings["LogFileSize_DB"].ToString()),
                rollOnFileSizeLimit: true, shared: true)
                .CreateLogger();

        }
        public static void WriteInformation(string message, Exception ex = null, LogEventLevel? logEventLevel = null)
        {
            logEventLevel = logEventLevel != null ? logEventLevel : LogEventLevel.Information;
            log.ForContext("UserName", HttpContext.Current.User.Identity.Name).Write(logEventLevel.Value, ex, message);
        }
        public static void DBInfo(string message)
        {
            logDB.ForContext("UserName", HttpContext.Current.User.Identity.Name).Write(LogEventLevel.Information, message);
        }
        public static void Info(string message)
        {
            log.ForContext("UserName", HttpContext.Current.User.Identity.Name).Write(LogEventLevel.Information, message);
        }
        public static void Info(string message, params object[] propertyValues)
        {
            log.ForContext("UserName", HttpContext.Current.User.Identity.Name).Write(LogEventLevel.Information, message, propertyValues);
        }
        public static void Error(string message)
        {
            log.ForContext("UserName", HttpContext.Current.User.Identity.Name).Write(LogEventLevel.Error, message);
        }
        public static void Error(string message, Exception ex)
        {
            log.ForContext("UserName", HttpContext.Current.User.Identity.Name).Write(LogEventLevel.Error, ex, message);
        }
    }
