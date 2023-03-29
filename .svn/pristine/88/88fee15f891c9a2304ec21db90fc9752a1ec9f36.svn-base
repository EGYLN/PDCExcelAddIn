using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using log4net;
using log4net.Core;
using System.Runtime.InteropServices;


namespace PDCOpenLibrary
{
    /// <summary>
    /// Client side logger
    /// </summary>
    [ComVisible(false)]
    public class PDCLogger
    {
        /// <summary>
        /// Should be used from classes in PDCLib as context
        /// </summary>
        public const string LOG_NAME_LIB = "PDC_LIB";

        /// <summary>
        /// Should be used from classes in PDCExcelAddIn as context
        /// </summary>
        public const string LOG_NAME_EXCEL = "PDC_Excel";

        /// <summary>
        /// Should be used from classes in PDCExcelCOMAddIn as context
        /// </summary>
        public const string LOG_NAME_COM = "PDC_COM";

        /// <summary>
        /// Should be used for performance related log messages
        /// </summary>
        public const string LOG_NAME_PERFORMANCE = "PDC_PERF";

        private static ILog log = null;
        private static PDCLogger theLogger = new PDCLogger();

        private Dictionary<string, Stopwatch> startTimes = new Dictionary<string, Stopwatch>();
        /// <summary>
        /// Returns the logger singleton
        /// </summary>
        public static PDCLogger TheLogger
        {
            get
            {
                return theLogger;
            }
        }

        private PDCLogger()
        {
            try
            {
                log = LogManager.GetLogger(typeof(PDCLogger));
            }
#pragma warning disable 0168
            catch (Exception e)
            {
                //System.Windows.Forms.MessageBox.Show("PDC Logger Init Failed: " + e.ToString());
            }
#pragma warning restore 0168
        }
        /// <summary>
        /// Returns true if logging on debug level is enabled, otherwise false.
        /// Should be used to avoid time consuming debug print outs
        /// </summary>
        public bool IsDebugEnabled {
            get {
                return true;//  log != null && log.IsDebugEnabled;
            }
        }
        /// <summary>
        /// Logs a message with the specified severeness
        /// </summary>
        /// <param name="aContext">Usually one of the LOG_NAME constants</param>
        /// <param name="aMessage"></param>
        public void LogMessage(string aContext, string aMessage)
        {
            LogMessage(aContext, aMessage, Level.Info);
        }

        /// <summary>
        /// Logs a warning
        /// </summary>
        /// <param name="aContext">Usually one of the LOG_NAME constants</param>
        /// <param name="aMessage"></param>
        public void LogWarning(string aContext, string aMessage)
        {
            LogMessage(aContext, aMessage, Level.Warn);
        }

        /// <summary>
        /// Logs a trace
        /// </summary>
        /// <param name="aContext">Usually one of the LOG_NAME constants</param>
        /// <param name="aMessage"></param>
        public void LogTrace(string aContext, string aMessage)
        {
            LogMessage(aContext, aMessage, Level.Trace);
        }

        /// <summary>
        /// Logs a severe problem
        /// </summary>
        /// <param name="aContext">Usually one of the LOG_NAME constants</param>
        /// <param name="aMessage"></param>
        public void LogSevere(string aContext, string aMessage)
        {
            LogMessage(aContext, aMessage, Level.Severe);
        }

        /// <summary>
        /// Logs an error message
        /// </summary>
        /// <param name="aContext">Usually one of the LOG_NAME constants</param>
        /// <param name="aMessage"></param>
        public void LogError(string aContext, string aMessage)
        {
            LogMessage(aContext, aMessage, Level.Error);
        }

        /// <summary>
        /// Logs an info
        /// </summary>
        /// <param name="aContext">Usually one of the LOG_NAME constants</param>
        /// <param name="aMessage"></param>
        /// <param name="aLogLevel"></param>
        public void LogMessage(string aContext, string aMessage, Level aLogLevel)
        {
            try
            {
                log.Logger.Log(typeof(PDCLogger), aLogLevel, aMessage, null);
            }
#pragma warning disable 0168
            catch (Exception e) { }
#pragma warning restore 0168
        }

        /// <summary>
        /// Logs a debug message
        /// </summary>
        /// <param name="aContext">Usually one of the LOG_NAME constants</param>
        /// <param name="aMessage"></param>
        public void LogDebugMessage(string aContext, string aMessage)
        {
            LogMessage(aContext, aMessage, Level.Debug);
        }
        /// <summary>
        /// Logs an exception with severness Error
        /// </summary>
        /// <param name="aContext">Usually one of the LOG_NAME constants</param>
        /// <param name="aMessage"></param>
        /// <param name="anException"></param>
        public void LogException(string aContext, string aMessage, Exception anException)
        {
            try
            {
                log.Error(aMessage, anException);
            }
#pragma warning disable 0168
            catch (Exception e) { }
#pragma warning restore 0168
        }

        public void LogStoptime(string aContext, string aMessage)
        {
            try
            {
                string tmpElapsed = "";
                if (startTimes.ContainsKey(aContext))
                {
                    Stopwatch tmpWatch = startTimes[aContext];
                    tmpWatch.Stop();
                    tmpElapsed = tmpWatch.Elapsed.ToString();
                }
                LogDebugMessage(LOG_NAME_PERFORMANCE, aMessage + ":Stop(" + tmpElapsed + ")");
            }
            catch (Exception e)
            {
            }
        }

        public void LogStarttime(string aContext, string aMessage)
        {
            try
            {
                Stopwatch tmpWatch = new Stopwatch();
                if (startTimes.ContainsKey(aContext))
                {
                    startTimes[aContext] = tmpWatch;
                }
                else
                {
                    startTimes.Add(aContext, tmpWatch);
                }
                LogDebugMessage(LOG_NAME_PERFORMANCE, aMessage + ":Start");
            }
            catch (Exception e)
            {
            }
        }
    }
}
