using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using log4net;
using log4net.Appender;
using log4net.Config;
using log4net.Core;
using log4net.Repository;

namespace BBS.ST.BHC.BSP.PDC.Lib.Util
{
  /// <summary>
  /// Client side logger
  /// </summary>
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

    private static ILog log;
    private static PDCLogger theLogger = new PDCLogger();

    private Dictionary<string, Stopwatch> startTimes = new Dictionary<string, Stopwatch>();

    #region constructor

    private PDCLogger()
    {
      try
      {
        init();
        log = LogManager.GetLogger("PDC");
      }
#pragma warning disable 0168
      catch (Exception e)
      {
        //System.Windows.Forms.MessageBox.Show("PDC Logger Init Failed: " + e.ToString());
      }
#pragma warning restore 0168
    }

    private void init()
    {
      Uri tmpUri = new Uri(GetType().Assembly.CodeBase);
      string tmpCodeBase = Path.GetDirectoryName(tmpUri.LocalPath);
      string tmpFileName = Path.Combine(tmpCodeBase, "PDC.log4net");
      FileInfo fi = new FileInfo(tmpFileName);
      XmlConfigurator.Configure(fi);
      tmpFileName = Path.Combine(tmpCodeBase, "pdc-log.txt");
      ChangeLogFile("RollingFileAppender", tmpFileName);
    }

    #endregion

    #region methods

#region switch directory

    private void ChangeLogFile(string appenderName, string fileName)
    {

      ILoggerRepository loggerRepository;

      loggerRepository = LogManager.GetRepository();

      foreach (IAppender appender in loggerRepository.GetAppenders())
      {
        if (appender.Name.CompareTo(appenderName) == 0 && appender is FileAppender)
        {

          FileAppender fileAppender = (FileAppender)appender;
          string oldFileName = fileAppender.File;
          if (!oldFileName.Equals(fileName))
          {
            fileAppender.File = fileName;
            fileAppender.ActivateOptions();
            try
            {
              File.Delete(oldFileName);
            }
            catch (Exception) { };
          };

          return;

        }

      }
    }

#endregion


    #region LogDebugMessage
    /// <summary>
    /// Logs a debug message
    /// </summary>
    /// <param name="aContext">Usually one of the LOG_NAME constants</param>
    /// <param name="aMessage"></param>
    public void LogDebugMessage(string aContext, string aMessage)
    {
      LogMessage(aContext, aMessage, Level.Debug);
    }
    #endregion

    #region LogError
    /// <summary>
    /// Logs an error message
    /// </summary>
    /// <param name="aContext">Usually one of the LOG_NAME constants</param>
    /// <param name="aMessage"></param>
    public void LogError(string aContext, string aMessage)
    {
      LogMessage(aContext, aMessage, Level.Error);
    }
    #endregion

    #region LogException
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
      catch (Exception e)
      {
      }
#pragma warning restore 0168
    }
    #endregion

    #region LogMessage
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
    /// Logs an info
    /// </summary>
    /// <param name="aContext">Usually one of the LOG_NAME constants</param>
    /// <param name="aMessage"></param>
    /// <param name="aLogLevel"></param>
    public void LogMessage(string aContext, string aMessage, Level aLogLevel)
    {
      try
      {
        if (!log.Logger.Repository.Configured)
        {
          init();
        }
        log.Logger.Log(typeof(PDCLogger), aLogLevel, aMessage, null);
      }
#pragma warning disable 0168
      catch (Exception e)
      {
      }
#pragma warning restore 0168
    }
    #endregion

    #region LogSevere
    /// <summary>
    /// Logs a severe problem
    /// </summary>
    /// <param name="aContext">Usually one of the LOG_NAME constants</param>
    /// <param name="aMessage"></param>
    public void LogSevere(string aContext, string aMessage)
    {
      LogMessage(aContext, aMessage, Level.Severe);
    }
    #endregion

    #region LogStarttime
    /// <summary>
    /// Start a time measurement for the specified context key.
    /// Prints the given message to the log
    /// </summary>
    /// <param name="aContext">Context key for time measurement</param>
    /// <param name="aMessage">Message for the log</param>
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
        tmpWatch.Start();
      }
#pragma warning disable 0168
      catch (Exception e)
      {
      }
#pragma warning restore 0168
    }
    #endregion

    #region LogStoptime
    /// <summary>
    /// Log the stop time for an operation using the message after a call
    /// to LogStartTime with the same context.
    /// </summary>
    /// <param name="aContext">Identifier the logging context for the time measurement to
    /// enable nested time measurements
    /// </param>
    /// <param name="aMessage">Message printed in the log</param>
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
#pragma warning disable 0168
      catch (Exception e)
      {
      }
#pragma warning restore 0168
    }
    #endregion

    #region LogTrace
    /// <summary>
    /// Logs a trace
    /// </summary>
    /// <param name="aContext">Usually one of the LOG_NAME constants</param>
    /// <param name="aMessage"></param>
    public void LogTrace(string aContext, string aMessage)
    {
      LogMessage(aContext, aMessage, Level.Trace);
    }
    #endregion

    #region LogWarning
    /// <summary>
    /// Logs a warning
    /// </summary>
    /// <param name="aContext">Usually one of the LOG_NAME constants</param>
    /// <param name="aMessage"></param>
    public void LogWarning(string aContext, string aMessage)
    {
      LogMessage(aContext, aMessage, Level.Warn);
    }
    #endregion

    #endregion

    #region properties

    #region IsDebugEnabled
    /// <summary>
    /// Returns true if logging on debug level is enabled, otherwise false.
    /// Should be used to avoid time consuming debug print outs
    /// </summary>
    public bool IsDebugEnabled {
        get
        {
          return log != null && log.IsDebugEnabled;
        }
    }
    #endregion

    #region TheLogger
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
    #endregion

    #endregion
  }
}
