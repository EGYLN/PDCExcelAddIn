using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using BBS.ST.BHC.BSP.PDC.Lib.Exceptions;
using BBS.ST.BHC.BSP.PDC.Lib.Util;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    /// <summary>
    /// General Exception handler for the PDC Excel client.
    /// Presents a failure message to the user.
    /// </summary>
    [ComVisible(false)]
    public class ExceptionHandler
    {
        private static object LOCK = new object();
        private static ExceptionHandler SINGLETON;

        public static ExceptionHandler TheExceptionHandler {
            get
            {
                lock (LOCK)
                {
                    if (SINGLETON == null)
                    {
                        SINGLETON = new ExceptionHandler();
                    }
                    return SINGLETON;
                }
            }
       }        
        public void handleException(Exception e, Form aParent)
        {
            PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, e.Message,e);
            PDCLogger.TheLogger.LogError(PDCLogger.LOG_NAME_EXCEL, "Helplink to Exception is: " + e.HelpLink);
            Exception tmpInner = e.InnerException;
            while (tmpInner != null)
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "InnerException is: " + tmpInner.Message, tmpInner);
                PDCLogger.TheLogger.LogError(PDCLogger.LOG_NAME_EXCEL, "Helplink for InnerException is " + tmpInner.HelpLink);
                if (tmpInner is System.Net.Sockets.SocketException)
                {
                    System.Net.Sockets.SocketException tmpSE = (System.Net.Sockets.SocketException)tmpInner;
                    PDCLogger.TheLogger.LogError(PDCLogger.LOG_NAME_EXCEL, "SocketException with Errorcode " + tmpSE.ErrorCode + ",Native Error code " + tmpSE.NativeErrorCode + ",Socket error code" + tmpSE.SocketErrorCode);
                }
                tmpInner = tmpInner.InnerException;
            }
            
            string errorText = e.Message;



            string title = Properties.Resources.MSG_GeneralException_Title;

            if (e is COMException)
            {
                errorText += Environment.NewLine + Properties.Resources.MSG_COMException_Restart;
                title = Properties.Resources.MSG_COMException_Title;
            }
            if (e is TooManyParametersException)
            {
                title = Properties.Resources.MSG_TOO_MANY_SEARCH_PARAMS_TITLE;
                errorText = Properties.Resources.MSG_TOO_MANY_SEARCH_PARAMS_TEXT;
            }
            if (e is Exceptions.NoExperimentFoundForMeasurementException ||
              e is Exceptions.AmbitiousExperimentsException)
            {
              title = Properties.Resources.MSG_MeasurementException_Title;
            }
            if (e is System.Net.WebException)
            {
              if (((System.Net.WebException)e).Status == System.Net.WebExceptionStatus.Timeout)
              {
                errorText = Properties.Resources.MSG_ERROR_TIMEOUT;
              }
            }
 
            if (aParent != null && !aParent.IsDisposed)
            {
              MessageBox.Show(aParent, errorText, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
              MessageBox.Show(errorText, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
