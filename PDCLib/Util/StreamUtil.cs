using System;
using System.IO;
using System.Net;

namespace BBS.ST.BHC.BSP.PDC.Lib.Util
{
  /// <summary>
  /// Utility methods for local files and URLs
  /// </summary>
  public class StreamUtil
  {
    #region methods

    #region GetContents
    /// <summary>
    /// Returns the contents for the specified path, which may 
    /// be a url to a http resource
    /// </summary>
    /// <param name="aPath"></param>
    /// <returns></returns>
    public static byte[] GetContents(string aPath)
    {
      Stream tmpStream = null;
      WebResponse tmpResponse = null;
      try
      {
        if (IsUrl(aPath))
        {
          WebRequest tmpRequest = WebRequest.Create(aPath);
          tmpRequest.Proxy = new WebProxy();
          tmpResponse = tmpRequest.GetResponse();
          tmpStream = tmpResponse.GetResponseStream();
        }
        else
        {
          tmpStream = new FileStream(aPath, FileMode.Open);
        }
        byte[] tmpContents = null;
        using (MemoryStream tmpMemoryStream = new MemoryStream())
        {
          byte[] tmpChunks = new byte[8186];
          int tmpLength = 0;
          while ((tmpLength = tmpStream.Read(tmpChunks, 0, tmpChunks.Length))>0)
          {
            tmpMemoryStream.Write(tmpChunks,0, tmpLength);
          }
          tmpContents = tmpMemoryStream.ToArray();
        }
        return tmpContents;
      }
      finally
      {
        if (tmpResponse != null)
        {
          try
          {
            tmpResponse.Close();
          }

          catch (Exception e)
          {
            PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_LIB, "closing resource", e);
          }
        }
        try
        {
          if (tmpStream != null)
          {
            tmpStream.Dispose();
          }
        }
        catch (Exception e)
        {
          PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_LIB, "closing resource", e);
        }
      }
    }
    #endregion

    #region GetShortFileName
    /// <summary>
    /// Tries to extract a short file name without path information.
    /// Returns the input if the file name does not point to a local
    /// file.
    /// </summary>
    /// <param name="aFilename"></param>
    /// <returns></returns>
    public static string GetShortFileName(string aFilename)
    {
      try
      {
        FileInfo tmpFileInfo = new FileInfo(aFilename);
        return tmpFileInfo.Name;
      }
#pragma warning disable 0168
      catch (Exception e)
      {
        return aFilename;
      }
#pragma warning restore 0168
    }
    #endregion

    #region GetSize
    /// <summary>
    /// Returns the file size of the specified object.
    /// </summary>
    /// <param name="aPath"></param>
    /// <returns></returns>
    public static long GetSize(string aPath)
    {
      if (IsUrl(aPath))
      {
        WebRequest tmpRequest = WebRequest.Create(aPath);
        //tmpRequest.Proxy = new WebProxy();
        using (WebResponse tmpResponse = tmpRequest.GetResponse())
        {
          return tmpResponse.ContentLength;
        }
      }
        FileInfo tmpFileInfo = new FileInfo(aPath);
        return tmpFileInfo.Length;
    }
    #endregion

    #region IsUrl
    /// <summary>
    /// Returns true if the path is considered as a url
    /// (Currently if it starts with http:
    /// </summary>
    /// <param name="aPath"></param>
    /// <returns></returns>
    public static bool IsUrl(string aPath)
    {
      if (aPath.ToLower().StartsWith("http:") || aPath.ToLower().StartsWith("https:"))
      {
        return true;
      }
      return false;
    }
    #endregion

    #endregion
  }
}
