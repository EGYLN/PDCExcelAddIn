using System;
using System.IO;
using System.Xml.Serialization;
using BBS.ST.BHC.BSP.PDC.Lib.Util;
using Microsoft.Win32;

namespace BBS.ST.BHC.BSP.PDC.Lib
{
    public class RegistryUtil
    {
        /// <summary>
        /// Property for the login state in the registry
        /// </summary>
        public static UserInfo LoggedInUser
        {
            get
            {
                UserInfo user = null;
                try
                {
                    EncryptUtil.Key = (string) Registry.GetValue(PDCClientConstants.PDC_REGISTRY_KEY, "LastLogin", "");
                    var userString = (string) Registry.GetValue(PDCClientConstants.PDC_REGISTRY_KEY, "Login", "");
                    userString = EncryptUtil.Decrypt(userString);

                    if (string.IsNullOrEmpty(userString))
                    {
                        return null;
                    }

                    XmlSerializer serializer = new XmlSerializer(typeof(UserInfo));
                    using (StringReader reader = new StringReader(userString))
                    {
                        user = (UserInfo) serializer.Deserialize(reader);
                    }
                }
                catch (Exception e)
                {
                    PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Error decrypting user", e);
                    LoggedInUser = null;
                }

                return user;
            }
            set
            {
                string serializedValue = null;
                if (value != null)
                {
                    String dateTime = DateTime.Now.ToString();
                    Registry.SetValue(PDCClientConstants.PDC_REGISTRY_KEY, "LastLogin", dateTime);
                    EncryptUtil.Key = dateTime;
                    XmlSerializer serializer = new XmlSerializer(typeof(UserInfo));
                    using (TextWriter writer = new StringWriter())
                    {
                        serializer.Serialize(writer, value);
                        serializedValue = writer.ToString();
                    }

                }
                Registry.SetValue(PDCClientConstants.PDC_REGISTRY_KEY, "Login",
                    EncryptUtil.Encrypt(serializedValue));
            }
        }
    }
}