//-----------------------------------------------------------------------
// 
//  Copyright (C) Microsoft Corporation.  All rights reserved.
// 
// THIS CODE AND INFORMATION ARE PROVIDED AS IS WITHOUT WARRANTY OF ANY
// KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
// IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//-----------------------------------------------------------------------

using System;
using System.ComponentModel;
using System.Configuration.Install;
using System.IO;
using Microsoft.Win32;
using ICSharpCode.SharpZipLib.Zip;
namespace CustomActions
{
    [RunInstaller(true)]
    [System.Security.Permissions.PermissionSetAttribute(System.Security.Permissions.SecurityAction.Demand, Name = "FullTrust")]
    public sealed partial class SetSecurity : Installer
    {
        #region constructor
        public SetSecurity()
        {
            InitializeComponent();
        }
        #endregion

        #region methods

        #region GetSubKeyName
        private static string GetSubKeyName(string aGUID, string subKeyName)
        {
            System.Text.StringBuilder s = new System.Text.StringBuilder();
            s.Append(@"CLSID\{");
            s.Append(aGUID.ToUpper());
            s.Append(@"}\");
            s.Append(subKeyName);
            return s.ToString();
        }
        private static string GetInitial()
        {
            return "PdcInitial.zip";
        }

        private static string GetSubKeyName(Type type, string subKeyName)
        {
            return GetSubKeyName(type.GUID.ToString().ToUpper(), subKeyName);
        }
        #endregion

        #region Install
        public override void Install(System.Collections.IDictionary stateSaver)
        {
            // Call the base implementation.
            base.Install(stateSaver);

            RegisterFunction(typeof(PDCOpenLibrary.PDCOpenLib));
        }

        #endregion

        #region RegisterFunction
        public static void RegisterFunction(Type type)
        {
            Registry.ClassesRoot.CreateSubKey(GetSubKeyName(type, "Programmable"));
            RegistryKey key = Registry.ClassesRoot.OpenSubKey(GetSubKeyName(type, "InprocServer32"), true);
            key.SetValue("", Environment.SystemDirectory + @"\mscoree.dll", RegistryValueKind.String);
        }
        #endregion

        #region Rollback
        public override void Rollback(System.Collections.IDictionary savedState)
        {
            // Check whether the "allUsers" property is saved.
            // If it is not set, the Install method did not set the security policy.
            if ((savedState == null) || (savedState["allUsers"] == null))
                return;

            base.Rollback(savedState);
            UnregisterFunction(typeof(PDCOpenLibrary.PDCOpenLib));
        }

        #endregion

        #region Uninstall
        public override void Uninstall(System.Collections.IDictionary savedState)
        {
            // Call the base implementation.
            base.Uninstall(savedState);

            // Check whether the "allUsers" property is saved.
            // If it is not set, the Install method did not set the security policy.
            if ((savedState == null) || (savedState["allUsers"] == null))
                return;

            UnregisterFunction(typeof(PDCOpenLibrary.PDCOpenLib));


        }

        private void UndeleteUpdateFolder(string updateFolder)
        {

            if (Directory.Exists(updateFolder))
            {
                foreach (string fileName in Directory.GetFiles(updateFolder))
                {
                    try
                    {
                        File.Delete(fileName);
                    }
                    catch
                    {
                        //Some files produce an exception if they cannot be deleted
                        //throw Exception ex; 
                    }
                }
                Directory.Delete(updateFolder, true);
            }



        }

        #endregion

        #region OnAfterInstall

        protected override void OnAfterInstall(System.Collections.IDictionary savedState)
        {
            base.OnAfterInstall(savedState);

            try
            {
                string targetDir = Context.Parameters["targetDir"];
                string file = Path.Combine(targetDir, GetInitial());

                ZipFile zip = ZipFile.Create(file);


                zip.BeginUpdate();

                AddZipFiles(zip, targetDir);

                zip.CommitUpdate();
                zip.Close();
            }
            catch
            {
            }
        }

        #endregion

        #region OnBeforeUninstall

        protected override void OnBeforeUninstall(System.Collections.IDictionary savedState)
        {
            base.OnBeforeUninstall(savedState);

            try
            {
                string targetDir = Context.Parameters["targetDir"];
                ExtractZipFile(targetDir, GetInitial());

                File.Delete(Path.Combine(targetDir, GetInitial()));
                File.Delete(Path.Combine(targetDir, "PDCVersion.dll"));

            }
            catch
            {
            }
        }

        #endregion

        protected override void OnAfterUninstall(System.Collections.IDictionary savedState)
        {
            base.OnAfterUninstall(savedState);

            // as targetDir is not a Parameter here, get the DLL name
            string targetDir = Context.Parameters["assemblypath"];
            // remove the file name to get the install directory
            targetDir = Directory.GetParent(targetDir).FullName;
            // if there is the actual something as install directory ... :-)
            if (targetDir != null)
            {
                try
                {
                    // delete the "update" subdirectory
                    string undelete = Path.Combine(targetDir, "update");
                    UndeleteUpdateFolder(undelete);

                }
                catch (Exception)
                {
                }
                try
                {
                    string initial = Path.Combine(targetDir, GetInitial());
                    if (File.Exists(initial))
                    {
                        File.Delete(initial);
                    }
                }
                catch (Exception)
                {
                }
            }
        }
        #region UnregisterFunction
        /// <summary>
        /// deregisters the COM functions of this class
        /// </summary>
        /// <param name="type"></param>
        public static void UnregisterFunction(Type type)
        {
            try
            {
                Registry.ClassesRoot.DeleteSubKey(GetSubKeyName(type, "Programmable"), false);
            }
#pragma warning disable 0168
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine("Could not unregister COM AddIn");
            }
#pragma warning restore 0168
        }
        #endregion

        #endregion

        #region AddZipFiles

        private void AddZipFiles(ZipFile zipFile, String directory)
        {
            foreach (String file in Directory.GetFiles(directory))
            {
                if (file.EndsWith(GetInitial())) continue;

                zipFile.Add(file, Path.GetFileName(file));
            }

            foreach (String dir in Directory.GetDirectories(directory))
                AddZipFiles(zipFile, dir);
        }

        #endregion

        #region ExtractZipFile

        /// <summary>
        ///   Extracts all auto updater files from the specified zip file into the specified folder.
        /// </summary>
        /// <param name="instructions">
        ///   If not set all files are extracted, if set only autoupdater files are extracted
        ///   Contains info which files are autoupdater files.</param>
        /// <param name="zipFile">
        ///   The name of the zip file to extract.
        /// </param>
        /// <param name="folder">
        ///   Target directory for the unzip operation.
        /// </param>
        /// <returns>
        ///   True if the extract process succeeded; otherwise false.
        /// </returns>
        protected Boolean ExtractZipFile(String folder, String zipFile)
        {
            ZipEntry entry;
            FileStream stream = null;
            ZipInputStream zipStream = null;
            String path;
            String file;


            try
            {
                file = Path.Combine(folder, zipFile);

                stream = File.OpenRead(file);

                zipStream = new ZipInputStream(stream);


                // loop through all files
                while ((entry = zipStream.GetNextEntry()) != null)
                {
                    if (entry.IsDirectory) continue;

                    // create path of the file to extract if it does not exist.  
                    file = Path.Combine(folder, entry.Name);

                    path = Path.GetDirectoryName(file);

                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);

                    Write(zipStream, file);
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (zipStream != null) zipStream.Close();

                if (stream != null) stream.Close();
            }
        }

        #endregion

        #region Write

        /// <summary>
        ///   Writes the file from the zip stream into the specified file location.
        /// </summary>
        /// <param name="logger">
        ///   The logger to use for logging the process.
        /// </param>
        /// <param name="zipStream">
        ///   The stream to use for read the bytes to write.
        /// </param>
        /// <param name="file">
        ///   The name of the file to write.
        /// </param>
        /// <returns>
        ///   True if the process succeeded; otherwise false.
        /// </returns>
        private void Write(ZipInputStream zipStream, String file)
        {
            FileStream fileStream = null;
            Int32 size = 2048;
            Byte[] data;


            try
            {
                fileStream = new FileStream(file, FileMode.Create);

                data = new Byte[size];

                while ((size = zipStream.Read(data, 0, data.Length)) > 0)
                    fileStream.Write(data, 0, size);
            }
            catch
            {
                throw;
            }
            finally
            {
                if (fileStream != null)
                {
                    fileStream.Flush();
                    fileStream.Close();
                }
            }
        }

        #endregion
    }
}