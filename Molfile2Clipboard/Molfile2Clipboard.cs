using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Xml.Serialization;
using BBS.ST.BHC.BSP.PDC.Lib.Util;
using BBS.ST.IVY.Chemistry.FileFormats;
using BBS.ST.IVY.Chemistry.Util;

namespace Molfile2Clipboard
{
    class Molfile2Clipboard
    {
        [STAThread]
        static void Main(string[] args)
        {
            Form form = new Form();

            form.Load += (s, e) =>
            {
                form.ShowInTaskbar = false;
                form.WindowState = FormWindowState.Minimized;
            };
            form.Shown += (s,e) =>
            {
                if (args.Length != 1)
                {
                    Console.Out.WriteLine("Usersettings missing");
                }
                UserSettings userSettings = ReadUsersettings(args);
                while (true)
                {
                    string molfile = ReadMolfileFromInput();
                    if (string.IsNullOrEmpty(molfile))
                    {
                        Application.Exit();
                        return;
                    }

                    string result = RenderMolfile2Clipboard(molfile, userSettings);
                    Console.WriteLine(result);
                }
            };
            Application.Run(form);
        }

        private static UserSettings ReadUsersettings(string[] args)
        {
            FileStream file = new FileStream(args[0], FileMode.Open);
            return (UserSettings) new XmlSerializer(typeof(UserSettings)).Deserialize(file);
        }

        private static string ReadMolfileFromInput()
        {
            StringBuilder molfile = new StringBuilder();
            while (true)
            {
                string line = Console.In.ReadLine();
                if (line == null || line.Trim().Equals("$$$$", StringComparison.OrdinalIgnoreCase))
                {
                    return molfile.ToString();
                }

                molfile.AppendLine(line);
            }
        }

        public static string RenderMolfile2Clipboard(string molFile, UserSettings userSettings)
    {
        if (molFile == null) return null;
        if (molFile.Trim().Equals("")) return null;
        PDCLogger.TheLogger.LogStarttime("PDCLib.InsertISISObject", "InsertISISObject - Method start");
        try
        {
            // insert Structure OLE object
            var formatConverter = new FormatConverter();

            DisplayPreferences displayPreferences = DisplayPreferencesByUserSetting(userSettings);
            PDCLogger.TheLogger.LogStarttime("PDCLib.InsertISISObject.Copy", "InsertISISObject.Copy - Copying to clipboard");
            formatConverter.CopyToClipboard(molFile, displayPreferences);
            PDCLogger.TheLogger.LogStarttime("PDCLib.InsertISISObject.Copy", "InsertISISObject.Copy - Copied to clipboard");
        }
        catch (Exception e)
        {
            PDCLogger.TheLogger.LogException(nameof(RenderMolfile2Clipboard), "Render2Clipboard failed", e);
            return e.Message;
        }
        finally
        {
            PDCLogger.TheLogger.LogStoptime("PDCLib.InsertISISObject", "InsertISISObject - Method end");
        }

        return "ok";
    }


        private static DisplayPreferences DisplayPreferencesByUserSetting(UserSettings userSettings)
        {
            DisplayPreferences displayPreferences =DisplayPreferences.GetSystemDefaultPreferences();
            // set the display preferences as selected in the user settings
            displayPreferences.BondLength = userSettings.BondLength / 2.54 * 720;
            displayPreferences.ChemLabelFont = userSettings.ChemLabelFont;
            displayPreferences.DisplayCarbonLabels = userSettings.DisplayCarbonLabels;
            displayPreferences.HydrogenDisplayMode = userSettings.HydrogenDisplayMode;
            displayPreferences.TextFont = userSettings.TextFont;
            displayPreferences.TransparentBackground = userSettings.TransparentBackground;
            displayPreferences.ColorAtomsByType = userSettings.AtomColor;
            return displayPreferences;
        }
    }
}
