using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;

using databaseAPI;
using gnaDataClasses;
using GNA_CommercialLicenseValidator;
using GNAgeneraltools;
using GNAspreadsheettools;
using OfficeOpenXml;

namespace GNAexportCoordinates
{
    class Program
    {
        static void Main()
        {
            try
            {
                int headingNo = 1;
                const string strTab1 = "     ";

                #region Instantiate core classes
                gnaTools gnaT = new();
                dbAPI gnaDBAPI = new();
                spreadsheetAPI gnaSpreadsheetAPI = new();
                gnaDataClass gnaDC = new();
                #endregion

                #region Banner
                gnaT.WelcomeMessage($"GNAcoordinateExporter {BuildInfo.BuildDateString()}");
                #endregion

                Console.WriteLine($"{headingNo++}. System Check");
                Console.Out.Flush();

                #region Config validation
                try
                {
                    gnaT.VerifyLocalConfig();
                    Console.WriteLine($"{strTab1}VerifyLocalConfig returned OK");
                    Console.Out.Flush();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("\nVerifyLocalConfig threw:");
                    Console.WriteLine(ex);
                    Console.Out.Flush();
                    throw;
                }
                #endregion

                #region Read config
                var config = ConfigurationManager.AppSettings;
                #endregion

                #region License validation (do not echo product tag)
                Console.WriteLine($"{headingNo++}. Validating the software license");

                string licenseCode = (config["LicenseCode"] ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(licenseCode))
                {
                    Console.WriteLine($"{strTab1}License code is not set in the configuration file.");
                    return;
                }

                LicenseValidator.ValidateLicense("CRDEXP", licenseCode);
                Console.WriteLine($"{strTab1}Validated");
                #endregion

                #region EPPlus license
                gnaT.epplusLicense();
                #endregion

                Console.WriteLine($"{headingNo++}. Variables");

                #region General variables
                Console.WriteLine($"{strTab1}General variables");

                Console.OutputEncoding = System.Text.Encoding.Unicode;

                string strFreezeScreen = (config["freezeScreen"] ?? "Yes").Trim();
                string strSystemLogsFolder = (config["SystemStatusFolder"] ?? @"C:\__SystemLogs\").Trim();
                string strAlarmfolder = (config["SystemAlarmFolder"] ?? @"C:\__SystemAlarms\").Trim();

                Directory.CreateDirectory(strSystemLogsFolder);
                Directory.CreateDirectory(strAlarmfolder);

                string strTimeBlockType = (config["TimeBlockType"] ?? "Schedule").Trim();
                string strManualBlockStart = (config["manualBlockStart"] ?? string.Empty).Trim();
                string strManualBlockEnd = (config["manualBlockEnd"] ?? string.Empty).Trim();
                string strBlockSizeHrs = (config["BlockSizeHrs"] ?? "6").Trim();

                string strTimeBlockStartLocal = "";
                string strTimeBlockEndLocal = "";
                string strTimeBlockStartUTC = "";
                string strTimeBlockEndUTC = "";
                string strEmailTime = "";
                string logFileMessage = "";
                string strManualEmailTime = "";
                #endregion

                #region Email settings
                Console.WriteLine($"{strTab1}Email settings");

                string strEmailLogin = (config["EmailLogin"] ?? string.Empty).Trim();
                string strEmailPassword = (config["EmailPassword"] ?? string.Empty).Trim();
                string strEmailFrom = (config["EmailFrom"] ?? string.Empty).Trim();
                string strEmailRecipients = (config["EmailRecipients"] ?? string.Empty).Trim();

                EmailCredentials emailCreds = gnaT.BuildEmailCredentials(
                    strEmailLogin,
                    strEmailPassword,
                    strEmailFrom,
                    strEmailRecipients);
                #endregion

                #region SMS settings
                Console.WriteLine($"{strTab1}SMS settings");

                List<string> smsMobile = new();
                string strMobileList = "";

                foreach (string key in config.AllKeys.Where(k => !string.IsNullOrWhiteSpace(k) &&
                                                                k.StartsWith("RecipientPhone", StringComparison.OrdinalIgnoreCase)))
                {
                    string value = (config[key] ?? string.Empty).Trim();
                    if (value.Length == 0) continue;

                    smsMobile.Add(value);

                    if (strMobileList.Length > 0) strMobileList += ",";
                    strMobileList += value;
                }
                #endregion

                #region Timeblocks
                Console.WriteLine($"{strTab1}Timeblocks");

                List<Tuple<string, string>> subBlocks = new();
                string strColumnHeaderTime = "";

                switch (strTimeBlockType)
                {
                    case "Historic":
                        subBlocks = gnaT.prepareTimeBlocks(
                            "Historic",
                            strBlockSizeHrs,
                            strManualBlockStart,
                            strManualBlockEnd);
                        break;

                    case "Manual":
                        subBlocks = gnaT.prepareTimeBlocks(
                            "Manual",
                            strManualBlockStart,
                            strManualBlockEnd);

                        {
                            string tmp = strManualBlockEnd.Replace("-", "")
                                                          .Replace(" ", "_")
                                                          .Replace(":", "h") + "m";
                            strManualEmailTime = tmp.Length >= 14 ? tmp.Substring(0, 14) : tmp;
                        }
                        break;

                    case "Schedule":
                        subBlocks = gnaT.prepareTimeBlocks(
                            "Schedule",
                            strBlockSizeHrs);
                        break;

                    default:
                        Console.WriteLine("\nError in Timeblock Type");
                        Console.WriteLine($"{strTab1}TimeBlockType: {strTimeBlockType}");
                        Console.WriteLine($"{strTab1}Must be Manual, Schedule or Historic");
                        Console.WriteLine("\nPress key to exit...");
                        TryReadKey();
                        Environment.Exit(1);
                        break;
                }
                #endregion

                #region Main program
                // Code goes here
                #endregion

                Console.WriteLine("\nGNAcoordinateExporter export completed...\n\n");
                gnaT.freezeScreen(strFreezeScreen);
                return;
            }
            catch (Exception ex)
            {
                try { File.WriteAllText("fatal_crash.log", ex.ToString()); } catch { }

                try
                {
                    Console.WriteLine("Fatal crash:");
                    Console.WriteLine(ex);
                    Console.Out.Flush();
                }
                catch { }
            }
        }

        #region Helpers
        private static void TryReadKey()
        {
            try
            {
                if (Environment.UserInteractive && !Console.IsInputRedirected)
                    Console.ReadKey(intercept: true);
            }
            catch { }
        }
        #endregion
    }
}


