using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Globalization;
using System.IO;
using System.Linq;

using databaseAPI;

using GNA_CommercialLicenseValidator;

using gnaDataClasses;

using GNAgeneraltools;

using GNAspreadsheettools;

using OfficeOpenXml;

using T4Dlibrary;

namespace GNAexportCoordinates
{
    class Program
    {
#pragma warning disable CS0219
#pragma warning disable CS8321
#pragma warning disable CS8600
#pragma warning disable CS8604

        static void Main()
        {
            try
            {
                #region Setting state
                // Applied (8): set once, early
                Console.OutputEncoding = System.Text.Encoding.Unicode;

                int headingNo = 1;
                const string strTab1 = "     ";
                const string strTab2 = "        ";
                const string strTab3 = "           ";

                #region Instantiate core classes
                gnaTools gnaT = new();
                dbAPI gnaDBAPI = new();
                spreadsheetAPI gnaSpreadsheetAPI = new();
                gnaDataClass gnaDC = new();
                T4Dapi t4dapi = new();
                t4dapi.SetCommercial("Dm4eGwoTaGxqY2hv"); // parked (7): remains hard-coded for now
                #endregion

                #region Read config
                NameValueCollection config = ConfigurationManager.AppSettings;
                #endregion
                string strFreezeScreen = CleanConfig(config["freezeScreen"]);
                if (strFreezeScreen.Length == 0) strFreezeScreen = "Yes";


                // Applied (2): remove goto, use a single controlled exit path
                void FinishAndExit()
                {
                    Console.WriteLine("\nGNAcoordinateExporter export completed...\n\n");
                    gnaT.freezeScreen(strFreezeScreen);
                }

                #region Header
                gnaT.WelcomeMessage($"GNAcoordinateExporter {BuildInfo.BuildDateString()}");
                #endregion

                #region Config validation
                Console.WriteLine($"{headingNo++}. System Check");
                Console.Out.Flush();
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



                #region License validation (do not echo product tag)
                Console.WriteLine($"{headingNo++}. Validating the software license");

                string licenseCode = CleanConfig(config["LicenseCode"]);
                if (licenseCode.Length == 0)
                {
                    Console.WriteLine($"{strTab1}License code is not set in the configuration file.");
                    return; // ok: truly fatal for the program
                }

                LicenseValidator.ValidateLicense("CRDEXP", licenseCode);
                Console.WriteLine($"{strTab1}Validated");
                #endregion

                #region EPPlus license
                gnaT.epplusLicense();
                #endregion

                #region General variables
                Console.WriteLine($"{headingNo++}. Variables");
                Console.WriteLine($"{strTab1}General variables");



                string strComputeMeanDeltas = CleanConfig(config["computeMeanDeltas"]);
                if (strComputeMeanDeltas.Length == 0) strComputeMeanDeltas = "No";

                string strUpdateSensorList = CleanConfig(config["updateSensorList"]);
                if (strUpdateSensorList.Length == 0) strUpdateSensorList = "No";

                string strSystemLogsFolder = CleanConfig(config["SystemStatusFolder"]);
                if (strSystemLogsFolder.Length == 0) strSystemLogsFolder = @"C:\__SystemLogs\";

                string strAlarmfolder = CleanConfig(config["SystemAlarmFolder"]);
                if (strAlarmfolder.Length == 0) strAlarmfolder = @"C:\__SystemAlarms\";

                Directory.CreateDirectory(strSystemLogsFolder);
                Directory.CreateDirectory(strAlarmfolder);

                string strTimeBlockType = CleanConfig(config["TimeBlockType"]);
                if (strTimeBlockType.Length == 0) strTimeBlockType = "Schedule";

                string strManualBlockStart = CleanConfig(config["manualBlockStart"]);
                string strManualBlockEnd = CleanConfig(config["manualBlockEnd"]);

                string strBlockSizeHrs = CleanConfig(config["BlockSizeHrs"]);
                if (strBlockSizeHrs.Length == 0) strBlockSizeHrs = "6";

                string strTimeBlockStartLocal = "";
                string strTimeBlockEndLocal = "";
                string strTimeBlockStartUTC = "";
                string strTimeBlockEndUTC = "";
                string strEmailTime = "";
                string logFileMessage = "";
                string strManualEmailTime = "";

                var cs = ConfigurationManager.ConnectionStrings["DBconnectionString"];
                if (cs == null || string.IsNullOrWhiteSpace(cs.ConnectionString))
                    throw new ConfigurationErrorsException("Missing connection string 'DBconnectionString'.");
                string strDBconnection = cs.ConnectionString;

                string strProjectTitle = GetRequired(config, "ProjectTitle");
                string strContractTitle = GetRequired(config, "ContractTitle");

                string strExcelPath = GetRequired(config, "ExcelPath");
                string strExcelFile = GetRequired(config, "ExcelFile");
                string strFTPSubdirectory = GetRequired(config, "FTPSubdirectory");

                string strReferenceWorksheet = GetRequired(config, "ReferenceWorksheet");
                string strSurveyWorksheet = GetRequired(config, "SurveyWorksheet");

                string strFirstDataRow = GetRequired(config, "FirstDataRow");
                int iFirstDataRow = GetRequiredInt(config, "FirstDataRow", 1, 1000000);

                string strExcelWorkbookFullPath = Path.Combine(strExcelPath, strExcelFile);
                if (!File.Exists(strExcelWorkbookFullPath))
                    throw new FileNotFoundException("Excel workbook not found.", strExcelWorkbookFullPath);
                #endregion

                #region Email settings
                Console.WriteLine($"{strTab1}Email settings");

                string strEmailLogin = CleanConfig(config["EmailLogin"]);
                string strEmailPassword = CleanConfig(config["EmailPassword"]);
                string strEmailFrom = CleanConfig(config["EmailFrom"]);
                string strEmailRecipients = CleanConfig(config["EmailRecipients"]);

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
                    string value = CleanConfig(config[key]);
                    if (value.Length == 0) continue;

                    smsMobile.Add(value);

                    if (strMobileList.Length > 0) strMobileList += ",";
                    strMobileList += value;
                }
                #endregion

                #region Environment check
                Console.WriteLine($"{headingNo++}. Check system environment");

                // Applied (4): use IsYes everywhere
                if (IsYes(strFreezeScreen))
                {
                    Console.WriteLine($"{strTab1}Project:{strProjectTitle}");
                    Console.WriteLine($"{strTab1}Check DB connection");
                    gnaDBAPI.testDBconnection(strDBconnection);
                    Console.WriteLine($"{strTab1}Master workbook:{strExcelWorkbookFullPath}");
                    gnaSpreadsheetAPI.checkWorkbookExists(strExcelWorkbookFullPath);
                    Console.WriteLine($"{strTab1}Checking worksheets:");
                    gnaSpreadsheetAPI.checkWorksheetExists(strExcelWorkbookFullPath, strSurveyWorksheet);
                    gnaSpreadsheetAPI.checkWorksheetExists(strExcelWorkbookFullPath, strReferenceWorksheet);
                }
                else
                {
                    Console.WriteLine($"{strTab2}Workbook & worksheets not checked");
                }
                #endregion

                #region Timeblocks
                Console.WriteLine($"{strTab1}Timeblocks");

                List<Tuple<string, string>> subBlocks = new();
                string strColumnHeaderTime = "";

                switch (strTimeBlockType.Trim().ToUpperInvariant())
                {
                    case "HISTORIC":
                        subBlocks = gnaT.prepareTimeBlocks(
                            "Historic",
                            strBlockSizeHrs,
                            strManualBlockStart,
                            strManualBlockEnd);
                        break;

                    case "MANUAL":
                        subBlocks = gnaT.prepareTimeBlocks(
                            "Manual",
                            strManualBlockStart,
                            strManualBlockEnd);

                        // Applied (9): moved to helper
                        strManualEmailTime = BuildManualEmailTime(strManualBlockEnd);
                        break;

                    case "SCHEDULE":
                        subBlocks = gnaT.prepareTimeBlocks(
                            "Schedule",
                            strBlockSizeHrs);
                        break;

                    default:
                        // Applied (10): no Environment.Exit; throw a config exception
                        throw new ConfigurationErrorsException(
                            $"Invalid TimeBlockType '{strTimeBlockType}'. Must be Manual, Schedule or Historic.");
                }
                #endregion

                #region Read configuration values
                Console.WriteLine($"{strTab1}Configuration values");
                string ProjectTitle = strProjectTitle;
                string ContractTitle = strContractTitle;

                string ReportType = GetRequired(config, "ReportType");
                string CoordinateOrder = GetRequired(config, "CoordinateOrder");

                string PrepareCoordinateExportWorkbook = CleanConfig(config["PrepareCoordinateExportWorkbook"]);
                if (PrepareCoordinateExportWorkbook.Length == 0) PrepareCoordinateExportWorkbook = "No";

                string includeHeader = CleanConfig(config["includeHeader"]);
                if (includeHeader.Length == 0) includeHeader = "Yes";

                string ReplacementNames = CleanConfig(config["ReplacementNames"]);
                if (ReplacementNames.Length == 0) ReplacementNames = "Yes";

                double dblDataJumpTriggerLevel = GetRequiredDoubleInvariant(config, "DataJumpTriggerLevel");
                #endregion

                #endregion

                #region Main program

                // Applied (3): removed unconditional UTC conversion + unconditional mean-deltas fetch.
                //             Work is now performed only inside the relevant branches below.

                List<Points> coordinateList = new();



                if (IsYes(PrepareCoordinateExportWorkbook))
                {
                    #region Prepare workbook

                    Console.WriteLine($"{headingNo++}. Workbook preparation");

                    if (strManualBlockStart.Length == 0 || strManualBlockEnd.Length == 0)
                    {
                        Console.WriteLine("\nFor PrepareCoordinateExportWorkbook = Yes, you must set manualBlockStart and manualBlockEnd in the config file");
                        TryReadKey();
                        FinishAndExit();
                        return;
                    }

                    strTimeBlockStartUTC = gnaT.convertLocalToUTC(strManualBlockStart).Trim();
                    strTimeBlockEndUTC = gnaT.convertLocalToUTC(strManualBlockEnd).Trim();

                    Console.WriteLine($"{strTab2}Extract sensor list");
                    coordinateList = t4dapi.GetSensorList(
                        strDBconnection,
                        strProjectTitle);

                    Console.WriteLine($"{strTab2}Extract deltas");
                    coordinateList = t4dapi.UpdatePointsWithMeanDeltas(
                        strDBconnection,
                        strProjectTitle,
                        coordinateList,
                        strTimeBlockStartUTC,
                        strTimeBlockEndUTC);


                    if (strUpdateSensorList == "Yes") { 
                    Console.WriteLine($"{strTab2}Write sensor list to {strSurveyWorksheet}");
                    gnaSpreadsheetAPI.WritePointsToWorksheet(
                        strExcelWorkbookFullPath,
                        strSurveyWorksheet,
                        coordinateList,
                        strFirstDataRow);
                    }

                    Console.WriteLine($"{strTab2}Write reference deltas to {strReferenceWorksheet}");
                    gnaSpreadsheetAPI.writeDeltasList(
                        strExcelWorkbookFullPath,
                        strReferenceWorksheet,
                        strDBconnection,
                        strProjectTitle,
                        coordinateList,
                        iFirstDataRow);

                    Console.WriteLine($"{strTab2}Write default time");
                    gnaSpreadsheetAPI.writeDefaultTimeUTC(
                        strExcelWorkbookFullPath,
                        strReferenceWorksheet,
                        iFirstDataRow);


                    string[] strPointNames = gnaSpreadsheetAPI.readPointNames(strExcelWorkbookFullPath, strSurveyWorksheet, strFirstDataRow);
                    Console.WriteLine($"{strTab2}Extract SensorID");
                    string[,] strSensorID = new string[5000, 2];
                    strSensorID = gnaDBAPI.getSensorIDfromDB(strDBconnection, strPointNames, strProjectTitle);
                    Console.WriteLine($"{strTab2}Write SensorID to workbook");
                    gnaSpreadsheetAPI.writeSensorID(strExcelWorkbookFullPath, strSurveyWorksheet, strSensorID, strFirstDataRow);
                    Console.WriteLine($"{strTab1}Preparation complete");

                    FinishAndExit();
                    return;
                    #endregion
                }
                else
                {
                    #region Export coordinates

                    string defaultStartUTC = gnaT.convertLocalToUTC(strManualBlockStart).Trim();

                    Console.WriteLine($"{headingNo++}. Export Coordinates to CSV file");
                    coordinateList = gnaSpreadsheetAPI.readPointDataToList(
                        strExcelWorkbookFullPath,
                        strReferenceWorksheet,
                        strFirstDataRow);


                    if (strComputeMeanDeltas == "No")
                    {
                        // pointDataList will contain ONE record per delta observation (append-only)
                        List<Points> pointDataList = new();

                        // Read the reference / master point list from Excel
                        // (includes per-point last TimeBlockEndUTC from column 41)
                        List<Points> pointMasterList = gnaSpreadsheetAPI.readPointDataToList(
                            strExcelWorkbookFullPath,
                            strReferenceWorksheet,
                            strFirstDataRow);

                        // Defensive: nothing to do
                        if (pointMasterList == null || pointMasterList.Count == 0)
                            return;

                        // Iterate over prepared sub-blocks
                        foreach (var block in subBlocks)
                        {
                            string blockEndUTC = block.Item2;
                            Console.WriteLine(
                                $"{strTab2}Retrieving ALL deltas (per-point start) up to {blockEndUTC}");

                            // ONE DB call per block, with per-point start times taken from column 41
                            List<Points> blockResults =
                                t4dapi.GetAllPointsAllDeltas_PerPointStart_OnePass(
                                    strDBconnection,
                                    pointMasterList,
                                    strManualBlockStart,   // default start if column 41 is blank
                                    blockEndUTC);

                            // Append results (one record per delta)
                            if (blockResults != null && blockResults.Count > 0)
                            {
                                pointDataList.AddRange(blockResults);

                                // Determine which points actually returned data in THIS block
                                HashSet<string> pointsWithData = new(
                                    blockResults
                                        .Where(p => !string.IsNullOrWhiteSpace(p.Name))
                                        .Select(p => p.Name!)
                                        .Distinct(StringComparer.OrdinalIgnoreCase),
                                    StringComparer.OrdinalIgnoreCase);

                                // Update column 41 (TimeBlockEndUTC) ONLY for points that had data
                                gnaSpreadsheetAPI.UpdateLastRetrievedTimeByPoint(
                                    strExcelWorkbookFullPath,
                                    strReferenceWorksheet,
                                    strFirstDataRow,
                                    pointsWithData,
                                    blockEndUTC);

                                // Keep the in-memory master list aligned with Excel for subsequent blocks
                                foreach (var p in pointMasterList)
                                {
                                    if (p.Name != null && pointsWithData.Contains(p.Name))
                                    {
                                        p.TimeBlockEndUTC = blockEndUTC;
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine($"{strTab2}No deltas retrieved up to {blockEndUTC}");
                            }
                        }


                        Console.WriteLine("\nEcho pointDataList (one record per delta):");

                        if (pointDataList == null || pointDataList.Count == 0)
                        {
                            Console.WriteLine($"{strTab1}<empty>");
                        }
                        else
                        {
                            int i = 1;
                            foreach (Points p in pointDataList)
                            {
                                if (p == null) continue;

                                Console.WriteLine(
                                    $"{i++:D4} | " +
                                    $"Name={p.Name ?? "<null>"} | " +
                                    $"SensorID={p.SensorID ?? "<null>"} | " +
                                    $"UTC={p.UTCtime ?? "<null>"} | " +
                                    $"dE={p.dE:F4} dN={p.dN:F4} dH={p.dH:F4} | " +
                                    $"dEcor={p.dEcor:F4} dNcor={p.dNcor:F4} dHcor={p.dHcor:F4} | " +
                                    $"Eref={p.Eref:F4} Nref={p.Nref:F4} Href={p.Href:F4} | " +
                                    $"Block={p.TimeBlockStartUTC ?? "<null>"} → {p.TimeBlockEndUTC ?? "<null>"}"
                                );
                            }
                        }

                        Console.Out.Flush();
                        // Optional pause for inspection
                        // Console.ReadKey();





                        // pointDataList now contains ALL retrieved delta observations
                        // Column 41 in Excel has been advanced per-point, per-block, only where data existed
                    }

                    #endregion
                }

                FinishAndExit();
                return;




                #endregion
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

        #region Config helpers
        static string CleanConfig(string s) => (s ?? string.Empty).Trim().Trim('\'', '"');

        static string GetRequired(NameValueCollection cfg, string key)
        {
            string v = CleanConfig(cfg[key]);
            if (v.Length == 0)
                throw new ConfigurationErrorsException($"Missing/empty config key '{key}'.");
            return v;
        }

        static int GetRequiredInt(NameValueCollection cfg, string key, int minValueInclusive = int.MinValue, int maxValueInclusive = int.MaxValue)
        {
            string s = GetRequired(cfg, key);
            if (!int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out int v))
                throw new ConfigurationErrorsException($"Config key '{key}' is invalid (expected integer). Value='{s}'.");
            if (v < minValueInclusive || v > maxValueInclusive)
                throw new ConfigurationErrorsException($"Config key '{key}' is out of range. Value={v}.");
            return v;
        }

        static double GetRequiredDoubleInvariant(NameValueCollection cfg, string key)
        {
            string s = GetRequired(cfg, key);
            if (!double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out double v))
                throw new ConfigurationErrorsException($"Config key '{key}' is invalid (expected invariant decimal, e.g. 0.030). Value='{s}'.");
            return v;
        }

        static bool IsYes(string s) => string.Equals(CleanConfig(s), "Yes", StringComparison.OrdinalIgnoreCase);
        #endregion

        #region General Helpers
        private static void TryReadKey()
        {
            try
            {
                if (Environment.UserInteractive && !Console.IsInputRedirected)
                    Console.ReadKey(intercept: true);
            }
            catch { }
        }

        // Applied (9): extracted from inline surgery
        private static string BuildManualEmailTime(string manualBlockEnd)
        {
            if (string.IsNullOrWhiteSpace(manualBlockEnd))
                return string.Empty;

            string tmp = manualBlockEnd.Replace("-", "")
                                       .Replace(" ", "_")
                                       .Replace(":", "h") + "m";
            return tmp.Length >= 14 ? tmp.Substring(0, 14) : tmp;
        }
        #endregion
    }
}

