using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
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
            gnaTools? gnaT = null;
            string strFreezeScreen = "Yes";
            string strSystemLogsFolder = @"C:\__SystemLogs\";
            int exitCode = 0;

            try
            {
                #region Setting state
                Console.OutputEncoding = System.Text.Encoding.Unicode;
                if (Environment.UserInteractive && !Console.IsOutputRedirected) Console.Clear();

                int headingNo = 1;
                const string strTab1 = "     ";
                const string strTab2 = "        ";
                const string strTab3 = "           ";
                #endregion

                #region Instantiate core classes
                gnaT = new gnaTools();
                dbAPI gnaDBAPI = new();
                spreadsheetAPI gnaSpreadsheetAPI = new();
                T4Dapi t4dapi = new();
                t4dapi.SetCommercial("Dm4eGwoTaGxqY2hv");
                #endregion

                #region Read config early
                NameValueCollection config = ConfigurationManager.AppSettings;

                strFreezeScreen = CleanConfig(config["freezeScreen"]);
                if (strFreezeScreen.Length == 0) strFreezeScreen = "Yes";
                #endregion

                #region Header
                gnaT.WelcomeMessage($"GNAcoordinateExporter {BuildInfo.BuildDateString()}");
                #endregion

                #region Config validation
                Console.WriteLine($"{headingNo++}. System Check");
                Console.Out.Flush();
                gnaT.VerifyLocalConfig();
                Console.WriteLine($"{strTab1}VerifyLocalConfig returned OK");
                Console.Out.Flush();
                #endregion

                #region License validation
                Console.WriteLine($"{headingNo++}. Validating the software license");
                string licenseCode = CleanConfig(config["LicenseCode"]);
                if (licenseCode.Length == 0)
                    throw new ConfigurationErrorsException("LicenseCode missing/empty.");
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

                strSystemLogsFolder = CleanConfig(config["SystemLogsFolder"]);
                if (strSystemLogsFolder.Length == 0) strSystemLogsFolder = @"C:\__SystemLogs\";

                string strAlarmfolder = CleanConfig(config["SystemAlarmFolder"]);
                if (strAlarmfolder.Length == 0) strAlarmfolder = @"C:\__SystemAlarms\";

                Directory.CreateDirectory(strSystemLogsFolder);
                Directory.CreateDirectory(strAlarmfolder);

                string strTimeBlockType = CleanConfig(config["TimeBlockType"]);
                if (strTimeBlockType.Length == 0) strTimeBlockType = "Schedule";

                string strManualBlockStart = gnaT.NormalizeTimeStampToString(CleanConfig(config["manualBlockStart"]));
                string strManualBlockEnd = gnaT.NormalizeTimeStampToString(CleanConfig(config["manualBlockEnd"]));

                string strBlockSizeHrs = CleanConfig(config["BlockSizeHrs"]);
                if (strBlockSizeHrs.Length == 0) strBlockSizeHrs = "6";

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

                int iFirstDataRow = GetRequiredInt(config, "FirstDataRow", 1, 1000000);
                string strFirstDataRow = iFirstDataRow.ToString(CultureInfo.InvariantCulture);

                string strExcelWorkbookFullPath = Path.Combine(strExcelPath, strExcelFile);
                if (!File.Exists(strExcelWorkbookFullPath))
                    throw new FileNotFoundException("Excel workbook not found.", strExcelWorkbookFullPath);

                string PrepareCoordinateExportWorkbook = CleanConfig(config["PrepareCoordinateExportWorkbook"]);
                if (PrepareCoordinateExportWorkbook.Length == 0) PrepareCoordinateExportWorkbook = "No";

                if (!Directory.Exists(strFTPSubdirectory))
                    Directory.CreateDirectory(strFTPSubdirectory);
                #endregion

                #region CSV settings
                string CoordinateOrder = GetRequired(config, "CoordinateOrder");
                string includeHeader = CleanConfig(config["includeHeader"]);
                if (includeHeader.Length == 0) includeHeader = "Yes";

                string OutputFileExtension = CleanConfig(config["OutputFileExtension"]);
                if (OutputFileExtension.Length == 0) OutputFileExtension = "csv";

                string CSVseparator = CleanConfig(config["CSVseparator"]);
                if (CSVseparator.Length == 0) CSVseparator = ",";

                string CSVformat = CleanConfig(config["CSVformat"]);
                if (CSVformat.Length == 0) CSVformat = "Standard";

                string[] allowedFormats = { "Standard", "Datum", "Dywidag", "MissionOS" };
                if (!allowedFormats.Contains(CSVformat, StringComparer.OrdinalIgnoreCase))
                    throw new ConfigurationErrorsException($"CSVformat invalid. Value='{CSVformat}'. Allowed: {string.Join(", ", allowedFormats)}.");
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
                foreach (string key in config.AllKeys.Where(k => !string.IsNullOrWhiteSpace(k) &&
                                                                k.StartsWith("RecipientPhone", StringComparison.OrdinalIgnoreCase)))
                {
                    string value = CleanConfig(config[key]);
                    if (value.Length == 0) continue;
                    smsMobile.Add(value);
                }
                #endregion

                #region Run header log
                {
                    string runHeader =
                        $"Run start | Build={BuildInfo.BuildDateString()} | Project='{strProjectTitle}' | Contract='{strContractTitle}' | " +
                        $"Mode={(IsYes(PrepareCoordinateExportWorkbook) ? "PrepareWorkbook" : "Export")} | TimeBlockType='{strTimeBlockType}' | " +
                        $"ManualStart='{strManualBlockStart}' | ManualEnd='{strManualBlockEnd}' | BlockSizeHrs='{strBlockSizeHrs}' | " +
                        $"ComputeMeanDeltas='{strComputeMeanDeltas}' | CSVformat='{CSVformat}' | Sep='{CSVseparator}' | Header='{includeHeader}' | " +
                        $"Workbook='{strExcelWorkbookFullPath}' | RefWS='{strReferenceWorksheet}' | SurveyWS='{strSurveyWorksheet}' | FirstRow={iFirstDataRow} | " +
                        $"OutDir='{strFTPSubdirectory}'";
                    gnaT.updateSystemLogFile(strSystemLogsFolder, runHeader);
                }
                #endregion

                #region Timeblocks
                Console.WriteLine($"{strTab1}Timeblocks");

                List<Tuple<string, string>> subBlocks;
                string strManualEmailTime = "";

                // prepareTimeBlocks() returns UTC timestamps as strings (yyyy-MM-dd HH:mm:ss)
                switch (strTimeBlockType.Trim().ToUpperInvariant())
                {
                    case "HISTORIC":
                        subBlocks = gnaT.prepareTimeBlocks("Historic", strBlockSizeHrs, strManualBlockStart, strManualBlockEnd);
                        break;

                    case "MANUAL":
                        subBlocks = gnaT.prepareTimeBlocks("Manual", strManualBlockStart, strManualBlockEnd);
                        strManualEmailTime = BuildManualEmailTime(strManualBlockEnd);
                        break;

                    case "SCHEDULE":
                        subBlocks = gnaT.prepareTimeBlocks("Schedule", strBlockSizeHrs);
                        break;

                    default:
                        throw new ConfigurationErrorsException($"Invalid TimeBlockType '{strTimeBlockType}'. Must be Manual, Schedule or Historic.");
                }
                #endregion

                #region Main program
                if (IsYes(PrepareCoordinateExportWorkbook))
                {
                    #region Prepare workbook
                    Console.WriteLine($"{headingNo++}. Workbook preparation");

                    if (strManualBlockStart.Length == 0 || strManualBlockEnd.Length == 0)
                        throw new ConfigurationErrorsException("For PrepareCoordinateExportWorkbook=Yes, manualBlockStart and manualBlockEnd must be set.");

                    string strTimeBlockStartUTC = gnaT.convertLocalToUTC(strManualBlockStart).Trim();
                    string strTimeBlockEndUTC = gnaT.convertLocalToUTC(strManualBlockEnd).Trim();

                    Console.WriteLine($"{strTab2}Extract sensor list");
                    List<Points> coordinateList = t4dapi.GetSensorList(strDBconnection, strProjectTitle);

                    Console.WriteLine($"{strTab2}Extract deltas");
                    coordinateList = t4dapi.UpdatePointsWithMeanDeltas(
                        strDBconnection,
                        strProjectTitle,
                        coordinateList,
                        strTimeBlockStartUTC,
                        strTimeBlockEndUTC);

                    if (IsYes(strUpdateSensorList))
                    {
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
                    string[,] strSensorID = gnaDBAPI.getSensorIDfromDB(strDBconnection, strPointNames, strProjectTitle);

                    Console.WriteLine($"{strTab2}Write SensorID to workbook");
                    gnaSpreadsheetAPI.writeSensorID(strExcelWorkbookFullPath, strSurveyWorksheet, strSensorID, strFirstDataRow);

                    Console.WriteLine($"{strTab1}Preparation complete");
                    exitCode = 0;
                    #endregion
                }
                else
                {
                    #region Export coordinates
                    Console.WriteLine($"{headingNo++}. Export Coordinates: {strTimeBlockType}: {CSVformat} format");
                    Console.WriteLine($"{strTab1}Compute means: {strComputeMeanDeltas}");

                    Console.WriteLine($"{strTab1}Read point data to list");
                    List<Points> pointMasterList = gnaSpreadsheetAPI.readPointDataToList(
                        strExcelWorkbookFullPath,
                        strReferenceWorksheet,
                        strFirstDataRow);

                    if (pointMasterList == null || pointMasterList.Count == 0)
                        throw new InvalidOperationException("Reference point list is empty.");

                    Console.WriteLine($"{strTab1}Iterate over time blocks:");

                    foreach (var block in subBlocks)
                    {
                        string blockStartUTC = gnaT.NormalizeTimeStampToString(block.Item1);
                        string blockEndUTC = gnaT.NormalizeTimeStampToString(block.Item2);

                        // Deterministic filename: ContractTitle + formatted block end time (UTC)
                        string formattedTime = FormatUtcForFilename(blockEndUTC);
                        string expectedCsvPath = Path.Combine(
                            strFTPSubdirectory,
                            $"{SanitizeForFilename(strContractTitle)}_{formattedTime}.{OutputFileExtension}");

                        // Idempotency: skip if already exists
                        if (File.Exists(expectedCsvPath))
                        {
                            Console.WriteLine($"{strTab3}CSV exists (skip): {expectedCsvPath}");
                            gnaT.updateSystemLogFile(strSystemLogsFolder, $"Skipped (exists): {expectedCsvPath}");
                            continue;
                        }

                        Console.WriteLine($"{strTab2}Retrieving deltas: {blockStartUTC} to {blockEndUTC}");

                        List<Points> blockResults = t4dapi.GetAllPointsAllDeltas_PerPointStart_OnePass(
                            strDBconnection,
                            pointMasterList,
                            strTimeBlockType,
                            blockStartUTC,
                            blockEndUTC,
                            strComputeMeanDeltas);

                        if (blockResults == null || blockResults.Count == 0)
                        {
                            Console.WriteLine($"{strTab3}No deltas retrieved up to {blockEndUTC}");
                            continue;
                        }

                        HashSet<string> pointsWithData = new(
                            blockResults
                                .Where(p => !string.IsNullOrWhiteSpace(p.Name))
                                .Select(p => p.Name!)
                                .Distinct(StringComparer.OrdinalIgnoreCase),
                            StringComparer.OrdinalIgnoreCase);

                        gnaSpreadsheetAPI.UpdateLastRetrievedTimeByPoint(
                            strExcelWorkbookFullPath,
                            strReferenceWorksheet,
                            strFirstDataRow,
                            pointsWithData,
                            blockEndUTC);

                        foreach (var p in pointMasterList)
                        {
                            if (p.Name != null && pointsWithData.Contains(p.Name))
                                p.TimeBlockEndUTC = blockEndUTC;
                        }

                        // Option A: generate CSV from blockResults only
                        string generatedCsvPath = gnaT.generateCoordinateCSVfile(
                            blockResults,
                            strFTPSubdirectory,
                            strContractTitle,
                            blockEndUTC,
                            CSVformat,
                            CoordinateOrder,
                            includeHeader,
                            OutputFileExtension,
                            CSVseparator,
                            4);

                        // Enforce deterministic name (rename/move)
                        try
                        {
                            if (!string.Equals(generatedCsvPath, expectedCsvPath, StringComparison.OrdinalIgnoreCase))
                            {
                                if (File.Exists(generatedCsvPath))
                                {
                                    Directory.CreateDirectory(Path.GetDirectoryName(expectedCsvPath)!);

                                    if (File.Exists(expectedCsvPath))
                                        File.Delete(expectedCsvPath);

                                    File.Move(generatedCsvPath, expectedCsvPath);
                                }
                            }
                        }
                        catch (Exception rx)
                        {
                            gnaT.updateSystemLogFile(strSystemLogsFolder, $"Rename failed: '{generatedCsvPath}' -> '{expectedCsvPath}' | {rx.Message}");
                        }

                        string finalPathToReport = File.Exists(expectedCsvPath) ? expectedCsvPath : generatedCsvPath;

                        string strMessage = $"Generated coordinate CSV file: {finalPathToReport}";
                        gnaT.updateSystemLogFile(strSystemLogsFolder, strMessage);
                        Console.WriteLine($"{strTab2}CSV created:\n{strTab3}{finalPathToReport}");
                        Console.Out.Flush();
                    }
                    exitCode = 0;
                    #endregion
                }
                #endregion
            }
            catch (Exception ex)
            {
                exitCode = 1;

                try { File.WriteAllText("fatal_crash.log", ex.ToString()); } catch { }

                try
                {
                    Console.WriteLine("Fatal crash:");
                    Console.WriteLine(ex);
                    Console.Out.Flush();
                }
                catch { }

                try
                {
                    if (gnaT != null)
                        gnaT.updateSystemLogFile(strSystemLogsFolder, "Fatal crash: " + ex);
                }
                catch { }
            }
            finally
            {
                try
                {
                    if (gnaT != null)
                    {
                        gnaT.updateSystemLogFile(strSystemLogsFolder, "Run end | ExitCode=" + exitCode.ToString(CultureInfo.InvariantCulture));
                        Console.WriteLine("\nGNAcoordinateExporter export completed...\n\n");
                        gnaT.freezeScreen(strFreezeScreen);
                    }
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

        static bool IsYes(string s) => string.Equals(CleanConfig(s), "Yes", StringComparison.OrdinalIgnoreCase);
        #endregion

        #region General helpers
        private static string BuildManualEmailTime(string manualBlockEnd)
        {
            if (string.IsNullOrWhiteSpace(manualBlockEnd))
                return string.Empty;

            string tmp = manualBlockEnd.Replace("-", "")
                                       .Replace(" ", "_")
                                       .Replace(":", "h") + "m";
            return tmp.Length >= 14 ? tmp.Substring(0, 14) : tmp;
        }

        private static string SanitizeForFilename(string s)
        {
            s = (s ?? string.Empty).Trim();
            if (s.Length == 0) return "empty";

            foreach (char c in Path.GetInvalidFileNameChars())
                s = s.Replace(c, '_');

            s = s.Replace(" ", "_").Replace(":", "-");
            return s;
        }

        private static string FormatUtcForFilename(string utcTimestamp)
        {
            if (!DateTime.TryParseExact(
                    utcTimestamp,
                    new[] { "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd HH:mm" },
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                    out DateTime dt))
            {
                throw new FormatException($"Invalid UTC timestamp format: '{utcTimestamp}'");
            }

            return dt.ToString("yyyyMMdd_HHmm", CultureInfo.InvariantCulture)
                     .Insert(11, "h");
        }
        #endregion
    }
}
