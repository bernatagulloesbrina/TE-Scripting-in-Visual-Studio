using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TabularEditor.TOMWrapper;
using TabularEditor.Scripting;
using System.Reflection.Emit;
using Microsoft.VisualBasic;
using System.Windows.Forms;
using Report.DTO;
using System.IO;
using Newtonsoft.Json;
using GeneralFunctions;
//using static Report.DTO.VisualDto;
using System.Xml;

namespace ReportFunctions
{

    //copy from the following line up to ****** and remove the // before the closing bracket
    //after the class declaration add all the #r and using statements necessary for the custom class code to run in Tabular Editor
    //these directives will be combined with the ones from the macro when using the CopyMacro script
    public static class Rx
    {


        // NOCOPY    in TE2 (at least up to 2.17.2) any method that accesses or modifies
        // NOCOPY    the model needs a reference to the model 
        // NOCOPY    the following is an example method where you can build extra logic
        
        
        public static VisualExtended DuplicateVisual(VisualExtended visualExtended)
        {
            // Generate a clean 16-character name from a GUID (no dashes or slashes)
            string newVisualName = Guid.NewGuid().ToString("N").Substring(0, 16);
            string sourceFolder = Path.GetDirectoryName(visualExtended.VisualFilePath);
            string targetFolder = Path.Combine(Path.GetDirectoryName(sourceFolder), newVisualName);
            if (Directory.Exists(targetFolder))
            {
                Error(string.Format("Folder already exists: {0}", targetFolder));
                return null;
            }
            Directory.CreateDirectory(targetFolder);

            // Deep clone the VisualDto.Root object
            string originalJson = JsonConvert.SerializeObject(visualExtended.Content, Newtonsoft.Json.Formatting.Indented);
            VisualDto.Root clonedContent = 
                JsonConvert.DeserializeObject<VisualDto.Root>(
                    originalJson, 
                    new JsonSerializerSettings {
                        DefaultValueHandling = DefaultValueHandling.Ignore,
                        NullValueHandling = NullValueHandling.Ignore

                    });

            // Update the name property if it exists
            if (clonedContent != null && clonedContent.Name != null)
            {
                clonedContent.Name = newVisualName;
            }

            // Set the new file path
            string newVisualFilePath = Path.Combine(targetFolder, "visual.json");

            // Create the new VisualExtended object
            VisualExtended newVisual = new VisualExtended
            {
                Content = clonedContent,
                VisualFilePath = newVisualFilePath
            };

            return newVisual;
        }

        public static VisualExtended GroupVisuals(List<VisualExtended> visualsToGroup, string groupName = null, string groupDisplayName = null)
        {
            if (visualsToGroup == null || visualsToGroup.Count == 0)
            {
                Error("No visuals to group.");
                return null;
            }
            // Generate a clean 16-character name from a GUID (no dashes or slashes) if no group name is provided
            if (string.IsNullOrEmpty(groupName))
            {
                groupName = Guid.NewGuid().ToString("N").Substring(0, 16);
            }
            if (string.IsNullOrEmpty(groupDisplayName))
            {
                groupDisplayName = groupName;
            }

            // Find minimum X and Y
            double minX = visualsToGroup.Min(v => v.Content.Position != null ? (double)v.Content.Position.X : 0);
            double minY = visualsToGroup.Min(v => v.Content.Position != null ? (double)v.Content.Position.Y : 0);

           //Info("minX:" + minX.ToString() + ", minY: " + minY.ToString());

            // Calculate width and height
            double groupWidth = 0;
            double groupHeight = 0;
            foreach (var v in visualsToGroup)
            {
                if (v.Content != null && v.Content.Position != null)
                {
                    double visualWidth = v.Content.Position != null ? (double)v.Content.Position.Width : 0;
                    double visualHeight = v.Content.Position != null ? (double)v.Content.Position.Height : 0;
                    double xOffset = (double)v.Content.Position.X - (double)minX;
                    double yOffset = (double)v.Content.Position.Y - (double)minY;
                    double totalWidth = xOffset + visualWidth;
                    double totalHeight = yOffset + visualHeight;
                    if (totalWidth > groupWidth) groupWidth = totalWidth;
                    if (totalHeight > groupHeight) groupHeight = totalHeight;
                }
            }

            // Create the group visual content
            var groupContent = new VisualDto.Root
            {
                Schema = visualsToGroup.FirstOrDefault().Content.Schema,
                Name = groupName,
                Position = new VisualDto.Position
                {
                    X = minX,
                    Y = minY,
                    Width = groupWidth,
                    Height = groupHeight
                },
                VisualGroup = new VisualDto.VisualGroup
                {
                    DisplayName = groupDisplayName,
                    GroupMode = "ScaleMode"
                }
            };

            // Set VisualFilePath for the group visual
            // Use the VisualFilePath of the first visual as a template
            string groupVisualFilePath = null;
            var firstVisual = visualsToGroup.FirstOrDefault(v => !string.IsNullOrEmpty(v.VisualFilePath));
            if (firstVisual != null && !string.IsNullOrEmpty(firstVisual.VisualFilePath))
            {
                string originalPath = firstVisual.VisualFilePath;
                string parentDir = Path.GetDirectoryName(Path.GetDirectoryName(originalPath)); // up to 'visuals'
                if (!string.IsNullOrEmpty(parentDir))
                {
                    string groupFolder = Path.Combine(parentDir, groupName);
                    groupVisualFilePath = Path.Combine(groupFolder, "visual.json");
                }
            }

            // Create the new VisualExtended for the group
            var groupVisual = new VisualExtended
            {
                Content = groupContent,
                VisualFilePath = groupVisualFilePath // Set as described
            };

            // Update grouped visuals: set parentGroupName and adjust X/Y
            foreach (var v in visualsToGroup)
            {
                
                if (v.Content == null) continue;
                v.Content.ParentGroupName = groupName;

                if (v.Content.Position != null)
                {
                    v.Content.Position.X = v.Content.Position.X - minX + 0;
                    v.Content.Position.Y = v.Content.Position.Y - minY + 0;
                }
            }

            return groupVisual;
        }

        

        private static readonly string RecentPathsFile = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "Tabular Editor Macro Settings", "recentPbirPaths.json");

        public static string GetPbirFilePathWithHistory(string label = "Select definition.pbir file")
        {
            // Detect reports currently open in Power BI Desktop (PBIR/PBIP format, via pbir CLI)
            PbirDesktopStatus pbirStatus;
            string pbirMessage;
            List<string> openReports = GetOpenPbirReportPaths(out pbirStatus, out pbirMessage);

            // Inform the user when pbir is missing or the Desktop preview feature is off
            if (!string.IsNullOrEmpty(pbirMessage))
            {
                Info(pbirMessage);
            }

            // Load recent paths
            List<string> recentPaths = LoadRecentPbirPaths();

            // Filter out non-existing files
            recentPaths = recentPaths.Where(File.Exists).ToList();

            // If exactly one report is open in Desktop, use it directly (skip manual selection)
            if (openReports.Count == 1)
            {
                string autoPath = openReports[0];
                UpdateRecentPbirPaths(autoPath, recentPaths);
                return autoPath;
            }

            // More than one PBIP report is open => the right target is ambiguous. Warn the user
            // so they consciously pick the one matching the model they are working on.
            if (openReports.Count > 1)
            {
                Info(MultipleInstancesMessage(openReports.Count));
            }

            // Present options to the user: currently-open reports first (clearly marked),
            // then recent paths, then browse. Several Desktop instances => several open reports.
            const string openPrefix = "\u25CF [Open in Power BI Desktop]  ";
            var displayToPath = new Dictionary<string, string>();
            var options = new List<string>();

            foreach (string p in openReports)
            {
                string display = openPrefix + p;
                if (!displayToPath.ContainsKey(display))
                {
                    displayToPath[display] = p;
                    options.Add(display);
                }
            }
            foreach (string p in recentPaths)
            {
                if (!options.Contains(p) && !displayToPath.ContainsValue(p)) options.Add(p);
            }
            options.Add("Browse for new file...");

            string selected = Fx.ChooseString(options,label:label, customWidth:600, customHeight:300);

            if (selected == null) return null;

            string chosenPath = null;
            if (selected == "Browse for new file..." )
            {
                chosenPath = GetPbirFilePath(label);
            }
            else if (displayToPath.ContainsKey(selected))
            {
                chosenPath = displayToPath[selected];
            }
            else
            {
                chosenPath = selected;
            }

            if (!string.IsNullOrEmpty(chosenPath))
            {
                // Update recent paths
                UpdateRecentPbirPaths(chosenPath, recentPaths);
            }

            return chosenPath;
        }

        private static List<string> LoadRecentPbirPaths()
        {
            try
            {
                if (File.Exists(RecentPathsFile))
                {
                    string json = File.ReadAllText(RecentPathsFile);
                    return JsonConvert.DeserializeObject<List<string>>(json) ?? new List<string>();
                }
            }
            catch { }
            return new List<string>();
        }

        private static void UpdateRecentPbirPaths(string newPath, List<string> recentPaths)
        {
            // Remove if already exists, insert at top
            recentPaths.RemoveAll(p => string.Equals(p, newPath, StringComparison.OrdinalIgnoreCase));
            recentPaths.Insert(0, newPath);

            // Keep only the latest 10
            while (recentPaths.Count > 10)
                recentPaths.RemoveAt(recentPaths.Count - 1);

            // Ensure directory exists
            Directory.CreateDirectory(Path.GetDirectoryName(RecentPathsFile));
            File.WriteAllText(RecentPathsFile, JsonConvert.SerializeObject(recentPaths, Newtonsoft.Json.Formatting.Indented));
        }

        // Outcome of probing Power BI Desktop through the pbir CLI.
        public enum PbirDesktopStatus
        {
            Ready,           // pbir ran and the local API responded (feature enabled)
            NotInstalled,    // the pbir executable could not be found
            FeatureDisabled, // pbir ran but the Desktop local-API preview feature is off / unavailable
            ThickReportOnly, // pbir ran but the only open report(s) are thick (.pbix), not PBIR
            Error            // any other failure (could not read or parse the response)
        }

        // Captures the result of launching an external process.
        private sealed class PbirProcResult
        {
            public bool Started;   // false => executable not found / could not start
            public string StdOut;
            public string StdErr;
        }

        // Returns the definition.pbir path(s) of the report(s) currently open in Power BI Desktop.
        // Uses the 'pbir' CLI (https://github.com/Kurt-Buhler/pbir). Returns an empty list if the
        // CLI is unavailable, the Desktop preview feature is off, no Desktop instance is open, or
        // only thick (.pbix) reports are open.
        public static List<string> GetOpenPbirReportPaths()
        {
            PbirDesktopStatus status;
            string message;
            return GetOpenPbirReportPaths(out status, out message);
        }

        // Overload that also reports why detection succeeded or failed via 'status' and a
        // user-facing 'message' (null when no message is needed, e.g. detection worked).
        public static List<string> GetOpenPbirReportPaths(out PbirDesktopStatus status, out string message)
        {
            var result = new List<string>();
            status = PbirDesktopStatus.Error;
            message = null;

            // 1) Check that the pbir CLI is installed
            PbirProcResult run = RunPbir("desktop list --json");
            if (run == null || !run.Started)
            {
                status = PbirDesktopStatus.NotInstalled;
                message =
                    "Auto-detection of the open report is unavailable because the 'pbir' CLI was not found.\n\n" +
                    "Install it with:  uv tool install pbir-cli\n\n" +
                    "Falling back to manual selection.";
                return result;
            }

            // 2) Check that Power BI Desktop's local API responded (preview feature enabled)
            string output = run.StdOut;
            if (!LooksLikeJson(output))
            {
                string combined = ((output ?? "") + " " + (run.StdErr ?? "")).ToLowerInvariant();
                if (combined.Contains("secure local") || combined.Contains("external tool") ||
                    combined.Contains("preview") || combined.Contains("not enabled"))
                {
                    status = PbirDesktopStatus.FeatureDisabled;
                    message = FeatureDisabledMessage();
                }
                else
                {
                    status = PbirDesktopStatus.Error;
                    message = "Could not read Power BI Desktop status from pbir. Falling back to manual selection.";
                }
                return result;
            }

            try
            {
                var root = Newtonsoft.Json.Linq.JObject.Parse(output);
                string apiStatus = (string)root["status"];

                if (!string.IsNullOrEmpty(apiStatus) &&
                    !apiStatus.Equals("ready", StringComparison.OrdinalIgnoreCase))
                {
                    status = PbirDesktopStatus.FeatureDisabled;
                    message = FeatureDisabledMessage();
                    return result;
                }

                status = PbirDesktopStatus.Ready;

                var instances = root["instances"] as Newtonsoft.Json.Linq.JArray;
                if (instances != null)
                {
                    foreach (var inst in instances)
                    {
                        string reportDir = (string)inst["reportDir"];
                        string currentFilePath = (string)inst["currentFilePath"];
                        string pbir = DerivePbirPath(reportDir, currentFilePath);
                        if (pbir != null && !result.Contains(pbir)) result.Add(pbir);
                    }
                }

                // pbir's 'desktop list' only surfaces PBIP reports; thick (.pbix) reports never
                // appear there. Detect them by comparing the number of open Power BI Desktop
                // report windows against the number of PBIP reports pbir reported.
                int openWindowCount = CountOpenDesktopReportWindows();
                int thickCount = openWindowCount - result.Count;
                if (thickCount < 0) thickCount = 0;

                if (result.Count == 0 && thickCount > 0)
                {
                    // Only thick (.pbix) report(s) are open - none can be used
                    status = PbirDesktopStatus.ThickReportOnly;
                    message = ThickReportMessage();
                }
                else if (result.Count > 0 && thickCount > 0)
                {
                    // A mix of PBIR and thick (.pbix) reports is open - warn that .pbix are ignored
                    message = MixedReportsMessage(thickCount);
                }
            }
            catch
            {
                status = PbirDesktopStatus.Error;
                message = "Could not parse Power BI Desktop status from pbir. Falling back to manual selection.";
            }

            return result;
        }

        private static string FeatureDisabledMessage()
        {
            return
                "Auto-detection of the open report is unavailable because Power BI Desktop's local API did not respond.\n\n" +
                "Please make sure that:\n" +
                "  1. Power BI Desktop is running with the report open, and\n" +
                "  2. The preview feature is enabled:\n" +
                "       File > Options and settings > Options > Preview features >\n" +
                "       'Enable external tool access to Power BI Desktop through secure local APIs'\n\n" +
                "Falling back to manual selection.";
        }

        private static string ThickReportMessage()
        {
            return
                "The report currently open in Power BI Desktop is a thick report (.pbix), " +
                "so it cannot be auto-detected or edited by the report-layer scripts.\n\n" +
                "These scripts require the PBIR (enhanced report format) on disk. To use them:\n" +
                "  1. In Power BI Desktop: File > Save as > Power BI project (.pbip), and\n" +
                "  2. Enable PBIR under File > Options > Preview features > 'Store reports using enhanced metadata format (PBIR)'.\n\n" +
                "Falling back to manual selection.";
        }

        private static string MixedReportsMessage(int thickCount)
        {
            string plural = thickCount == 1 ? "report is" : "reports are";
            return String.Format(
                "Note: {0} thick (.pbix) {1} also open in Power BI Desktop.\n\n" +
                "Thick reports cannot be edited by the report-layer scripts, so only the " +
                "PBIR (.pbip) report(s) are being considered.",
                thickCount, plural);
        }

        private static string MultipleInstancesMessage(int openCount)
        {
            return String.Format(
                "Warning: {0} Power BI Desktop reports (PBIP) are currently open.\n\n" +
                "Auto-detection cannot tell which one matches the model you are working on, " +
                "so please pick the intended report from the list. Reports currently open in " +
                "Power BI Desktop are marked with a bullet at the top.",
                openCount);
        }

        // Counts the report windows currently open in Power BI Desktop. Each open report is a
        // PBIDesktop.exe process with a non-empty main window title; thick (.pbix) reports are
        // counted here even though pbir's 'desktop list' does not report them.
        private static int CountOpenDesktopReportWindows()
        {
            try
            {
                return System.Diagnostics.Process.GetProcessesByName("PBIDesktop")
                    .Select(p => { try { return p.MainWindowTitle; } catch { return null; } })
                    .Where(t => !string.IsNullOrEmpty(t))
                    .Distinct()
                    .Count();
            }
            catch
            {
                return 0;
            }
        }
        // location (~/.local/bin/pbir.exe). Started=false means pbir is not installed.
        private static PbirProcResult RunPbir(string arguments)
        {
            PbirProcResult r = RunCommandCapture("pbir", arguments);
            if (r.Started) return r;

            string fallback = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                ".local", "bin", "pbir.exe");
            if (File.Exists(fallback))
            {
                return RunCommandCapture(fallback, arguments);
            }

            return r; // Started = false
        }

        private static PbirProcResult RunCommandCapture(string fileName, string arguments)
        {
            var r = new PbirProcResult();
            try
            {
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (var proc = System.Diagnostics.Process.Start(psi))
                {
                    r.Started = true;
                    r.StdOut = proc.StandardOutput.ReadToEnd();
                    r.StdErr = proc.StandardError.ReadToEnd();
                    proc.WaitForExit(15000);
                }
            }
            catch (System.ComponentModel.Win32Exception)
            {
                r.Started = false; // executable not found
            }
            catch
            {
                r.Started = false;
            }
            return r;
        }

        private static bool LooksLikeJson(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            string t = s.TrimStart();
            return t.StartsWith("{") || t.StartsWith("[");
        }

        // Maps a pbir instance (reportDir / currentFilePath) to an existing definition.pbir path.
        private static string DerivePbirPath(string reportDir, string currentFilePath)
        {
            // Case 1: pbir already resolved the report folder
            if (!string.IsNullOrEmpty(reportDir))
            {
                if (reportDir.EndsWith(".pbir", StringComparison.OrdinalIgnoreCase) && File.Exists(reportDir))
                    return reportDir;

                string candidate = Path.Combine(reportDir, "definition.pbir");
                if (File.Exists(candidate)) return candidate;
            }

            // Case 2: derive the report folder from an open .pbip project
            if (!string.IsNullOrEmpty(currentFilePath) &&
                currentFilePath.EndsWith(".pbip", StringComparison.OrdinalIgnoreCase))
            {
                string dir = Path.GetDirectoryName(currentFilePath);
                string name = Path.GetFileNameWithoutExtension(currentFilePath);

                string candidate = Path.Combine(dir, name + ".Report", "definition.pbir");
                if (File.Exists(candidate)) return candidate;

                try
                {
                    foreach (string reportFolder in Directory.GetDirectories(dir, "*.Report"))
                    {
                        string c = Path.Combine(reportFolder, "definition.pbir");
                        if (File.Exists(c)) return c;
                    }
                }
                catch { }
            }

            return null;
        }


        public static ReportExtended InitReport(string label = "Please select definition.pbir file of the target report")
        {
            // Get the base path from the user  
            string basePath = Rx.GetPbirFilePathWithHistory(label:label);
            if (basePath == null) return null; 
            
            // Define the target path  
            string baseDirectory = Path.GetDirectoryName(basePath);
            string targetPath = Path.Combine(baseDirectory, "definition", "pages");

            // Check if the target path exists  
            if (!Directory.Exists(targetPath))
            {
                Error(String.Format("The path '{0}' does not exist.", targetPath));
                return null;
            }

            // Get all subfolders in the target path  
            List<string> subfolders = Directory.GetDirectories(targetPath).ToList();

            string pagesFilePath = Path.Combine(targetPath, "pages.json");
            string pagesJsonContent = File.ReadAllText(pagesFilePath);
            
            if (string.IsNullOrEmpty(pagesJsonContent))
            {
                Error(String.Format("The file '{0}' is empty or does not exist.", pagesFilePath));
                return null;
            }

            PagesDto pagesDto = JsonConvert.DeserializeObject<PagesDto>(pagesJsonContent);

            ReportExtended report = new ReportExtended();
            report.PagesFilePath = pagesFilePath;
            report.PagesConfig = pagesDto;

            // Process each folder  
            foreach (string folder in subfolders)
            {
                string pageJsonPath = Path.Combine(folder, "page.json");
                if (File.Exists(pageJsonPath))
                {
                    try
                    {
                        string jsonContent = File.ReadAllText(pageJsonPath);
                        PageDto page = JsonConvert.DeserializeObject<PageDto>(jsonContent);

                        PageExtended pageExtended = new PageExtended();
                        pageExtended.Page = page;
                        pageExtended.PageFilePath = pageJsonPath;

                        pageExtended.ParentReport = report;

                        string visualsPath = Path.Combine(folder, "visuals");

                        if (!Directory.Exists(visualsPath))
                        {
                            report.Pages.Add(pageExtended); // still add the page
                            continue; // skip visual loading
                        }

                        List<string> visualSubfolders = Directory.GetDirectories(visualsPath).ToList();

                        foreach (string visualFolder in visualSubfolders)
                        {
                            string visualJsonPath = Path.Combine(visualFolder, "visual.json");
                            if (File.Exists(visualJsonPath))
                            {
                                try
                                {
                                    string visualJsonContent = File.ReadAllText(visualJsonPath);
                                    VisualDto.Root visual = JsonConvert.DeserializeObject<VisualDto.Root>(visualJsonContent);

                                    VisualExtended visualExtended = new VisualExtended();
                                    visualExtended.Content = visual;
                                    visualExtended.VisualFilePath = visualJsonPath;
                                    visualExtended.ParentPage = pageExtended; // Set parent page reference
                                    pageExtended.Visuals.Add(visualExtended);
                                }
                                catch (Exception ex2)
                                {
                                    Output(String.Format("Error reading or deserializing '{0}': {1}", visualJsonPath, ex2.Message));
                                    return null;
                                }

                            }
                        }

                        report.Pages.Add(pageExtended);

                    }
                    catch (Exception ex)
                    {
                        Output(String.Format("Error reading or deserializing '{0}': {1}", pageJsonPath, ex.Message));
                    }
                }

            }
            return report;
        }


        public static VisualExtended SelectTableVisual(ReportExtended report)
        {
            List<string> visualTypes = new List<string>
            {
                "tableEx","pivotTable"
            };
            return SelectVisual(report: report, visualTypes);
        }



        public static VisualExtended SelectVisual(ReportExtended report, List<string> visualTypeList = null)
        {
            return SelectVisualInternal(report, Multiselect: false, visualTypeList:visualTypeList) as VisualExtended;
        }

        public static List<VisualExtended> SelectVisuals(ReportExtended report, List<string> visualTypeList = null)
        {
            return SelectVisualInternal(report, Multiselect: true, visualTypeList:visualTypeList) as List<VisualExtended>;
        }

        private static object SelectVisualInternal(ReportExtended report, bool Multiselect, List<string> visualTypeList = null)
        {
            // Step 1: Build selection list
            var visualSelectionList = 
                report.Pages
                .SelectMany(p => p.Visuals
                    .Where(v =>
                        v?.Content != null &&
                        (
                            // If visualTypeList is null, do not filter at all
                            (visualTypeList == null) ||
                            // If visualTypeList is provided and not empty, filter by it
                            (visualTypeList.Count > 0 && v.Content.Visual != null && visualTypeList.Contains(v.Content?.Visual?.VisualType))
                            // Otherwise, include all visuals and visual groups
                            || (visualTypeList.Count == 0)
                        )
                    )
                    .Select(v => new
                        {
                            // Use visual type for regular visuals, displayname for groups
                            Display = string.Format(
                                "{0} - {1} ({2}, {3})",
                                p.Page.DisplayName,
                                v?.Content?.Visual?.VisualType
                                    ?? v?.Content?.VisualGroup?.DisplayName,
                                (int)(v.Content.Position?.X ?? 0),
                                (int)(v.Content.Position?.Y ?? 0)
                            ),
                            Page = p,
                            Visual = v
                        }
                    )
                )
                .ToList();

            if (visualSelectionList.Count == 0)
            {
                if (visualTypeList != null)
                {
                    string types = string.Join(", ", visualTypeList);
                    Error(string.Format("No visual of type {0} were found", types));

                }else
                {
                    Error("No visuals found in the report.");
                }


                return null;
            }

            // Step 2: Let user choose a visual
            var options = visualSelectionList.Select(v => v.Display).ToList();

            if (Multiselect)
            {
                // For multiselect, use ChooseStringMultiple
                var multiSelelected = Fx.ChooseStringMultiple(options);
                if (multiSelelected == null || multiSelelected.Count == 0)
                {
                    Info("You cancelled.");
                    return null;
                }
                // Find all selected visuals
                var selectedVisuals = visualSelectionList.Where(v => multiSelelected.Contains(v.Display)).Select(v => v.Visual).ToList();

                return selectedVisuals;
            }
            else
            {
                string selected = Fx.ChooseString(options);

                if (string.IsNullOrEmpty(selected))
                {
                    Info("You cancelled.");
                    return null;
                }

                // Step 3: Find the selected visual
                var selectedVisual = visualSelectionList.FirstOrDefault(v => v.Display == selected);

                if (selectedVisual == null)
                {
                    Error("Selected visual not found.");
                    return null;
                }

                return selectedVisual.Visual;
            }
        }

        public static PageExtended ReplicateFirstPageAsBlank(ReportExtended report, bool showMessages = false)
        {
            if (report.Pages == null || !report.Pages.Any())
            {
                Error("No pages found in the report.");
                return null;
            }

            PageExtended firstPage = report.Pages[0];

            // Generate a clean 16-character name from a GUID (no dashes or slashes)
            string newPageName = Guid.NewGuid().ToString("N").Substring(0, 16);
            string newPageDisplayName = firstPage.Page.DisplayName + " - Copy";

            string sourceFolder = Path.GetDirectoryName(firstPage.PageFilePath);
            string targetFolder = Path.Combine(Path.GetDirectoryName(sourceFolder), newPageName);
            string visualsFolder = Path.Combine(targetFolder, "visuals");

            if (Directory.Exists(targetFolder))
            {
                Error($"Folder already exists: {targetFolder}");
                return null;
            }

            Directory.CreateDirectory(targetFolder);
            Directory.CreateDirectory(visualsFolder);

            var newPageDto = new PageDto
            {
                Name = newPageName,
                DisplayName = newPageDisplayName,
                DisplayOption = firstPage.Page.DisplayOption,
                Height = firstPage.Page.Height,
                Width = firstPage.Page.Width,
                Schema = firstPage.Page.Schema
            };

            var newPage = new PageExtended
            {
                Page = newPageDto,
                PageFilePath = Path.Combine(targetFolder, "page.json"),
                Visuals = new List<VisualExtended>() // empty visuals
            };

            File.WriteAllText(newPage.PageFilePath, JsonConvert.SerializeObject(newPageDto, Newtonsoft.Json.Formatting.Indented));

            report.Pages.Add(newPage);

            if(showMessages) Info($"Created new blank page: {newPageName}");

            return newPage; 
        }


        public static void SaveVisual(VisualExtended visual)
        {

            // Save new JSON, ignoring nulls
            string newJson = JsonConvert.SerializeObject(
                visual.Content,
                Newtonsoft.Json.Formatting.Indented,
                new JsonSerializerSettings
                {
                    //DefaultValueHandling = DefaultValueHandling.Ignore,
                    NullValueHandling = NullValueHandling.Ignore

                }
            );
            // Ensure the directory exists before saving
            string visualFolder = Path.GetDirectoryName(visual.VisualFilePath);
            if (!Directory.Exists(visualFolder))
            {
                Directory.CreateDirectory(visualFolder);
            }
            File.WriteAllText(visual.VisualFilePath, newJson);
        }

        // Pushes the just-saved on-disk report changes into the open Power BI Desktop instance via the
        // pbir CLI (https://github.com/Kurt-Buhler/pbir), so report-layer scripts reflect their edits on
        // the canvas without a manual reopen. The report is assumed valid (it was already initialised).
        public static void UpdateDesktop(ReportExtended report)
        {
            // PagesFilePath is <Name>.Report/definition/pages/pages.json; walk up three levels to reach
            // the report folder (<Name>.Report), which is what 'pbir desktop reload' expects as its path.
            string reportDir =
                Path.GetDirectoryName(
                    Path.GetDirectoryName(
                        Path.GetDirectoryName(report.PagesFilePath)));

            PbirProcResult run = RunPbir(String.Format("desktop reload \"{0}\"", reportDir));

            // Note: do NOT show a modal dialog (Info) here. Reloading the Power BI Desktop instance
            // that Tabular Editor is live-connected to queues a "connection changed" event. A modal
            // dialog pumps the message loop, letting TE process that event mid-macro, which resets
            // TE's undo manager and causes "EndBatch() called before BeginBatch()" when the macro
            // ends. Writing to the output pane avoids the nested message loop.
            if (run != null && run.Started)
            {
                Output("Power BI Desktop reloaded the report canvas from disk.");
            }
            else
            {
                Output("Changes saved to disk. Reopen the report in Power BI Desktop to see them " +
                    "(the 'pbir' CLI was not found for automatic reload; install it with 'uv tool install pbir-cli').");
            }
        }


        public static string ReplacePlaceholders(string pageContents, Dictionary<string, string> placeholders)
        {
            if (placeholders != null)
            {
                foreach (string placeholder in placeholders.Keys)
                {
                    string valueToReplace = placeholders[placeholder];

                    pageContents = pageContents.Replace(placeholder, valueToReplace);

                }
            }


            return pageContents;
        }


        public static String GetPbirFilePath(string label = "Please select definition.pbir file of the target report")
        {

            // Create an instance of the OpenFileDialog
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = label,
                // Set filter options and filter index.
                Filter = "PBIR Files (*.pbir)|*.pbir",
                FilterIndex = 1
            };
            // Call the ShowDialog method to show the dialog box.
            DialogResult result = openFileDialog.ShowDialog();
            // Process input if the user clicked OK.
            if (result != DialogResult.OK)
            {
                Error("You cancelled");
                return null;
            }
            return openFileDialog.FileName;

        }

        // NOCOPY add other methods always as "public static" followed by the data type they will return or void if they do not return anything.

        //}

        //******************
        //do not copy from this line below, and remove the // before the closing bracket above to close the class definition


        //Model and Selected cannot be accessed directly. Always pass a reference to the requited objects. 
        //static readonly Model Model;
        //static readonly TabularEditor.UI.UITreeSelection Selected;


        //These functions replicate the ScriptHelper functions so that they can be
        //used inside the script without the ScriptHelper prefix which cannot be used inside tabular editor
        //the list is not complete and does not include all overloads, complete as necessary. 
        public static void Error(string message, int lineNumber = -1, bool suppressHeader = false)
        {
            ScriptHelper.Error(message: message, lineNumber: lineNumber, suppressHeader: suppressHeader);
        }

        public static void Info(string message, int lineNumber = -1)
        {
            ScriptHelper.Info(message: message, lineNumber: lineNumber);
        }

        public static Table SelectTable(IEnumerable<Table> tables, Table preselect = null, string label = "Select Table")
        {
            return ScriptHelper.SelectTable(tables: tables, preselect: preselect, label: label);
        }
        public static Measure SelectMeasure(Measure preselect = null, string label = "Select Measure")
        {
            return ScriptHelper.SelectMeasure(preselect: preselect, label: label);
        }
        public static void Output(object value, int lineNumber = -1)
        {
            ScriptHelper.Output(value: value, lineNumber: lineNumber);
        }
    }
}
