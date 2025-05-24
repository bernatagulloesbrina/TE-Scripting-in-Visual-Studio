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

        public static ReportExtended InitReport()
        {
            // Get the base path from the user  
            string basePath = Rx.GetPbirFilePath();
            if (basePath == null)
            {
                Error("Operation canceled by the user.");
                return null;
            }

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

            ReportExtended report = new ReportExtended();
            report.PagesFilePath = Path.Combine(targetPath, "pages.json");

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

                        string visualsPath = Path.Combine(folder, "visuals");
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

        public static VisualExtended SelectVisual(ReportExtended report)
        {
            // Step 1: Build selection list
            var visualSelectionList = report.Pages
                .SelectMany(p => p.Visuals.Select(v => new
                {
                    Display = string.Format("{0} - {1} ({2}, {3})", p.Page.DisplayName, v.Content.Visual.VisualType, (int)v.Content.Position.X, (int)v.Content.Position.Y),
                    Page = p,
                    Visual = v
                }))
                .ToList();

            // Step 2: Let user choose a visual
            var options = visualSelectionList.Select(v => v.Display).ToList();
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

        public static void SaveVisual(VisualExtended visual)
        {

            // Save new JSON, ignoring nulls
            string newJson = JsonConvert.SerializeObject(
                visual.Content,
                Newtonsoft.Json.Formatting.Indented,
                new JsonSerializerSettings
                {
                    DefaultValueHandling = DefaultValueHandling.Ignore,
                    NullValueHandling = NullValueHandling.Ignore

                }
            );
            File.WriteAllText(visual.VisualFilePath, newJson);
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


        public static String GetPbirFilePath()
        {

            // Create an instance of the OpenFileDialog
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Please select definition.pbir file of the target report",
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
