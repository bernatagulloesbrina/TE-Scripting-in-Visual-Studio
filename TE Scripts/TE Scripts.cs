using System;
using System.Collections.Generic;
using System.Linq;
using TabularEditor.TOMWrapper;
using TabularEditor.Scripting;
using System.Windows.Forms;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.CSharp;

using System.IO;
using GeneralFunctions; //Uncomment if you use the custom class, add reference to the project too.
using ReportFunctions;
using Report.DTO;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;



// '2023-05-06 / B.Agullo / 
//coding environment for Tabular Editor C# Scripts
// see https://www.esbrina-ba.com/c-scripting-nirvana-effortlessly-use-visual-studio-as-your-coding-environment/ for reference on how to use it.

namespace TE_Scripting
{
    public class TE_Scripts
    {

        //using GeneralFunctions;
        //using Report.DTO;
        //using System.IO;
        //using Newtonsoft.Json.Linq;

        //using GeneralFunctions; //using Report.DTO; //using System.IO;//using Newtonsoft.Json.Linq;
        void openVisualJsonFile()
        {
            //using GeneralFunctions;
            //using Report.DTO;
            //using System.IO;
            //using Newtonsoft.Json.Linq;

            //2025-05-25/B.Agullo
            //this script allows the user to open the JSON file of one or more visuals in the report.
            //see https://www.esbrina-ba.com/pbir-scripts-to-replace-field-and-open-visual-json-files/ for reference on how to use it

            // Step 1: Initialize the report object
            ReportExtended report = Rx.InitReport();
            if (report == null) return;

            // Step 2: Gather all visuals with page info
            var allVisuals = report.Pages
                .SelectMany(p => p.Visuals.Select(v => new { Page = p.Page, Visual = v }))
                .ToList();

            if (allVisuals.Count == 0)
            {
                Info("No visuals found in the report.");
                return;
            }

            // Step 3: Prepare display names for selection
            var visualDisplayList = allVisuals.Select(x =>
                String.Format(@"{0} - {1} ({2}, {3})", x.Page.DisplayName, x.Visual.Content.Visual.VisualType, (int)x.Visual.Content.Position.X, (int)x.Visual.Content.Position.Y)
            ).ToList();

            // Step 4: Let the user select one or more visuals
            List<string> selected = Fx.ChooseStringMultiple(OptionList: visualDisplayList, Label: "Select visuals to open JSON files");
            if (selected == null || selected.Count == 0)
            {
                Info("No visuals selected.");
                return;
            }

            // Step 5: For each selected visual, open its JSON file
            foreach (var visualEntry in allVisuals)
            {
                string display = String.Format(@"{0} - {1} ({2}, {3})", visualEntry.Page.DisplayName, visualEntry.Visual.Content.Visual.VisualType, (int)visualEntry.Visual.Content.Position.X, (int)visualEntry.Visual.Content.Position.Y);
                if (selected.Contains(display))
                {
                    string jsonPath = visualEntry.Visual.VisualFilePath;
                    if (!File.Exists(jsonPath))
                    {
                        Error(String.Format(@"JSON file not found: {0}", jsonPath));
                        continue;
                    }
                    System.Diagnostics.Process.Start(jsonPath);
                }
            }
        }
        
        void replaceField() 
        {

            //2025-05-25/B.Agullo
            //provided a definition.pbir file, this script allows the user to replace a measure in all visuals that use it with another measure.
            //when executing the script you must be connected to the semantic model to which the report is connected to or one that is identical. 
            //see https://www.esbrina-ba.com/pbir-scripts-to-replace-field-and-open-visual-json-files/ for reference on how to use it

            //using GeneralFunctions;
            //using Report.DTO;
            //using System.IO;
            //using Newtonsoft.Json.Linq;

            ReportExtended report = Rx.InitReport();
            if (report == null) return;

            var modifiedVisuals = new HashSet<VisualExtended>();

            var allVisuals = report.Pages
             .SelectMany(p => p.Visuals.Select(v => new { Page = p.Page, Visual = v }))
             .ToList();


            IList<string> allReportMeasures = allVisuals
                .SelectMany(x => x.Visual.GetAllReferencedMeasures())
                .Distinct()
                .ToList();

            string measureToReplace = Fx.ChooseString(
                OptionList: allReportMeasures,
                "Select a measure to replace"
            );

            if (string.IsNullOrEmpty(measureToReplace))
            {
                Error("No measure selected.");
                return;
            }

            Measure replacementMeasure = SelectMeasure(
                label: $"Select a replacement for '{measureToReplace}'"
            );

            if (replacementMeasure == null)
            {
                Error("No replacement measure selected.");
                return;
            }

            var visualsUsingMeasure = allVisuals
                .Where(x => x.Visual.GetAllReferencedMeasures().Contains(measureToReplace))
                .Select(x => new
                {
                    Display = $"{x.Page.DisplayName} - {x.Visual.Content.Visual.VisualType} ({(int)x.Visual.Content.Position.X}, {(int)x.Visual.Content.Position.Y})",
                    Visual = x.Visual
                })
                .ToList();

            if (visualsUsingMeasure.Count == 0)
            {
                Info($"No visuals use the measure '{measureToReplace}'.");
                return;
            }

            // Step 2: Let the user choose one or more visuals
            var options = visualsUsingMeasure.Select(v => v.Display).ToList();
            List<string> selected = Fx.ChooseStringMultiple(options, "Select visuals to update");

            if (selected == null || selected.Count == 0)
            {
                Info("No visuals selected.");
                return;
            }

            // Step 3: Apply replacement only to selected visuals
            foreach (var visualEntry in visualsUsingMeasure)
            {
                if (selected.Contains(visualEntry.Display))
                {
                    visualEntry.Visual.ReplaceMeasure(measureToReplace, replacementMeasure);
                    modifiedVisuals.Add(visualEntry.Visual);
                }
            }

            // Save modified visuals
            foreach (var visual in modifiedVisuals)
            {
                Rx.SaveVisual(visual);
            }

            Output($"{modifiedVisuals.Count} visuals were modified.");




        }


        void fixBrokenFields()
        {
            //using GeneralFunctions;
            //using Report.DTO;
            //using System.IO;
            //using Newtonsoft.Json.Linq;
            
            ReportExtended report = Rx.InitReport();
            if (report == null) return;

            var modifiedVisuals = new HashSet<VisualExtended>();

            // Gather all visuals and all fields used in them
            IList<VisualExtended> allVisuals = (report.Pages ?? new List<PageExtended>())
                .SelectMany(p => p.Visuals ?? Enumerable.Empty<VisualExtended>())
                .ToList();

            IList<string> allReportMeasures = allVisuals
                .SelectMany(v => v.GetAllReferencedMeasures())
                .Distinct()
                .ToList();

            IList<string> allReportColumns = allVisuals
                .SelectMany(v => v.GetAllReferencedColumns())
                .Distinct()
                .ToList();

            IList<string> allModelMeasures = Model.AllMeasures
                .Select(m => $"{m.Table.DaxObjectFullName}[{m.Name}]")
                .ToList();

            IList<string> allModelColumns = Model.AllColumns
                .Select(c => c.DaxObjectFullName)
                .ToList();

            IList<string> brokenMeasures = allReportMeasures
                .Where(m => !allModelMeasures.Contains(m))
                .ToList();

            IList<string> brokenColumns = allReportColumns
                .Where(c => !allModelColumns.Contains(c))
                .ToList();

            if(!brokenMeasures.Any() && !brokenColumns.Any())
            {
                Info("No broken measures or columns found.");
                return;
            }


            // Replacement maps for filterConfig patch
            var tableReplacementMap = new Dictionary<string, string>();
            var fieldReplacementMap = new Dictionary<string, string>();

            foreach (string brokenMeasure in brokenMeasures)
            {
                Measure replacement = 
                    SelectMeasure(label: $"{brokenMeasure} was not found in the model. What's the new measure?");
                if (replacement == null) { Error("You Cancelled"); return; }

                string oldTable = brokenMeasure.Split('[')[0].Trim('\'');
                string oldField = brokenMeasure.Split('[', ']')[1];

                tableReplacementMap[oldTable] = replacement.Table.Name;
                fieldReplacementMap[oldField] = replacement.Name;

                foreach (var visual in allVisuals)
                {
                    if (visual.GetAllReferencedMeasures().Contains(brokenMeasure))
                    {
                        visual.ReplaceMeasure(brokenMeasure, replacement, modifiedVisuals);
                    }
                }
            }

            foreach (string brokenColumn in brokenColumns)
            {
                Column replacement = SelectColumn(Model.AllColumns, label: $"{brokenColumn} was not found in the model. What's the new column?");
                if (replacement == null) { Error("You Cancelled"); return; }

                string oldTable = brokenColumn.Split('[')[0].Trim('\'');
                string oldField = brokenColumn.Split('[', ']')[1];

                tableReplacementMap[oldTable] = replacement.Table.Name;
                fieldReplacementMap[oldField] = replacement.Name;

                foreach (var visual in allVisuals)
                {
                    if (visual.GetAllReferencedColumns().Contains(brokenColumn))
                    {
                        visual.ReplaceColumn(brokenColumn, replacement, modifiedVisuals);
                        
                    }
                }

                
            }

            // Apply raw text-based replacement to filterConfig JSON strings
            foreach (var visual in allVisuals)
            {
                visual.ReplaceInFilterConfigRaw(tableReplacementMap, fieldReplacementMap, modifiedVisuals);
            }

            // Save modified visuals
            foreach (var visual in modifiedVisuals)
            {
                Rx.SaveVisual(visual);
            }

            Output($"{modifiedVisuals.Count} visuals were modified.");
        }

        
        

        void changeCoordinatesOfVisual()
        {
            //using GeneralFunctions;
            //using Report.DTO;
            //using System.IO;

            ReportExtended report = Rx.InitReport();
            VisualExtended visual = Rx.SelectVisual(report);

            //Ask user for new X and Y values
            string currentX = ((int)visual.Content.Position.X).ToString();
            string currentY = ((int)visual.Content.Position.Y).ToString();

            string newXStr = Interaction.InputBox("Enter new X position:", "Modify Visual", currentX, 740, 400);
            string newYStr = Interaction.InputBox("Enter new Y position:", "Modify Visual", currentY, 740, 400);

            int newX, newY;
            if (!int.TryParse(newXStr, out newX) || !int.TryParse(newYStr, out newY))
            {
                Error("Invalid input. Please enter numeric values.");
                return;
            }

            // Step 5: Update the visual
            visual.Content.Position.X = newX;
            visual.Content.Position.Y = newY;

            Rx.SaveVisual(visual);

        }
                          
               
            
        
        void myNewScript()
        {
            //create a measure for each of the selected columns
            foreach(Column c in Selected.Columns)
            {
                string mName = "Sum of " + c.Name;
                string mExpression = String.Format("SUM({0})", c.DaxObjectFullName);
                Measure m = c.Table.AddMeasure(mName, mExpression);
                m.FormatString = c.FormatString; 
                //put the measure in a subfolder 
                
            }
            
        }


        void myScript()
        {

            //using GeneralFunctions; 

            /*uncomment in TE3 to avoid wating cursor infront of dialogs*/
            //bool waitCursor = Application.UseWaitCursor;
            //Application.UseWaitCursor = false;
            
            Fx.CreateCalcTable(Model, "myMeasures", "{0}");

            //Application.UseWaitCursor = waitCursor;
        }



        //code snippets
        void userChooseName()
        {
            //#r "Microsoft.VisualBasic"
            //using Microsoft.VisualBasic;

            /*uncomment in TE3 to avoid wating cursor infront of dialogs*/
            //bool waitCursor = Application.UseWaitCursor;
            //Application.UseWaitCursor = false;

            string calcGroupName = Interaction.InputBox("Provide a name for your Calc Group", "Calc Group Name", "Time Intelligence", 740, 400);
            
            //sample code using the variable
            Output(calcGroupName);

            //Application.UseWaitCursor = waitCursor;

        }

        void userChooseYesNo()
        {

            //using System.Windows.Forms;

            /*uncomment in TE3 to avoid wating cursor infront of dialogs*/
            //bool waitCursor = Application.UseWaitCursor;
            //Application.UseWaitCursor = false;

            DialogResult dialogResult = MessageBox.Show(text:"Generate Field Parameter?", caption:"Field Parameter", buttons:MessageBoxButtons.YesNo);
            bool generateFieldParameter = (dialogResult == DialogResult.Yes);
            
            //sample code using the variable
            Output(generateFieldParameter);

            //Application.UseWaitCursor = waitCursor;

        }

        void userChooseString()
        {

            //using System.Windows.Forms;

            /*uncomment in TE3 to avoid wating cursor infront of dialogs*/
            //bool waitCursor = Application.UseWaitCursor;
            //Application.UseWaitCursor = false;


            List<string> sampleList = new List<string>();

            sampleList.Add("a");
            sampleList.Add("b");
            sampleList.Add("c");


            // Code that defines a local function "SelectString", which pops up a listbox allowing the user to select a 
            // string from a number of options:
            Func<IList<string>, string, string> SelectString = (IList<string> options, string title) =>
            {
                var form = new Form();
                form.Text = title;
                var buttonPanel = new Panel();
                buttonPanel.Dock = DockStyle.Bottom;
                buttonPanel.Height = 30;
                var okButton = new System.Windows.Forms.Button() { DialogResult = DialogResult.OK, Text = "OK" };
                var cancelButton = new System.Windows.Forms.Button() { DialogResult = DialogResult.Cancel, Text = "Cancel", Left = 80 };
                var listbox = new ListBox();
                listbox.Dock = DockStyle.Fill;
                listbox.Items.AddRange(options.ToArray());
                listbox.SelectedItem = options[0];

                form.Controls.Add(listbox);
                form.Controls.Add(buttonPanel);
                buttonPanel.Controls.Add(okButton);
                buttonPanel.Controls.Add(cancelButton);

                var result = form.ShowDialog();
                if (result == DialogResult.Cancel) return null;
                return listbox.SelectedItem.ToString();
            };

            

            //let the user select the name of the macro to copy
            String select = SelectString(sampleList, "Choose a macro");

            //check that indeed one macro was selected
            if (select == null)
            {
                Info("You cancelled!");
                return;
            }

            //code using "select" variable
            Output(select);

            //Application.UseWaitCursor = waitCursor;
        }
    
        void CopyMacroFromVSFileWithDll()
        {

            // NOCOPY replace <PROJECT FOLDER> (both instances) with the path to the folder where the .sln file is stored.
            //#r "<PROJECT FOLDER>\TE Scripts\bin\Debug\TE Scripts.dll"
            //using TE_Scripting;

            string baseFolderPath = @"<PROJECT FOLDER>";
            string macroFilePath = String.Format(@"{0}\TE Scripts\TE Scripts.cs", baseFolderPath);
            string generalFunctionsClassFilePath = String.Format(@"{0}\GeneralFunctions\GeneralFunctions.cs", baseFolderPath);
            string reportClassFilePath = String.Format(@"{0}\Report\Report.cs", baseFolderPath);
            string reportFunctionsClassFilePath = String.Format(@"{0}\ReportFunctions\ReportFunctions.cs", baseFolderPath);

            TE_Scripting.TE_Scripts.CopyMacroFromVSFile(
                macroFilePath, generalFunctionsClassFilePath, reportClassFilePath,reportFunctionsClassFilePath
            );
        }


        public static void CopyMacroFromVSFile(string macroFilePath, string generalFunctionsClassFilePath, string reportClassFilePath, string reportFunctionsClassFilePath)
        {
            //#r "System.IO"
            //#r "Microsoft.CodeAnalysis"
            //using System.IO;
            //using System.Windows.Forms;
            //using Microsoft.CodeAnalysis;
            //using Microsoft.CodeAnalysis.CSharp;
            //using Microsoft.CodeAnalysis.CSharp.Syntax;
            //using System.Text.RegularExpressions;

            // '2023-11-25 / B.Agullo / Fixed the code to combine references from general functions correctly
            // '2023-05-06 / B.Agullo / 
            // this macro copies the code of any of the methods defined in the TE_Scripts.cs File
            // if the macro is using the custom class it must include de following commented directive
            //           //using GeneralFunctions;
            // if this line is found the macro will copy the code also from the class defined in GeneralFunctions
            // and will combine the commented references of the class with those of the macro
            // once the macro finishes the code is in the clipboard so it can be pasted
            // in a new c# script tab in Tabular Editor, using CTRL+V 
            // see further detail at -- 

            //config
            // NOCOPY -- TO USE IN TE3 without creating a the dll file, uncomment the four following lines and complete the full path of both class files.
            //string baseFolderPath = @"<BASE FOLDER PATH OF THE SOLUTION>";
            //String macroFilePath = String.Format(@"{0}\TE Scripts\TE Scripts.cs", baseFolderPath);
            //String generalFunctionsClassFilePath = String.Format(@"{0}\GeneralFunctions\GeneralFunctions.cs", baseFolderPath);
            //String reportClassFilePath = String.Format(@"{0}\Report\Report.cs", baseFolderPath);
            //String reportFunctionsClassFilePath = String.Format(@"{0}\ReportFunctions\ReportFunctions.cs", baseFolderPath);
            String codeIndent = "            ";
            String noCopyMark = "NOCOPY";
            ////these libraries are already loaded in Tabular Editor and must not be specified
            //String[] tabularEditorLibraries = { "#r \"System.Windows.Forms\"" };
            //get file structure
            SyntaxTree tree = CSharpSyntaxTree.ParseText(File.ReadAllText(macroFilePath));

            List<string> macroNames = new List<string>();

            //extract method names that are not public static (just macro names) 
            macroNames = tree.GetRoot().DescendantNodes().OfType<MethodDeclarationSyntax>()
                                            .Where(m => m.Modifiers.ToString() != "public static")
                                            .Select(m => m.Identifier.ToString()).ToList();

            // Code that defines a local function "SelectString", which pops up a listbox allowing the user to select a 
            // string from a number of options:
            Func<IList<string>, string, string> SelectString = (IList<string> options, string title) =>
            {
                var form = new Form();
                form.Text = title;
                var buttonPanel = new Panel();
                buttonPanel.Dock = DockStyle.Bottom;
                buttonPanel.Height = 30;
                var okButton = new System.Windows.Forms.Button() { DialogResult = DialogResult.OK, Text = "OK" };
                var cancelButton = new System.Windows.Forms.Button() { DialogResult = DialogResult.Cancel, Text = "Cancel", Left = 80 };
                var listbox = new ListBox();
                listbox.Dock = DockStyle.Fill;
                listbox.Items.AddRange(options.ToArray());
                listbox.SelectedItem = options[0];

                form.Controls.Add(listbox);
                form.Controls.Add(buttonPanel);
                buttonPanel.Controls.Add(okButton);
                buttonPanel.Controls.Add(cancelButton);

                var result = form.ShowDialog();
                if (result == DialogResult.Cancel) return null;
                return listbox.SelectedItem.ToString();
            };

            //check that macros were found
            if (macroNames.Count == 0)
            {
                Error("No macros found in " + macroFilePath);
                return;
            }

            //let the user select the name of the macro to copy
            String select = SelectString(macroNames, "Choose a macro");

            //check that indeed one macro was selected
            if (select == null)
            {
                Info("You cancelled!");
                return;
            }

            //get the method
            MethodDeclarationSyntax method = tree.GetRoot().DescendantNodes().OfType<MethodDeclarationSyntax>()
                .First(m => m.Identifier.ToString() == select);

            //fix the code
            String macroCode = method.Body.ToFullString().Replace("//using", "using").Replace("//#r", "#r");

            bool usingGeneralFunctions = macroCode.Contains("using GeneralFunctions;");
            if (usingGeneralFunctions)
            {
                macroCode = macroCode.Replace("using GeneralFunctions;", "");
            }

            bool usingReportDTO = macroCode.Contains("using Report.DTO;");
            if (usingReportDTO)
            {
                macroCode = macroCode.Replace("using Report.DTO;", "");
            }

            int firstCurlyBracket = macroCode.IndexOf("{");
            int lastCurlyBracket = macroCode.LastIndexOf("}");

            macroCode = macroCode.Substring(firstCurlyBracket + 1, lastCurlyBracket - firstCurlyBracket - 1);

            string macroCodeClean = "";
            string[] macroCodeLines = macroCode.Split('\n');
            foreach (string macroCodeLine in macroCodeLines)
            {
                if (macroCodeLine.Contains(noCopyMark))
                {
                    //do nothing
                }
                else if (macroCodeLine.Contains("using") && macroCodeLine.Contains("Report.DTO"))
                {
                    //do nothing
                }
                else if (macroCodeLine.StartsWith(codeIndent))
                {
                    macroCodeClean += macroCodeLine.Substring(codeIndent.Length) + '\n';
                }
                else if (macroCodeLine.Contains("#r") || macroCodeLine.Contains("using"))
                {
                    macroCodeClean += macroCodeLine.Trim() + '\n';
                }
                else
                {
                    macroCodeClean += macroCodeLine + '\n';
                }
            }

            Func<string, string, string> CombineWithCustomClass = (string previousCode, string customClassFilePath) =>
            {
                string customClassEndMark = @"//******************";
                string customClassIndent = "    ";
                //these libraries are already loaded in Tabular Editor and must not be specified
                string[] tabularEditorLibraries = { "#r \"System.Windows.Forms\"" };


                string codeToAppend = "";

                //check the custom className 
                SyntaxTree customClassTree = CSharpSyntaxTree.ParseText(File.ReadAllText(customClassFilePath));

                string customClassNamespaceName = customClassTree.GetRoot().DescendantNodes().OfType<NamespaceDeclarationSyntax>().First().Name.ToString();

                ClassDeclarationSyntax customClass = customClassTree.GetRoot().DescendantNodes().OfType<ClassDeclarationSyntax>().First();

                String customClassCode = customClass.ToString();

                int endMarkIndex = customClassCode.IndexOf(customClassEndMark);

                //crop the last part and uncomment the closing bracket
                customClassCode = customClassCode.Substring(0, endMarkIndex - 1).Replace("        //}", "}").Replace("//using", "using").Replace("//#r", "#r");


                string[] customClassCodeLines = customClassCode.Split('\n');

                foreach (string customClassCodeLine in customClassCodeLines)
                {
                    if (customClassCodeLine.Contains(noCopyMark))
                    {
                        //do nothing
                    }
                    else if (customClassCodeLine.StartsWith(customClassIndent))
                    {
                        codeToAppend += customClassCodeLine.Substring(customClassIndent.Length) + Environment.NewLine;
                    }
                    else
                    {
                        codeToAppend += customClassCodeLine + Environment.NewLine;
                    }
                }


                int hashrFirstMacroCode = Math.Max(previousCode.IndexOf("#r"), 0);
                int hashrFirstCustomClass = codeToAppend.IndexOf("#r");

                if (hashrFirstCustomClass != -1)
                {
                    int hashrLastCustomClass = codeToAppend.LastIndexOf("#r");
                    int endOfHashrCustomClass = codeToAppend.IndexOf(Environment.NewLine, hashrLastCustomClass);

                    string[] hashrLines = codeToAppend.Substring(hashrFirstCustomClass, endOfHashrCustomClass - hashrFirstCustomClass).Split('\n');



                    foreach (String hashrLine in hashrLines)
                    {



                        if (tabularEditorLibraries.Contains(hashrLine.Trim()))
                        {
                            //do nothing
                        }
                        //if #r directive not present
                        else if (!previousCode.Contains(hashrLine.Trim()))
                        {
                            //insert in the code right before the first one
                            previousCode = previousCode.Substring(0, Math.Max(hashrFirstMacroCode - 1, 0))
                                + hashrLine.Trim() + Environment.NewLine
                                + previousCode.Substring(hashrFirstMacroCode);

                            //update the position of the first #r
                            hashrFirstMacroCode = Math.Max(previousCode.IndexOf("#r"), 0);
                        }


                    }



                    //remove #r directives from custom class 
                    codeToAppend = codeToAppend.Replace(codeToAppend.Substring(hashrFirstCustomClass, endOfHashrCustomClass - hashrFirstCustomClass), "");

                    int usingFirstMacroCode = Math.Max(previousCode.IndexOf("using"), 0);
                    int usingFirstCustomClass = codeToAppend.IndexOf("using");

                    if (usingFirstCustomClass != -1)
                    {
                        int usingLastCustomClass = codeToAppend.LastIndexOf("using");
                        int endOfusingCustomClass = codeToAppend.IndexOf(Environment.NewLine, usingLastCustomClass);

                        string[] usingLines = codeToAppend.Substring(usingFirstCustomClass, endOfusingCustomClass - usingFirstCustomClass).Split('\n');

                        foreach (String usingLine in usingLines)
                        {
                            //if using directive not present
                            if (!previousCode.Contains(usingLine))
                            {
                                //insert in the code right before the first one
                                previousCode = previousCode.Substring(0, Math.Max(usingFirstMacroCode - 1, 0))
                                    + Environment.NewLine + usingLine.Trim() + Environment.NewLine
                                    + previousCode.Substring(usingFirstMacroCode);

                                usingFirstMacroCode = Math.Max(previousCode.IndexOf("using"), 0);
                            }
                        }

                        //remove using directives from custom class 
                        codeToAppend = codeToAppend
                                                   .Replace(codeToAppend
                                                        .Substring(usingFirstCustomClass, endOfusingCustomClass - usingFirstCustomClass) + Environment.NewLine,
                                                        "");

                    }

                    //remove empty lines
                    previousCode = Regex.Replace(previousCode, @"^\s*$\n|\r", string.Empty, RegexOptions.Multiline);
                    codeToAppend = Regex.Replace(codeToAppend, @"^\s*$\n|\r", string.Empty, RegexOptions.Multiline);


                }

                string outputCode = previousCode += Environment.NewLine + codeToAppend;

                int lastUsingFinal = outputCode.IndexOf("using");

                if (lastUsingFinal != -1)
                {
                    int endOfDirective = outputCode.IndexOf(";", lastUsingFinal) + 1;
                    outputCode = outputCode.Substring(0, endOfDirective)
                        + Environment.NewLine
                        + Environment.NewLine
                        + outputCode.Substring(endOfDirective + 1);

                }

                return outputCode;
            };



            //remove empty lines
            macroCodeClean = Regex.Replace(macroCodeClean, @"^\s*$\n|\r", string.Empty, RegexOptions.Multiline);

            string macroCodeClean2 = macroCodeClean;

            if (usingGeneralFunctions)
            {
                macroCodeClean2 = CombineWithCustomClass(macroCodeClean2, generalFunctionsClassFilePath);
            }

            string macroCodeClean3 = macroCodeClean2;

            //check if macro is using custom class
            if (usingReportDTO)
            {

                macroCodeClean3 = CombineWithCustomClass(macroCodeClean2, reportFunctionsClassFilePath);

                // Parse the syntax tree
                SyntaxTree reportClassTree = CSharpSyntaxTree.ParseText(File.ReadAllText(reportClassFilePath));
                var root = reportClassTree.GetRoot();

                // Get namespace node
                var namespaceNode = root.DescendantNodes().OfType<NamespaceDeclarationSyntax>().First();

                // Get classes that are direct children of the namespace (not nested classes)
                var reportClass = namespaceNode.Members
                    .OfType<ClassDeclarationSyntax>();

                // Concatenate class codes
                string reportClassCode = string.Join(
                    Environment.NewLine,
                    reportClass.Select(c => c.ToFullString())
                );

                macroCodeClean3 += Environment.NewLine + reportClassCode;

            }



            //copy the code to the clipboard
            Clipboard.SetText(macroCodeClean3);


        }
            
        


        //public static string CombineWithCustomClass(string previousCode,string customClassFilePath)
        //{
        //    string customClassEndMark = @"//******************";
        //    string customClassIndent = "    ";
        //    string noCopyMark = "NOCOPY";
        //    //these libraries are already loaded in Tabular Editor and must not be specified
        //    string[] tabularEditorLibraries = { "#r \"System.Windows.Forms\"" };


        //    string codeToAppend = "";

        //    //check the custom className 
        //    SyntaxTree customClassTree = CSharpSyntaxTree.ParseText(File.ReadAllText(customClassFilePath));

        //    string customClassNamespaceName = customClassTree.GetRoot().DescendantNodes().OfType<NamespaceDeclarationSyntax>().First().Name.ToString();

        //    ClassDeclarationSyntax customClass = customClassTree.GetRoot().DescendantNodes().OfType<ClassDeclarationSyntax>().First();

        //    String customClassCode = customClass.ToString();

        //    int endMarkIndex = customClassCode.IndexOf(customClassEndMark);

        //    //crop the last part and uncomment the closing bracket
        //    customClassCode = customClassCode.Substring(0, endMarkIndex - 1).Replace("        //}", "}").Replace("//using", "using").Replace("//#r", "#r");

            
        //    string[] customClassCodeLines = customClassCode.Split('\n');

        //    foreach (string customClassCodeLine in customClassCodeLines)
        //    {
        //        if (customClassCodeLine.Contains(noCopyMark))
        //        {
        //            //do nothing
        //        }
        //        else if (customClassCodeLine.StartsWith(customClassIndent))
        //        {
        //            codeToAppend += customClassCodeLine.Substring(customClassIndent.Length) + Environment.NewLine;
        //        }
        //        else
        //        {
        //            codeToAppend += customClassCodeLine + Environment.NewLine;
        //        }
        //    }


        //    int hashrFirstMacroCode = Math.Max(previousCode.IndexOf("#r"), 0);
        //    int hashrFirstCustomClass = codeToAppend.IndexOf("#r");

        //    if (hashrFirstCustomClass != -1)
        //    {
        //        int hashrLastCustomClass = codeToAppend.LastIndexOf("#r");
        //        int endOfHashrCustomClass = codeToAppend.IndexOf(Environment.NewLine, hashrLastCustomClass);

        //        string[] hashrLines = codeToAppend.Substring(hashrFirstCustomClass, endOfHashrCustomClass - hashrFirstCustomClass).Split('\n');



        //        foreach (String hashrLine in hashrLines)
        //        {



        //            if (tabularEditorLibraries.Contains(hashrLine.Trim()))
        //            {
        //                //do nothing
        //            }
        //            //if #r directive not present
        //            else if (!previousCode.Contains(hashrLine.Trim()))
        //            {
        //                //insert in the code right before the first one
        //                previousCode = previousCode.Substring(0, Math.Max(hashrFirstMacroCode - 1, 0))
        //                    + hashrLine.Trim() + Environment.NewLine
        //                    + previousCode.Substring(hashrFirstMacroCode);

        //                //update the position of the first #r
        //                hashrFirstMacroCode = Math.Max(previousCode.IndexOf("#r"), 0);
        //            }


        //        }



        //        //remove #r directives from custom class 
        //        codeToAppend = codeToAppend.Replace(codeToAppend.Substring(hashrFirstCustomClass, endOfHashrCustomClass - hashrFirstCustomClass), "");

        //        int usingFirstMacroCode = Math.Max(previousCode.IndexOf("using"), 0);
        //        int usingFirstCustomClass = codeToAppend.IndexOf("using");

        //        if (usingFirstCustomClass != -1)
        //        {
        //            int usingLastCustomClass = codeToAppend.LastIndexOf("using");
        //            int endOfusingCustomClass = codeToAppend.IndexOf(Environment.NewLine, usingLastCustomClass);

        //            string[] usingLines = codeToAppend.Substring(usingFirstCustomClass, endOfusingCustomClass - usingFirstCustomClass).Split('\n');

        //            foreach (String usingLine in usingLines)
        //            {
        //                //if using directive not present
        //                if (!previousCode.Contains(usingLine))
        //                {
        //                    //insert in the code right before the first one
        //                    previousCode = previousCode.Substring(0, Math.Max(usingFirstMacroCode - 1, 0))
        //                        + Environment.NewLine + usingLine.Trim() + Environment.NewLine
        //                        + previousCode.Substring(usingFirstMacroCode);

        //                    usingFirstMacroCode = Math.Max(previousCode.IndexOf("using"), 0);
        //                }
        //            }

        //            //remove using directives from custom class 
        //            codeToAppend = codeToAppend
        //                                       .Replace(codeToAppend
        //                                            .Substring(usingFirstCustomClass, endOfusingCustomClass - usingFirstCustomClass) + Environment.NewLine,
        //                                            "");

        //        }

        //        //remove empty lines
        //        previousCode = Regex.Replace(previousCode, @"^\s*$\n|\r", string.Empty, RegexOptions.Multiline);
        //        codeToAppend = Regex.Replace(codeToAppend, @"^\s*$\n|\r", string.Empty, RegexOptions.Multiline);

                
        //    }

        //    string outputCode = previousCode += Environment.NewLine + codeToAppend;

        //    int lastUsingFinal = outputCode.IndexOf("using");

        //    if (lastUsingFinal != -1)
        //    {
        //        int endOfDirective = outputCode.IndexOf(";", lastUsingFinal) + 1;
        //        outputCode = outputCode.Substring(0, endOfDirective)
        //            + Environment.NewLine
        //            + Environment.NewLine
        //            + outputCode.Substring(endOfDirective + 1);

        //    }

        //    return outputCode;
        //}


        

        //these two are necessary to have the Model and Selected objects available in the script
        static readonly Model Model;
        static readonly TabularEditor.UI.UITreeSelection Selected;


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

        public static Table SelectTable(Table preselect = null, string label = "Select Table")
        {
            return ScriptHelper.SelectTable(preselect: preselect, label: label);
        }

        public static Column SelectColumn(Table table, Column preselect = null, string label = "Select Column")
        {
            return ScriptHelper.SelectColumn(table: table, preselect: preselect, label: label);
        }

        public static Column SelectColumn(IEnumerable<Column> columns, Column preselect = null, string label = "Select Column")
        {
            return ScriptHelper.SelectColumn(columns: columns, preselect: preselect, label: label);
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
