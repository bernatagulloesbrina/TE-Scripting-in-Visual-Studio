using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TabularEditor.TOMWrapper;
using TabularEditor.Scripting;
using System.Windows.Forms;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.CSharp;
using System.IO;
using GeneralFunctions; //Uncomment if you use the custom class, add reference to the project too.
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;

// '2023-05-06 / B.Agullo / 
//coding environment for Tabular Editor C# Scripts
// see https://www.esbrina-ba.com/c-scripting-nirvana-effortlessly-use-visual-studio-as-your-coding-environment/ for reference on how to use it.

namespace TE_Scripting
{
    public class TE_Scripts
    {

        void myScript()
        {

            //using GeneralFunctions; 

            bool waitCursor = Application.UseWaitCursor;
            Application.UseWaitCursor = false;
            
            Fx.CreateCalcTable(Model, "myMeasures", "{0}");

            Application.UseWaitCursor = waitCursor;
        }



        //code snippets
        void userChooseName()
        {
            //#r "Microsoft.VisualBasic"
            //using Microsoft.VisualBasic;

            bool waitCursor = Application.UseWaitCursor;
            Application.UseWaitCursor = false;

            string calcGroupName = Interaction.InputBox("Provide a name for your Calc Group", "Calc Group Name", "Time Intelligence", 740, 400);
            
            //sample code using the variable
            Output(calcGroupName);

            Application.UseWaitCursor = waitCursor;

        }

        void userChooseYesNo()
        {

            //using System.Windows.Forms;

            bool waitCursor = Application.UseWaitCursor;
            Application.UseWaitCursor = false;

            DialogResult dialogResult = MessageBox.Show(text:"Generate Field Parameter?", caption:"Field Parameter", buttons:MessageBoxButtons.YesNo);
            bool generateFieldParameter = (dialogResult == DialogResult.Yes);
            
            //sample code using the variable
            Output(generateFieldParameter);

            Application.UseWaitCursor = waitCursor;

        }

        void userChooseString()
        {

            //using System.Windows.Forms;

            bool waitCursor = Application.UseWaitCursor;
            Application.UseWaitCursor = false;


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
                var okButton = new Button() { DialogResult = DialogResult.OK, Text = "OK" };
                var cancelButton = new Button() { DialogResult = DialogResult.Cancel, Text = "Cancel", Left = 80 };
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

            Application.UseWaitCursor = waitCursor;
        }
    
        void CopyMacroFromVSFileWithDll()
        {

            // NOCOPY Compile the project and get the path to TE Scripts.dll in <PROJECT FOLDER>\bin\Debug\TE Scripts.dll
            // NOCOPY replace 
            //#r "<HERE FULL PATH TO 'TE Scripts.dll' file>"
            //using TE_Scripting;

            // NOCOPY replace <HERE FULL PATH TO TE_Scripts.cs FILE> by the full path to the TE_Scripts.cs file
            // NOCOPY replace <HERE FULL PATH TO GeneralFunctions.cs FILE> by the full path to the GeneralFunctions.cs file
            TE_Scripting.TE_Scripts.CopyMacroFromVSFile(
                @"<HERE FULL PATH TO TE_Scripts.cs FILE>",
                @"<HERE FULL PATH TO GeneralFunctions.cs FILE>"
            );
        }


        public static void CopyMacroFromVSFile(string macroFilePath, string customClassFilePath)
        {
            //#r "System.IO"
            //#r "Microsoft.CodeAnalysis"
            //using System.IO;
            //using System.Windows.Forms;
            //using Microsoft.CodeAnalysis;
            //using Microsoft.CodeAnalysis.CSharp;
            //using Microsoft.CodeAnalysis.CSharp.Syntax;
            //using System.Text.RegularExpressions;


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
            // NOCOPY -- TO USE IN TE3 without creating a the dll file, uncomment the two following lines and complete the full path of both class files.
            //String macroFilePath = @"<HERE FULL PATH TO TE_Scripts.cs FILE>";
            //String customClassFilePath = @"<HERE FULL PATH TO GeneralFunctions.cs FILE>";
            String codeIndent = "            ";
            String customClassEndMark = @"//******************";
            String customClassIndent = "    ";
            String noCopyMark = "NOCOPY";

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
                var okButton = new Button() { DialogResult = DialogResult.OK, Text = "OK" };
                var cancelButton = new Button() { DialogResult = DialogResult.Cancel, Text = "Cancel", Left = 80 };
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
                else if (macroCodeLine.StartsWith(codeIndent))
                {
                    macroCodeClean += macroCodeLine.Substring(codeIndent.Length) + '\n';
                } 
                else if (macroCodeLine.Contains("#r") || macroCodeLine.Contains("using")) {
                    macroCodeClean += macroCodeLine.Trim() + '\n';
                }
                else
                {
                    macroCodeClean += macroCodeLine + '\n';
                }
            }
            
            //remove empty lines
            macroCodeClean = Regex.Replace(macroCodeClean, @"^\s*$\n|\r", string.Empty, RegexOptions.Multiline);


            //check the custom className 
            SyntaxTree customClassTree = CSharpSyntaxTree.ParseText(File.ReadAllText(customClassFilePath));

            string customClassNamespaceName = customClassTree.GetRoot().DescendantNodes().OfType<NamespaceDeclarationSyntax>().First().Name.ToString();


            //check if macro is using custom class
            if (macroCodeClean.Contains("using " + customClassNamespaceName))
            {

                ClassDeclarationSyntax customClass = customClassTree.GetRoot().DescendantNodes().OfType<ClassDeclarationSyntax>().First();

                String customClassCode = customClass.ToString();

                int endMarkIndex = customClassCode.IndexOf(customClassEndMark);

                //crop the last part and uncomment the closing bracket
                customClassCode = customClassCode.Substring(0, endMarkIndex - 1).Replace("        //}", "}").Replace("//using", "using").Replace("//#r", "#r");

                string customClassCodeClean = ""; 
                string[] customClassCodeLines = customClassCode.Split('\n'); 
                
                foreach(string customClassCodeLine in customClassCodeLines)
                {
                    if (customClassCodeLine.Contains(noCopyMark))
                    {
                        //do nothing
                    }
                    else if (customClassCodeLine.StartsWith(customClassIndent))
                    {
                        customClassCodeClean += customClassCodeLine.Substring(customClassIndent.Length) + Environment.NewLine;
                    }
                    else
                    {
                        customClassCodeClean += customClassCodeLine + Environment.NewLine;
                    }
                }


                int hashrFirstMacroCode = Math.Max(macroCodeClean.IndexOf("#r"), 0);
                int hashrFirstCustomClass = customClassCodeClean.IndexOf("#r");

                if (hashrFirstCustomClass != -1)
                {
                    int hashrLastCustomClass = customClassCodeClean.LastIndexOf("#r");
                    int endOfHashrCustomClass = customClassCodeClean.IndexOf(Environment.NewLine, hashrLastCustomClass);

                    string[] hashrLines = customClassCodeClean.Substring(hashrFirstCustomClass, endOfHashrCustomClass - hashrFirstCustomClass).Split('\n');

                    foreach (String hashrLine in hashrLines)
                    {
                        //if #r directive not present
                        if (!macroCodeClean.Contains(hashrLine))
                        {
                            //insert in the code right before the first one
                            macroCodeClean = macroCodeClean.Substring(0, Math.Max(hashrFirstMacroCode - 1, 0))
                                + hashrLine.Trim() + Environment.NewLine
                                + macroCodeClean.Substring(hashrFirstMacroCode);

                            //update the position of the first #r
                            hashrFirstMacroCode = Math.Max(customClassCode.IndexOf("#r"), 0);
                        }
                    }

                    //remove #r directives from custom class 
                    customClassCodeClean = customClassCodeClean.Replace(customClassCodeClean.Substring(hashrLastCustomClass, endOfHashrCustomClass - hashrLastCustomClass), "");

                }

                int usingFirstMacroCode = Math.Max(macroCodeClean.IndexOf("using"), 0);
                int usingFirstCustomClass = customClassCodeClean.IndexOf("using");

                if (usingFirstCustomClass != -1)
                {
                    int usingLastCustomClass = customClassCodeClean.LastIndexOf("using");
                    int endOfusingCustomClass = customClassCodeClean.IndexOf(Environment.NewLine, usingLastCustomClass);

                    string[] usingLines = customClassCodeClean.Substring(usingFirstCustomClass, endOfusingCustomClass - usingFirstCustomClass).Split('\n');
                   
                    foreach (String usingLine in usingLines)
                    {
                        //if using directive not present
                        if (!macroCodeClean.Contains(usingLine))
                        {
                            //insert in the code right before the first one
                            macroCodeClean = macroCodeClean.Substring(0, Math.Max(usingFirstMacroCode - 1, 0))
                                + Environment.NewLine + usingLine.Trim() + Environment.NewLine
                                + macroCodeClean.Substring(usingFirstMacroCode);

                            usingFirstMacroCode = Math.Max(macroCodeClean.IndexOf("using"), 0);
                        }
                    }

                    //remove using directives from custom class 
                    customClassCodeClean = customClassCodeClean
                                               .Replace(customClassCodeClean
                                                    .Substring(usingFirstCustomClass, endOfusingCustomClass - usingFirstCustomClass) + Environment.NewLine, 
                                                    "");

                }

                //remove the using directive since it is an in-script custom class
                macroCodeClean = macroCodeClean.Replace("using " + customClassNamespaceName + ";", "");


                //remove empty lines
                macroCodeClean = Regex.Replace(macroCodeClean, @"^\s*$\n|\r", string.Empty, RegexOptions.Multiline);
                customClassCodeClean = Regex.Replace(customClassCodeClean, @"^\s*$\n|\r", string.Empty, RegexOptions.Multiline);

                //append custom class to macro with some space
                macroCodeClean += Environment.NewLine + customClassCodeClean;
            }

            string macroCodeClean2 = macroCodeClean;

            int lastUsingFinal = macroCodeClean2.IndexOf("using");

            if (lastUsingFinal != -1) {
                int endOfDirective = macroCodeClean2.IndexOf(";", lastUsingFinal) + 1;
                macroCodeClean2 = macroCodeClean2.Substring(0,endOfDirective) 
                    + Environment.NewLine
                    + Environment.NewLine
                    + macroCodeClean2.Substring(endOfDirective+1);

            }
            //copy the code to the clipboard
            Clipboard.SetText(macroCodeClean2);


        }


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

        public static void Output(object value, int lineNumber = -1)
        {
            ScriptHelper.Output(value: value, lineNumber: lineNumber);
        }


    }
}
