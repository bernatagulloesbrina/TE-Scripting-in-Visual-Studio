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
//using GeneralFunctions; //Uncomment if you use the custom class, add reference to the project too.

// '2023-05-06 / B.Agullo / 
//coding environment for Tabular Editor C# Scripts
// see --- for reference on how to use it.

namespace TE_Scripting
{
    public class TE_Scripts
    {

        void myScript()
        {
           //your code goes here, only code should be copied to TabularEditor, along with any necessary references. 


        }


        void copyMacroFromVSFile()
        {
            //#r "System.IO"
            //#r "Microsoft.CodeAnalysis"
            //using System.IO;
            //using System.Windows.Forms;
            //using Microsoft.CodeAnalysis;
            //using Microsoft.CodeAnalysis.CSharp;
            //using Microsoft.CodeAnalysis.CSharp.Syntax;

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
            String macroFilePath = @"<HERE FULL PATH TO TE_Scripts.cs FILE>";
            String customClassFilePath = @"<HERE FULL PATH TO GeneralFunctions.cs FILE>";
            String codeIndent = "            ";
            String customClassEndMark = @"//******************";

            //get file structure
            SyntaxTree tree = CSharpSyntaxTree.ParseText(File.ReadAllText(macroFilePath));

            //extract method names that are not public static (just macro names) 
            List<string> macroNames = tree.GetRoot().DescendantNodes().OfType<MethodDeclarationSyntax>()
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


            //check the custom className 
            SyntaxTree customClassTree = CSharpSyntaxTree.ParseText(File.ReadAllText(customClassFilePath));

            string customClassNamespaceName = customClassTree.GetRoot().DescendantNodes().OfType<NamespaceDeclarationSyntax>().First().Name.ToString();


            //check if macro is using custom class
            if (macroCode.Contains("using " + customClassNamespaceName))
            {

                ClassDeclarationSyntax customClass = customClassTree.GetRoot().DescendantNodes().OfType<ClassDeclarationSyntax>().First();

                String customClassCode = customClass.ToString();
                int endMarkIndex = customClassCode.IndexOf(customClassEndMark);

                //crop the last part and uncomment the closing bracket
                customClassCode = customClassCode.Substring(0, endMarkIndex - 1).Replace("//}", "}").Replace("//using", "using").Replace("//#r", "#r");


                int hashrFirstMacroCode = Math.Max(macroCode.IndexOf("#r"), 0);
                int hashrFirstCustomClass = customClassCode.IndexOf("#r");

                if (hashrFirstCustomClass != -1)
                {
                    int hashrLastCustomClass = customClassCode.LastIndexOf("#r");
                    int endOfHashrCustomClass = customClassCode.IndexOf(Environment.NewLine, hashrLastCustomClass);

                    string[] hashrLines = customClassCode.Substring(hashrFirstCustomClass, endOfHashrCustomClass - hashrFirstCustomClass).Split('\n');

                    foreach (String hashrLine in hashrLines)
                    {
                        //if #r directive not present
                        if (!macroCode.Contains(hashrLine))
                        {
                            //insert in the code right before the first one
                            macroCode = macroCode.Substring(0, Math.Max(hashrFirstMacroCode - 1, 0))
                                + hashrLine + Environment.NewLine
                                + macroCode.Substring(hashrFirstMacroCode);

                            //update the position of the first #r
                            hashrFirstMacroCode = Math.Max(customClassCode.IndexOf("#r"), 0);
                        }
                    }

                    //remove #r directives from custom class 
                    customClassCode = customClassCode.Replace(customClassCode.Substring(hashrLastCustomClass, endOfHashrCustomClass - hashrLastCustomClass), "");

                }

                int usingFirstMacroCode = Math.Max(macroCode.IndexOf("using"), 0);
                int usingFirstCustomClass = customClassCode.IndexOf("using");

                if (usingFirstCustomClass != -1)
                {
                    int usingLastCustomClass = customClassCode.LastIndexOf("using");
                    int endOfusingCustomClass = customClassCode.IndexOf(Environment.NewLine, usingLastCustomClass);

                    string[] usingLines = customClassCode.Substring(usingFirstCustomClass, endOfusingCustomClass - usingFirstCustomClass).Split('\n');

                    foreach (String usingLine in usingLines)
                    {


                        //if #r directive not present
                        if (!macroCode.Contains(usingLine))
                        {
                            //insert in the code right before the first one
                            macroCode = macroCode.Substring(0, Math.Max(usingFirstMacroCode - 1, 0))
                                + usingLine + Environment.NewLine
                                + macroCode.Substring(usingFirstMacroCode);

                            usingFirstMacroCode = Math.Max(macroCode.IndexOf("using"), 0);
                        }
                    }

                    //remove using directives from custom class 
                    customClassCode = customClassCode.Replace(customClassCode.Substring(usingFirstCustomClass, endOfusingCustomClass - usingFirstCustomClass), "");

                }

                //remove the using directive since it is an in-script custom class
                macroCode = macroCode.Replace("using " + customClassNamespaceName, "");

                //append custom class to macro 
                macroCode += customClassCode;
            }

            string macroCodeClean = "";
            string[] macroCodeLines = macroCode.Split('\n');
            foreach (string macroCodeLine in macroCodeLines)
            {
                if (macroCodeLine.StartsWith(codeIndent))
                {
                    macroCodeClean += macroCodeLine.Substring(codeIndent.Length);
                }
                else
                {
                    macroCodeClean += macroCodeLine;
                }
            }

            //copy the code to the clipboard
            Clipboard.SetText(macroCodeClean);


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

        public static Column SelectColumn(Table table, Column preselect = null, string label = "Select Column")
        {
            return ScriptHelper.SelectColumn(table: table, preselect: preselect, label: label); 
        }

        public static void Output(object value, int lineNumber = -1)
        {
            ScriptHelper.Output(value: value, lineNumber: lineNumber);
        }


    }
}
