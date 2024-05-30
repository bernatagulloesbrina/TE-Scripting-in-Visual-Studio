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

namespace GeneralFunctions
{

    //copy from the following line up to ****** and remove the // before the closing bracket
    //after the class declaration add all the #r and using statements necessary for the custom class code to run in Tabular Editor
    //these directives will be combined with the ones from the macro when using the CopyMacro script
    public static class Fx
    {
        //#r "Microsoft.VisualBasic"
        //#r "System.Windows.Forms"
        //using Microsoft.VisualBasic;
        //using System.Windows.Forms;


        // NOCOPY    in TE2 (at least up to 2.17.2) any method that accesses or modifies
        // NOCOPY    the model needs a reference to the model 
        // NOCOPY    the following is an example method where you can build extra logic
        public static Table CreateCalcTable(Model model, string tableName, string tableExpression)
        {
            if(!model.Tables.Any(t => t.Name == tableName))
            {
                return model.AddCalculatedTable(tableName, tableExpression);
            }
            else
            {
                return model.Tables.Where(t => t.Name == tableName).First();
            }
        }

        public static string GetNameFromUser(string Prompt, string Title, string DefaultResponse)
        {    
            string response = Interaction.InputBox(Prompt, Title, DefaultResponse, 740, 400);
            return response;
        }


        public static string ChooseString(IList<string> OptionList)
        {
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
            String select = SelectString(OptionList, "Choose a macro");

            //check that indeed one macro was selected
            if (select == null)
            {
                Info("You cancelled!");
                
            }

            return select;

        }

        public static IEnumerable<Table> GetDateTables(Model model)
        {
            IEnumerable<Table> dateTables = null as IEnumerable<Table>;

            if (model.Tables.Any(t => t.DataCategory == "Time" && t.Columns.Any(c => c.IsKey == true)))
            {
                dateTables = model.Tables.Where(t => t.DataCategory == "Time" && t.Columns.Any(c => c.IsKey == true && c.DataType == DataType.DateTime));
            }
            else
            {
                Error("No date table detected in the model. Please mark your date table(s) as date table");

            }

            return dateTables;
        }


        public static Table GetTablesWithAnnotation(IEnumerable<Table> tables, string annotationLabel, string annotationValue)
        {
            Func<Table, bool> lambda = t => t.GetAnnotation(annotationLabel) == annotationValue;

            IEnumerable<Table> matchTables = GetFilteredTables(tables, lambda);

            if (matchTables == null)
            {
                return null;
            }
            else
            {
                return matchTables.First();
            }

        }

        public static IEnumerable<Table> GetFilteredTables(IEnumerable<Table> tables, Func<Table, bool> lambda)
        {
            if (tables.Any(t => lambda(t)))
            {
                return tables.Where(t => lambda(t));
            }
            else
            {
                return null as IEnumerable<Table>;
            }
        }

        public static IEnumerable<Column> GetFilteredColumns(IEnumerable<Column> columns, Func<Column, bool> lambda, bool returnAllIfNoneFound = true)
        {
            if (columns.Any(c => lambda(c)))
            {
                return columns.Where(c => lambda(c));
            }
            else
            {
                if (returnAllIfNoneFound)
                {
                    return columns;
                }
                else
                {
                    return null as IEnumerable<Column>;
                }
            }

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
