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
using static Report.DTO.VisualDto;

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
            return model.Tables.FirstOrDefault(t =>
                                string.Equals(t.Name, tableName, StringComparison.OrdinalIgnoreCase)) //case insensitive search
                                ?? model.AddCalculatedTable(tableName, tableExpression);
        }

        public static string GetNameFromUser(string Prompt, string Title, string DefaultResponse)
        {
            string response = Interaction.InputBox(Prompt, Title, DefaultResponse, 740, 400);
            return response;
        }


        public static string ChooseString(IList<string> OptionList, string label = "Choose item", int customWidth = 400, int customHeight = 500)
        {
            return ChooseStringInternal(OptionList, MultiSelect: false, label: label, customWidth: customWidth, customHeight:customHeight) as string;
        }

        public static List<string> ChooseStringMultiple(IList<string> OptionList, string label = "Choose item(s)", int customWidth = 400, int customHeight = 500)
        {
            return ChooseStringInternal(OptionList, MultiSelect:true, label:label, customWidth: customWidth, customHeight: customHeight) as List<string>;
        }

        private static object ChooseStringInternal(IList<string> OptionList, bool MultiSelect, string label = "Choose item(s)", int customWidth = 400, int customHeight = 500)
        {
            Form form = new Form
            {
                Text =label,
                Width = customWidth,
                Height = customHeight,
                StartPosition = FormStartPosition.CenterScreen,
                Padding = new Padding(20)
            };

            ListBox listbox = new ListBox
            {
                Dock = DockStyle.Fill,
                SelectionMode = MultiSelect ? SelectionMode.MultiExtended : SelectionMode.One
            };
            listbox.Items.AddRange(OptionList.ToArray());
            if (!MultiSelect && OptionList.Count > 0)
                listbox.SelectedItem = OptionList[0];

            FlowLayoutPanel buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 40,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(10)
            };

            Button selectAllButton = new Button { Text = "Select All", Visible = MultiSelect };
            Button selectNoneButton = new Button { Text = "Select None", Visible = MultiSelect };
            Button okButton = new Button { Text = "OK", DialogResult = DialogResult.OK };
            Button cancelButton = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel };

            selectAllButton.Click += delegate
            {
                for (int i = 0; i < listbox.Items.Count; i++)
                    listbox.SetSelected(i, true);
            };

            selectNoneButton.Click += delegate
            {
                for (int i = 0; i < listbox.Items.Count; i++)
                    listbox.SetSelected(i, false);
            };

            buttonPanel.Controls.Add(selectAllButton);
            buttonPanel.Controls.Add(selectNoneButton);
            buttonPanel.Controls.Add(okButton);
            buttonPanel.Controls.Add(cancelButton);

            form.Controls.Add(listbox);
            form.Controls.Add(buttonPanel);

            DialogResult result = form.ShowDialog();
            if (result == DialogResult.Cancel)
            {
                Info("You Cancelled!");
                return null;
            }

            if (MultiSelect)
            {
                List<string> selectedItems = new List<string>();
                foreach (object item in listbox.SelectedItems)
                    selectedItems.Add(item.ToString());
                return selectedItems;
            }
            else
            {
                return listbox.SelectedItem != null ? listbox.SelectedItem.ToString() : null;
            }
        }


        public static IEnumerable<Table> GetDateTables(Model model)
        {
            var dateTables = model.Tables
                .Where(t => t.DataCategory == "Time" &&
                       t.Columns.Any(c => c.IsKey && c.DataType == DataType.DateTime))
                .ToList();

            if (!dateTables.Any())
            {
                Error("No date table detected in the model. Please mark your date table(s) as date table");
                return null;
            }

            return dateTables;
        }


        public static Table GetTablesWithAnnotation(IEnumerable<Table> tables, string annotationLabel, string annotationValue)
        {
            Func<Table, bool> lambda = t => t.GetAnnotation(annotationLabel) == annotationValue;

            IEnumerable<Table> matchTables = GetFilteredTables(tables, lambda);

            return GetFilteredTables(tables, lambda).FirstOrDefault();
        }

        public static IEnumerable<Table> GetFilteredTables(IEnumerable<Table> tables, Func<Table, bool> lambda)
        {
            var filteredTables = tables.Where(t => lambda(t));
            return filteredTables.Any() ? filteredTables : null;
        }

        public static IEnumerable<Column> GetFilteredColumns(IEnumerable<Column> columns, Func<Column, bool> lambda, bool returnAllIfNoneFound = true)
        {
            var filteredColumns = columns.Where(c => lambda(c));

            return filteredColumns.Any() || returnAllIfNoneFound ? filteredColumns : null;

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
        public static Measure SelectMeasure(Measure preselect = null, string label = "Select Table")
        {
            return ScriptHelper.SelectMeasure(preselect: preselect, label: label);
        }
        public static void Output(object value, int lineNumber = -1)
        {
            ScriptHelper.Output(value: value, lineNumber: lineNumber);
        }
    }
}
