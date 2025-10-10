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

        public static Function CreateFunction(
            Model model,
            string name,
            string expression,
            out bool functionCreated,
            string description = null,
            string annotationLabel = null,
            string annotationValue = null,
            string outputType = null,
            string nameTemplate = null,
            string formatString = null,
            string displayFolder = null,
            string outputDestination = null)
        {

            Function function = null as Function;
            functionCreated = false;

            var matchingFunctions = model.Functions.Where(f => f.GetAnnotation(annotationLabel) == annotationValue);
            if (matchingFunctions.Count() == 1)
            {
                return matchingFunctions.First();
            }
            else if (matchingFunctions.Count() == 0)
            {
                function = model.AddFunction(name);
                function.Expression = expression;
                function.Description = description;
                functionCreated = true;


            }
            else
            {
                Error("More than one function found with annoation " + annotationLabel + " value " + annotationValue);
                return null as Function;
            }

            if (!string.IsNullOrEmpty(annotationLabel) && !string.IsNullOrEmpty(annotationValue))
            {
                function.SetAnnotation(annotationLabel, annotationValue);
            }
            if (!string.IsNullOrEmpty(outputType))
            {
                function.SetAnnotation("outputType", outputType);
            }
            if (!string.IsNullOrEmpty(nameTemplate))
            {
                function.SetAnnotation("nameTemplate", nameTemplate);
            }
            if (!string.IsNullOrEmpty(formatString))
            {
                function.SetAnnotation("formatString", formatString);
            }
            if (!string.IsNullOrEmpty(displayFolder))
            {
                function.SetAnnotation("displayFolder", displayFolder);
            }
            if (!string.IsNullOrEmpty(outputDestination))
            {
                function.SetAnnotation("outputDestination", outputDestination);
            }
            return function;

        }
        public static Table CreateCalcTable(Model model, string tableName, string tableExpression = "FILTER({0},FALSE)")
        {
            return model.Tables.FirstOrDefault(t =>
                                string.Equals(t.Name, tableName, StringComparison.OrdinalIgnoreCase)) //case insensitive search
                                ?? model.AddCalculatedTable(tableName, tableExpression);
        }

        public static Measure CreateMeasure(
            Table table, 
            string measureName, 
            string measureExpression,
            out bool measureCreated,
            string formatString = null,
            string displayFolder = null,
            string description = null,
            string annotationLabel = null, 
            string annotationValue = null,
            bool isHidden = false)
        {
            measureCreated = false;
            var matchingMeasures = table.Measures.Where(m => m.GetAnnotation(annotationLabel) == annotationValue);
            if (matchingMeasures.Count() == 1)
            {
                return matchingMeasures.First();
            }
            else if (matchingMeasures.Count() == 0)
            {
                Measure measure = table.AddMeasure(measureName, measureExpression);
                measure.Description = description;
                measure.DisplayFolder = displayFolder;
                measure.FormatString = formatString;
                measureCreated = true;

                if (!string.IsNullOrEmpty(annotationLabel) && !string.IsNullOrEmpty(annotationValue))
                {
                    measure.SetAnnotation(annotationLabel, annotationValue);
                }
                measure.IsHidden = isHidden;
                return measure;
            }
            else
            {
                Error("More than one measure found with annoation " + annotationLabel + " value " + annotationValue);
                Output(matchingMeasures);
                return null as Measure;
            }
        }

        public static string GetNameFromUser(string Prompt, string Title, string DefaultResponse)
        {
            string response = Interaction.InputBox(Prompt, Title, DefaultResponse, 740, 400);
            return response;
        }

        public static bool IsAnswerYes(string question, string title = "Please confirm")
        {
            var result = MessageBox.Show(question, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            return result == DialogResult.Yes;
        }
        public static (IList<string> Values, string Type) SelectAnyObjects(Model model, string selectionType = null, string prompt1 = "select item type", string prompt2 = "select item(s)", string placeholderValue = "")
        {

            
            var returnEmpty = (Values: new List<string>(), Type: (string)null);

            if (prompt1.Contains("{0}"))
                prompt1 = string.Format(prompt1, placeholderValue ?? "");

            if(prompt2.Contains("{0}"))
                prompt2 = string.Format(prompt2, placeholderValue ?? "");


            if (selectionType == null)
            {
                IList<string> selectionTypeOptions = new List<string> { "Table", "Column", "Measure", "Scalar" };
                selectionType = ChooseString(selectionTypeOptions, label: prompt1, customWidth: 600);
            }

            if (selectionType == null) return returnEmpty;

            IList<string> selectedValues = new List<string>();
            switch (selectionType)
            {
                case "Table":
                    selectedValues = SelectTableMultiple(model, label: prompt2);
                    break;
                case "Column":
                    selectedValues = SelectColumnMultiple(model, label: prompt2);
                    break;
                case "Measure":

                    selectedValues = SelectMeasureMultiple(model: model, label: prompt2);
                    break;
                case "Scalar":
                    IList<string> scalarList = new List<string>();
                    scalarList.Add(GetNameFromUser(prompt2, "Scalar value", "0"));
                    selectedValues = scalarList;
                    break;
                default:
                    Error("Invalid selection type");
                    return returnEmpty;

            }
            if (selectedValues.Count == 0) return returnEmpty; 
            return (Values:selectedValues, Type:selectionType);
        }


        public static string ChooseString(IList<string> OptionList, string label = "Choose item", int customWidth = 400, int customHeight = 500)
        {
            return ChooseStringInternal(OptionList, MultiSelect: false, label: label, customWidth: customWidth, customHeight:customHeight) as string;
        }

        public static List<string> ChooseStringMultiple(IList<string> OptionList, string label = "Choose item(s)", int customWidth = 650, int customHeight = 550)
        {
            return ChooseStringInternal(OptionList, MultiSelect:true, label:label, customWidth: customWidth, customHeight: customHeight) as List<string>;
        }

        private static object ChooseStringInternal(IList<string> OptionList, bool MultiSelect, string label = "Choose item(s)", int customWidth = 400, int customHeight = 500)
        {
            Form form = new Form
            {
                Text =label,
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
                Height = 70,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(10)
            };

            Button selectAllButton = new Button { Text = "Select All", Visible = MultiSelect , Height = 50, Width = 150};
            Button selectNoneButton = new Button { Text = "Select None", Visible = MultiSelect, Height = 50, Width = 150 };
            Button okButton = new Button { Text = "OK", DialogResult = DialogResult.OK, Height = 50, Width = 100 };
            Button cancelButton = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Height = 50, Width = 100 };

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

            form.Width = customWidth;
            form.Height = customHeight;

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

        public static Table GetDateTable(Model model, string prompt = "Select Date Table")
        {
            var dateTables = GetDateTables(model);
            if (dateTables == null) {
                Table t = SelectTable(model.Tables, label: prompt);
                if(t == null)
                {
                    Error("No table selected");
                    return null;
                }
                if (IsAnswerYes(String.Format("Mark {0} as date table?",t.DaxObjectFullName)))
                {
                    t.DataCategory = "Time";
                    var dateColumns = t.Columns
                        .Where(c => c.DataType == DataType.DateTime)
                        .ToList();
                    if(dateColumns.Count == 0)
                    {
                        Error(String.Format(@"No date column detected in the table {0}. Please check that the table contains a date column",t.Name));
                        return null;
                    }
                    var keyColumn = SelectColumn(dateColumns, preselect:dateColumns.First(), label: "Select Date Column to be used as key column");
                    if(keyColumn == null)
                    {
                        Error("No key column selected");
                        return null;
                    }
                    keyColumn.IsKey = true;
                }

                return t;
            };
            if (dateTables.Count() == 1)
                return dateTables.First();

            Table dateTable = SelectTable(dateTables, label: prompt);
            if(dateTable == null)
            {
                Error("No table selected");
                return null;
            }
            return dateTable;
        }

        public static Column GetDateColumn(Table dateTable, string prompt = "Select Date Column")
        {
            var dateColumns = dateTable.Columns
                .Where(c => c.DataType == DataType.DateTime)
                .ToList();
            if(dateColumns.Count == 0)
            {
                Error(String.Format(@"No date column detected in the table {0}. Please check that the table contains a date column", dateTable.Name));
                return null;
            }

            if(dateColumns.Any(c => c.IsKey))
            {
                return dateColumns.First(c => c.IsKey);
            }

            Column dateColumn = null;
            if (dateColumns.Count() == 1)
            {
                dateColumn = dateColumns.First();
            }
            else
            {
                dateColumn = SelectColumn(dateColumns, label: prompt);
                if (dateColumn == null)
                {
                    Error("No column selected");
                    return null;
                }
            }

            return dateColumn;

        }


        public static IEnumerable<Table> GetFactTables(Model model)
        {
            IEnumerable<Table> factTables = model.Tables.Where(
                x => model.Relationships.Where(r => r.ToTable == x)
                        .All(r => r.ToCardinality == RelationshipEndCardinality.Many)
                    && model.Relationships.Where(r => r.FromTable == x)
                        .All(r => r.FromCardinality == RelationshipEndCardinality.Many)
                    && model.Relationships.Where(r => r.ToTable == x || r.FromTable == x).Any()); // at least one relationship

            if (!factTables.Any())
            {
                Error("No fact table detected in the model. Please check that the model contains relationships");
                return null;
            }
            return factTables;
        }

        public static Table GetFactTable(Model model, string prompt = "Select Fact Table")
        {
            Table factTable = null;
            var factTables = GetFactTables(model);
            if (factTables == null)
            {
               factTable = SelectTable(model.Tables, label: "This does not look like a star schema. Choose your fact table manually");
                if (factTable == null)
                {
                    Error("No table selected");
                    return null;
                }
                return factTable;
            };
            if (factTables.Count() == 1)
                return factTables.First();
            factTable = SelectTable(factTables, label: prompt);
            if (factTable == null)
            {
                Error("No table selected");
                return null;
            }
            return factTable;
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

        public static IList<string> SelectMeasureMultiple(Model model, IEnumerable<Measure> measures = null, string label = "Select Measure(s)")
        {
            measures ??= model.AllMeasures;
            IList<string> measureNames = measures.Select(m => m.DaxObjectFullName).ToList();
            IList<string> selectedMeasureNames = ChooseStringMultiple(measureNames, label: label);
            return selectedMeasureNames; 
            
        }

        public static IList<string> SelectColumnMultiple(Model model, IEnumerable<Column> columns = null, string label = "Select Columns(s)")
        {
            columns ??= model.AllColumns;
            IList<string> columnNames = columns.Select(m => m.DaxObjectFullName).ToList();
            IList<string> selectedColumnNames = ChooseStringMultiple(columnNames, label: label);
            return selectedColumnNames;

        }

        public static IList<string> SelectTableMultiple(Model model, IEnumerable<Table> Tables = null, string label = "Select Tables(s)", int customWidth = 400)
        {
            Tables ??= model.Tables;
            IList<string> TableNames = Tables.Select(m => m.DaxObjectFullName).ToList();
            IList<string> selectedTableNames = ChooseStringMultiple(TableNames, label: label, customWidth: customWidth);
            return selectedTableNames;

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
