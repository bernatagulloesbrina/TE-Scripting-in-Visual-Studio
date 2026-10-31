using Dax.Formatter;
using Dax.Formatter.Models;
using DaxUserDefinedFunction;
using GeneralFunctions; //Uncomment if you use the custom class, add reference to the project too.
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.VisualBasic;
using Newtonsoft.Json.Linq;
using Report.DTO;
using ReportFunctions;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.SqlTypes;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using TabularEditor;
using TabularEditor.Scripting;
using TabularEditor.TextServices;
using TabularEditor.TOMWrapper;
using TabularEditor.TOMWrapper.Utils;
using TabularEditor.UI;
using static DaxTokenizer;
using static Report.DTO.VisualDto;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;


// '2023-05-06 / B.Agullo / 
//coding environment for Tabular Editor C# Scripts
// see https://www.esbrina-ba.com/c-scripting-nirvana-effortlessly-use-visual-studio-as-your-coding-environment/ for reference on how to use it.

namespace TE_Scripting
{
    public class TE_Scripts
    {
        void SUM()
        {
            //using GeneralFunctions;
            // MACRO: Name: Sum of Columns
            // MACRO: ValidContexts: Column
            // MACRO: Tooltip: Creates a "Sum of <column>" measure for each selected column and hides the column
            // SKILL: name: sum-of-columns
            // SKILL: description: Create "Sum of <column>" measures for numeric columns of a semantic model. Use when the user asks to create sum or base measures from columns.
            // SKILL:
            // SKILL: Creates one measure "Sum of <column>" per column, in the column's table and display folder, with Properties annotations.
            // SKILL:
            // SKILL: ## Inputs (environment variables)
            // SKILL: Set these variables before running the macro. Work out each value from the conversation and the model (inspect the model if needed); if a value is unclear, ask the user instead of guessing.
            // SKILL: - SUM_COLUMNS (required): 'Table'[Column] refs separated by ';'. The numeric columns to sum.
            // SKILL:
            // SKILL: ## Notes
            // SKILL: Copies FormatString and DisplayFolder from the column and hides the base column.
            // '2026-10-28 / B.Agullo / Agent ready: in CLI mode the columns come from the SUM_COLUMNS environment variable
            // '2026-02-20 / B.Agullo / Added Properties annotation
            // '1990-01-01 / D.Otykier

#if TECLI
            var columns = Fx.GetColumnsFromEnv(Model, "SUM_COLUMNS");
            if (columns == null) return;
#else
            var columns = Selected.Columns.ToList();
#endif

            // Loop through all target columns:
            foreach(var c in columns)
            {
                var newMeasure = c.Table.AddMeasure(
                    "Sum of " + c.Name,                    // Name
                    "SUM(" + c.DaxObjectFullName + ")",    // DAX expression
                    c.DisplayFolder                        // Display Folder
                );
                
                // Set the format string on the new measure:
                newMeasure.FormatString = c.FormatString;

                // Provide some documentation:
                newMeasure.Description = "This measure is the sum of column " + c.DaxObjectFullName;
                newMeasure.SetAnnotation("Properties","SUM|" + c.Name + "|" + c.Table.Name); 
                newMeasure.SetAnnotation("PropertyNames","Aggregation|baseColumn|baseColumnTable"); 
                
                // Hide the base column:
                c.IsHidden = true;
            }
        }
        void managePropertyAnnotations()
        {
            // 2026-06-28 / B.Agullo / Grid editor to manage Properties / PropertyNames annotations across selected measures.
            //                         Rows = measures, columns = property names. Supports column rename (with merge on
            //                         duplicate name), adding columns, and copy/paste. On save every measure stores the
            //                         full column list in PropertyNames (empty segments where it has no value); completely
            //                         empty columns are dropped and '|' characters are sanitised out of names/values.

            //using System.Windows.Forms;
            //using System.Drawing;

#if TE3
                ScriptHelper.WaitFormVisible = false;
#endif

            const string propertiesAnnotationLabel = "Properties";
            const string propertyNamesAnnotationLabel = "PropertyNames";

            // Rows of the grid: only the selected measures
            List<Measure> orderedMeasures = Selected.Measures.ToList();
            if (orderedMeasures.Count == 0)
            {
                Error("Select one or more measures and try again.");
                return;
            }

            // Backing model: ordered column names + per-measure (columnName -> value) dictionary
            List<string> columnNames = new List<string>();
            Dictionary<Measure, Dictionary<string, string>> values = new Dictionary<Measure, Dictionary<string, string>>();
            bool sanitizedAny = false;

            string Sanitize(string s)
            {
                if (s == null) return "";
                if (s.IndexOf('|') >= 0)
                {
                    sanitizedAny = true;
                    s = s.Replace("|", "/");
                }
                return s;
            }

            // Initialise backing model from existing annotations
            foreach (Measure measure in orderedMeasures)
            {
                string propertiesAnnotation = measure.GetAnnotation(propertiesAnnotationLabel) ?? "";
                string propertyNamesAnnotation = measure.GetAnnotation(propertyNamesAnnotationLabel) ?? "";

                string[] props = propertiesAnnotation.Length == 0 ? new string[0] : propertiesAnnotation.Split('|');
                string[] names = propertyNamesAnnotation.Length == 0 ? new string[0] : propertyNamesAnnotation.Split('|');

                int count = Math.Max(props.Length, names.Length);
                Dictionary<string, string> measureValues = new Dictionary<string, string>();

                for (int i = 0; i < count; i++)
                {
                    string name = (i < names.Length && !string.IsNullOrWhiteSpace(names[i]))
                        ? names[i]
                        : String.Format("Property{0}", i + 1);
                    string value = i < props.Length ? props[i] : "";

                    if (!columnNames.Contains(name)) columnNames.Add(name);
                    measureValues[name] = value; // duplicate name within a measure: last value wins
                }

                values[measure] = measureValues;
            }

            // Build the dialog
            Form form = new Form();
            form.Text = "Manage Property Annotations";
            form.Width = 900;
            form.Height = 520;
            form.StartPosition = FormStartPosition.CenterScreen;

            System.Windows.Forms.Label info = new System.Windows.Forms.Label();
            info.Dock = DockStyle.Top;
            info.Height = 40;
            info.Padding = new System.Windows.Forms.Padding(8, 6, 8, 0);
            info.Text = "Each row is a measure, each column a property name. Double-click a column header (or use 'Rename Column') to rename it; "
                + "renaming to an existing column merges them. Use 'Add Column' to add a property. Copy/paste with Ctrl+C / Ctrl+V.";

            DataGridView grid = new DataGridView();
            grid.Dock = DockStyle.Fill;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.RowHeadersVisible = false;
            grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            grid.MultiSelect = true;
            grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            grid.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;

            // (Re)build the grid from the backing model
            void RebuildGrid()
            {
                grid.Columns.Clear();
                grid.Rows.Clear();

                DataGridViewTextBoxColumn measureColumn = new DataGridViewTextBoxColumn();
                measureColumn.HeaderText = "Measure";
                measureColumn.ReadOnly = true;
                measureColumn.Frozen = true;
                measureColumn.Width = 220;
                measureColumn.DefaultCellStyle.BackColor = System.Drawing.SystemColors.Control;
                grid.Columns.Add(measureColumn);

                foreach (string columnName in columnNames)
                {
                    DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
                    col.HeaderText = columnName;
                    col.Width = 140;
                    grid.Columns.Add(col);
                }

                foreach (Measure measure in orderedMeasures)
                {
                    int rowIndex = grid.Rows.Add();
                    DataGridViewRow row = grid.Rows[rowIndex];
                    row.Cells[0].Value = measure.Name;

                    Dictionary<string, string> measureValues = values[measure];
                    for (int c = 0; c < columnNames.Count; c++)
                    {
                        string columnName = columnNames[c];
                        string value;
                        row.Cells[c + 1].Value = measureValues.TryGetValue(columnName, out value) ? value : "";
                    }
                }
            }

            // Read the grid back into the backing model (grid columns currently match columnNames 1:1)
            void SyncGridToModel()
            {
                for (int r = 0; r < grid.Rows.Count; r++)
                {
                    Measure measure = orderedMeasures[r];
                    Dictionary<string, string> measureValues = new Dictionary<string, string>();

                    for (int c = 0; c < columnNames.Count; c++)
                    {
                        string columnName = columnNames[c];
                        object cellValue = grid.Rows[r].Cells[c + 1].Value;
                        string value = Sanitize(cellValue == null ? "" : cellValue.ToString());

                        if (measureValues.ContainsKey(columnName))
                        {
                            // merged (duplicate header) column: keep the first non-empty value
                            if (string.IsNullOrEmpty(measureValues[columnName]) && !string.IsNullOrEmpty(value))
                                measureValues[columnName] = value;
                        }
                        else
                        {
                            measureValues[columnName] = value;
                        }
                    }

                    values[measure] = measureValues;
                }
            }

            // Small input prompt (self-contained, no Error popup on cancel)
            string PromptText(string title, string prompt, string defaultValue)
            {
                using (Form inputForm = new Form())
                {
                    inputForm.Text = title;
                    inputForm.Width = 420;
                    inputForm.Height = 170;
                    inputForm.StartPosition = FormStartPosition.CenterParent;
                    inputForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                    inputForm.MinimizeBox = false;
                    inputForm.MaximizeBox = false;

                    System.Windows.Forms.Label label = new System.Windows.Forms.Label();
                    label.Text = prompt;
                    label.Left = 10;
                    label.Top = 10;
                    label.Width = 390;

                    TextBox textBox = new TextBox();
                    textBox.Left = 10;
                    textBox.Top = 35;
                    textBox.Width = 390;
                    textBox.Text = defaultValue;

                    Button okButton = new Button();
                    okButton.Text = "OK";
                    okButton.DialogResult = DialogResult.OK;
                    okButton.Left = 230;
                    okButton.Top = 75;
                    okButton.Width = 80;
                    okButton.Height = 30;

                    Button cancelButton = new Button();
                    cancelButton.Text = "Cancel";
                    cancelButton.DialogResult = DialogResult.Cancel;
                    cancelButton.Left = 320;
                    cancelButton.Top = 75;
                    cancelButton.Width = 80;
                    cancelButton.Height = 30;

                    inputForm.Controls.Add(label);
                    inputForm.Controls.Add(textBox);
                    inputForm.Controls.Add(okButton);
                    inputForm.Controls.Add(cancelButton);
                    inputForm.AcceptButton = okButton;
                    inputForm.CancelButton = cancelButton;

                    textBox.SelectAll();
                    return inputForm.ShowDialog() == DialogResult.OK ? textBox.Text : null;
                }
            }

            // Rename the column at gridColumnIndex (1-based against the data columns, i.e. excludes Measure column)
            void RenameColumn(int gridColumnIndex)
            {
                if (gridColumnIndex <= 0 || gridColumnIndex > columnNames.Count) return;

                SyncGridToModel();

                int modelIndex = gridColumnIndex - 1;
                string oldName = columnNames[modelIndex];

                string newNameRaw = PromptText("Rename Column", "Enter the new property name:", oldName);
                if (newNameRaw == null) return;

                string newName = Sanitize(newNameRaw).Trim();
                if (string.IsNullOrEmpty(newName) || newName == oldName) return;

                // Update the per-measure values: rename oldName -> newName, merging when needed
                foreach (Measure measure in orderedMeasures)
                {
                    Dictionary<string, string> measureValues = values[measure];
                    string oldValue;
                    bool hadOld = measureValues.TryGetValue(oldName, out oldValue);
                    measureValues.Remove(oldName);

                    if (measureValues.ContainsKey(newName))
                    {
                        if (string.IsNullOrEmpty(measureValues[newName]) && hadOld && !string.IsNullOrEmpty(oldValue))
                            measureValues[newName] = oldValue;
                    }
                    else
                    {
                        measureValues[newName] = hadOld ? oldValue : "";
                    }
                }

                // Update the column list: merge if the new name already exists elsewhere
                bool existsElsewhere = false;
                for (int i = 0; i < columnNames.Count; i++)
                {
                    if (i != modelIndex && columnNames[i] == newName) { existsElsewhere = true; break; }
                }

                if (existsElsewhere)
                    columnNames.RemoveAt(modelIndex);
                else
                    columnNames[modelIndex] = newName;

                RebuildGrid();
            }

            // Paste clipboard (tab/newline delimited) starting at the current cell
            void PasteClipboard()
            {
                if (!Clipboard.ContainsText()) return;
                string text = Clipboard.GetText();
                if (string.IsNullOrEmpty(text)) return;

                string[] lines = text.Replace("\r\n", "\n").Replace("\r", "\n").TrimEnd('\n').Split('\n');

                // Single copied value pasted across a multi-cell selection: fill every selected (editable) cell
                bool singleValue = lines.Length == 1 && lines[0].IndexOf('\t') < 0;
                if (singleValue && grid.SelectedCells.Count > 1)
                {
                    string single = Sanitize(lines[0]);
                    foreach (DataGridViewCell selectedCell in grid.SelectedCells)
                    {
                        if (selectedCell.ColumnIndex == 0) continue; // skip the read-only Measure column
                        selectedCell.Value = single;
                    }
                    return;
                }

                int startRow = grid.CurrentCell != null ? grid.CurrentCell.RowIndex : 0;
                int startCol = grid.CurrentCell != null ? grid.CurrentCell.ColumnIndex : 1;
                if (startCol < 1) startCol = 1; // never write into the read-only Measure column

                for (int r = 0; r < lines.Length; r++)
                {
                    int rowIdx = startRow + r;
                    if (rowIdx >= grid.Rows.Count) break;

                    string[] cells = lines[r].Split('\t');
                    for (int c = 0; c < cells.Length; c++)
                    {
                        int colIdx = startCol + c;
                        if (colIdx >= grid.Columns.Count) break;
                        if (colIdx == 0) continue;
                        grid.Rows[rowIdx].Cells[colIdx].Value = Sanitize(cells[c]);
                    }
                }
            }

            grid.ColumnHeaderMouseDoubleClick += (sender, e) =>
            {
                if (e.ColumnIndex >= 1) RenameColumn(e.ColumnIndex);
            };

            grid.KeyDown += (sender, e) =>
            {
                if (e.Control && e.KeyCode == Keys.V)
                {
                    PasteClipboard();
                    e.Handled = true;
                }
            };

            // Button bar
            Panel buttonPanel = new Panel();
            buttonPanel.Dock = DockStyle.Bottom;
            buttonPanel.Height = 52;

            Button addColumnButton = new Button();
            addColumnButton.Text = "Add Column";
            addColumnButton.Left = 8;
            addColumnButton.Top = 10;
            addColumnButton.Width = 110;
            addColumnButton.Height = 32;
            addColumnButton.Click += (sender, e) =>
            {
                SyncGridToModel();
                string defaultName = String.Format("Property{0}", columnNames.Count + 1);
                string nameRaw = PromptText("Add Column", "Enter the new property name:", defaultName);
                if (nameRaw == null) return;

                string name = Sanitize(nameRaw).Trim();
                if (string.IsNullOrEmpty(name)) return;

                if (columnNames.Contains(name))
                {
                    MessageBox.Show("A column with that name already exists.", "Add Column", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                columnNames.Add(name);
                foreach (Measure measure in orderedMeasures) values[measure][name] = "";
                RebuildGrid();
            };

            Button renameColumnButton = new Button();
            renameColumnButton.Text = "Rename Column";
            renameColumnButton.Left = 126;
            renameColumnButton.Top = 10;
            renameColumnButton.Width = 110;
            renameColumnButton.Height = 32;
            renameColumnButton.Click += (sender, e) =>
            {
                int colIndex = grid.CurrentCell != null ? grid.CurrentCell.ColumnIndex : -1;
                if (colIndex >= 1) RenameColumn(colIndex);
                else MessageBox.Show("Select a cell in the column you want to rename.", "Rename Column", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            Button saveButton = new Button();
            saveButton.Text = "Save";
            saveButton.DialogResult = DialogResult.OK;
            saveButton.Width = 90;
            saveButton.Top = 10;
            saveButton.Height = 32;

            Button cancelButton2 = new Button();
            cancelButton2.Text = "Cancel";
            cancelButton2.DialogResult = DialogResult.Cancel;
            cancelButton2.Width = 90;
            cancelButton2.Top = 10;
            cancelButton2.Height = 32;

            buttonPanel.Resize += (sender, e) =>
            {
                cancelButton2.Left = buttonPanel.Width - cancelButton2.Width - 8;
                saveButton.Left = cancelButton2.Left - saveButton.Width - 8;
            };

            buttonPanel.Controls.Add(addColumnButton);
            buttonPanel.Controls.Add(renameColumnButton);
            buttonPanel.Controls.Add(saveButton);
            buttonPanel.Controls.Add(cancelButton2);

            form.Controls.Add(grid);
            form.Controls.Add(info);
            form.Controls.Add(buttonPanel);
            form.AcceptButton = saveButton;
            form.CancelButton = cancelButton2;

            RebuildGrid();
            cancelButton2.Left = buttonPanel.Width - cancelButton2.Width - 8;
            saveButton.Left = cancelButton2.Left - saveButton.Width - 8;

            if (form.ShowDialog() != DialogResult.OK) return;

            // Persist: read final state, drop completely-empty columns, write full column list to every measure
            SyncGridToModel();

            List<string> keptColumns = new List<string>();
            foreach (string columnName in columnNames)
            {
                bool anyValue = orderedMeasures.Any(m =>
                {
                    string v;
                    return values[m].TryGetValue(columnName, out v) && !string.IsNullOrEmpty(v);
                });
                if (anyValue) keptColumns.Add(columnName);
            }

            int updatedCount = 0;
            foreach (Measure measure in orderedMeasures)
            {
                if (keptColumns.Count == 0)
                {
                    measure.RemoveAnnotation(propertiesAnnotationLabel);
                    measure.RemoveAnnotation(propertyNamesAnnotationLabel);
                    continue;
                }

                Dictionary<string, string> measureValues = values[measure];
                List<string> valuesList = new List<string>();
                foreach (string columnName in keptColumns)
                {
                    string v;
                    valuesList.Add(Sanitize(measureValues.TryGetValue(columnName, out v) ? v : ""));
                }

                string propertyNamesAnnotation = string.Join("|", keptColumns.Select(c => Sanitize(c)));
                string propertiesAnnotation = string.Join("|", valuesList);

                measure.SetAnnotation(propertyNamesAnnotationLabel, propertyNamesAnnotation);
                measure.SetAnnotation(propertiesAnnotationLabel, propertiesAnnotation);
                updatedCount++;
            }

            string summary = keptColumns.Count == 0
                ? String.Format("Cleared Property annotations on {0} measure(s).", orderedMeasures.Count)
                : String.Format("Updated {0} measure(s) with {1} column(s): {2}", updatedCount, keptColumns.Count, string.Join(", ", keptColumns));

            if (sanitizedAny)
                summary += Environment.NewLine + Environment.NewLine + "Note: '|' characters are reserved as separators and were replaced with '/'.";

            Info(summary);
        }
        void referentialIntegrity()
        {

            //Select the desired table to store all data quality measures
            //See https://www.esbrina-ba.com/easy-management-of-referential-integrity/

            // 2026-02-20 / B.Agulló / GetSelectedTable to use without creating first the calc table
            // 2024-07-13 / B.Agulló / Add annotations for the report-layer script, possible to execute without selected table, subtitles measures
            // 2023-12-22 / B.Agulló / Added suggestions by Ed Hansberry 
            // 2023-02-15 / B.Agulló / Added useRelationship in the expression to check also inactive relationships

            // Instructions:
            // Select the desired table  to store all data quality measures
            // or execute just on the model and a new table will be created for you
            //
            // See https://www.esbrina-ba.com/easy-management-of-referential-integrity/
            //
            // To create the report sheet check out:
            // https://www.esbrina-ba.com/building-a-referential-integrity-report-page-with-a-c-script/

            //change the resulting variable names if you want

#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif

            string overallCounterName = "Total Unmapped Items";
            string overallDetailName = "Data Problems";
            string targetTableNameIfCreated = "Referential Integrity";
            //do not modify the script below
            string overallCounterExpression = "";
            string overallDetailExpression = "\"\"";
            string annLabel = "ReferencialIntegrityMeasures";
            string annValueTotal = "TotalUnmappedItems";
            string annValueDetail = "DataProblems";
            string annValueDataQualityMeasures = "DataQualityMeasure";
            string annValueDataQualityTitles = "DataQualityTitle";
            string annValueDataQualitySubtitles = "DataQualitySubitle";
            string annValueFactColumn = "FactColumn";
            Table tableToStoreMeasures = null as Table;
            if (Selected.Tables.Count() == 0)
            {
                tableToStoreMeasures = Model.AddCalculatedTable(targetTableNameIfCreated, "{0}");
            }
            else
            {
                tableToStoreMeasures = Selected.Tables.First();
            }
            int measureIndex = 0;
            foreach (var r in Model.Relationships)
            {
                bool isOneToMany =
                    r.FromCardinality == RelationshipEndCardinality.One
                    && r.ToCardinality == RelationshipEndCardinality.Many;
                bool isManyToOne =
                    r.FromCardinality == RelationshipEndCardinality.Many
                    && r.ToCardinality == RelationshipEndCardinality.One;
                Column manyColumn = null as Column;
                Column oneColumn = null as Column;
                bool isOneToManyOrManyToOne = true;
                if (isOneToMany)
                {
                    manyColumn = r.ToColumn;
                    oneColumn = r.FromColumn;
                }
                else if (isManyToOne)
                {
                    manyColumn = r.FromColumn;
                    oneColumn = r.ToColumn;
                }
                else
                {
                    isOneToManyOrManyToOne = false;
                }
                if (isOneToManyOrManyToOne)
                {
                    measureIndex++; //increment index
                                    //add measure counting how many different items in the fact table are not present in the dimension
                    string orphanCountExpression =
                        "CALCULATE("
                            + "SUMX(VALUES(" + manyColumn.DaxObjectFullName + "),1),"
                            + "ISBLANK(" + oneColumn.DaxObjectFullName + "),"
                            + "USERELATIONSHIP(" + manyColumn.DaxObjectFullName + "," + oneColumn.DaxObjectFullName + "),"
                            + "ALLEXCEPT(" + manyColumn.Table.DaxObjectFullName + "," + manyColumn.DaxObjectFullName + ")"
                        + ")";
                    string orphanMeasureName =
                        manyColumn.Name + " not mapped in " + manyColumn.Table.Name;
                    Measure newCounter = tableToStoreMeasures.AddMeasure(name: orphanMeasureName, expression: orphanCountExpression, displayFolder: "_Data quality Measures");
                    newCounter.FormatString = "#,##0";
                    newCounter.FormatDax();
                    newCounter.SetAnnotation(annLabel, annValueDataQualityMeasures + "_" + measureIndex.ToString());
                    //add annotation to trace back the fact table when building the report
                    manyColumn.SetAnnotation(annLabel, annValueFactColumn + "_" + measureIndex.ToString());
                    //add measure saying how many are not mapped in the fact table
                    string orphanTableTitleMeasureExpression = "FORMAT(" + newCounter.DaxObjectFullName + "+0,\"" + newCounter.FormatString + "\") & \" " + newCounter.Name + "\"";
                    string orphanTableTitleMeasureName = newCounter.Name + " Title";
                    Measure newTitle = tableToStoreMeasures.AddMeasure(name: orphanTableTitleMeasureName, expression: orphanTableTitleMeasureExpression, displayFolder: "_Data quality Titles");
                    newTitle.FormatDax();
                    newTitle.SetAnnotation(annLabel, annValueDataQualityTitles + "_" + measureIndex.ToString());
                    //add measure for subtitle saying how many need to be added to the dimension table
                    string orphanTableSubtitleMeasureExpression =
                        String.Format(
                            @"FORMAT({0}+0,""{1}"") & "" values missing in "" & ""{2}""",
                            newCounter.DaxObjectFullName,
                            newCounter.FormatString,
                            oneColumn.Table.Name);
                    string orphanTableSubitleMeasureName = newCounter.Name + " Subtitle";
                    Measure newSubtitle = tableToStoreMeasures.AddMeasure(name: orphanTableSubitleMeasureName, expression: orphanTableSubtitleMeasureExpression, displayFolder: "_Data quality Subtitles");
                    newSubtitle.FormatDax();
                    newSubtitle.SetAnnotation(annLabel, annValueDataQualitySubtitles + "_" + measureIndex.ToString());
                    overallCounterExpression = overallCounterExpression + "+" + newCounter.DaxObjectFullName;
                    overallDetailExpression = overallDetailExpression
                            + " & IF(" + newCounter.DaxObjectFullName + "> 0,"
                                        + newTitle.DaxObjectFullName + " & UNICHAR(10))";
                }
                ;
            }
            ;
            Measure counter = tableToStoreMeasures.AddMeasure(name: overallCounterName, expression: overallCounterExpression);
            counter.FormatString = "#,##0";
            counter.FormatDax();
            counter.SetAnnotation(annLabel, annValueTotal);
            Measure descr = tableToStoreMeasures.AddMeasure(name: overallDetailName, expression: overallDetailExpression);
            descr.FormatDax();
            descr.SetAnnotation(annLabel, annValueDetail);

        }

        void createFilteredMeasures()
        {
            // 2026-05-06 / B.Agullo / fixed null-reference bug
            //  2026-03-29 / B.Agullo / added support for propertyNames annotation
            // '2026-03-20 / B.Agullo / added template for name, displayfolder, description and properties annotation
            // '2026-02-20 / B.Agullo / added properties annotation for fieldParameter and dynamic measure creation
            // '2022-05-21 / B.Agullo / 
            // FILTERED MEASURES BY COLUMN VALUES SCRIPT 
            // creates a measure for each of the values in a column filtering the selected base measure
            // step by step instructions at https://www.esbrina-ba.com/creating-filtered-measures-or-how-to-show-the-total-along-with-the-detail-in-a-chart/

            //using GeneralFunctions;
#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif
            var measures = Selected.Measures;

            if (measures.Count == 0)
            {
                Error("Select one or more measures");
            }

            Table table = SelectTable();
            if (table == null)
            {
                Error("Select a table");
                return;
            }
            Column column = SelectColumn(table);
            if (column == null)
            {
                Error("Select a column");
                return;
            }

            string measureNameTemplate = Fx.GetNameFromUser("Enter measure name template. You can use measureName, columnName, columnValue in the expression", "template", "measureName - columnValue");
            if (string.IsNullOrEmpty(measureNameTemplate)) return;

            string measureNameReplaceExpression =
                measureNameTemplate
                    .Replace("measureName", "{0}")
                    .Replace("columnName", "{1}")
                    .Replace("columnValue", "{2}");

            string measureDescriptionTemplate = Fx.GetNameFromUser("Enter measure description template. You can use measureName, columnName, columnValue in the expression", "template", "measureName filtered by columnName = columnValue");
            if (string.IsNullOrEmpty(measureDescriptionTemplate)) return;

            string measureDescriptionReplaceExpression =
                measureDescriptionTemplate
                    .Replace("measureName", "{0}")
                    .Replace("columnName", "{1}")
                    .Replace("columnValue", "{2}");

            string displayFolderTemplate = Fx.GetNameFromUser("Enter display folder template. You can use measureName, columnName, columnValue in the expression", "template", "measureName - columnName");
            if (string.IsNullOrEmpty(displayFolderTemplate)) return;

            string displayFolderReplaceExpression =
                displayFolderTemplate
                    .Replace("measureName", "{0}")
                    .Replace("columnName", "{1}")
                    .Replace("columnValue", "{2}");



            string propertiesAnnotationTemplate = Fx.GetNameFromUser("Enter properties annotation template. You can use measureName, measureProperties, columnName, columnValue in the expression", "template", "measureName|columnName|columnValue");
            if (string.IsNullOrEmpty(propertiesAnnotationTemplate)) return;

            string propertiesAnnotationReplaceExpression =
                propertiesAnnotationTemplate
                    .Replace("measureName", "{0}")
                    .Replace("measureProperties", "{1}")
                    .Replace("columnName", "{2}")
                    .Replace("columnValue", "{3}");

            string query = "EVALUATE DISTINCT(" + column.DaxObjectFullName + ")";

            using (var reader = Model.Database.ExecuteReader(query))
            {
                // Create a loop for every row in the resultset
                while (reader.Read())
                {
                    var rawValue = reader.GetValue(0);
                    if (rawValue == null || rawValue == DBNull.Value) continue;
                    string columnValue = rawValue.ToString();

                    string formulaColumnValue = columnValue;




                    if (column.DataType.Equals(DataType.String))
                    {
                        formulaColumnValue = "\"" + columnValue + "\"";
                    }


                    foreach (Measure measure in measures)
                    {
                        string measureName = string.Format(measureNameReplaceExpression, measure.Name, column.Name, columnValue);
                        string measureExpression =
                            string.Format("CALCULATE({0},{1}={2})",
                                measure.DaxObjectName,
                                column.DaxObjectFullName,
                                formulaColumnValue
                            );
                        string measureDescription =
                            string.Format(measureDescriptionReplaceExpression,
                                measure.Name,
                                column.Name,
                                columnValue
                           );
                        string displayFolderName =
                            string.Format(displayFolderReplaceExpression,
                                measure.Name,
                                column.Name,
                                columnValue
                            );

                        string propertiesAnnotation =
                            string.Format(propertiesAnnotationReplaceExpression,
                                measure.Name,
                                measure.GetAnnotation("Properties") ?? "",
                                column.Name,
                                columnValue
                            );

                        Measure newMeasure =
                            measure.Table.AddMeasure(
                                name: measureName,
                                expression: measureExpression,
                                displayFolder: displayFolderName
                            );
                        newMeasure.Description = measureDescription;
                        newMeasure.FormatDax();
                        newMeasure.FormatString = measure.FormatString;
                        newMeasure.SetAnnotation("Properties", propertiesAnnotation);
                        newMeasure.SetAnnotation("PropertyNames", propertiesAnnotationTemplate);

                    }
                }
            }
        }
        void createDynamicMeasure()
        {
            //using GeneralFunctions;
            // 2026-05-30 / B.Agullo / FormatString values with embedded double-quotes are now escaped (doubled) when written into a DAX string literal
            // 2026-05-30 / B.Agullo / Fixed missing comma between SWITCH condition and its result value in all generated DAX expressions
            // 2026-05-30 / B.Agullo / PropertyNames with spaces are now sanitised (spaces removed) before use as DAX variable names; display names are unchanged
            // 2026-04-23 / B.Agullo /  semidynamic measures have also propertyNames annotation
            // 2026-04-22 / B.Agullo / improvements in the disconnected tables, with distinct key and display values if needed. 
            // 2026-03-29 / B.Agullo / added support for propertyNames annotation
            // 2025-01-07 / B.Agullo
            // Creates a dynamic measure and disconnected tables based on Properties annotations of selected measures.
            // Each property position gets its own disconnected table with distinct values across all measures.
            // The dynamic measure returns the appropriate measure based on selected values in each disconnected table.

#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif

            // Configuration: Prefix for disconnected tables (to group them together in Power BI Desktop)
            string tablePrefix = "Selection ";

            // Validate selection
            if (Selected.Measures.Count() < 2)
            {
                Error("Select at least two measures with Properties annotations and try again.");
                return;
            }

            // Step 1: Extract and validate Properties annotations
            var measuresWithProperties = new List<(Measure measure, List<string> properties)>();

            foreach (var measure in Selected.Measures)
            {
                string propertiesAnnotation = measure.GetAnnotation("Properties");
                if (string.IsNullOrEmpty(propertiesAnnotation))
                {
                    Error(String.Format("Measure '{0}' does not have a Properties annotation.", measure.Name));
                    return;
                }

                var properties = propertiesAnnotation.Split('|').ToList();
                measuresWithProperties.Add((measure, properties));
            }

            // Step 2: Validate all have the same length
            int propertyCount = measuresWithProperties[0].properties.Count;
            if (measuresWithProperties.Any(m => m.properties.Count != propertyCount))
            {
                Error("All Properties annotations must have the same number of pipe-separated values.");
                return;
            }

            // Step 3: Validate all Properties annotations are unique
            var allPropertiesStrings = measuresWithProperties.Select(m => string.Join("|", m.properties)).ToList();
            if (allPropertiesStrings.Count != allPropertiesStrings.Distinct().Count())
            {
                Error("All Properties annotations must be unique. Some measures have duplicate Properties.");
                return;
            }

            // Step 3.5: Retrieve PropertyNames annotations from selected measures
            var propertyNamesAnnotations = new List<string>();
            foreach (var measure in Selected.Measures)
            {
                string propertyNamesAnnotation = measure.GetAnnotation("PropertyNames") ?? "";

                if (!string.IsNullOrEmpty(propertyNamesAnnotation))
                {
                    propertyNamesAnnotations.Add(propertyNamesAnnotation);
                }
            }

            // Determine property names for disconnected tables and variables
            List<string> propertyNames = new List<string>();

            if (propertyNamesAnnotations.Count > 0)
            {
                // Get distinct PropertyNames values
                var distinctPropertyNames = propertyNamesAnnotations.Distinct().ToList();

                if (distinctPropertyNames.Count == 1)
                {
                    // All PropertyNames are the same, use them
                    propertyNames = distinctPropertyNames[0].Split('|').ToList();
                }
                else
                {
                    // Different PropertyNames found, let user choose
                    var optionsWithDefine = distinctPropertyNames.ToList();
                    optionsWithDefine.Add("<Define now>");

                    string chosenPropertyNames = Fx.ChooseString(
                        OptionList: optionsWithDefine,
                        label: "Different PropertyNames found. Choose one or define new:"
                    );

                    if (string.IsNullOrEmpty(chosenPropertyNames)) return;

                    if (chosenPropertyNames == "<Define now>")
                    {
                        // Show input box with first PropertyNames as default
                        chosenPropertyNames = Fx.GetNameFromUser(
                            Prompt: "Enter property names (pipe-separated):",
                            Title: "Property Names",
                            DefaultResponse: distinctPropertyNames[0]
                        );

                        if (string.IsNullOrEmpty(chosenPropertyNames)) return;
                    }

                    propertyNames = chosenPropertyNames.Split('|').ToList();
                }
            }
            else
            {
                // No PropertyNames found, ask one by one
                for (int i = 0; i < propertyCount; i++)
                {
                    // Collect distinct values for this property position (up to 3)
                    var sampleValues = measuresWithProperties
                        .Select(m => m.properties[i])
                        .Distinct()
                        .Take(3)
                        .ToList();

                    string sampleText = sampleValues.Count > 0
                        ? String.Format(" (found values: '{0}')", string.Join("', '", sampleValues))
                        : "";

                    string propertyName = Fx.GetNameFromUser(
                        Prompt: String.Format("Enter name for property {0}{1}:", i + 1, sampleText),
                        Title: "Property Name",
                        DefaultResponse: String.Format("Property{0}", i + 1)
                    );

                    if (string.IsNullOrEmpty(propertyName)) return;

                    propertyNames.Add(propertyName);
                }
            }

            // Step 4: Ask user for dynamic measure name with smart default based on PropertyNames
            string defaultDynamicMeasureName = propertyNames.Count > 0
                ? "Dynamic Measure " + string.Join(" ", propertyNames)
                : "Dynamic Measure";

            string dynamicMeasureName = Fx.GetNameFromUser(
                Prompt: "Enter the name for the dynamic measure:",
                Title: "Dynamic Measure Name",
                DefaultResponse: defaultDynamicMeasureName
            );
            if (dynamicMeasureName == null) return;

            // DAX variable names cannot contain spaces – strip them from propertyNames
            List<string> daxSafePropertyNames = propertyNames.Select(p => p.Replace(" ", "")).ToList();

            // Step 5: Get destination table (use first measure's table)
            Table destinationTable = measuresWithProperties[0].measure.Table;

            // Step 6: Create disconnected tables for each property position
            var propertyTables = new List<Table>();
            var columnBaseNames = new List<string>();

            for (int i = 0; i < propertyCount; i++)
            {
                // Collect distinct values for this property position across all measures
                var distinctValues = measuresWithProperties
                    .Select(m => m.properties[i])
                    .Distinct()
                    .OrderBy(v => v)
                    .ToList();

                // Create table name with prefix and column base name without prefix
                string columnBaseName = propertyNames.Count > i ? propertyNames[i] : String.Format("Property{0}", i + 1);
                string tableName = tablePrefix + columnBaseName;

                // Build DAX expression for calculated table with 3 columns (Label, Key, Order)
                string tableDax = "{" + Environment.NewLine;
                for (int j = 0; j < distinctValues.Count; j++)
                {
                    tableDax += String.Format("    (\"{0}\", \"{1}\", {2})", distinctValues[j], distinctValues[j], j);
                    if (j < distinctValues.Count - 1)
                        tableDax += "," + Environment.NewLine;
                }
                tableDax += Environment.NewLine + "}";

                // Create the calculated table
                var propTable = Model.AddCalculatedTable(tableName, tableDax);

                // Setup columns (handle both TE2 and TE3) - use columnBaseName without prefix
                var te2 = propTable.Columns.Count == 0;
                var labelColumn = te2
                    ? propTable.AddCalculatedTableColumn(columnBaseName, "[Value1]")
                    : propTable.Columns["Value1"] as CalculatedTableColumn;
                var keyColumn = te2
                    ? propTable.AddCalculatedTableColumn(columnBaseName + " Key", "[Value2]")
                    : propTable.Columns["Value2"] as CalculatedTableColumn;
                var orderColumn = te2
                    ? propTable.AddCalculatedTableColumn(columnBaseName + " Order", "[Value3]")
                    : propTable.Columns["Value3"] as CalculatedTableColumn;

                if (!te2)
                {
                    labelColumn.IsNameInferred = false;
                    labelColumn.Name = columnBaseName;
                    keyColumn.IsNameInferred = false;
                    keyColumn.Name = columnBaseName + " Key";
                    orderColumn.IsNameInferred = false;
                    orderColumn.Name = columnBaseName + " Order";
                }

                labelColumn.SortByColumn = orderColumn;
                orderColumn.IsHidden = true;

                propertyTables.Add(propTable);
                columnBaseNames.Add(columnBaseName);
            }

            // Step 6.5: Create partially dynamic measures in each property table
            int totalPartialMeasures = 0;

            for (int tableIndex = 0; tableIndex < propertyTables.Count; tableIndex++)
            {
                var currentPropertyTable = propertyTables[tableIndex];

                // Get distinct values for this property table
                var distinctValuesForThisProperty = measuresWithProperties
                    .Select(m => m.properties[tableIndex])
                    .Distinct()
                    .ToList();

                foreach (var propertyValue in distinctValuesForThisProperty)
                {
                    // Filter measures that have this property value at this position
                    var matchingMeasures = measuresWithProperties
                        .Where(m => m.properties[tableIndex] == propertyValue)
                        .ToList();

                    // Build measure expression
                    string partialMeasureExpression = "";

                    // Add variable declarations for OTHER properties (not current table)
                    for (int i = 0; i < propertyCount; i++)
                    {
                        if (i == tableIndex) continue;

                        string varName = daxSafePropertyNames.Count > i ? daxSafePropertyNames[i] : columnBaseNames[i].Replace(" ", "");
                        partialMeasureExpression += String.Format(
                            "VAR __{0} = SELECTEDVALUE( '{1}'[{2} Key] )" + Environment.NewLine,
                            varName,
                            propertyTables[i].Name,
                            columnBaseNames[i]
                        );
                    }

                    if (propertyCount > 1)
                    {
                        partialMeasureExpression += "RETURN" + Environment.NewLine;
                        partialMeasureExpression += "SWITCH(" + Environment.NewLine + "    TRUE()," + Environment.NewLine;

                        // Add cases only for matching measures
                        foreach (var item in matchingMeasures)
                        {
                            string condition = "    ";
                            bool firstCondition = true;

                            for (int i = 0; i < propertyCount; i++)
                            {
                                if (i == tableIndex) continue;

                                if (!firstCondition)
                                    condition += Environment.NewLine + "        && ";

                                string varName = daxSafePropertyNames.Count > i ? daxSafePropertyNames[i] : columnBaseNames[i].Replace(" ", "");
                                condition += String.Format(
                                    "__{0} = \"{1}\"",
                                    varName,
                                    item.properties[i]
                                );

                                firstCondition = false;
                            }

                            partialMeasureExpression += condition + "," + Environment.NewLine;
                            partialMeasureExpression += String.Format("        {0}," + Environment.NewLine, item.measure.DaxObjectFullName);
                        }

                        partialMeasureExpression += "    BLANK()" + Environment.NewLine + ")";
                    }
                    else
                    {
                        // Only one property, just return the measure directly
                        partialMeasureExpression += "RETURN" + Environment.NewLine;
                        partialMeasureExpression += matchingMeasures[0].measure.DaxObjectFullName;
                    }

                    // Check if measure name already exists and add suffix if needed
                    string partialMeasureName = propertyValue;
                    if (Model.AllMeasures.Any(m => m.Name == propertyValue))
                    {
                        partialMeasureName = propertyValue + " (dynamic)";
                    }

                    // Create the measure
                    Measure partialMeasure = currentPropertyTable.AddMeasure(partialMeasureName, partialMeasureExpression);
                    partialMeasure.FormatDax();
                    partialMeasure.SetAnnotation("Properties", propertyValue);
                    partialMeasure.SetAnnotation("PropertyNames", propertyNames.Count > tableIndex ? propertyNames[tableIndex] : String.Format("Property{0}", tableIndex + 1));

                    // Build format string expression
                    string partialFormatExpression = "";

                    // Add variable declarations for OTHER properties (not current table)
                    for (int i = 0; i < propertyCount; i++)
                    {
                        if (i == tableIndex) continue;

                        string varName = daxSafePropertyNames.Count > i ? daxSafePropertyNames[i] : columnBaseNames[i].Replace(" ", "");
                        partialFormatExpression += String.Format(
                            "VAR __{0} = SELECTEDVALUE( '{1}'[{2} Key] )" + Environment.NewLine,
                            varName,
                            propertyTables[i].Name,
                            columnBaseNames[i]
                        );
                    }

                    if (propertyCount > 1)
                    {
                        partialFormatExpression += "RETURN" + Environment.NewLine;
                        partialFormatExpression += "SWITCH(" + Environment.NewLine + "    TRUE()," + Environment.NewLine;

                        // Add cases only for matching measures
                        foreach (var item in matchingMeasures)
                        {
                            string condition = "    ";
                            bool firstCondition = true;

                            for (int i = 0; i < propertyCount; i++)
                            {
                                if (i == tableIndex) continue;

                                if (!firstCondition)
                                    condition += Environment.NewLine + "        && ";

                                string varName = daxSafePropertyNames.Count > i ? daxSafePropertyNames[i] : columnBaseNames[i].Replace(" ", "");
                                condition += String.Format(
                                    "__{0} = \"{1}\"",
                                    varName,
                                    item.properties[i]
                                );

                                firstCondition = false;
                            }

                            partialFormatExpression += condition + "," + Environment.NewLine;

                            // Use format string or format string expression
                            string formatValue = !string.IsNullOrEmpty(item.measure.FormatStringExpression)
                                ? item.measure.FormatStringExpression
                                : String.Format("\"{0}\"", item.measure.FormatString.Replace("\"", "\"\""));

                            partialFormatExpression += String.Format("        {0}," + Environment.NewLine, formatValue);
                        }

                        partialFormatExpression += "    \"\"" + Environment.NewLine + ")";
                    }
                    else
                    {
                        // Only one property, use the measure's format directly
                        partialFormatExpression += "RETURN" + Environment.NewLine;
                        string formatValue = !string.IsNullOrEmpty(matchingMeasures[0].measure.FormatStringExpression)
                            ? matchingMeasures[0].measure.FormatStringExpression
                            : String.Format("\"{0}\"", matchingMeasures[0].measure.FormatString.Replace("\"", "\"\""));
                        partialFormatExpression += formatValue;
                    }

                    // Set format string expression
                    partialMeasure.FormatStringExpression = partialFormatExpression;

                    totalPartialMeasures++;
                }
            }

            // Step 7: Build dynamic measure expression
            string measureExpression = "";

            // Add variable declarations for each property
            for (int i = 0; i < propertyCount; i++)
            {
                string varName = daxSafePropertyNames.Count > i ? daxSafePropertyNames[i] : columnBaseNames[i].Replace(" ", "");
                measureExpression += String.Format(
                    "VAR __{0} = SELECTEDVALUE( '{1}'[{2} Key] )" + Environment.NewLine,
                    varName,
                    propertyTables[i].Name,
                    columnBaseNames[i]
                );
            }

            measureExpression += "RETURN" + Environment.NewLine;
            measureExpression += "SWITCH(" + Environment.NewLine + "    TRUE()," + Environment.NewLine;

            foreach (var item in measuresWithProperties)
            {
                // Build condition for this measure with variables
                string condition = "    ";
                for (int i = 0; i < propertyCount; i++)
                {
                    if (i > 0)
                        condition += Environment.NewLine + "        && ";

                    string varName = daxSafePropertyNames.Count > i ? daxSafePropertyNames[i] : columnBaseNames[i].Replace(" ", "");
                    condition += String.Format(
                        "__{0} = \"{1}\"",
                        varName,
                        item.properties[i]
                    );
                }

                measureExpression += condition + "," + Environment.NewLine;
                measureExpression += String.Format("        {0}," + Environment.NewLine, item.measure.DaxObjectFullName);
            }

            measureExpression += "    BLANK()" + Environment.NewLine + ")";

            // Step 8: Create the dynamic measure
            Measure dynamicMeasure = destinationTable.AddMeasure(dynamicMeasureName, measureExpression);
            dynamicMeasure.FormatDax();

            // Step 9: Build dynamic format string expression
            string formatExpression = "";

            // Add variable declarations for each property
            for (int i = 0; i < propertyCount; i++)
            {
                string varName = daxSafePropertyNames.Count > i ? daxSafePropertyNames[i] : columnBaseNames[i].Replace(" ", "");
                formatExpression += String.Format(
                    "VAR __{0} = SELECTEDVALUE( '{1}'[{2} Key] )" + Environment.NewLine,
                    varName,
                    propertyTables[i].Name,
                    columnBaseNames[i]
                );
            }

            formatExpression += "RETURN" + Environment.NewLine;
            formatExpression += "SWITCH(" + Environment.NewLine + "    TRUE()," + Environment.NewLine;

            foreach (var item in measuresWithProperties)
            {
                // Build condition for this measure with variables
                string condition = "    ";
                for (int i = 0; i < propertyCount; i++)
                {
                    if (i > 0)
                        condition += Environment.NewLine + "        && ";

                    string varName = daxSafePropertyNames.Count > i ? daxSafePropertyNames[i] : columnBaseNames[i].Replace(" ", "");
                    condition += String.Format(
                        "__{0} = \"{1}\"",
                        varName,
                        item.properties[i]
                    );
                }

                formatExpression += condition + "," + Environment.NewLine;

                // Use format string or format string expression
                string formatValue = !string.IsNullOrEmpty(item.measure.FormatStringExpression)
                    ? item.measure.FormatStringExpression
                    : String.Format("\"{0}\"", item.measure.FormatString.Replace("\"", "\"\""));

                formatExpression += String.Format("        {0}," + Environment.NewLine, formatValue);
            }

            formatExpression += "    \"\"" + Environment.NewLine + ")";

            // Set format string expression
            dynamicMeasure.FormatStringExpression = formatExpression;

            Output(String.Format(
                "Created dynamic measure '{0}' with {1} property tables, {2} partial measures, and {3} source measures.",
                dynamicMeasureName,
                propertyTables.Count,
                totalPartialMeasures,
                measuresWithProperties.Count
            ));
        }


        void SingleMeasureParameterFunction()
        {
            //2026-10-15 / B.Agullo / added support for dataCategory annotation (measures only)
            //2026-03-29 / B.Agullo / added support for propertyNames annotation
            //2026-02-15 / B.Agullo / improved default function annotation values and added support for output properties annotation.
            //2025-09-26/B.Agullo/ fixed bug that would not store annotations if initialized during runtime
            //2025-09-16/B.Agullo/
            //Creates measures based on DAX UDFs 
            //Check the blog post for futher information: https://www.esbrina-ba.com/automatically-create-measures-with-dax-user-defined-functions/

            //using GeneralFunctions;
            //using DaxUserDefinedFunction;
            //using System.Text.RegularExpressions;


#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif

            // PSEUDOCODE / PLAN:
            // 1. Verify that the user has selected one or more functions (Selected.Functions).
            // 2. If none selected, show error and abort.
            // 3. Create FunctionExtended objects for each selected function and keep them in a list.
            // 4. Extract all parameters from the selected functions and build a distinct list by name.
            // 5. For each distinct parameter name:
            //      - Prompt the user once with Fx.SelectAnyObjects to choose the objects to iterate for that parameter.
            //      - If the user cancels or selects nothing, abort the whole operation.
            //      - Store the resulting IList<string> in a dictionary keyed by parameter name so it can be retrieved later.
            // 6. Example usage: create a sample FunctionExtended (as the original example does),
            //    then when iterating the function parameters use the previously built dictionary to obtain the list
            //    of objects for each parameter name (do not prompt again).
            // 7. Build measure names/expressions by iterating the parameter-object combinations and create measures.
            //
            // NOTES:
            // - All Fx.SelectAnyObjects calls must use parameter names on the call.
            // - The dictionary is Dictionary<string, IList<string>> parameterObjectsMap.
            // - Abort execution if any required selection is cancelled.

            // Validate selection
            if (Selected.Functions.Count() == 0)
            {
                Error("Select one or more functions and try again.");
                return;
            }

            // Create FunctionExtended objects for each selected function and store them for later iteration
            IList<FunctionExtended> selectedFunctions = new List<FunctionExtended>();

            foreach (var f in Selected.Functions)
            {
                // Create the FunctionExtended and add to list
                FunctionExtended fe = FunctionExtended.CreateFunctionExtended(f);
                if (fe == null) return;
                selectedFunctions.Add(fe);
            }

            // Flatten all parameters from selected functions
            var allParametersFlat = selectedFunctions
                .SelectMany(sf => sf.Parameters ?? new List<FunctionParameter>())
                .ToList();

            // Build distinct FunctionParameter objects (first occurrence per name)
            IList<FunctionParameter> distinctParameters = allParametersFlat
                .GroupBy(p => p.Name)
                .Select(g => g.First())
                .ToList();

            // For each distinct parameter, ask the user once which objects should be iterated and store mapping

            var parameterObjectsMap = new Dictionary<string, (IList<string> Values, string Type)>();
            foreach (var param in distinctParameters)
            {
                string selectionType = null;
                if (param.Type == "CALENDARREF" || param.Name.ToUpper().EndsWith("CALENDAR"))
                {
                    selectionType = "Calendar";
                }
                else if (param.Type == "MEASUREREF" || param.Name.ToUpper().Contains("MEASURE"))
                {
                    selectionType = "Measure";
                }
                else if (param.Type == "COLUMNREF" || param.Name.ToUpper().Contains("COLUMN"))
                {
                    selectionType = "Column";
                }
                else if (param.Type == "TABLEREF" || param.Name.ToUpper().Contains("TABLE"))
                {
                    selectionType = "Table";
                }
                else if (param.Name.ToUpper().Contains("TEXT") || param.Name.ToUpper().Contains("STRING"))
                {
                    selectionType = "Text";
                }


                (IList<string> Values, string Type) selectedObjectsForParam = Fx.SelectAnyObjects(
                    Model,
                    selectionType: selectionType,
                    prompt1: String.Format(@"Select object type for {0} parameter", param.Name),
                    prompt2: String.Format(@"Select item for {0} parameter", param.Name),
                    placeholderValue: param.Name,
                    defaultValue: param.DefaultValue,
                    description: param.Description

                );

                if (selectedObjectsForParam.Values == null || selectedObjectsForParam.Type == null) return; //user cancelled


                parameterObjectsMap[param.Name] = selectedObjectsForParam;
            }

            int totalMeasuresCreated = 0;
            int totalColumnsCreated = 0;

            foreach (var func in selectedFunctions)
            {
                string delimiter = "";

                IList<string> previousList = new List<string>() { func.Name + "(" };
                IList<string> currentList = new List<string>();

                IList<string> previousListNames = new List<string>() { func.OutputNameTemplate };
                IList<string> currentListNames = new List<string>();

                IList<string> previousDestinations = new List<string>() { func.OutputDestination };
                IList<string> currentDestinations = new List<string>();

                IList<string> previousDisplayFolders = new List<string>() { func.OutputDisplayFolder };
                IList<string> currentDisplayFolders = new List<string>();

                IList<string> previousFormatStrings = new List<string>() { func.OutputFormatString };
                IList<string> currentFormatStrings = new List<string>();

                IList<string> previousOutputProperties = new List<string>() { func.OutputProperties };
                IList<string> currentOutputProperties = new List<string>();

                IList<string> previousOutputPropertyNames = new List<string>() { func.OutputPropertyNames };
                IList<string> currentOutputPropertyNames = new List<string>();

                IList<string> previousDataCategories = new List<string>() { func.OutputDataCategory ?? "" };
                IList<string> currentDataCategories = new List<string>();

                // When iterating the parameters of this specific function, use the mapping created for distinct parameters.
                foreach (var param in func.Parameters)
                {
                    currentList = new List<string>(); //reset current list
                    currentListNames = new List<string>();
                    currentFormatStrings = new List<string>();
                    currentDestinations = new List<string>();
                    currentDisplayFolders = new List<string>();
                    currentOutputProperties = new List<string>();
                    currentOutputPropertyNames = new List<string>();
                    currentDataCategories = new List<string>();

                    // Retrieve the objects list for this parameter name from the map (prompting was done earlier)
                    (IList<string> Values, string Type) paramObject;
                    if (!parameterObjectsMap.TryGetValue(param.Name, out paramObject) || paramObject.Type == null || paramObject.Values.Count == 0)
                    {
                        Error(String.Format("No objects were selected earlier for parameter '{0}'.", param.Name));
                        return;
                    }

                    for (int i = 0; i < previousList.Count; i++)
                    {
                        string s = previousList[i];
                        string sName = previousListNames[i];
                        string sFormatString = previousFormatStrings[i];
                        string sDisplayFolder = previousDisplayFolders[i];
                        string sDestination = previousDestinations[i];
                        string sOutputProperties = previousOutputProperties[i];
                        string sOutputPropertyNames = previousOutputPropertyNames[i];
                        string sDataCategory = previousDataCategories[i];

                        foreach (var o in paramObject.Values)
                        {
                            //extract original name and format string if the parameter is a measure
                            string paramName = o;
                            string paramFormatStringFull = "";
                            string paramFormatStringRoot = "";
                            string paramDisplayFolder = "";
                            string paramTable = "";
                            string paramProperties = "";
                            string paramPropertyNames = "";
                            string paramDataCategory = "";

                            //prepare placeholder
                            string paramNamePlaceholder = param.Name + "Name";
                            string paramFormatStringRootPlaceholder = param.Name + "FormatStringRoot";
                            string paramFormatStringFullPlaceholder = param.Name + "FormatStringFull";
                            string paramDisplayFolderPlaceholder = param.Name + "DisplayFolder";
                            string paramPropertiesPlaceholder = param.Name + "Properties";
                            string paramPropertyNamesPlaceholder = param.Name + "PropertyNames";
                            string paramDataCategoryPlaceholder = param.Name + "DataCategory";


                            string paramTablePlaceholder = "";

                            // Empty value means the parameter default is used: keep the comma (blank arg)
                            bool useDefault = string.IsNullOrEmpty(o);

                            if (useDefault)
                            {
                                paramName = "";
                                paramTablePlaceholder = param.Name;
                            }
                            else if (paramObject.Type == "Measure")
                            {
                                Measure m = Model.AllMeasures.FirstOrDefault(m => m.DaxObjectFullName == o);
                                paramName = m.Name;
                                paramFormatStringFull = m.FormatString;
                                paramDisplayFolder = m.DisplayFolder;
                                paramTable = m.Table.DaxObjectFullName;
                                paramProperties = m.GetAnnotation("Properties") ?? "";
                                paramPropertyNames = m.GetAnnotation("PropertyNames") ?? "";
                                paramDataCategory = m.DataCategory ?? "";

                                paramTablePlaceholder = param.Name + "Table";

                            }
                            else if (paramObject.Type == "Column")
                            {
                                Column c = Model.AllColumns.FirstOrDefault(c => c.DaxObjectFullName == o);
                                paramName = c.Name;
                                paramFormatStringFull = c.FormatString;
                                paramDisplayFolder = c.DisplayFolder;
                                paramTable = c.Table.DaxObjectFullName;
                                paramProperties = c.GetAnnotation("Properties") ?? "";
                                paramPropertyNames = c.GetAnnotation("PropertyNames") ?? "";
                                paramDataCategory = c.DataCategory ?? "";

                                paramTablePlaceholder = param.Name + "Table";
                            }
                            else if (paramObject.Type == "Table")
                            {
                                Table t = Model.Tables.FirstOrDefault(t => t.DaxObjectFullName == o);
                                paramName = t.Name;
                                paramFormatStringFull = "";
                                paramDisplayFolder = "";
                                paramTable = t.DaxObjectFullName;
                                paramProperties = t.GetAnnotation("Properties") ?? "";
                                paramPropertyNames = t.GetAnnotation("PropertyNames") ?? "";

                                paramTablePlaceholder = param.Name;
                            }
                            else if (paramObject.Type == "Calendar")
                            {
                                // o is the quoted calendar reference ('Name'); strip quotes for the display name
                                paramName = o.Trim('\'');
                                paramTablePlaceholder = param.Name;
                            }
                            else
                            {
                                // Scalar / Text: use the value itself (strip quotes) as the display name
                                paramName = o.Trim('"');
                                paramTablePlaceholder = param.Name;
                            }

                            if (paramFormatStringFull.Contains(";"))
                            {
                                //keep the first part of the format string, strip it of any + sign
                                paramFormatStringRoot = paramFormatStringFull.Split(';')[0].Replace("+", "");
                            }
                            else
                            {
                                paramFormatStringRoot = paramFormatStringFull;
                            }



                            currentList.Add(s + delimiter + o);
                            currentListNames.Add(sName.Replace(paramNamePlaceholder, paramName));

                            currentFormatStrings.Add(
                                sFormatString
                                    .Replace(paramFormatStringFullPlaceholder, paramFormatStringFull)
                                    .Replace(paramFormatStringRootPlaceholder, paramFormatStringRoot));

                            currentDisplayFolders.Add(
                                sDisplayFolder
                                    .Replace(paramNamePlaceholder, paramName)
                                    .Replace(paramDisplayFolderPlaceholder, paramDisplayFolder));

                            currentDestinations.Add(
                                sDestination.Replace(paramTablePlaceholder, paramTable));

                            currentOutputProperties.Add(
                                sOutputProperties
                                    .Replace(paramNamePlaceholder, paramName)
                                    .Replace(paramPropertiesPlaceholder, paramProperties));

                            currentOutputPropertyNames.Add(
                                sOutputPropertyNames
                                    .Replace(paramPropertyNamesPlaceholder, paramPropertyNames));

                            currentDataCategories.Add(
                                sDataCategory.Replace(paramDataCategoryPlaceholder, paramDataCategory));


                        }


                    }

                    delimiter = ", ";
                    previousList = currentList;
                    previousListNames = currentListNames;
                    previousDestinations = currentDestinations;
                    previousDisplayFolders = currentDisplayFolders;
                    previousFormatStrings = currentFormatStrings;
                    previousOutputProperties = currentOutputProperties;
                    previousOutputPropertyNames = currentOutputPropertyNames;
                    previousDataCategories = currentDataCategories;
                }



                IList<Table> currentDestinationTables = new List<Table>();


                if (func.OutputType == "Measure" || func.OutputType == "Column")
                {
                    for (int i = 0; i < currentDestinations.Count; i++)
                    {

                        //transform to actual tables, initialize if necessary
                        Table destinationTable = Model.Tables.Where(
                            t => t.DaxObjectFullName == currentDestinations[i])
                            .FirstOrDefault();

                        if (destinationTable == null)
                        {
                            destinationTable = SelectTable(preselect: null, label: $"Select destinatoin table for {func.OutputType} {currentListNames[i]}");
                            if (destinationTable == null) return;
                        }

                        currentDestinationTables.Add(destinationTable);
                    }
                }



                if (func.OutputType == "Measure")
                {

                    for (int i = 0; i < currentList.Count; i++)
                    {

                        //It normalizes a folder/display-folder string by collapsing repeated slashes, removing leading/trailing backslashes and trimming whitespace.
                        string cleanCurrentDisplayFolder = Regex.Replace(currentDisplayFolders[i], @"[/]+", @"").Trim('\\').Trim();

                        Measure measure = currentDestinationTables[i].AddMeasure(currentListNames[i], currentList[i] + ")");
                        measure.FormatDax();
                        measure.Description = String.Format("Measure created with {0} function. Check function for details.", func.Name);
                        measure.DisplayFolder = cleanCurrentDisplayFolder;
                        measure.FormatString = currentFormatStrings[i];

                        // Set Properties annotation with the transformed outputProperties value
                        if (!string.IsNullOrEmpty(currentOutputProperties[i]))
                        {
                            measure.SetAnnotation("Properties", currentOutputProperties[i]);
                        }

                        // Set PropertyNames annotation with the transformed outputPropertyNames value
                        if (!string.IsNullOrEmpty(currentOutputPropertyNames[i]))
                        {
                            measure.SetAnnotation("PropertyNames", currentOutputPropertyNames[i]);
                        }

                        // Set data category unless Uncategorized (empty data category in the model)
                        if (!string.IsNullOrEmpty(currentDataCategories[i]) && currentDataCategories[i] != "Uncategorized")
                        {
                            measure.DataCategory = currentDataCategories[i];
                        }

                        totalMeasuresCreated++;
                    }

                }
                else if (func.OutputType == "Column")
                {

                    for (int i = 0; i < currentList.Count; i++)
                    {

                        //It normalizes a folder/display-folder string by collapsing repeated slashes, removing leading/trailing backslashes and trimming whitespace.
                        string cleanCurrentDisplayFolder = Regex.Replace(currentDisplayFolders[i], @"[/]+", @"").Trim('\\').Trim();

                        Column column = currentDestinationTables[i].AddCalculatedColumn(currentListNames[i], currentList[i] + ")");
                        //column.FormatDax();
                        column.Description = String.Format("Column created with {0} function. Check function for details.", func.Name);
                        column.DisplayFolder = cleanCurrentDisplayFolder;
                        column.FormatString = currentFormatStrings[i];

                        // Set Properties annotation with the transformed outputProperties value
                        if (!string.IsNullOrEmpty(currentOutputProperties[i]))
                        {
                            column.SetAnnotation("Properties", currentOutputProperties[i]);
                        }

                        // Set PropertyNames annotation with the transformed outputPropertyNames value
                        if (!string.IsNullOrEmpty(currentOutputPropertyNames[i]))
                        {
                            column.SetAnnotation("PropertyNames", currentOutputPropertyNames[i]);
                        }
                    }

                }
                else
                {
                    Info("Not implemented yet for output types other than Measure.");
                }
            }

            // Display summary message
            if (totalMeasuresCreated > 0 || totalColumnsCreated > 0)
            {
                string message = "";
                if (totalMeasuresCreated > 0)
                {
                    message += String.Format("{0} measure{1} created", totalMeasuresCreated, totalMeasuresCreated == 1 ? "" : "s");
                }
                if (totalColumnsCreated > 0)
                {
                    if (totalMeasuresCreated > 0) message += " and ";
                    message += String.Format("{0} column{1} created", totalColumnsCreated, totalColumnsCreated == 1 ? "" : "s");
                }
                Info(message + ".");
            }



        }



        void createFieldParameter()
        {
            //#r "Microsoft.VisualBasic"
            //#r "System.Drawing"
            //using Microsoft.VisualBasic;
            //using GeneralFunctions;

            // 2026-05-13 / B.Agullo / fixed bug to execute in TE2
            //2026-04-23 / B.Agullo / improvements to the default parameter name suggestion logic (based on property names if available) 
            //2026-03-29 / B.Agullo / added support for propertyNames annotation
            //80% comes from Daniel Otykier --> https://github.com/TabularEditor/TabularEditor3/issues/541#issuecomment-1129228481
            //20% B.Agullo --> pop-up to choose parameter name + all the properties annotation support (split by pipe delimiter and added as additional columns in the field parameter table)

            // Before running the script, select the measures or columns that you
            // would like to use as field parameters (hold down CTRL to select multiple
            // objects). Also, you may change the name of the field parameter table
            // below. NOTE: If used against Power BI Desktop, you must enable unsupported
            // features under File > Preferences (TE2) or Tools > Preferences (TE3).

#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif

            if (Selected.Columns.Count == 0 && Selected.Measures.Count == 0) throw new Exception("No columns or measures selected!");

            // Get objects collection
            var objects = Selected.Columns.Any() ? Selected.Columns.Cast<ITabularTableObject>() : Selected.Measures;

            if (Fx.IsAnswerYes("Do you want to order measures/columns by name?"))
            {
                objects = objects.OrderBy(o => o.Name).ToList();
            }

            // Retrieve Properties annotations and split them by pipe delimiter
            var objectPropertiesLists = new List<List<string>>();

            foreach (var obj in objects)
            {
                string propertiesAnnotation = "";
                if (obj is Measure)
                {
                    propertiesAnnotation = ((Measure)obj).GetAnnotation("Properties") ?? "";
                }
                else if (obj is Column)
                {
                    propertiesAnnotation = ((Column)obj).GetAnnotation("Properties") ?? "";
                }

                var propertiesList = string.IsNullOrEmpty(propertiesAnnotation)
                    ? new List<string>()
                    : propertiesAnnotation.Split('|').ToList();

                objectPropertiesLists.Add(propertiesList);
            }

            // Find maximum length to normalize all lists
            int maxPropertiesCount = objectPropertiesLists.Any() ? objectPropertiesLists.Max(pl => pl.Count) : 0;

            // Normalize all lists to the same length by adding empty strings
            for (int i = 0; i < objectPropertiesLists.Count; i++)
            {
                while (objectPropertiesLists[i].Count < maxPropertiesCount)
                {
                    objectPropertiesLists[i].Add("");
                }
            }

            // Retrieve PropertyNames annotations from selected objects
            var propertyNamesAnnotations = new List<string>();
            foreach (var obj in objects)
            {
                string propertyNamesAnnotation = "";
                if (obj is Measure)
                {
                    propertyNamesAnnotation = ((Measure)obj).GetAnnotation("PropertyNames") ?? "";
                }
                else if (obj is Column)
                {
                    propertyNamesAnnotation = ((Column)obj).GetAnnotation("PropertyNames") ?? "";
                }

                if (!string.IsNullOrEmpty(propertyNamesAnnotation))
                {
                    propertyNamesAnnotations.Add(propertyNamesAnnotation);
                }
            }

            // Determine property names for field parameter columns
            List<string> propertyNames = new List<string>();

            if (propertyNamesAnnotations.Count > 0)
            {
                // Get distinct PropertyNames values
                var distinctPropertyNames = propertyNamesAnnotations.Distinct().ToList();

                if (distinctPropertyNames.Count == 1)
                {
                    // All PropertyNames are the same, use them
                    propertyNames = distinctPropertyNames[0].Split('|').ToList();
                }
                else
                {
                    // Different PropertyNames found, let user choose
                    var optionsWithDefine = distinctPropertyNames.ToList();
                    optionsWithDefine.Add("<Define now>");

                    string chosenPropertyNames = Fx.ChooseString(
                        OptionList: optionsWithDefine,
                        label: "Different PropertyNames found. Choose one or define new:"
                    );

                    if (string.IsNullOrEmpty(chosenPropertyNames)) return;

                    if (chosenPropertyNames == "<Define now>")
                    {
                        // Show input box with first PropertyNames as default
                        chosenPropertyNames = Fx.GetNameFromUser(
                            Prompt: "Enter property names (pipe-separated):",
                            Title: "Property Names",
                            DefaultResponse: distinctPropertyNames[0]
                        );

                        if (string.IsNullOrEmpty(chosenPropertyNames)) return;
                    }

                    propertyNames = chosenPropertyNames.Split('|').ToList();
                }
            }
            else
            {
                // No PropertyNames found, ask one by one
                for (int i = 0; i < maxPropertiesCount; i++)
                {
                    // Collect distinct values for this property position (up to 3)
                    var sampleValues = objectPropertiesLists
                        .Where(pl => pl.Count > i && !string.IsNullOrEmpty(pl[i]))
                        .Select(pl => pl[i])
                        .Distinct()
                        .Take(3)
                        .ToList();

                    string sampleText = sampleValues.Count > 0
                        ? String.Format(" (found values: '{0}')", string.Join("', '", sampleValues))
                        : "";

                    string propertyName = Fx.GetNameFromUser(
                        Prompt: String.Format("Enter name for property {0}{1}:", i + 1, sampleText),
                        Title: "Property Name",
                        DefaultResponse: String.Format("Property{0}", i + 1)
                    );

                    if (string.IsNullOrEmpty(propertyName)) return;

                    propertyNames.Add(propertyName);
                }
            }

            // Build default parameter name based on PropertyNames
            string defaultParameterName = propertyNames.Count > 0
                ? "Parameter " + string.Join(" ", propertyNames)
                : "Parameter";

            // Ask user for parameter name with smart default
            var name = Fx.GetNameFromUser("Provide the name for the field parameter", "Parameter name", defaultParameterName);

            if (string.IsNullOrEmpty(name)) return;

            // Construct the DAX for the calculated table based on the current selection:
            string dax = "";
            if (maxPropertiesCount > 0)
            {
                // Build DAX with additional property columns
                var daxRows = objects.Select((c, i) =>
                {
                    var baseValues = string.Format("\"{0}\", NAMEOF('{1}'[{0}]), {2}", c.Name, c.Table.Name, i);
                    var propertyValues = string.Join(", ", objectPropertiesLists[i].Select(p => string.Format("\"{0}\"", p)));
                    return string.Format("({0}, {1})", baseValues, propertyValues);
                });

                dax = "{\n    " + string.Join(",\n    ", daxRows) + "\n}";
            }
            else
            {
                // No properties found, use original format
                dax = "{\n    " + string.Join(",\n    ", objects.Select((c, i) => string.Format("(\"{0}\", NAMEOF('{1}'[{0}]), {2})", c.Name, c.Table.Name, i))) + "\n}";
            }

            // Add the calculated table to the model:
            var table = Model.AddCalculatedTable(name, dax);

            // In TE2 columns are not created automatically from a DAX expression, so 
            // we will have to add them manually:
            var te2 = table.Columns.Count == 0;
            var nameColumn = te2 ? table.AddCalculatedTableColumn(name, "[Value1]") : table.Columns["Value1"] as CalculatedTableColumn;
            var fieldColumn = te2 ? table.AddCalculatedTableColumn(name + " Fields", "[Value2]") : table.Columns["Value2"] as CalculatedTableColumn;
            var orderColumn = te2 ? table.AddCalculatedTableColumn(name + " Order", "[Value3]") : table.Columns["Value3"] as CalculatedTableColumn;

            // Create additional columns for each property level
            var propertyColumns = new List<CalculatedTableColumn>();
            for (int i = 0; i < maxPropertiesCount; i++)
            {
                int valueIndex = 4 + i; // Value4, Value5, etc.
                string columnName = propertyNames.Count > i ? propertyNames[i] : String.Format("Property{0}", i + 1);
                var propColumn = te2
                    ? table.AddCalculatedTableColumn(columnName, "[Value" + valueIndex + "]")
                    : table.Columns["Value" + valueIndex] as CalculatedTableColumn;

                if (!te2)
                {
                    propColumn.IsNameInferred = false;
                    propColumn.Name = columnName;
                }

                propertyColumns.Add(propColumn);
            }

            if (!te2)
            {
                // Rename the columns that were added automatically in TE3:
                nameColumn.IsNameInferred = false;
                nameColumn.Name = name;
                fieldColumn.IsNameInferred = false;
                fieldColumn.Name = name + " Fields";
                orderColumn.IsNameInferred = false;
                orderColumn.Name = name + " Order";
            }

            // Set remaining properties for field parameters to work
            // See: https://twitter.com/markbdi/status/1526558841172893696
            nameColumn.SortByColumn = orderColumn;
            nameColumn.GroupByColumns.Add(fieldColumn);
            fieldColumn.SortByColumn = orderColumn;
            fieldColumn.SetExtendedProperty("ParameterMetadata", "{\"version\":3,\"kind\":2}", ExtendedPropertyType.Json);
            fieldColumn.IsHidden = true;
            orderColumn.IsHidden = true;
        }
        void applyCalcGroupsFormatExpression()
        {
            //this script only executes in TE3

            //using Dax.Analyzer;
            //using Dax.Formatter; 
            //using System.IO;

            //using GeneralFunctions; 

#if TE3
    ScriptHelper.WaitFormVisible = false;
#endif

            Measure targetMeasure = Fx.GetSelectedMeasure(Selected.Measures);
            if (targetMeasure == null) return;

            var selectedItems = Fx.SelectCalculationItems(Model, "Select one calculation item per group");
            if (selectedItems == null || selectedItems.Count == 0) return;

            // Get only calculation groups that have selected items, ordered by precedence (descending)
            var calcGroupsByPrecedence = Model.Tables
                .OfType<CalculationGroupTable>()
                .Where(cg => selectedItems.ContainsKey(cg.Name))
                .OrderByDescending(cg => cg.CalculationGroup.Precedence)
                .ToList();


            string previousFormatExpression = @"SELECTEDMEASUREFORMATSTRING()";
            string currentFormatExpression = @"";

            foreach (var calcGroup in calcGroupsByPrecedence)
            {
                string selectedItemName = selectedItems[calcGroup.Name];
                //Output($"Calc Group: {calcGroup.Name}, Precedence: {calcGroup.CalculationGroup.Precedence}, Selected Item: {selectedItemName}");

                // Get the actual calculation item object
                var calcItem = calcGroup.CalculationItems.FirstOrDefault(ci => ci.Name == selectedItemName);
                if (calcItem == null) { Error(String.Format("Calc Item {0} not found", selectedItemName)); return; }

                var calcItemFormatTokens = DaxTokenizer.Tokenize(calcItem.FormatStringExpression);

                string calcItemFormatReplaceExpression = String.Format(@"/*{0} - {1}*/" + Environment.NewLine, calcGroup.Name, calcItem.Name);
                bool closeParenthesisFound = true;
                bool isSelectedMeasureArgs = false;
                //{0} SELECTEDMEASURE
                //{1} SELECTEDMEASURENAME
                //{2} SELECTEDMEASUREFORMATSTRING
                foreach (var calcItemFormatToken in calcItemFormatTokens)
                {
                    if (!closeParenthesisFound)
                    {
                        if (calcItemFormatToken.Type == DaxToken.CLOSE_PARENS)
                        {
                            closeParenthesisFound = true;
                        }
                        if (isSelectedMeasureArgs)
                        {
                            if (calcItemFormatToken.Type == DaxToken.CLOSE_PARENS)
                            {
                                calcItemFormatReplaceExpression += "}}";
                                isSelectedMeasureArgs = false;

                            }
                            else if (calcItemFormatToken.Type == DaxToken.COLUMN_OR_MEASURE)
                            {
                                calcItemFormatReplaceExpression += @"""" + calcItemFormatToken.Text + @"""";

                            }
                            else if (calcItemFormatToken.Type == DaxToken.COMMA)
                            {
                                calcItemFormatReplaceExpression += calcItemFormatToken.Text;
                            }

                        }


                    }
                    else if (calcItemFormatToken.Type == DaxToken.TABLE)
                    {
                        calcItemFormatReplaceExpression += "'" + calcItemFormatToken.Text + "'";

                    }
                    else if (calcItemFormatToken.Type == DaxToken.COLUMN_OR_MEASURE)
                    {
                        calcItemFormatReplaceExpression += "[" + calcItemFormatToken.Text + "]";
                    }
                    else if (calcItemFormatToken.Type == DaxToken.SELECTEDMEASURE)
                    {
                        calcItemFormatReplaceExpression += "{0}";
                        closeParenthesisFound = false;

                    }
                    else if (calcItemFormatToken.Type == DaxToken.ISSELECTEDMEASURE)
                    {
                        calcItemFormatReplaceExpression += @"""{1}"" IN {{";
                        isSelectedMeasureArgs = true;
                        closeParenthesisFound = false;

                    }
                    else if (calcItemFormatToken.Type == DaxToken.SELECTEDMEASURENAME)
                    {
                        calcItemFormatReplaceExpression += @"""{1}""";
                        closeParenthesisFound = false;

                    }
                    //selectedmeasureformatstring should stay as is in format expression
                    //else if (calcItemFormatToken.Type == DaxToken.SELECTEDMEASUREFORMATSTRING)
                    //{
                    //    calcItemFormatReplaceExpression += @"{2}";
                    //    closeParenthesisFound = false;
                    //}
                    else if (calcItemFormatToken.Type == DaxToken.STRING_LITERAL)
                    {
                        calcItemFormatReplaceExpression += @"""" + calcItemFormatToken.Text + @"""";
                    }
                    else
                    {
                        calcItemFormatReplaceExpression += calcItemFormatToken.Text;
                    }

                }






                var expressionFormatTokens = DaxTokenizer.Tokenize(previousFormatExpression);

                currentFormatExpression = "";
                bool closeParenthesisFound2 = true;
                bool isColumn = false;

                foreach (var token in expressionFormatTokens)
                {

                    // Check if previous token indicated this is a column
                    if (!closeParenthesisFound2)
                    {
                        //'do nothing'
                        if (token.Type == DaxToken.CLOSE_PARENS)
                        {
                            closeParenthesisFound2 = true;
                        }

                    }
                    else if (isColumn)
                    {
                        currentFormatExpression += "[" + token.Text + "]";
                        isColumn = false;
                    }
                    // Check for Table token followed by Column/Measure token
                    else if (token.Type == DaxToken.TABLE)
                    {
                        Table table = Model.Tables.FirstOrDefault(t => t.Name == token.Text);
                        if (table == null)
                        {
                            Error(token.Text + " table not found in the model.");
                            return;
                        }

                        if (token.Next.Type == DaxToken.COLUMN_OR_MEASURE)
                        {
                            // Set flag to indicate next token is a column
                            isColumn = true;
                        }

                        currentFormatExpression += table.DaxObjectFullName;
                    }
                    else if (token.Type == DaxToken.COLUMN_OR_MEASURE)
                    {
                        //currentFormatExpression += "[" + token.Text + "]";
                        Output(currentFormatExpression);
                        currentFormatExpression += String.Format(
                            calcItemFormatReplaceExpression,
                            targetMeasure.DaxObjectFullName,
                            targetMeasure.Name);

                        closeParenthesisFound2 = false;

                    }
                    else if (token.Type == DaxToken.SELECTEDMEASURE)
                    {
                        currentFormatExpression += "{0}";
                        closeParenthesisFound2 = false;
                    }
                    else if (token.Type == DaxToken.SELECTEDMEASURENAME)
                    {
                        currentFormatExpression += @"""{1}""";
                        closeParenthesisFound2 = false;
                    }
                    else if (token.Type == DaxToken.SELECTEDMEASUREFORMATSTRING)
                    {
                        currentFormatExpression += String.Format(
                            calcItemFormatReplaceExpression,
                            targetMeasure.DaxObjectFullName,
                            targetMeasure.Name);

                        closeParenthesisFound2 = false;
                    }
                    else if (token.Type == DaxToken.STRING_LITERAL)
                    {
                        currentFormatExpression += @"""" + token.Text + @"""";
                    }
                    else
                    {
                        currentFormatExpression += token.Text;
                    }
                }





                previousFormatExpression = currentFormatExpression;

            }


            //replace selectedmeasureformatstring by the format string of the target measure now 
            previousFormatExpression = currentFormatExpression;
            var finalTokens = DaxTokenizer.Tokenize(previousFormatExpression);
            currentFormatExpression = "";
            bool closeParenthesisFound3 = true;
            foreach (var token in finalTokens)
            {
                if (!closeParenthesisFound3)
                {
                    if (token.Type == DaxToken.CLOSE_PARENS)
                    {
                        closeParenthesisFound3 = true;
                    }
                }
                else if (token.Type == DaxToken.SELECTEDMEASUREFORMATSTRING)
                {
                    if (targetMeasure.FormatStringExpression != null)
                    {
                        currentFormatExpression += targetMeasure.FormatStringExpression;
                    }
                    else
                    {
                        currentFormatExpression += @"""" + targetMeasure.FormatString + @"""";
                    }

                    closeParenthesisFound3 = false;

                }
                else if (token.Type == DaxToken.STRING_LITERAL)
                {
                    currentFormatExpression += @"""" + token.Text + @"""";
                }
                else if (token.Type == DaxToken.TABLE)
                {
                    currentFormatExpression += "'" + token.Text + "'";
                }
                else if (token.Type == DaxToken.COLUMN_OR_MEASURE)
                {
                    currentFormatExpression += "[" + token.Text + "]";
                }
                else
                {
                    currentFormatExpression += token.Text;
                }



            }

            currentFormatExpression = FormatDax(currentFormatExpression);

            string fileName = targetMeasure.Name + "_" + calcGroupsByPrecedence.Select(cg => cg.Name + " " + selectedItems[cg.Name]).Aggregate((a, b) => a + "_" + b) + ".dax";

            string tempFile = Path.Combine(Path.GetTempPath(), "AppliedCalcGroupFormatExpression_" + fileName + ".dax");
            File.WriteAllText(tempFile, currentFormatExpression);

            System.Diagnostics.Process.Start("notepad.exe", tempFile);

            Info("Expansion complete. Result saved to: " + tempFile);


        }
        void applyCalcGroups()
        {
            //this script only executes in TE3

            //using Dax.Analyzer;
            //using Dax.Formatter; 
            //using System.IO;

            //using GeneralFunctions; 

#if TE3
    ScriptHelper.WaitFormVisible = false;
#endif

            Measure targetMeasure = Fx.GetSelectedMeasure(Selected.Measures);
            if (targetMeasure == null) return;

            var selectedItems = Fx.SelectCalculationItems(Model, "Select one calculation item per group");
            if (selectedItems == null || selectedItems.Count == 0) return;

            // Get only calculation groups that have selected items, ordered by precedence (descending)
            var calcGroupsByPrecedence = Model.Tables
                .OfType<CalculationGroupTable>()
                .Where(cg => selectedItems.ContainsKey(cg.Name))
                .OrderByDescending(cg => cg.CalculationGroup.Precedence)
                .ToList();

            string previousExpression = targetMeasure.DaxObjectFullName;
            string currentExpression = @"";

            foreach (var calcGroup in calcGroupsByPrecedence)
            {
                string selectedItemName = selectedItems[calcGroup.Name];
                //Output($"Calc Group: {calcGroup.Name}, Precedence: {calcGroup.CalculationGroup.Precedence}, Selected Item: {selectedItemName}");

                // Get the actual calculation item object
                var calcItem = calcGroup.CalculationItems.FirstOrDefault(ci => ci.Name == selectedItemName);
                if (calcItem == null) { Error(String.Format("Calc Item {0} not found", selectedItemName)); return; }

                var calcItemTokens = calcItem.Tokenize();

                string calcItemReplaceExpression = String.Format(@"/*{0} - {1}*/", calcGroup.Name, selectedItemName) + Environment.NewLine;
                bool closeParenthesisFound = true;
                bool isSelectedMeasureArgs = false;
                //{0} SELECTEDMEASURE
                //{1} SELECTEDMEASURENAME
                //{2} SELECTEDMEASUREFORMATSTRING
                foreach (var calcItemToken in calcItemTokens)
                {
                    if (!closeParenthesisFound)
                    {
                        if (calcItemToken.Type == DaxToken.CLOSE_PARENS)
                        {
                            closeParenthesisFound = true;
                        }
                        if (isSelectedMeasureArgs)
                        {
                            if (calcItemToken.Type == DaxToken.CLOSE_PARENS)
                            {
                                calcItemReplaceExpression += "}}";
                                isSelectedMeasureArgs = false;

                            }
                            else if (calcItemToken.Type == DaxToken.COLUMN_OR_MEASURE)
                            {
                                calcItemReplaceExpression += @"""" + calcItemToken.Text + @"""";

                            }
                            else if (calcItemToken.Type == DaxToken.COMMA)
                            {
                                calcItemReplaceExpression += calcItemToken.Text;
                            }

                        }


                    }
                    else if (calcItemToken.Type == DaxToken.TABLE)
                    {
                        calcItemReplaceExpression += "'" + calcItemToken.Text + "'";

                    }
                    else if (calcItemToken.Type == DaxToken.COLUMN_OR_MEASURE)
                    {
                        calcItemReplaceExpression += "[" + calcItemToken.Text + "]";
                    }
                    else if (calcItemToken.Type == DaxToken.SELECTEDMEASURE)
                    {
                        calcItemReplaceExpression += "{0}";
                        closeParenthesisFound = false;

                    }
                    else if (calcItemToken.Type == DaxToken.ISSELECTEDMEASURE)
                    {
                        calcItemReplaceExpression += @"""{1}"" IN {{";
                        isSelectedMeasureArgs = true;
                        closeParenthesisFound = false;

                    }
                    else if (calcItemToken.Type == DaxToken.SELECTEDMEASURENAME)
                    {
                        calcItemReplaceExpression += @"""{1}""";
                        closeParenthesisFound = false;

                    }
                    else if (calcItemToken.Type == DaxToken.SELECTEDMEASUREFORMATSTRING)
                    {
                        calcItemReplaceExpression += @"{2}";
                        closeParenthesisFound = false;
                    }
                    else if (calcItemToken.Type == DaxToken.STRING_LITERAL)
                    {
                        calcItemReplaceExpression += @"""" + calcItemToken.Text + @"""";
                    }
                    else
                    {
                        calcItemReplaceExpression += calcItemToken.Text;
                    }

                }

                var expressionTokens = DaxTokenizer.Tokenize(previousExpression);
                currentExpression = "";
                bool isColumn = false;

                foreach (var token in expressionTokens)
                {

                    // Check if previous token indicated this is a column
                    if (isColumn)
                    {
                        currentExpression += "[" + token.Text + "]";
                        isColumn = false;
                    }
                    // Check for Table token followed by Column/Measure token
                    else if (token.Type == DaxToken.TABLE)
                    {
                        Table table = Model.Tables.FirstOrDefault(t => t.Name == token.Text);
                        if (table == null)
                        {
                            Error(token.Text + " table not found in the model.");
                            return;
                        }

                        if (token.Next.Type == DaxToken.COLUMN_OR_MEASURE)
                        {
                            // Set flag to indicate next token is a column
                            isColumn = true;
                        }

                        currentExpression += table.DaxObjectFullName;
                    }
                    else if (token.Type == DaxToken.COLUMN_OR_MEASURE)
                    {
                        Measure measure = Model.AllMeasures.FirstOrDefault(m => m.Name == token.Text);
                        if (measure == null)
                        {
                            Error(token.Text + " measure not found in the model.");
                            return;
                        }



                        currentExpression += String.Format(
                            calcItemReplaceExpression,
                            measure.DaxObjectFullName,
                            measure.Name,
                            measure.FormatString ?? measure.FormatStringExpression);
                    }
                    else if (token.Type == DaxToken.STRING_LITERAL)
                    {
                        currentExpression += @"""" + token.Text + @"""";
                    }
                    else
                    {
                        currentExpression += token.Text;
                    }
                }

                currentExpression = FormatDax(currentExpression);
                previousExpression = currentExpression;

            }


            Output(currentExpression.ReplaceLineEndings());

            //string tempFile = Path.Combine(Path.GetTempPath(), "ExpandedDAX_test.dax");
            //File.WriteAllText(tempFile, currentExpression);

            //System.Diagnostics.Process.Start("notepad.exe", tempFile);

            //Info("Expansion complete. Result saved to: " + tempFile);


        }
        void evaluatedDax()
        {

            //this script only executes in TE3

            //using Dax.Analyzer;
            //using Dax.Formatter; 
            //using System.IO;
            //using GeneralFunctions;
            //using DaxUserDefinedFunction;

#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif

            Measure targetMeasure = null;
            string measureName = "Clipboard Expression";
            string previousExpression = "";

            // Check if clipboard contains text (and not just whitespace)
            string clipboardContent = Clipboard.ContainsText() ? Clipboard.GetText() : string.Empty;

            if (!string.IsNullOrWhiteSpace(clipboardContent))
            {
                // Show clipboard content in a read-only window and ask for confirmation
                var form = new Form();
                form.Text = "Clipboard Content";
                form.Width = 800;
                form.Height = 600;
                form.StartPosition = FormStartPosition.CenterScreen;

                var textBox = new TextBox();
                textBox.Multiline = true;
                textBox.ScrollBars = ScrollBars.Both;
                textBox.WordWrap = false;
                textBox.ReadOnly = true;
                textBox.Text = clipboardContent;
                textBox.Dock = DockStyle.Fill;
                textBox.Font = new System.Drawing.Font("Consolas", 10);

                var buttonPanel = new Panel();
                buttonPanel.Dock = DockStyle.Bottom;
                buttonPanel.Height = 40;

                var okButton = new Button();
                okButton.Text = "Expand this expression";
                okButton.DialogResult = DialogResult.OK;
                okButton.Width = 150;
                okButton.Left = 10;
                okButton.Top = 5;

                var cancelButton = new Button();
                cancelButton.Text = "Cancel";
                cancelButton.DialogResult = DialogResult.Cancel;
                cancelButton.Width = 80;
                cancelButton.Left = 170;
                cancelButton.Top = 5;

                buttonPanel.Controls.Add(okButton);
                buttonPanel.Controls.Add(cancelButton);

                form.Controls.Add(textBox);
                form.Controls.Add(buttonPanel);

                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    previousExpression = clipboardContent;
                }
                else
                {
                    Info("Operation cancelled.");
                    return;
                }
            }
            else
            {
                // No clipboard content, ask if user wants to select a measure
                if (!Fx.IsAnswerYes("No expression found in clipboard. Do you want to select a measure instead?"))
                {
                    Info("Operation cancelled.");
                    return;
                }

                // Select measure as before
                targetMeasure = Fx.GetSelectedMeasure(Selected.Measures);
                if (targetMeasure == null) return;

                measureName = targetMeasure.Name;
                previousExpression = targetMeasure.Expression;
            }

            // Parse all UDFs in the model and create expression templates
            var udfTemplates = new Dictionary<string, (string template, List<FunctionParameter> parameters)>();

            foreach (var function in Model.Functions/*.Where(f => f.Name == "Local.TimeIntel.YOYPCT")*/)
            {
                FunctionExtended fe = FunctionExtended.CreateFunctionExtended(function, completeMetadata: false);
                if (fe == null) continue;

                // Extract the expression body (after =>)
                string funcExpression = function.Expression;
                int arrowIndex = funcExpression.IndexOf("=>");
                if (arrowIndex == -1) continue;

                string expressionBody = funcExpression.Substring(arrowIndex + 2).Trim();

                // Tokenize the function expression to build the template
                var funcTokens = DaxTokenizer.Tokenize(expressionBody);
                string template = "";

                // Build ANYVAL variable declarations for parameters that are not EXPR/ANYREF
                var anyvalParams = fe.Parameters.Where(p =>
                    p.ParameterMode != "EXPR" &&
                    p.Type != "ANYREF").ToList();

                if (anyvalParams.Count > 0)
                {
                    template += String.Format(@"/*{0}*/" + Environment.NewLine, function.Name);
                    foreach (var param in anyvalParams)
                    {
                        int paramIndex = fe.Parameters.IndexOf(param);
                        template += String.Format("VAR __param_{0} = {{{1}}}" + Environment.NewLine, param.Name, paramIndex);
                    }
                    template += "RETURN" + Environment.NewLine;
                }
                else
                {
                    template += String.Format(@"/*{0}*/" + Environment.NewLine, function.Name);
                }

                // Process tokens and replace parameter names with placeholders
                foreach (var token in funcTokens)
                {
                    if (token.Type == DaxToken.TABLE_OR_VARIABLE)
                    {
                        // Check if this is a parameter name
                        var matchingParam = fe.Parameters.FirstOrDefault(p => p.Name == token.Text);
                        if (matchingParam != null)
                        {
                            int paramIndex = fe.Parameters.IndexOf(matchingParam);

                            // If ANYVAL, use the variable instead
                            if (matchingParam.ParameterMode != "EXPR" && matchingParam.Type != "ANYREF")
                            {
                                template += "__param_" + matchingParam.Name;
                            }
                            else
                            {
                                template += "{" + paramIndex + "}";
                            }
                        }
                        else if (Model.Tables.Any(t => t.Name == token.Text))
                        {
                            template += "'" + token.Text + "'";
                        }
                        else
                        {
                            template += token.Text;
                        }
                    }
                    else if (token.Type == DaxToken.TABLE)
                    {
                        template += "'" + token.Text + "'";
                    }
                    else if (token.Type == DaxToken.STRING_LITERAL)
                    {
                        template += @"""" + token.Text + @"""";
                    }
                    else if (token.Type == DaxToken.COLUMN_OR_MEASURE)
                    {
                        template += "[" + token.Text + "]";
                    }


                    else
                    {
                        template += token.Text;
                    }
                }

                udfTemplates[function.Name] = (template, fe.Parameters);
            }

            //initialize variables for the iterations    
            var tokens = DaxTokenizer.Tokenize(previousExpression);

            string currentExpression = "";
            bool isColumn = false;
            int i = 0;

            bool isDaxUdf = false;
            string daxUdfName = null;
            IList<string> daxParams = new List<string>();
            string paramValue = null;
            bool daxUdfFirstParensFound = false;
            int daxUdfParensCount = 0;
            int paramParensCount = 0;

            //Check if there are any column_or_measure tokens that are actual measures in the model and not preceded by a table token,
            // or a column_or_measure as first token that is an actual measure, 
            // or any token whose text matches a DAX UDF name
            // and limit to 10 iterations
            while (i < 10 && tokens.Any(t =>
                (
                    (
                        t.Type == DaxToken.COLUMN_OR_MEASURE
                        && (t.Prev != null && t.Prev.Type != DaxToken.TABLE)
                        && Model.AllMeasures.Any(m => m.Name == t.Text)
                    )
                    ||
                    (
                        t.Type == DaxToken.COLUMN_OR_MEASURE
                        && t.TokenIndex == 0
                        && Model.AllMeasures.Any(m => m.Name == t.Text)
                    )
                    ||
                    udfTemplates.ContainsKey(t.Text)
                )
            ))
            {

                currentExpression = "";


                foreach (DaxToken token in tokens)
                {
                    // Check if previous token indicated this is a column
                    if (isColumn)
                    {
                        currentExpression += "[" + token.Text + "]";
                        isColumn = false;
                    }
                    else if (isDaxUdf)
                    {
                        if (token.Type == DaxToken.OPEN_PARENS)
                        {
                            if (!daxUdfFirstParensFound)
                            {
                                daxUdfFirstParensFound = true;
                                paramValue = "";
                            }
                            else
                            {
                                paramValue += token.Text;
                            }

                            daxUdfParensCount += 1;
                            paramParensCount += 1;

                        }

                        else if (token.Type == DaxToken.CLOSE_PARENS)
                        {
                            daxUdfParensCount -= 1;
                            paramParensCount -= 1;

                            if (daxUdfParensCount == 0)
                            {
                                daxParams.Add(paramValue);
                                paramValue = null;

                                // Expand the UDF expression with the collected parameters
                                if (udfTemplates.ContainsKey(daxUdfName))
                                {
                                    var udfInfo = udfTemplates[daxUdfName];

                                    // Create array of parameter values for String.Format
                                    object[] paramValues = daxParams.Select(p => (object)p).ToArray();

                                    // Format the template with actual parameter values
                                    string expandedUdf = String.Format(udfInfo.template, paramValues);

                                    currentExpression += expandedUdf;
                                }
                                else
                                {
                                    // UDF not found in templates, keep as-is
                                    currentExpression += daxUdfName + "(" + string.Join(",", daxParams) + ")";
                                }

                                // Reset UDF tracking variables
                                isDaxUdf = false;
                                daxUdfName = null;
                                daxParams = new List<string>();
                                daxUdfFirstParensFound = false;
                            }
                            else
                            {
                                paramValue += token.Text;
                            }
                        }


                        else if (token.Type == DaxToken.COMMA)
                        {
                            if (paramParensCount == 0)
                            {
                                daxParams.Add(paramValue);
                                paramValue = "";

                            }
                            else
                            {
                                paramValue += token.Text;
                            }


                        }
                        else
                        {
                            if (token.Type == DaxToken.COLUMN_OR_MEASURE)
                            {
                                paramValue += "[" + token.Text + "]";
                            }
                            else if (token.Type == DaxToken.TABLE)
                            {
                                paramValue += "'" + token.Text + "'";
                            }
                            else if (token.Type == DaxToken.STRING_LITERAL)
                            {
                                paramValue += @"""" + token.Text + @"""";
                            }
                            else if (token.CommentOrWhitespace)
                            {
                                /*do nothing*/
                            }
                            else
                            {
                                paramValue += token.Text;
                            }
                        }
                    }
                    // Check for Table token followed by Column/Measure token
                    else if (token.Type == DaxToken.TABLE)
                    {
                        // First check if it's a Calendar
                        if (Model.AllCalendars.Any(f => f.Name == token.Text))
                        {
                            // It's a Calendar, keep with single quotes
                            currentExpression += "'" + token.Text + "'";
                        }
                        else
                        {
                            // It's a table, proceed with existing logic
                            Table table = Model.Tables.FirstOrDefault(t => t.Name == token.Text);
                            if (table == null)
                            {
                                Error(token.Text + " table not found in the model.");
                                return;
                            }

                            if (token.Next.Type == DaxToken.COLUMN_OR_MEASURE)
                            {
                                // Check if it's actually a column in this table
                                Column column = table.Columns.FirstOrDefault(c => c.Name == token.Next.Text);
                                if (column != null)
                                {
                                    // It's a column, set flag to add brackets in next iteration
                                    isColumn = true;
                                    currentExpression += table.DaxObjectFullName;
                                }
                                else
                                {
                                    // Not a column, check if it's a measure (will be expanded in next token)
                                    Measure measure = Model.AllMeasures.FirstOrDefault(m => m.Name == token.Next.Text);
                                    if (measure != null)
                                    {
                                        // It's a measure, don't add table name (measure will be expanded next)
                                    }
                                    else
                                    {
                                        // Check if it's a special FormatString measure pattern
                                        bool isFormatStringMeasure = false;
                                        if (token.Next.Text.StartsWith("_") && token.Next.Text.EndsWith(" FormatString"))
                                        {
                                            string baseMeasureName = token.Next.Text.Substring(1, token.Next.Text.Length - " FormatString".Length - 1);
                                            Measure baseMeasure = Model.AllMeasures.FirstOrDefault(m => m.Name == baseMeasureName);
                                            if (baseMeasure != null)
                                            {
                                                // It's a valid FormatString measure, don't add table name (will be expanded next)
                                                isFormatStringMeasure = true;
                                            }
                                        }

                                        if (!isFormatStringMeasure)
                                        {
                                            Error(String.Format("'{0}' not found as column in table '{1}' or as measure in the model.", token.Next.Text, table.Name));
                                            return;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                currentExpression += table.DaxObjectFullName;
                            }
                        }
                    }
                    else if (token.Type == DaxToken.COLUMN_OR_MEASURE)
                    {
                        Measure measure = Model.AllMeasures.FirstOrDefault(m => m.Name == token.Text);
                        if (measure != null)
                        {
                            // It's a measure, expand it
                            currentExpression += String.Format(
                                @"/*{0}*/
                            CALCULATE({1})",
                                measure.Name,
                                measure.Expression);
                        }
                        else
                        {
                            // Check if it follows the pattern "_<MeasureName> FormatString"
                            if (token.Text.StartsWith("_") && token.Text.EndsWith(" FormatString"))
                            {
                                // Extract the base measure name
                                string baseMeasureName = token.Text.Substring(1, token.Text.Length - " FormatString".Length - 1);
                                Measure baseMeasure = Model.AllMeasures.FirstOrDefault(m => m.Name == baseMeasureName);

                                if (baseMeasure != null)
                                {
                                    // Replace with FormatStringExpression or FormatString
                                    if (!string.IsNullOrEmpty(baseMeasure.FormatStringExpression))
                                    {
                                        currentExpression += String.Format(
                                            @"/*{0}*/
                            ({1})",
                                            token.Text,
                                            baseMeasure.FormatStringExpression);
                                    }
                                    else
                                    {
                                        currentExpression += String.Format(
                                            @"/*{0}*/
                            ""{1}""",
                                            token.Text,
                                            baseMeasure.FormatString);
                                    }
                                }
                                else
                                {
                                    // Base measure not found, keep as-is with brackets
                                    currentExpression += "[" + token.Text + "]";
                                }
                            }
                            else
                            {
                                // Not a measure and doesn't match the pattern, keep as-is with brackets
                                currentExpression += "[" + token.Text + "]";
                            }
                        }
                    }
                    else if (token.Type == 114 /*other identifier*/)
                    {
                        // Check if this UDF exists in the model before setting the flag
                        if (udfTemplates.ContainsKey(token.Text))
                        {
                            isDaxUdf = true;
                            daxUdfName = token.Text;
                        }
                        else
                        {
                            // Not a UDF, keep as-is
                            currentExpression += token.Text;
                        }

                    }
                    else if (token.Type == DaxToken.STRING_LITERAL)
                    {
                        currentExpression += @"""" + token.Text + @"""";
                    }
                    else
                    {
                        currentExpression += token.Text;
                    }
                }

                currentExpression = FormatDax(currentExpression);

                previousExpression = currentExpression;
                tokens = DaxTokenizer.Tokenize(previousExpression);
                i++;

            }

            Output(currentExpression.ReplaceLineEndings());
            Clipboard.SetText(currentExpression);
            Info("Expansion complete. Result copied to clipboard.");
        }
        void disableAutoDateTime()
        {

            // 2025-11-30 / B.Agullo
            // Disables the Auto Date/Time feature in the model by removing local date tables and date table templates.

            IList<string> LocalDateTables =
                Model.Tables.Where(
                    t => t.GetAnnotation("__PBI_LocalDateTable") == "true")
                .Select(t => t.Name).ToList();

            foreach (string tName in LocalDateTables)
            {
                Model.Tables[tName].Delete();
            }

            IList<string> DateTableTemplate =
                Model.Tables.Where(
                    t => t.GetAnnotation("__PBI_TemplateDateTable") == "true")
                .Select(t => t.Name).ToList();

            foreach (string tName in DateTableTemplate)
            {
                Model.Tables[tName].Delete();
            }

            Model.SetAnnotation("__PBI_TimeIntelligenceEnabled", "0");
        }
        void applyDisplayNames()
        {
            // 2026-06-21 / B.Agullo: now updates the report directly in Power BI Desktop without reopening it.
            //using GeneralFunctions;
            //using Report.DTO;
            //using System.IO;
            //using Newtonsoft.Json.Linq;

            // 2025-10-31 / B.Agullo
            // Initializes a report, lets user select visuals, iterates through projections,
            // and applies DisplayName annotations from the model to the projections in the report.

#if TE3
    ScriptHelper.WaitFormVisible = false;
#endif

            // Step 1: Initialize report
            ReportExtended report = Rx.InitReport();
            if (report == null) return;

            // Step 2: Let user select visuals
            IList<VisualExtended> selectedVisuals = Rx.SelectVisuals(report);
            if (selectedVisuals == null || selectedVisuals.Count == 0)
            {
                Info("No visuals selected.");
                return;
            }

            int updatedCount = 0;
            int visualsUpdated = 0;

            // Step 3: Process each selected visual
            foreach (var visual in selectedVisuals)
            {
                var queryState = visual.Content?.Visual?.Query?.QueryState;
                if (queryState == null) continue;

                // Create list of all projection sets to iterate
                var projectionSets = new List<VisualDto.ProjectionsSet>
                {
                    queryState.Values,
                    queryState.Y,
                    queryState.Y2,
                    queryState.Category,
                    queryState.Series,
                    queryState.Data,
                    queryState.Rows
                };

                bool visualModified = false;

                // Iterate through each projection set
                foreach (var projectionSet in projectionSets)
                {
                    if (projectionSet?.Projections == null) continue;

                    foreach (var projection in projectionSet.Projections)
                    {
                        if (projection?.Field == null) continue;

                        string displayNameFromModel = null;

                        // Check if it's a measure
                        if (projection.Field.Measure != null)
                        {
                            var measureExpr = projection.Field.Measure;
                            if (measureExpr.Expression?.SourceRef?.Entity != null
                                && measureExpr.Property != null)
                            {
                                string fullName = String.Format(
                                    "'{0}'[{1}]",
                                    measureExpr.Expression.SourceRef.Entity,
                                    measureExpr.Property);

                                var measure = Model.AllMeasures.FirstOrDefault(
                                    m => m.Table.DaxObjectFullName
                                        + m.DaxObjectFullName == fullName);

                                if (measure != null)
                                {
                                    displayNameFromModel =
                                        measure.GetAnnotation("DisplayName");
                                }
                            }
                        }
                        // Check if it's a column
                        else if (projection.Field.Column != null)
                        {
                            var columnExpr = projection.Field.Column;
                            if (columnExpr.Expression?.SourceRef?.Entity != null && columnExpr.Property != null)
                            {
                                string fullName = String.Format("'{0}'[{1}]", columnExpr.Expression.SourceRef.Entity, columnExpr.Property);
                                var column = Model.AllColumns.FirstOrDefault(c => c.DaxObjectFullName == fullName);

                                if (column != null)
                                {
                                    displayNameFromModel = column.GetAnnotation("DisplayName");
                                }
                            }
                        }

                        // Apply display name if found in model
                        if (!string.IsNullOrEmpty(displayNameFromModel))
                        {
                            // Check if projection already has a display name
                            if (string.IsNullOrEmpty(projection.DisplayName)
                                || projection.DisplayName != displayNameFromModel)
                            {
                                projection.DisplayName = displayNameFromModel;
                                updatedCount++;
                                visualModified = true;
                            }
                        }
                    }
                }

                // Save visual if it was modified
                if (visualModified)
                {
                    Rx.SaveVisual(visual);
                    visualsUpdated++;
                }
            }

            Output(String.Format("Updated {0} display names in {1} visuals of the report.", updatedCount, visualsUpdated));

            Rx.UpdateDesktop(report);
        }
        void storeDisplayNames()
        {
            //using GeneralFunctions;
            //using Report.DTO;
            //using System.IO;
            //using Newtonsoft.Json.Linq;

            // 2025-10-30 / B.Agullo
            // Initializes a report, lets user select visuals, iterates through projections to extract display names,
            // and stores them as "DisplayName" annotations in the model. Handles multiple display names by prompting user.

#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif

            // Step 1: Initialize report
            ReportExtended report = Rx.InitReport();
            if (report == null) return;

            // Step 2: Let user select visuals
            IList<VisualExtended> selectedVisuals = Rx.SelectVisuals(report);
            if (selectedVisuals == null || selectedVisuals.Count == 0)
            {
                Info("No visuals selected.");
                return;
            }

            // Step 3: Collect display names from selected visuals
            var measureDisplayNamesDict = new Dictionary<string, HashSet<string>>();
            var columnDisplayNamesDict = new Dictionary<string, HashSet<string>>();

            foreach (var visual in selectedVisuals)
            {
                var queryState = visual.Content?.Visual?.Query?.QueryState;
                if (queryState == null) continue;

                // Create list of all projection sets to iterate
                var projectionSets = new List<VisualDto.ProjectionsSet>
                {
                    queryState.Values,
                    queryState.Y,
                    queryState.Y2,
                    queryState.Category,
                    queryState.Series,
                    queryState.Data,
                    queryState.Rows
                };

                // Iterate through each projection set
                foreach (var projectionSet in projectionSets)
                {
                    if (projectionSet?.Projections == null) continue;

                    foreach (var projection in projectionSet.Projections)
                    {
                        if (projection?.Field == null) continue;

                        string displayName = projection.DisplayName;

                        if (string.IsNullOrEmpty(displayName)) continue;

                        // Check if it's a measure
                        if (projection.Field.Measure != null)
                        {
                            var measureExpr = projection.Field.Measure;
                            if (measureExpr.Expression?.SourceRef?.Entity != null && measureExpr.Property != null)
                            {
                                string fullName = String.Format("'{0}'[{1}]", measureExpr.Expression.SourceRef.Entity, measureExpr.Property);

                                if (!measureDisplayNamesDict.ContainsKey(fullName))
                                {
                                    measureDisplayNamesDict[fullName] = new HashSet<string>();
                                }

                                // Only add if it's different from the default field name
                                if (displayName != measureExpr.Property)
                                {
                                    measureDisplayNamesDict[fullName].Add(displayName);
                                }
                                else if (measureDisplayNamesDict[fullName].Count == 0)
                                {
                                    // If no custom display name yet, use the property name
                                    measureDisplayNamesDict[fullName].Add(measureExpr.Property);
                                }
                            }
                        }
                        // Check if it's a column
                        else if (projection.Field.Column != null)
                        {
                            var columnExpr = projection.Field.Column;
                            if (columnExpr.Expression?.SourceRef?.Entity != null && columnExpr.Property != null)
                            {
                                string fullName = String.Format("'{0}'[{1}]", columnExpr.Expression.SourceRef.Entity, columnExpr.Property);

                                if (!columnDisplayNamesDict.ContainsKey(fullName))
                                {
                                    columnDisplayNamesDict[fullName] = new HashSet<string>();
                                }

                                // Only add if it's different from the default field name
                                if (displayName != columnExpr.Property && displayName != fullName)
                                {
                                    columnDisplayNamesDict[fullName].Add(displayName);
                                }

                            }
                        }
                    }
                }
            }

            // Step 4: Resolve conflicts (multiple display names for same field)
            var measureDisplayNames = new Dictionary<string, string>();
            var columnDisplayNames = new Dictionary<string, string>();

            foreach (var kvp in measureDisplayNamesDict)
            {
                string fieldKey = kvp.Key;
                var displayNames = kvp.Value.ToList();

                if (displayNames.Count == 0) continue;

                if (displayNames.Count == 1)
                {
                    measureDisplayNames[fieldKey] = displayNames[0];
                }
                else
                {
                    // Multiple display names found - ask user
                    string chosen = Fx.ChooseString(
                        OptionList: displayNames,
                        label: String.Format("Multiple display names found for measure '{0}'. Choose one:", fieldKey)
                    );
                    if (string.IsNullOrEmpty(chosen))
                    {
                        Info("Operation cancelled.");
                        return;
                    }
                    measureDisplayNames[fieldKey] = chosen;
                }
            }

            foreach (var kvp in columnDisplayNamesDict)
            {
                string fieldKey = kvp.Key;
                var displayNames = kvp.Value.ToList();

                if (displayNames.Count == 0) continue;

                if (displayNames.Count == 1)
                {
                    columnDisplayNames[fieldKey] = displayNames[0];
                }
                else
                {
                    // Multiple display names found - ask user
                    string chosen = Fx.ChooseString(
                        OptionList: displayNames,
                        label: String.Format("Multiple display names found for column '{0}'. Choose one:", fieldKey)
                    );
                    if (string.IsNullOrEmpty(chosen))
                    {
                        Info("Operation cancelled.");
                        return;
                    }
                    columnDisplayNames[fieldKey] = chosen;
                }
            }

            // Step 5: Apply annotations to model
            int measuresUpdated = 0;
            int columnsUpdated = 0;

            foreach (var kvp in measureDisplayNames)
            {
                var measure = Model.AllMeasures.FirstOrDefault(
                    m => m.Table.DaxObjectFullName + m.DaxObjectFullName == kvp.Key);
                if (measure != null)
                {
                    if (measure.GetAnnotation("DisplayName") == kvp.Value) continue;

                    measure.SetAnnotation("DisplayName", kvp.Value);
                    measuresUpdated++;
                }
            }

            foreach (var kvp in columnDisplayNames)
            {
                var column = Model.AllColumns.FirstOrDefault(c => c.DaxObjectFullName == kvp.Key);
                if (column != null)
                {
                    if (column.GetAnnotation("DisplayName") == kvp.Value) continue;
                    column.SetAnnotation("DisplayName", kvp.Value);
                    columnsUpdated++;
                }
            }

            Output(String.Format("Updated {0} measures and {1} columns with DisplayName annotations.", measuresUpdated, columnsUpdated));
        }
        void copyConditionalFormatting()
        {
            // 2026-06-21 / B.Agullo: now updates the report directly in Power BI Desktop without reopening it.

            //using GeneralFunctions;
            //using Report.DTO;
            //using System.IO;
            //using Newtonsoft.Json.Linq;
            // 2025-10-22 / B.Agullo
            // Copies conditional formatting from one header to multiple other headers in a table visual.

#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif

            // Step 1: Initialize report
            ReportExtended report = Rx.InitReport();
            if (report == null) return;

            VisualExtended selectedVisual = Rx.SelectTableVisual(report);
            if (selectedVisual == null) return;

            // Step 2: Extract all headers from projections (not just those with formatting)
            var projectionHeaders = selectedVisual.Content?.Visual?.Query?.QueryState?.Values?.Projections
                .Select(p => p.QueryRef)
                .Where(h => !string.IsNullOrEmpty(h))
                .Distinct()
                .ToList();

            if (projectionHeaders == null || projectionHeaders.Count == 0)
            {
                Error("No headers found in the visual projections.");
                return;
            }

            // Step 3: Extract all displayed headers (with formatting objects)
            var formattedHeaders = selectedVisual.Content?.Visual?.Objects?.Values?
                .Select(cf => cf.Selector?.Metadata)
                .Where(h => !string.IsNullOrEmpty(h))
                .Distinct()
                .ToList();

            // Step 4: Let user choose the source header for formatting (from all projection headers)
            string sourceHeader = Fx.ChooseString(
                OptionList: projectionHeaders,
                label: "Select the header to copy formatting from"
            );
            if (string.IsNullOrEmpty(sourceHeader)) return;

            // Step 5: Let user choose target headers (multi-select, exclude source)
            List<string> targetHeaders = Fx.ChooseStringMultiple(
                OptionList: projectionHeaders.Where(h => h != sourceHeader).ToList(),
                label: "Select headers to apply the formatting to"
            );
            if (targetHeaders == null || targetHeaders.Count == 0)
            {
                Info("No target headers selected.");
                return;
            }

            // Step 6: Get source formatting (excluding selector)
            var sourceFormatting = selectedVisual.Content.Visual.Objects.Values
                .FirstOrDefault(cf => cf.Selector?.Metadata == sourceHeader);

            if (sourceFormatting == null)
            {
                Error("Source header formatting not found.");
                return;
            }

            // Step 7: Apply formatting to target headers
            int updatedCount = 0;
            foreach (var targetHeader in targetHeaders)
            {
                var targetFormatting = selectedVisual.Content.Visual.Objects.Values
                    .FirstOrDefault(cf => cf.Selector != null && cf.Selector.Metadata == targetHeader);

                if (targetFormatting != null)
                {
                    // Copy all properties except Selector
                    var sourceProps = typeof(VisualDto.ObjectProperties).GetProperties();
                    foreach (var prop in sourceProps)
                    {
                        if (prop.Name == "Selector") continue;
                        prop.SetValue(targetFormatting, prop.GetValue(sourceFormatting, null), null);
                    }
                    updatedCount++;
                }
                else
                {
                    // Create new ObjectProperties and copy all except Selector
                    var newFormatting = new VisualDto.ObjectProperties();
                    var sourceProps = typeof(VisualDto.ObjectProperties).GetProperties();
                    foreach (var prop in sourceProps)
                    {
                        if (prop.Name == "Selector")
                        {
                            // Create new Selector and set Metadata to targetHeader
                            newFormatting.Selector = new VisualDto.Selector
                            {
                                Data = new List<VisualDto.DataObject>()
                                {
                                    new VisualDto.DataObject {
                                        DataViewWildcard = new VisualDto.DataViewWildcard()
                                        {
                                            MatchingOption = 1
                                        }

                                    }
                                },
                                Metadata = targetHeader
                            };
                        }
                        else
                        {
                            prop.SetValue(newFormatting, prop.GetValue(sourceFormatting, null), null);
                        }
                    }
                    if (selectedVisual.Content.Visual.Objects.Values == null)
                        selectedVisual.Content.Visual.Objects.Values = new List<VisualDto.ObjectProperties>();
                    selectedVisual.Content.Visual.Objects.Values.Add(newFormatting);
                    updatedCount++;
                }
            }

            Rx.SaveVisual(selectedVisual);
            Output(String.Format(@"{0} headers updated with formatting from '{1}'.", updatedCount, sourceHeader));

            Rx.UpdateDesktop(report);
        }
        void adjustLineOverColumns()
        {
            // 2026-06-21 / B.Agullo: now updates the report directly in Power BI Desktop without reopening it.
            //using GeneralFunctions;
            //using Report.DTO;
            //using ReportFunctions;
            //using System.IO;
            //using Newtonsoft.Json.Linq;

            // 2025-10-11 / B.Agullo
            // Adjusts the line and column axis min/max values in a Line and Clustered Column Combo Chart to ensure that the line is always fully visible.
            // It creates a "Formatting" calculation table with measures and functions to calculate the axis min/max values dynamically based on the data in the visual.
            // It assumes the visual has at least one field in the x-axis, one for the columns and one for the line.
            // It modifies the visual to use these calculations for the axis min/max values.
            // It is recommended to create a backup of the model before running this script as it modifies both the model and the report.
            // The script only works with "lineClusteredColumnComboChart" visuals.

#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif

            // Check model compatibility level
            Fx.CheckCompatibilityVersion(Model, 1702, "Time Intelligence functions are only supported in 1702 or higher.Do you want to change the compatibility level to 1702?");
            if (Model.Database.CompatibilityLevel < 1702) return;


            //check if there's a matching visual in the report
            ReportExtended report = Rx.InitReport();
            if (report == null) return;

            List<string> visualTypes = new List<string>() { "lineClusteredColumnComboChart" };
            VisualExtended selectedVisual = Rx.SelectVisual(report, visualTypes);
            if (selectedVisual == null) return;

            var queryState = selectedVisual.Content?.Visual?.Query?.QueryState;

            var categories = queryState?.Category?.Projections ?? new List<VisualDto.Projection>();
            var columns = queryState?.Y?.Projections ?? new List<VisualDto.Projection>();
            var lines = queryState?.Y2?.Projections ?? new List<VisualDto.Projection>();


            if (categories.Count() == 0 || columns.Count() == 0 || lines.Count() == 0)
            {
                Error("Chart not completely configured. Please configure at least a field for x-axis, one for the columns and one for the line");
                return;
            }

            Table formattingTable = Fx.CreateCalcTable(Model, "Formatting");
            if (formattingTable == null) return;

            bool globalPaddingCreated = false;
            bool lineChartHeightCreated = false;
            bool secondaryMaxFunctionCreated = false;
            bool secondaryMinFunctionCreated = false;
            bool primaryMaxFunctionCreated = false;

            Measure globalPaddingMeasure = Fx.CreateMeasure(
                table: formattingTable,
                measureName: "GlobalPadding",
                measureExpression: "0.1",
                out globalPaddingCreated,
                description: "Global padding to apply to secondary axis max calculation (10% by default)",
                annotationLabel: "@AgulloBernat",
                annotationValue: "GlobalPadding",
                isHidden: true);
            if (globalPaddingMeasure == null) return;

            Measure lineChartHeight = Fx.CreateMeasure(
                table: formattingTable,
                measureName: "LineChartHeight",
                measureExpression: "0.4",
                out lineChartHeightCreated,
                description: "Percent of the chart height for the line chart",
                annotationLabel: "@AgulloBernat",
                annotationValue: "LineChartHeight",
                isHidden: true);
            if (lineChartHeight == null) return;

            string secondaryMaxName = "Formatting.AxisMaxMin.SecondaryMax";
            string secondaryMaxExpression =
                @"(
                    lineMaxExpression: ANYREF EXPR,
                    xAxisColumn: ANYREF EXPR,
                    paddingScalar:ANYVAL
                ) =>

                EXPAND( MAXX( ROWS, lineMaxExpression ), xAxisColumn ) * ( 1 + paddingScalar )";

            string secondaryMaxAnnotationLabel = "@AgulloBernat";
            string secondaryMaxAnnotationValue = "Formatting.AxisMaxMin.SecondaryMax";

            Function secondaryMaxFunction = Fx.CreateFunction(
                model: Model,
                name: secondaryMaxName,
                expression: secondaryMaxExpression,
                out secondaryMaxFunctionCreated,
                annotationLabel: secondaryMaxAnnotationLabel,
                annotationValue: secondaryMaxAnnotationValue);


            string secondaryMinName = "Formatting.AxisMaxMin.SecondaryMin";
            string secondaryMinExpression =
                @"(
                    lineMinExpression: ANYREF EXPR,
                    xAxisColumn: ANYREF EXPR,
                    paddingScalar: ANYVAL DECIMAL,
                    secondaryAxisMaxValue: ANYVAL DECIMAL,
                    lineChartWeight: ANYVAL DECIMAL
                ) =>
                VAR _lineMinVal =
                    EXPAND(
                        MINX( ROWS, lineMinExpression ),
                        xAxisColumn
                    )
                        * ( 1 - paddingScalar )
                VAR _lineHeight = secondaryAxisMaxValue - _lineMinVal
                VAR _secondaryAxisHeight = _lineHeight / lineChartWeight
                VAR _result = secondaryAxisMaxValue - _secondaryAxisHeight
                RETURN
                    _result";

            string secondaryMinAnnotationLabel = "@AgulloBernat";
            string secondaryMinAnnotationValue = "Formatting.AxisMaxMin.SecondaryMin";
            Function secondaryMinFunction = Fx.CreateFunction(
                model: Model,
                name: secondaryMinName,
                expression: secondaryMinExpression,
                out secondaryMinFunctionCreated,
                annotationLabel: secondaryMinAnnotationLabel,
                annotationValue: secondaryMinAnnotationValue);


            string primaryMaxName = "Formatting.AxisMaxMin.PrimaryMax";
            string primaryMaxExpression =
                @"(
                    columnMaxExpression: ANYREF EXPR,
                    xAxisColumn: ANYREF EXPR,
                    paddingScalar: ANYVAL DECIMAL,
                    lineChartWeight: ANYVAL DECIMAL
                ) =>
                VAR _maxColumnValue =
                    EXPAND(
                        MAXX( ROWS, columnMaxExpression ),
                        xAxisColumn
                    )
                        * ( 1 + paddingScalar )
                VAR _result = _maxColumnValue / ( 1 - lineChartWeight )
                RETURN
                    _result";

            string primaryMaxAnnotationLabel = "@AgulloBernat";
            string primaryMaxAnnotationValue = "Formatting.AxisMaxMin.PrimaryMax";
            Function primaryMaxFunction = Fx.CreateFunction(
                model: Model,
                name: primaryMaxName,
                expression: primaryMaxExpression,
                out primaryMaxFunctionCreated,
                annotationLabel: primaryMaxAnnotationLabel,
                annotationValue: primaryMaxAnnotationValue);

            if (globalPaddingCreated
                || lineChartHeightCreated
                || secondaryMaxFunctionCreated
                || secondaryMinFunctionCreated
                || primaryMaxFunctionCreated)
            {
                Info("Some elements were added to the semantic model. Commit changes to the model, save your progress and run this script again");
                return;
            }

            //all elements were already in place, time to proceed with the report layer


            var paddingProjection = new VisualDto.Projection
            {
                Field = new VisualDto.Field
                {
                    Measure = new VisualDto.MeasureObject
                    {
                        Expression = new VisualDto.Expression
                        {
                            SourceRef = new VisualDto.SourceRef
                            {
                                Entity = globalPaddingMeasure.Table.Name

                            }
                        },
                        Property = globalPaddingMeasure.Name
                    }
                },
                QueryRef = globalPaddingMeasure.Table.Name + "." + globalPaddingMeasure.Name,
                NativeQueryRef = globalPaddingMeasure.Name,
                Hidden = true
            };




            var lineChartHeightProjection = new VisualDto.Projection
            {
                Field = new VisualDto.Field
                {
                    Measure = new VisualDto.MeasureObject
                    {
                        Expression = new VisualDto.Expression
                        {
                            SourceRef = new VisualDto.SourceRef
                            {
                                Entity = lineChartHeight.Table.Name
                            }
                        },
                        Property = lineChartHeight.Name
                    }
                },
                QueryRef = lineChartHeight.Table.Name + "." + lineChartHeight.Name,
                NativeQueryRef = lineChartHeight.Name,
                Hidden = true
            };


            if (categories.Count() > 1)
            {
                Error("Multiple fields found in x-axis. Not implemented yet");
                return;
            }

            //variables used in different visual calculations
            string xAxisColumn = "[" + categories[0].NativeQueryRef + "]";
            string paddingScalar = "[" + paddingProjection.NativeQueryRef + "]";
            string lineChartWeight = "[" + lineChartHeightProjection.NativeQueryRef + "]";


            string secondaryMaxLineMaxExpression = String.Format(
                "MAXX({{{0}}},[Value])",
                "[" + string.Join(
                    "],[",
                    lines.Select(l => l.NativeQueryRef)) + "]"
                );

            string secondaryMaxVisualCalcExpression =
                String.Format(
                    @"{0}({1},{2},{3})",
                    secondaryMaxFunction.Name,
                    secondaryMaxLineMaxExpression,
                    xAxisColumn,
                    paddingScalar
                );

            var secondaryMaxVisualCalcProjection = new VisualDto.Projection
            {
                Field = new VisualDto.Field
                {
                    NativeVisualCalculation = new VisualDto.NativeVisualCalculation
                    {
                        Language = "dax",
                        Expression = secondaryMaxVisualCalcExpression,
                        Name = "secondaryMax"
                    }
                },
                QueryRef = "secondaryMax",
                NativeQueryRef = "secondaryMax",
                Hidden = true
            };



            string secondaryAxisMaxValue = "[" + secondaryMaxVisualCalcProjection.NativeQueryRef + "]";
            string lineMinExpression = String.Format(
                "MINX({{{0}}},[Value])",
                "[" + string.Join(
                    "],[",
                    lines.Select(l => l.NativeQueryRef)) + "]"
                );

            string secondaryMinVisualCalcExpression =
                String.Format(
                    @"{0}({1},{2},{3},{4},{5})",
                    secondaryMinFunction.Name,
                    lineMinExpression,
                    xAxisColumn,
                    paddingScalar,
                    secondaryAxisMaxValue,
                    lineChartWeight

                );

            var secondaryMinVisualCalcProjection = new VisualDto.Projection
            {
                Field = new VisualDto.Field
                {
                    NativeVisualCalculation = new VisualDto.NativeVisualCalculation
                    {
                        Language = "dax",
                        Expression = secondaryMinVisualCalcExpression,
                        Name = "secondaryMin"
                    }
                },
                QueryRef = "secondaryMin",
                NativeQueryRef = "secondaryMin",
                Hidden = true
            };



            string primaryMaxColumnMaxExpression = String.Format(
                "MAXX({{{0}}},[Value])",
                "[" + string.Join(
                    "],[",
                    columns.Select(l => l.NativeQueryRef)) + "]"
                );
            string primaryMaxVisualCalcExpression =
                String.Format(
                    @"{0}({1},{2},{3},{4})",
                    primaryMaxFunction.Name,
                    primaryMaxColumnMaxExpression,
                    xAxisColumn,
                    paddingScalar,
                    lineChartWeight

                );

            var primaryMaxVisualCalcProjection = new VisualDto.Projection
            {
                Field = new VisualDto.Field
                {
                    NativeVisualCalculation = new VisualDto.NativeVisualCalculation
                    {
                        Language = "dax",
                        Expression = primaryMaxVisualCalcExpression,
                        Name = "primaryMax"
                    }
                },
                QueryRef = "primaryMax",
                NativeQueryRef = "primaryMax",
                Hidden = true
            };


            columns.Add(paddingProjection);
            columns.Add(lineChartHeightProjection);
            columns.Add(secondaryMaxVisualCalcProjection);
            columns.Add(secondaryMinVisualCalcProjection);
            columns.Add(primaryMaxVisualCalcProjection);

            if (selectedVisual.Content.Visual.Objects == null)
                selectedVisual.Content.Visual.Objects = new VisualDto.Objects();

            if (selectedVisual.Content.Visual.Objects.ValueAxis == null)
                selectedVisual.Content.Visual.Objects.ValueAxis = new List<VisualDto.ObjectProperties>();

            // Ensure there's at least one ObjectProperties entry
            if (selectedVisual.Content.Visual.Objects.ValueAxis.Count == 0)
            {
                selectedVisual.Content.Visual.Objects.ValueAxis.Add(new VisualDto.ObjectProperties
                {
                    Properties = new Dictionary<string, object>()
                });
            }

            var valueAxisProperties = selectedVisual.Content.Visual.Objects.ValueAxis[0].Properties;

            // secondary axis min
            valueAxisProperties["secStart"] = new VisualDto.VisualObjectProperty
            {
                Expr = new VisualDto.VisualPropertyExpr
                {
                    SelectRef = new VisualDto.SelectRefExpression
                    {
                        ExpressionName = "secondaryMin"
                    }
                }
            };

            // secondary axis max
            valueAxisProperties["secEnd"] = new VisualDto.VisualObjectProperty
            {
                Expr = new VisualDto.VisualPropertyExpr
                {
                    SelectRef = new VisualDto.SelectRefExpression
                    {
                        ExpressionName = "secondaryMax"
                    }
                }
            };

            //main axis min
            valueAxisProperties["start"] = new VisualDto.VisualObjectProperty
            {
                Expr = new VisualDto.VisualPropertyExpr
                {
                    Literal = new VisualDto.VisualLiteral
                    {
                        Value = "0D"
                    }
                }
            };

            //main axis max
            valueAxisProperties["end"] = new VisualDto.VisualObjectProperty
            {
                Expr = new VisualDto.VisualPropertyExpr
                {
                    SelectRef = new VisualDto.SelectRefExpression
                    {
                        ExpressionName = "primaryMax"
                    }
                }
            };

            Rx.SaveVisual(selectedVisual);
            Info("Visual on page '"
                + selectedVisual.ParentPage.Page.DisplayName
                + "' has been modified.");

            Rx.UpdateDesktop(report);
        }
        void createTimeIntelFunctions()
        {
            //using GeneralFunctions;

            // 2025-09-29/B.Agullo
            // Creates Time Intelligence functions (CY, PY, YOY, YOYPCT) in the model if they do not exist.
            // It also creates a hidden calculated column and measure in the date table to handle cases where the fact table has no data for some dates.
            // The script assumes there is a date table and a fact table in the model.
            // The script will prompt the user to select the main date column in the fact table if there are multiple date columns.

#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif

            if (Model.Database.CompatibilityLevel < 1702)
            {
                if (Fx.IsAnswerYes("The model compatibility level is below 1702. Time Intelligence functions are only supported in 1702 or higher. Do to change the compatibility level to 1702?"))
                {
                    Model.Database.CompatibilityLevel = 1702;
                }
                else
                {
                    Info("Operation cancelled.");
                    return;
                }
            }


            Table dateTable = Fx.GetDateTable(model: Model);
            if (dateTable == null) return;

            Column dateColumn = Fx.GetDateColumn(dateTable);
            if (dateColumn == null) return;

            Table factTable = Fx.GetFactTable(model: Model);
            if (factTable == null) return;

            Column factTableDateColumn = null;

            IEnumerable<Column> factTableDateColumns =
                factTable.Columns.Where(
                    c => c.DataType == DataType.DateTime);

            if (factTableDateColumns.Count() == 0)
            {
                Error("No Date columns found in fact table " + factTable.Name);
                return;
            }

            if (factTableDateColumns.Count() == 1)
            {
                factTableDateColumn = factTableDateColumns.First();
            }
            else
            {
                factTableDateColumn = factTableDateColumns.First(
                    c => Model.Relationships.Any(
                        r => ((r.FromColumn == dateColumn && r.ToColumn == c)
                            || (r.ToColumn == dateColumn && r.FromColumn == c)
                                && r.IsActive)));

                factTableDateColumn = SelectColumn(factTableDateColumns, factTableDateColumn, "Select main date column from the fact table");
            }
            if (factTableDateColumn == null) return;

            string dateTableAuxColumnName = "DateWith" + factTable.Name.Replace(" ", "");
            string dateTableAuxColumnExpression = String.Format(@"{0} <= MAX({1})", dateColumn.DaxObjectFullName, factTableDateColumn.DaxObjectFullName);

            CalculatedColumn dateTableAuxColumn = dateTable.AddCalculatedColumn(dateTableAuxColumnName, dateTableAuxColumnExpression);
            dateTableAuxColumn.FormatDax();

            dateTableAuxColumn.IsHidden = true;

            string dateTableAuxMeasureName = "ShowValueForDates";
            string dateTableAuxMeasureExpression =
                String.Format(
                    @"VAR LastDateWithData =
                        CALCULATE (
                            MAX ( {0} ),
                            REMOVEFILTERS ()
                        )
                    VAR FirstDateVisible =
                        MIN ( {1} )
                    VAR Result =
                        FirstDateVisible <= LastDateWithData
                    RETURN
                        Result",
                    factTableDateColumn.DaxObjectFullName,
                    dateColumn.DaxObjectFullName);

            Measure dateTableAuxMeasure = dateTable.AddMeasure(dateTableAuxMeasureName, dateTableAuxMeasureExpression);
            dateTableAuxMeasure.IsHidden = true;
            dateTableAuxMeasure.FormatDax();


            //CY --just for the sake of completion 
            string CYfunctionName = "Local.TimeIntel.CY";
            string CYfunctionExpression = "(baseMeasure) => baseMeasure";

            Function CYfunction = Model.AddFunction(CYfunctionName);
            CYfunction.Expression = CYfunctionExpression;
            CYfunction.FormatDax();

            CYfunction.SetAnnotation("displayFolder", @"baseMeasureDisplayFolder\baseMeasureName TimeIntel");
            CYfunction.SetAnnotation("formatString", "baseMeasureFormatStringFull");
            CYfunction.SetAnnotation("outputType", "Measure");
            CYfunction.SetAnnotation("nameTemplate", "baseMeasureName CY");
            CYfunction.SetAnnotation("outputDestination", "baseMeasureTable");

            //PY

            string PYfunctionName = "Local.TimeIntel.PY";
            string PYfunctionExpression =
                String.Format(
                    @"(baseMeasure: ANYREF) =>
                    IF(
                        {0},
                        CALCULATE(         
                            baseMeasure,
                            CALCULATETABLE(
                                DATEADD(
                                    {1},
                                    -1,
                                    YEAR
                                ),
                                {2} = TRUE
                            )
                        )
                    )",
                    dateTableAuxMeasure.DaxObjectFullName,
                    dateColumn.DaxObjectFullName,
                    dateTableAuxColumn.DaxObjectFullName);

            Function PYfunction = Model.AddFunction(PYfunctionName);
            PYfunction.Expression = PYfunctionExpression;
            PYfunction.FormatDax();

            PYfunction.SetAnnotation("displayFolder", @"baseMeasureDisplayFolder\baseMeasureName TimeIntel");
            PYfunction.SetAnnotation("formatString", "baseMeasureFormatStringFull");
            PYfunction.SetAnnotation("outputType", "Measure");
            PYfunction.SetAnnotation("nameTemplate", "baseMeasureName PY");
            PYfunction.SetAnnotation("outputDestination", "baseMeasureTable");

            //YOY
            string YOYfunctionName = "Local.TimeIntel.YOY";
            string YOYfunctionExpression =
                @"(baseMeasure: ANYREF) =>
                VAR ValueCurrentPeriod = Local.TimeIntel.CY(baseMeasure)
                VAR ValuePreviousPeriod = Local.TimeIntel.PY(baseMeasure)
                VAR Result =
	                IF(
		                NOT ISBLANK( ValueCurrentPeriod )
			                && NOT ISBLANK( ValuePreviousPeriod ),
		                ValueCurrentPeriod
			                - ValuePreviousPeriod
	                )
                RETURN
	                Result";

            Function YOYfunction = Model.AddFunction(YOYfunctionName);
            YOYfunction.Expression = YOYfunctionExpression;
            YOYfunction.FormatDax();

            YOYfunction.SetAnnotation("displayFolder", @"baseMeasureDisplayFolder\baseMeasureName TimeIntel");
            YOYfunction.SetAnnotation("formatString", "+baseMeasureFormatStringRoot;-baseMeasureFormatStringRoot;-");
            YOYfunction.SetAnnotation("outputType", "Measure");
            YOYfunction.SetAnnotation("nameTemplate", "baseMeasureName YOY");
            YOYfunction.SetAnnotation("outputDestination", "baseMeasureTable");

            //YOY%
            string YOYPfunctionName = "Local.TimeIntel.YOYPCT";
            string YOYPfunctionExpression =
                @"(baseMeasure: ANYREF) =>
                VAR ValueCurrentPeriod = Local.TimeIntel.CY(baseMeasure)
                VAR ValuePreviousPeriod = Local.TimeIntel.PY(baseMeasure)
                VAR CurrentMinusPreviousPeriod =
	                IF(
		                NOT ISBLANK( ValueCurrentPeriod )
			                && NOT ISBLANK( ValuePreviousPeriod ),
		                ValueCurrentPeriod
			                - ValuePreviousPeriod
	                )
                VAR Result =
	                DIVIDE(
		                CurrentMinusPreviousPeriod,
		                ValuePreviousPeriod
	                )
                RETURN
	                Result";
            Function YOYPfunction = Model.AddFunction(YOYPfunctionName);
            YOYPfunction.Expression = YOYPfunctionExpression;
            YOYPfunction.FormatDax();

            YOYPfunction.SetAnnotation("displayFolder", @"baseMeasureDisplayFolder\baseMeasureName TimeIntel");
            YOYPfunction.SetAnnotation("formatString", "+0.0%;-0.0%;-");
            YOYPfunction.SetAnnotation("outputType", "Measure");
            YOYPfunction.SetAnnotation("nameTemplate", "baseMeasureName YOY%");
            YOYPfunction.SetAnnotation("outputDestination", "baseMeasureTable");


        }

        void copyHeaderFormatting()
        {
            // 2026-06-21 / B.Agullo: now updates the report directly in Power BI Desktop without reopening it.

            //using GeneralFunctions;
            //using Report.DTO;
            //using System.IO;
            //using Newtonsoft.Json.Linq;

            // 2025-10-22/B.Agullo
            // This script copies column header formatting from one header to multiple other headers in a table visual.
#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif

            // Step 1: Initialize report
            ReportExtended report = Rx.InitReport();
            if (report == null) return;

            VisualExtended selectedVisual = Rx.SelectTableVisual(report);
            if (selectedVisual == null) return;

            // Step 2: Extract all headers from projections (not just those with formatting)
            var projectionHeaders = selectedVisual.Content?.Visual?.Query?.QueryState?.Values?.Projections
                .Select(p => p.QueryRef)
                .Where(h => !string.IsNullOrEmpty(h))
                .Distinct()
                .ToList();

            if (projectionHeaders == null || projectionHeaders.Count == 0)
            {
                Error("No headers found in the visual projections.");
                return;
            }

            // Step 3: Extract all displayed headers (with formatting objects)
            var formattedHeaders = selectedVisual.Content?.Visual?.Objects?.ColumnFormatting?
                .Select(cf => cf.Selector?.Metadata)
                .Where(h => !string.IsNullOrEmpty(h))
                .Distinct()
                .ToList();

            // Step 4: Let user choose the source header for formatting (from all projection headers)
            string sourceHeader = Fx.ChooseString(
                OptionList: projectionHeaders,
                label: "Select the header to copy formatting from"
            );
            if (string.IsNullOrEmpty(sourceHeader)) return;

            // Step 5: Let user choose target headers (multi-select, exclude source)
            List<string> targetHeaders = Fx.ChooseStringMultiple(
                OptionList: projectionHeaders.Where(h => h != sourceHeader).ToList(),
                label: "Select headers to apply the formatting to"
            );
            if (targetHeaders == null || targetHeaders.Count == 0)
            {
                Info("No target headers selected.");
                return;
            }

            // Step 6: Get source formatting (excluding selector)
            var sourceFormatting = selectedVisual.Content.Visual.Objects.ColumnFormatting
                .FirstOrDefault(cf => cf.Selector?.Metadata == sourceHeader);

            if (sourceFormatting == null)
            {
                Error("Source header formatting not found.");
                return;
            }

            // Step 7: Apply formatting to target headers
            int updatedCount = 0;
            foreach (var targetHeader in targetHeaders)
            {
                var targetFormatting = selectedVisual.Content.Visual.Objects.ColumnFormatting
                    .FirstOrDefault(cf => cf.Selector != null && cf.Selector.Metadata == targetHeader);

                if (targetFormatting != null)
                {
                    // Copy all properties except Selector
                    var sourceProps = typeof(VisualDto.ObjectProperties).GetProperties();
                    foreach (var prop in sourceProps)
                    {
                        if (prop.Name == "Selector") continue;
                        prop.SetValue(targetFormatting, prop.GetValue(sourceFormatting, null), null);
                    }
                    updatedCount++;
                }
                else
                {
                    // Create new ObjectProperties and copy all except Selector
                    var newFormatting = new VisualDto.ObjectProperties();
                    var sourceProps = typeof(VisualDto.ObjectProperties).GetProperties();
                    foreach (var prop in sourceProps)
                    {
                        if (prop.Name == "Selector")
                        {
                            // Create new Selector and set Metadata to targetHeader
                            newFormatting.Selector = new VisualDto.Selector { Metadata = targetHeader };
                        }
                        else
                        {
                            prop.SetValue(newFormatting, prop.GetValue(sourceFormatting, null), null);
                        }
                    }
                    if (selectedVisual.Content.Visual.Objects.ColumnFormatting == null)
                        selectedVisual.Content.Visual.Objects.ColumnFormatting = new List<VisualDto.ObjectProperties>();
                    selectedVisual.Content.Visual.Objects.ColumnFormatting.Add(newFormatting);
                    updatedCount++;
                }
            }

            Rx.SaveVisual(selectedVisual);
            Output(String.Format(@"{0} headers updated with formatting from '{1}'.", updatedCount, sourceHeader));

            Rx.UpdateDesktop(report);
        }





        void createTextMeasures()
        {
            //using GeneralFunctions;


            //2025-07-28/B.Agullo
            //This script creates text measures based on the selected measures in the model.
            //It prompts the user for a prefix and suffix to be added to the text measures.
            //It also allows the user to specify a suffix for the names of the new text measures.

#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif

            if (Selected.Measures.Count() == 0)
            {
                Error("No measures selected. Please select at least one measure.");
                return;
            }

            // Ask user for prefix
            string prefix = Fx.GetNameFromUser(
                Prompt: "Enter a prefix for the new text measures (use ### for current measure name):",
                Title: "Text Measure Prefix",
                DefaultResponse: ""
            );
            if (prefix == null) return;



            // Ask user for suffix
            string suffix = Fx.GetNameFromUser(
                Prompt: "Enter a suffix for the new text measures (use ### for current measure name):",
                Title: "Text Measure Suffix",
                DefaultResponse: ""
            );
            if (suffix == null) return;

            // Ask user for measure name suffix
            string measureNameSuffix = Fx.GetNameFromUser(
                Prompt: "Enter a suffix for the Name of the new text measures:",
                Title: "Suffix for names!",
                DefaultResponse: " Text"
            );
            if (measureNameSuffix == null) return;



            foreach (Measure m in Selected.Measures)
            {
                string newMeasureName = m.Name + measureNameSuffix;
                string newMeasureDisplayFolder = (m.DisplayFolder + measureNameSuffix).Trim();
                string newMeasureExpression =
                    String.Format(
                        @"""{2}"" & FORMAT([{0}], ""{1}"") & ""{3}""",
                        m.Name,
                        m.FormatString,
                        prefix.Replace("###", m.Name),
                        suffix.Replace("###", m.Name));
                Measure newMeasure = m.Table.AddMeasure(newMeasureName, newMeasureExpression, newMeasureDisplayFolder);
                newMeasure.FormatDax();
            }
        }

        void removeEmptyFolders()
        {
            // 2026-06-21 / B.Agullo: now updates the report directly in Power BI Desktop without reopening it.
            //using System.IO;
            //using Newtonsoft.Json.Linq;
            //using Report.DTO;
            //using GeneralFunctions;

#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif


            // Prompt user to select report
            var report = Rx.InitReport("Select PBIR file to clean up empty visual folders");
            if (report == null)
            {
                Info("Operation cancelled or failed to load report.");
                return;
            }

            int removedCount = 0;

            foreach (var page in report.Pages)
            {
                if (page == null || string.IsNullOrEmpty(page.PageFilePath))
                    continue;

                string pageFolder = Path.GetDirectoryName(page.PageFilePath);
                if (string.IsNullOrEmpty(pageFolder))
                    continue;

                string visualsFolder = Path.Combine(pageFolder, "visuals");
                if (!Directory.Exists(visualsFolder))
                    continue;

                var visualSubfolders = Directory.GetDirectories(visualsFolder);
                foreach (var visualFolder in visualSubfolders)
                {
                    string visualJsonPath = Path.Combine(visualFolder, "visual.json");
                    if (!File.Exists(visualJsonPath))
                    {
                        try
                        {
                            Directory.Delete(visualFolder, true);
                            removedCount++;
                        }
                        catch (Exception ex)
                        {
                            Output(String.Format("Failed to remove folder '{0}': {1}", visualFolder, ex.Message));
                        }
                    }
                }
            }

            Info(String.Format("Removed {0} empty visual folders.", removedCount));

            Rx.UpdateDesktop(report);
        }

        void AddBilingualLayerToReport()
        {
            // 2026-06-21 / B.Agullo: now updates the report directly in Power BI Desktop without reopening it.
            //using GeneralFunctions;
            //using Report.DTO;
            //using System.IO;
            //using Newtonsoft.Json.Linq;
            //2025-06-23/B.Agullo
            //this script adds a bilingual layer to the report, allowing the user to select the language of the report.
            //this will only prepare the report for an extraction of the definition as descrived in https://www.esbrina-ba.com/transforming-a-regular-report-into-a-bilingual-one-part-2-extracting-display-names-of-measures-and-field-prameters/

#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif

            ReportExtended report = Rx.InitReport();
            if (report == null) return;


            string altTextFlag = Fx.GetNameFromUser(
                Prompt: "Enter the flag for the original language (e.g., 'EN' for English):",
                Title: "Alternative Language Flag",
                DefaultResponse: "EN"
            );

            if (string.IsNullOrEmpty(altTextFlag))
            {
                Info("Operation cancelled.");
                return;
            }

            int totalCount = 0;

            // For each page, process visuals
            foreach (var pageExt in report.Pages)
            {
                var visuals = (pageExt.Visuals ?? new List<VisualExtended>())
                    .OrderBy(v => v.Content.Position.Y)
                    .ThenBy(v => v.Content.Position.X)
                    .ToList();

                int bilingualCounter = 1;

                foreach (var visual in visuals)
                {
                    // Skip if already in a bilingual group or if it's a group itself
                    if (visual.IsInBilingualVisualGroup()) continue;
                    if (visual.isVisualGroup) continue;

                    // Duplicate the visual (deep copy)
                    VisualExtended duplicate = Rx.DuplicateVisual(visual);

                    // Add the duplicate to the page
                    pageExt.Visuals.Add(duplicate);

                    // Prepare bilingual group name, ensure uniqueness
                    string pagePrefix = String.Format("P{0:00}", visual.ParentPage.PageIndex + 1);

                    string groupSuffix = String.Format("{0:000}", bilingualCounter);

                    string bilingualGroupDisplayName = pagePrefix + "-" + groupSuffix;

                    // Check for existing group with the same display name and increment counter if needed
                    while (pageExt.Visuals.Any(v =>
                        v.isVisualGroup &&
                        v.Content.VisualGroup != null &&
                        v.Content.VisualGroup.DisplayName == bilingualGroupDisplayName))
                    {
                        bilingualCounter++;
                        groupSuffix = String.Format("{0:000}", bilingualCounter);
                        bilingualGroupDisplayName = pagePrefix + "-" + groupSuffix;
                    }

                    string originalVisualGroupName = visual.Content.ParentGroupName;

                    List<VisualExtended> visualsToGroup = new List<VisualExtended> { visual, duplicate };

                    // Create bilingual visual group
                    VisualExtended visualGroup = Rx.GroupVisuals(visualsToGroup, groupDisplayName: bilingualGroupDisplayName);

                    //configure the original visual group if existed
                    if (originalVisualGroupName != null)
                    {
                        visualGroup.Content.ParentGroupName = originalVisualGroupName;
                    }

                    //set the altText flag 
                    string currentAltText = visual.AltText ?? "";
                    if (!currentAltText.StartsWith(altTextFlag))
                    {
                        visual.AltText = String.Format(@"{0} {1}", altTextFlag, currentAltText).Trim();
                    }

                    // Remove flag from duplicate's altText if present
                    string duplicateAltText = duplicate.AltText ?? "";
                    if (duplicateAltText.StartsWith(altTextFlag))
                    {
                        duplicate.AltText = duplicateAltText.Substring(altTextFlag.Length).TrimStart();
                    }

                    //hide the original visual
                    visual.Content.IsHidden = true;

                    Rx.SaveVisual(visual);
                    Rx.SaveVisual(duplicate);
                    Rx.SaveVisual(visualGroup);

                    bilingualCounter++;
                    totalCount++;
                }
            }

            Output(String.Format("Bilingual visual groups created for {0} visuals.", totalCount));

            Rx.UpdateDesktop(report);
        }


        void CopyVisual()
        {
            //using GeneralFunctions;
            //using Report.DTO;
            //using System.IO;
            //using Newtonsoft.Json.Linq;

            // 2025-07-05/B.Agullo
            // This script will copy a visual from a template report to the target report. 
            // Target report must be connected with the model that this instance of tabular editor is connected to. 
            // Both target report and template report must use PBIR format
            // If you are executing this in Tabular Editor 2 you need to 
            // configure Roslyn compiler as explained here:
            // https://docs.tabulareditor.com/te2/Advanced-Scripting.html#compiling-with-roslyn

#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif



            // Step 1: Initialize source and target reports
            ReportExtended sourceReport = Rx.InitReport(label: @"Select the SOURCE report");
            if (sourceReport == null) return;

            ReportExtended targetReport = Rx.InitReport(label: @"Select the TARGET report");
            if (targetReport == null) return;

            IList<VisualExtended> sourceVisuals = Rx.SelectVisuals(sourceReport);

            // If no visuals were selected, exit
            if (sourceVisuals == null || sourceVisuals.Count == 0) return;

            // Step 5: Ask in which page of the target report the new visual should be created
            var targetPages = targetReport.Pages.ToList();
            var pageDisplayList = targetPages.Select(p => p.Page.DisplayName).ToList();
            string newPageOption = @"<Create new page>";
            pageDisplayList.Add(newPageOption);
            string selectedPageDisplay = Fx.ChooseString(OptionList: pageDisplayList, label: @"Select target page for the new visual");
            if (String.IsNullOrEmpty(selectedPageDisplay))
            {
                Info(@"No target page selected.");
                return;
            }

            object targetPage = null;
            // Step 5.1: If the user selected the option to create a new page, replicate the first page as blank
            if (selectedPageDisplay == newPageOption)
            {
                targetPage = Rx.ReplicateFirstPageAsBlank(targetReport);
            }
            else
            {
                targetPage = targetPages.First(p => p.Page.DisplayName == selectedPageDisplay);
            }



            // Create a mapping from original visual names to new GUID-based names
            var visualNameMap = new Dictionary<string, string>();
            foreach (var vis in sourceVisuals)
            {
                string newGuidName = Guid.NewGuid().ToString().Replace("-", "").Substring(0, 20);
                visualNameMap[vis.Content.Name] = newGuidName;
            }

            // Prepare replacement maps. Once a reference is replaced, it will be stored in these maps to avoid re-selection.
            var measureReplacementMap = new Dictionary<string, Measure>();
            var columnReplacementMap = new Dictionary<string, Column>();


            // Replacement maps for filterConfig patch
            var tableReplacementMap = new Dictionary<string, string>();
            var fieldReplacementMap = new Dictionary<string, string>();

            int visualsCount = 0;

            // Step 2: Let user select a single visual from the source report
            foreach (VisualExtended sourceVisual in sourceVisuals)
            {
                if (sourceVisual == null) return;

                // Step 3: For each measure and column used, find equivalent in connected model and replace
                var referencedMeasures = sourceVisual.GetAllReferencedMeasures().ToList();
                var referencedColumns = sourceVisual.GetAllReferencedColumns().ToList();


                foreach (string measureRef in referencedMeasures)
                {
                    // If measureRef is already in the dictionary, use the existing replacement
                    Measure replacement;

                    if (measureReplacementMap.ContainsKey(measureRef))
                    {
                        replacement = measureReplacementMap[measureRef];
                    }
                    else
                    {
                        Measure preselect = Model.AllMeasures.FirstOrDefault(m =>
                            String.Format(@"{0}[{1}]", m.Table.DaxObjectFullName, m.Name) == measureRef
                        );
                        replacement = SelectMeasure(preselect: preselect, label: String.Format(@"Select replacement for measure {0}", measureRef));
                        if (replacement == null)
                        {
                            Error(String.Format(@"No replacement selected for measure {0}.", measureRef));
                            return;
                        }
                        measureReplacementMap[measureRef] = replacement;

                        string oldTable = measureRef.Split('[')[0].Trim('\'');
                        string oldField = measureRef.Split('[', ']')[1];

                        tableReplacementMap[oldTable] = replacement.Table.Name;
                        fieldReplacementMap[oldField] = replacement.Name;
                    }

                }

                foreach (string columnRef in referencedColumns)
                {
                    Column replacement;
                    if (columnReplacementMap.ContainsKey(columnRef))
                    {
                        replacement = columnReplacementMap[columnRef];
                    }
                    else
                    {


                        Column preselect = Model.AllColumns.FirstOrDefault(c =>
                        c.DaxObjectFullName == columnRef
                        );
                        replacement = SelectColumn(Model.AllColumns, preselect: preselect, label: String.Format(@"Select replacement for column {0}", columnRef));
                        if (replacement == null)
                        {
                            Error(String.Format(@"No replacement selected for column {0}.", columnRef));
                            return;
                        }
                        columnReplacementMap[columnRef] = replacement;

                        string oldTable = columnRef.Split('[')[0].Trim('\'');
                        string oldField = columnRef.Split('[', ']')[1];

                        tableReplacementMap[oldTable] = replacement.Table.Name;
                        fieldReplacementMap[oldField] = replacement.Name;

                    }
                }

                // Step 4: Replace fields in the visual object
                foreach (var kv in measureReplacementMap)
                {
                    sourceVisual.ReplaceMeasure(kv.Key, kv.Value);
                }
                foreach (var kv in columnReplacementMap)
                {
                    sourceVisual.ReplaceColumn(kv.Key, kv.Value);
                }



                // Step 5.2: Assign a new GUID as the visual name to avoid conflicts
                string newVisualName = visualNameMap[sourceVisual.Content.Name];
                sourceVisual.Content.Name = newVisualName;

                if (sourceVisual.Content.ParentGroupName != null)
                {
                    string newParentGroupName = visualNameMap.ContainsKey(sourceVisual.Content.ParentGroupName)
                        ? visualNameMap[sourceVisual.Content.ParentGroupName]
                        : null;

                    sourceVisual.Content.ParentGroupName = newParentGroupName;
                }

                // Step 6: Build new visual file path
                string targetPageFolder = Path.GetDirectoryName(((PageExtended)targetPage).PageFilePath);
                string visualsFolder = Path.Combine(targetPageFolder, "visuals");
                string newVisualJsonPath = Path.Combine(visualsFolder, sourceVisual.Content.Name, "visual.json");

                // Update visual's file path and parent page
                sourceVisual.VisualFilePath = newVisualJsonPath;

            }


            Output(String.Format(@"{0} Visuals copied to page '{1}' in target report.", visualsCount, ((PageExtended)targetPage).Page.DisplayName));

        }

        void openVisualJsonFile()
        {
            //using GeneralFunctions;
            //using Report.DTO;
            //using System.IO;
            //using Newtonsoft.Json.Linq;

#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif

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
                String.Format(
                    @"{0} - {1} ({2}, {3})",
                    x.Page.DisplayName,
                    x.Visual?.Content?.Visual?.VisualType
                        ?? x.Visual?.Content?.VisualGroup?.DisplayName,
                    (int)x.Visual.Content.Position.X,
                    (int)x.Visual.Content.Position.Y)
            ).ToList();

            // Step 4: Let the user select one or more visuals
            List<string> selected = Fx.ChooseStringMultiple(OptionList: visualDisplayList, label: "Select visuals to open JSON files");
            if (selected == null || selected.Count == 0)
            {
                Info("No visuals selected.");
                return;
            }

            // Step 5: For each selected visual, open its JSON file
            foreach (var visualEntry in allVisuals)
            {
                string display = String.Format
                    (@"{0} - {1} ({2}, {3})",
                    visualEntry.Page.DisplayName,
                    visualEntry?.Visual?.Content?.Visual?.VisualType
                        ?? visualEntry.Visual?.Content?.VisualGroup?.DisplayName,
                    (int)visualEntry.Visual.Content.Position.X,
                    (int)visualEntry.Visual.Content.Position.Y);

                if (selected.Contains(display))
                {
                    string jsonPath = visualEntry.Visual.VisualFilePath;
                    if (!File.Exists(jsonPath))
                    {
                        Error(String.Format(@"JSON file not found: {0}", jsonPath));
                        continue;
                    }
                    var psi = new System.Diagnostics.ProcessStartInfo()
                    {
                        FileName = jsonPath,
                        UseShellExecute = true  // This is the crucial part
                    };
                    System.Diagnostics.Process.Start(psi);
                }
            }
        }

        void replaceField()
        {
            // 2026-06-21 / B.Agullo: now updates the report directly in Power BI Desktop without reopening it.

            //2025-05-25/B.Agullo
            //provided a definition.pbir file, this script allows the user to replace a measure in all visuals that use it with another measure.
            //when executing the script you must be connected to the semantic model to which the report is connected to or one that is identical. 
            //see https://www.esbrina-ba.com/pbir-scripts-to-replace-field-and-open-visual-json-files/ for reference on how to use it

            //using GeneralFunctions;
            //using Report.DTO;
            //using System.IO;
            //using Newtonsoft.Json.Linq;

#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif

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

            Rx.UpdateDesktop(report);
        }


        void fixBrokenFields()
        {
            // 2026-06-21 / B.Agullo: now updates the report directly in Power BI Desktop without reopening it.
            //using GeneralFunctions;
            //using Report.DTO;
            //using System.IO;
            //using Newtonsoft.Json.Linq;

#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif

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

            if (!brokenMeasures.Any() && !brokenColumns.Any())
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



            // Save modified visuals
            foreach (var visual in modifiedVisuals)
            {
                Rx.SaveVisual(visual);
            }

            Output($"{modifiedVisuals.Count} visuals were modified.");

            Rx.UpdateDesktop(report);
        }




        void changeCoordinatesOfVisual()
        {
            // 2026-06-21 / B.Agullo: now updates the report directly in Power BI Desktop without reopening it.
            //using GeneralFunctions;
            //using Report.DTO;
            //using System.IO;

#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif

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

            Rx.UpdateDesktop(report);
        }



        void myNewScript()
        {
#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif

            //create a measure for each of the selected columns
            foreach (Column c in Selected.Columns)
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

#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif

            Fx.CreateCalcTable(Model, "myMeasures", "{0}");


        }



        //code snippets
        void userChooseName()
        {
            //#r "Microsoft.VisualBasic"
            //using Microsoft.VisualBasic;

#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif

            string calcGroupName = Interaction.InputBox("Provide a name for your Calc Group", "Calc Group Name", "Time Intelligence", 740, 400);

            //sample code using the variable
            Output(calcGroupName);

        }

        void userChooseYesNo()
        {

            //using System.Windows.Forms;

#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif

            DialogResult dialogResult = MessageBox.Show(text: "Generate Field Parameter?", caption: "Field Parameter", buttons: MessageBoxButtons.YesNo);
            bool generateFieldParameter = (dialogResult == DialogResult.Yes);

            //sample code using the variable
            Output(generateFieldParameter);

        }

        void userChooseString()
        {

            //using System.Windows.Forms;

#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif

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
                form.Width = 500;
                form.Height = 400;
                form.StartPosition = FormStartPosition.CenterScreen;
                form.Padding = new Padding(10);

                var buttonPanel = new Panel();
                buttonPanel.Dock = DockStyle.Bottom;
                buttonPanel.Height = 50;
                buttonPanel.Padding = new Padding(5);

                var okButton = new System.Windows.Forms.Button()
                {
                    DialogResult = DialogResult.OK,
                    Text = "OK",
                    Width = 100,
                    Height = 35,
                    Left = 10,
                    Top = 5
                };

                var cancelButton = new System.Windows.Forms.Button()
                {
                    DialogResult = DialogResult.Cancel,
                    Text = "Cancel",
                    Width = 100,
                    Height = 35,
                    Left = 120,
                    Top = 5
                };

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

        }
        public static void sayHelloWorld()
        {
            Info("Hello World");
        }
        public static void CopyMacroFromVSFileWithDll()
        {

            // NOCOPY replace <PROJECT FOLDER> (both instances) with the path to the folder where the .sln file is stored.
            //#r "<PROJECT FOLDER>\TE Scripts\bin\Debug\net48\TE Scripts.dll"
            //using TE_Scripting;

            string baseFolderPath = @"<PROJECT FOLDER>";
            string macroFilePath = String.Format(@"{0}\TE Scripts\TE Scripts.cs", baseFolderPath);
            string generalFunctionsClassFilePath = String.Format(@"{0}\GeneralFunctions\GeneralFunctions.cs", baseFolderPath);
            string reportClassFilePath = String.Format(@"{0}\Report\Report.cs", baseFolderPath);
            string daxUserDefinedFunctionClassFilePath = String.Format(@"{0}\DaxUserDefinedFunction\DaxUserDefinedFunction.cs", baseFolderPath);
            string reportFunctionsClassFilePath = String.Format(@"{0}\ReportFunctions\ReportFunctions.cs", baseFolderPath);

            TE_Scripting.TE_Scripts.CopyMacroFromVSFile(
                macroFilePath, generalFunctionsClassFilePath, reportClassFilePath, reportFunctionsClassFilePath, daxUserDefinedFunctionClassFilePath
            );
        }


        public static void CopyMacroFromVSFile
            (string macroFilePath,
            string generalFunctionsClassFilePath,
            string reportClassFilePath,
            string reportFunctionsClassFilePath,
            string daxUserDefinedFunctionClassFilePath
            )
        {
            //#r "System.IO"
            //#r "Microsoft.CodeAnalysis"
            //using System.IO;
            //using System.Windows.Forms;
            //using Microsoft.CodeAnalysis;
            //using Microsoft.CodeAnalysis.CSharp;
            //using Microsoft.CodeAnalysis.CSharp.Syntax;
            //using System.Text.RegularExpressions;

            // '2026-10-29 / B.Agullo / Helper methods not reached in CLI mode are wrapped in #if !TECLI so skill-ready macros compile in te
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
            //String daxUserDefinedFunctionClassFilePath = String.Format(@"{0}\DaxUserDefinedFunction\DaxUserDefinedFunction.cs", baseFolderPath);
            //String reportFunctionsClassFilePath = String.Format(@"{0}\ReportFunctions\ReportFunctions.cs", baseFolderPath);
            String codeIndent = "            ";
            String noCopyMark = "NOCOPY";
            ////these libraries are already loaded in Tabular Editor and must not be specified
            //String[] tabularEditorLibraries = { "#r \"System.Windows.Forms\"" };
            //get file structure
            SyntaxTree tree = CSharpSyntaxTree.ParseText(File.ReadAllText(macroFilePath));

            List<string> macroNames = new List<string>();

#if TE3
            ScriptHelper.WaitFormVisible = false;
#endif


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
                form.Width = 500;
                form.Height = 400;
                form.StartPosition = FormStartPosition.CenterScreen;
                form.Padding = new Padding(10);

                var buttonPanel = new Panel();
                buttonPanel.Dock = DockStyle.Bottom;
                buttonPanel.Height = 50;
                buttonPanel.Padding = new Padding(5);

                var okButton = new System.Windows.Forms.Button()
                {
                    DialogResult = DialogResult.OK,
                    Text = "OK",
                    Width = 100,
                    Height = 35,
                    Left = 10,
                    Top = 5
                };

                var cancelButton = new System.Windows.Forms.Button()
                {
                    DialogResult = DialogResult.Cancel,
                    Text = "Cancel",
                    Width = 100,
                    Height = 35,
                    Left = 120,
                    Top = 5
                };

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

            bool usingDaxUserDefinedFunction = macroCode.Contains("using DaxUserDefinedFunction;");
            if (usingDaxUserDefinedFunction)
            {
                macroCode = macroCode.Replace("using DaxUserDefinedFunction;", "");
            }

            bool usingReportFunctions = macroCode.Contains("using ReportFunctions;");
            if (usingReportFunctions)
            {
                macroCode = macroCode.Replace("using ReportFunctions;", "");
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



                int hashrFirstMacroCode = previousCode.IndexOf("#r");
                int hashrFirstCustomClass = codeToAppend.IndexOf("#r");

                int hasrLastMacroCode = previousCode.LastIndexOf("#r");
                int endOfHashrMacroCode = previousCode.IndexOf(Environment.NewLine, Math.Max(hasrLastMacroCode, 0));

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
                                + previousCode.Substring(Math.Max(hashrFirstMacroCode, 0));

                            //update the position of the first #r
                            hashrFirstMacroCode = previousCode.IndexOf("#r");
                            hasrLastMacroCode = previousCode.LastIndexOf("#r");
                            endOfHashrMacroCode = previousCode.IndexOf(Environment.NewLine, Math.Max(hasrLastMacroCode, 0));
                        }


                    }



                    //remove #r directives from custom class 
                    codeToAppend = codeToAppend.Replace(codeToAppend.Substring(hashrFirstCustomClass, endOfHashrCustomClass - hashrFirstCustomClass), "");

                    int usingFirstMacroCode = Math.Max(previousCode.Substring(0, Math.Min(200, previousCode.Length)).IndexOf("using"), endOfHashrMacroCode);
                    int usingFirstCustomClass = codeToAppend.Substring(0, Math.Min(200, codeToAppend.Length)).IndexOf("using");

                    if (usingFirstCustomClass != -1)
                    {
                        int usingLastCustomClass = codeToAppend.Substring(0, Math.Min(200, codeToAppend.Length)).LastIndexOf("using");
                        int endOfusingCustomClass = codeToAppend.IndexOf(Environment.NewLine, usingLastCustomClass);

                        string[] usingLines = codeToAppend.Substring(usingFirstCustomClass, endOfusingCustomClass - usingFirstCustomClass).Split('\n');


                        foreach (String usingLine in usingLines)
                        {

                            //if using directive not present
                            if (!previousCode.Contains(usingLine))
                            {
                                //insert in the code right before the first one
                                previousCode = previousCode.Substring(0, Math.Max(usingFirstMacroCode, 0))
                                    + Environment.NewLine + usingLine.Trim() + Environment.NewLine
                                    + previousCode.Substring(Math.Max(usingFirstMacroCode, 0));

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

                // Reorder directives: all #r lines first, then all using lines, then the rest of the code
                string[] allOutputLines = outputCode.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
                List<string> hashrDirectivesFinal = new List<string>();
                List<string> usingDirectivesFinal = new List<string>();
                List<string> remainingCodeLines = new List<string>();
                foreach (string outputLine in allOutputLines)
                {
                    string trimmedLine = outputLine.Trim();
                    if (trimmedLine.StartsWith("#r "))
                        hashrDirectivesFinal.Add(trimmedLine);
                    else if (trimmedLine.StartsWith("using ") && trimmedLine.EndsWith(";") && !trimmedLine.Contains("("))
                        usingDirectivesFinal.Add(trimmedLine);
                    else
                        remainingCodeLines.Add(outputLine);
                }
                // Remove leading blank lines from the code body
                while (remainingCodeLines.Count > 0 && string.IsNullOrWhiteSpace(remainingCodeLines[0]))
                    remainingCodeLines.RemoveAt(0);

                List<string> finalOutputLines = new List<string>();
                finalOutputLines.AddRange(hashrDirectivesFinal);
                finalOutputLines.AddRange(usingDirectivesFinal);
                if (hashrDirectivesFinal.Count > 0 || usingDirectivesFinal.Count > 0)
                    finalOutputLines.Add("");
                finalOutputLines.AddRange(remainingCodeLines);
                outputCode = string.Join(Environment.NewLine, finalOutputLines);

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

            if (usingDaxUserDefinedFunction)
            {


                // Parse the syntax tree
                SyntaxTree daxUserDefinedFunctionClassTree = CSharpSyntaxTree.ParseText(File.ReadAllText(daxUserDefinedFunctionClassFilePath));
                var root = daxUserDefinedFunctionClassTree.GetRoot();

                // Get namespace node
                var namespaceNode = root.DescendantNodes().OfType<NamespaceDeclarationSyntax>().First();

                // Get classes that are direct children of the namespace (not nested classes)
                var daxUserDefinedFunctionClass = namespaceNode.Members
                    .OfType<ClassDeclarationSyntax>();

                // Concatenate class codes
                string daxUserDefinedFunctionClassCode = string.Join(
                    Environment.NewLine,
                    daxUserDefinedFunctionClass.Select(c => c.ToFullString())
                );

                macroCodeClean3 += Environment.NewLine + daxUserDefinedFunctionClassCode;

            }

            // te (the Tabular Editor CLI) compiles the macro with TECLI defined and has no WinForms, so the helper
            // methods that only the interactive branch reaches are wrapped in #if !TECLI. TE3 still compiles them all.
            Func<string, string> GuardInteractiveOnlyMethods = (string code) =>
            {
                var cliOptions = new CSharpParseOptions(kind: SourceCodeKind.Script, preprocessorSymbols: new[] { "TECLI" });
                SyntaxNode cliRoot = CSharpSyntaxTree.ParseText(code, cliOptions).GetRoot();

                List<MethodDeclarationSyntax> helperMethods = cliRoot.DescendantNodes().OfType<ClassDeclarationSyntax>()
                    .Where(c => c.Modifiers.Any(SyntaxKind.StaticKeyword))
                    .SelectMany(c => c.Members.OfType<MethodDeclarationSyntax>())
                    .ToList();

                //names used by the macro itself in CLI mode (the #else branch is inactive, so it is not visited)
                HashSet<string> reached = new HashSet<string>(cliRoot.DescendantNodes().OfType<IdentifierNameSyntax>()
                    .Where(i => !i.Ancestors().OfType<MethodDeclarationSyntax>().Any(m => helperMethods.Contains(m)))
                    .Select(i => i.Identifier.ValueText));

                //follow the calls between helper methods until nothing new is reached
                bool added = true;
                while (added)
                {
                    added = false;
                    foreach (var m in helperMethods.Where(m => reached.Contains(m.Identifier.ValueText)).ToList())
                    {
                        foreach (var i in m.DescendantNodes().OfType<IdentifierNameSyntax>())
                        {
                            if (reached.Add(i.Identifier.ValueText)) added = true;
                        }
                    }
                }

                //FullSpan starts at the beginning of the method's first line (comments included) and ends after its last line break
                var unreached = helperMethods.Where(m => !reached.Contains(m.Identifier.ValueText))
                    .OrderBy(m => m.FullSpan.Start).ToList();

                //contiguous unreached methods share one #if block
                List<Tuple<int, int>> blocks = new List<Tuple<int, int>>();
                foreach (var m in unreached)
                {
                    if (blocks.Count > 0 && blocks[blocks.Count - 1].Item2 == m.FullSpan.Start)
                        blocks[blocks.Count - 1] = Tuple.Create(blocks[blocks.Count - 1].Item1, m.FullSpan.End);
                    else
                        blocks.Add(Tuple.Create(m.FullSpan.Start, m.FullSpan.End));
                }

                foreach (var b in blocks.OrderByDescending(b => b.Item1))
                {
                    string endif = code.Substring(0, b.Item2).EndsWith("\n") ? "#endif" + Environment.NewLine : Environment.NewLine + "#endif" + Environment.NewLine;
                    code = code.Insert(b.Item2, endif).Insert(b.Item1, "#if !TECLI" + Environment.NewLine);
                }
                return code;
            };

            if (macroCodeClean3.Contains("#if TECLI"))
            {
                macroCodeClean3 = GuardInteractiveOnlyMethods(macroCodeClean3);
            }

            //copy the code to the clipboard
            Clipboard.SetText(macroCodeClean3);


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

        public static Measure SelectMeasure(Measure preselect = null, string label = "Select Measure")
        {
            return ScriptHelper.SelectMeasure(preselect: preselect, label: label);
        }

        public static void Output(object value, int lineNumber = -1)
        {
            ScriptHelper.Output(value: value, lineNumber: lineNumber);
        }

        public static string FormatDax(string dax)
        {
            var client = new DaxFormatterClient();

            var req = new DaxFormatterSingleRequest
            {
                Dax = dax
            };

            var resp = client.FormatAsync(req).GetAwaiter().GetResult();
            return resp.Formatted;
        }

    }

    // Extension method to enable ReplaceLineEndings() syntax for .NET Framework 4.8
    // Note: This is only for Visual Studio compilation. Tabular Editor 3 uses .NET 6+ which has the native method.
    public static class StringExtensions
    {
        public static string ReplaceLineEndings(this string value)
        {
            // This implementation will never be called in Tabular Editor 3
            // It's only here to satisfy Visual Studio's compiler
            return value?.Replace("\r\n", Environment.NewLine).Replace("\n", Environment.NewLine).Replace("\r", Environment.NewLine);
        }

        public static string ReplaceLineEndings(this string value, string replacementText)
        {
            // Overload that accepts a replacement string
            return value?.Replace("\r\n", replacementText).Replace("\n", replacementText).Replace("\r", replacementText);
        }
    }
}
