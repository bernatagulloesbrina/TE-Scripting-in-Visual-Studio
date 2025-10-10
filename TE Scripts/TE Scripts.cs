using System;
using System.Collections.Generic;
using System.Linq;
using TabularEditor.TOMWrapper;
using TabularEditor.Scripting;
using System.Windows.Forms;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.CSharp;
using DaxUserDefinedFunction;
using System.IO;
using GeneralFunctions; //Uncomment if you use the custom class, add reference to the project too.
using ReportFunctions;
using Report.DTO;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;
using TabularEditor;
using static Report.DTO.VisualDto;



// '2023-05-06 / B.Agullo / 
//coding environment for Tabular Editor C# Scripts
// see https://www.esbrina-ba.com/c-scripting-nirvana-effortlessly-use-visual-studio-as-your-coding-environment/ for reference on how to use it.

namespace TE_Scripting
{
    public class TE_Scripts
    {

        void adjustLineOverColumns()
        {
            //using GeneralFunctions;
            //using Report.DTO;
            //using ReportFunctions;
            //using System.IO;
            //using Newtonsoft.Json.Linq;




            //check if there's a matching visual in the report
            ReportExtended report = Rx.InitReport();
            if (report == null) return;

            List<string> visualTypes = new List<string>() { "lineClusteredColumnComboChart" };
            VisualExtended selectedVisual = Rx.SelectVisual(report, visualTypes);
            if (selectedVisual == null) return;

            var queryState = selectedVisual.Content?.Visual?.Query?.QueryState;

            var categories = queryState.Category.Projections;
            var columns = queryState.Y.Projections;
            var lines = queryState.Y2.Projections;


            if (categories.Count() == 0 || columns.Count() == 0 || lines.Count() == 0)
            {
                Error("Chart not completely configured. Please configure at least a field for x-axis, one for the columns and one for the line");
                return;
            }

            Table formattingTable = Fx.CreateCalcTable(Model,"Formatting");
            if(formattingTable == null) return;

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
            if(globalPaddingMeasure == null) return;

            Measure lineChartHeight = Fx.CreateMeasure(
                table: formattingTable,
                measureName: "LineChartHeight",
                measureExpression: "0.4", 
                out lineChartHeightCreated,
                description: "Percent of the chart height for the line chart",
                annotationLabel: "@AgulloBernat",
                annotationValue: "LineChartHeight",
                isHidden: true);
            if(lineChartHeight == null) return;

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

            if (globalPaddingCreated || lineChartHeightCreated || secondaryMaxFunctionCreated || secondaryMinFunctionCreated || primaryMaxFunctionCreated)
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


            if(categories.Count() > 1)
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
                    lines.Select(l=> l.NativeQueryRef)) + "]"
                );
            
            string secondaryMaxVisualCalcExpression = 
                String.Format(
                    @"{3}(
                        {0},
                        {1},
                        {2}
                    )",
                    secondaryMaxLineMaxExpression,
                    xAxisColumn,
                    paddingScalar,
                    secondaryMaxFunction.Name
                );

            var secondaryMaxVisualCalcProjection = new VisualDto.Projection
            {
                Field = new VisualDto.Field
                {
                    NativeVisualCalculation = new VisualDto.NativeVisualCalculation
                    {
                        Language = "dax",
                        Expression= secondaryMaxVisualCalcExpression, 
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
                    @"{5}(
                        {0},
                        {1},
                        {2},
                        {3},
                        {4}
                    )",
                    lineMinExpression,
                    xAxisColumn,
                    paddingScalar,
                    secondaryAxisMaxValue,
                    lineChartWeight,
                    secondaryMinFunction.Name
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
                    @"{4}(
                        {0},
                        {1},
                        {2},
                        {3}
                    )",
                    primaryMaxColumnMaxExpression,
                    xAxisColumn,
                    paddingScalar,
                    lineChartWeight,
                    primaryMaxFunction.Name
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

        }

        void createTimeIntelFunctions()
        {
            //using GeneralFunctions;

            // 2025-09-29/B.Agullo
            // Creates Time Intelligence functions (CY, PY, YOY, YOYPCT) in the model if they do not exist.
            // It also creates a hidden calculated column and measure in the date table to handle cases where the fact table has no data for some dates.
            // The script assumes there is a date table and a fact table in the model.
            // The script will prompt the user to select the main date column in the fact table if there are multiple date columns.
            
            if(Model.Database.CompatibilityLevel < 1702)
            {
                if(Fx.IsAnswerYes("The model compatibility level is below 1702. Time Intelligence functions are only supported in 1702 or higher. Do to change the compatibility level to 1702?"))
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

            if(factTableDateColumns.Count() == 0)
            {
                Error("No Date columns found in fact table " + factTable.Name);
                return;
            }

            if(factTableDateColumns.Count() == 1)
            {
                factTableDateColumn = factTableDateColumns.First();
            }
            else
            {
                factTableDateColumn = factTableDateColumns.First(
                    c=> Model.Relationships.Any(
                        r => ((r.FromColumn == dateColumn && r.ToColumn == c)
                            || (r.ToColumn == dateColumn && r.FromColumn == c)
                                && r.IsActive)));

                factTableDateColumn = SelectColumn(factTableDateColumns, factTableDateColumn, "Select main date column from the fact table"); 
            }
            if(factTableDateColumn == null) return;

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

        void SingleMeasureParameterFunction()
        {

            //2025-09-26/B.Agullo/ fixed bug that would not store annotations if initialized during runtime
            //2025-09-16/B.Agullo/
            //Creates measures based on DAX UDFs 
            //Check the blog post for futher information: https://www.esbrina-ba.com/automatically-create-measures-with-dax-user-defined-functions/

            //using GeneralFunctions;
            //using DaxUserDefinedFunction;
            //using System.Text.RegularExpressions;


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
                if (param.Name.ToUpper().Contains("MEASURE"))
                {
                    selectionType = "Measure";
                }
                else if (param.Name.ToUpper().Contains("COLUMN"))
                {
                    selectionType = "Column";
                }
                else if (param.Name.ToUpper().Contains("TABLE"))
                {
                    selectionType = "Table";
                }


                (IList<string> Values,string Type) selectedObjectsForParam = Fx.SelectAnyObjects(
                    Model,
                    selectionType: selectionType,
                    prompt1: String.Format(@"Select object type for {0} parameter", param.Name),
                    prompt2: String.Format(@"Select item for {0} parameter", param.Name),
                    placeholderValue: param.Name

                );

                if (selectedObjectsForParam.Type == null || selectedObjectsForParam.Values.Count == 0)
                {
                    Info(String.Format("No objects selected for parameter '{0}'. Operation cancelled.", param.Name));
                    return;
                }

                parameterObjectsMap[param.Name] = selectedObjectsForParam;
            }

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

                // When iterating the parameters of this specific function, use the mapping created for distinct parameters.
                foreach (var param in func.Parameters)
                {
                    currentList = new List<string>(); //reset current list
                    currentListNames = new List<string>();
                    currentFormatStrings = new List<string>();
                    currentDestinations = new List<string>();
                    currentDisplayFolders = new List<string>();

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

                        foreach (var o in paramObject.Values)
                        {
                            //extract original name and format string if the parameter is a measure
                            string paramName = o;
                            string paramFormatStringFull = "";
                            string paramFormatStringRoot = "";
                            string paramDisplayFolder = "";
                            string paramTable = "";

                            //prepare placeholder
                            string paramNamePlaceholder = param.Name + "Name";
                            string paramFormatStringRootPlaceholder = param.Name + "FormatStringRoot";
                            string paramFormatStringFullPlaceholder = param.Name + "FormatStringFull";
                            string paramDisplayFolderPlaceholder = param.Name + "DisplayFolder";
                            string paramTablePlaceholder = "";


                            if (paramObject.Type == "Measure")
                            {
                                Measure m = Model.AllMeasures.FirstOrDefault(m => m.DaxObjectFullName == o);
                                paramName = m.Name;
                                paramFormatStringFull = m.FormatString;
                                paramDisplayFolder = m.DisplayFolder;
                                paramTable = m.Table.DaxObjectFullName;

                                paramTablePlaceholder = param.Name + "Table";

                            }
                            else if (paramObject.Type == "Column")
                            {
                                Column c = Model.AllColumns.FirstOrDefault(c => c.DaxObjectFullName == o);
                                paramName = c.Name;
                                paramFormatStringFull = c.FormatString;
                                paramDisplayFolder = c.DisplayFolder;
                                paramTable = c.Table.DaxObjectFullName;

                                paramTablePlaceholder = param.Name + "Table";
                            }
                            else if (paramObject.Type == "Table")
                            {
                                Table t = Model.Tables.FirstOrDefault(t => t.DaxObjectFullName == o);
                                paramName = t.Name;
                                paramFormatStringFull = "";
                                paramDisplayFolder = "";
                                paramTable = t.DaxObjectFullName;

                                paramTablePlaceholder = param.Name;
                            }

                            if (paramFormatStringFull.Contains(";"))
                            {
                                //keep the first part of the format string, strip it of any + sign
                                paramFormatStringRoot = paramFormatStringFull.Split(';')[0].Replace("+","");
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


                        }

                        
                    }

                    delimiter = ", ";
                    previousList = currentList;
                    previousListNames = currentListNames;
                    previousDestinations = currentDestinations;
                    previousDisplayFolders = currentDisplayFolders;
                    previousFormatStrings = currentFormatStrings;
                }



                IList<Table> currentDestinationTables = new List<Table>();


                if(func.OutputType == "Measure" || func.OutputType == "Column")
                {
                    for (int i = 0; i < currentDestinations.Count; i++)
                    {
                        //transform to actual tables, initialize if necessary
                        Table destinationTable = Model.Tables.Where(
                            t => t.DaxObjectFullName == currentDestinations[i])
                            .FirstOrDefault();

                        if (destinationTable == null)
                        {
                            destinationTable = SelectTable(label: "Select destinatoin table for " + func.OutputType + " " + currentListNames[i]);
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
                    }

                }
                else
                {
                    Info("Not implemented yet for output types other than Measure.");
                }
            }

            

        }


        void copyHeaderFormatting()
        {

            //using GeneralFunctions;
            //using Report.DTO;
            //using System.IO;
            //using Newtonsoft.Json.Linq;


            // Step 1: Initialize report
            ReportExtended report = Rx.InitReport();
            if (report == null) return;

            VisualExtended selectedVisual = Rx.SelectTableVisual(report);
            if(selectedVisual == null) return;

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
        }
        
        
        void testReportClass()
        {
            //using GeneralFunctions;
            //using Report.DTO;
            //using System.IO;
            //using Newtonsoft.Json.Linq;

            ReportExtended report = Rx.InitReport();
            if (report == null)
            {
                Info("Operation cancelled or failed to load report.");
                return;
            }

            VisualExtended visual = Rx.SelectVisual(report);
            if (visual == null)
            {
                Info("No visual selected.");
                return;
            }

            Rx.SaveVisual(visual);
            Output("Visual saved to visual.json.");
        }
        void createTextMeasures()
        {
            //using GeneralFunctions;


            //2025-07-28/B.Agullo
            //This script creates text measures based on the selected measures in the model.
            //It prompts the user for a prefix and suffix to be added to the text measures.
            //It also allows the user to specify a suffix for the names of the new text measures.

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
                        suffix.Replace("###",m.Name));
                Measure newMeasure = m.Table.AddMeasure(newMeasureName, newMeasureExpression,newMeasureDisplayFolder);
                newMeasure.FormatDax();
            }
        }
        
        void removeEmptyFolders()
        {
            //using System.IO;
            //using Newtonsoft.Json.Linq;
            //using Report.DTO;
            //using GeneralFunctions;

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
        }

        void AddBilingualLayerToReport()
        {
            //using GeneralFunctions;
            //using Report.DTO;
            //using System.IO;
            //using Newtonsoft.Json.Linq;
            //2025-06-23/B.Agullo
            //this script adds a bilingual layer to the report, allowing the user to select the language of the report.
            //this will only prepare the report for an extraction of the definition as descrived in https://www.esbrina-ba.com/transforming-a-regular-report-into-a-bilingual-one-part-2-extracting-display-names-of-measures-and-field-prameters/
            

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

            Output(String.Format("Bilingual visual groups created for {0} visuals.",totalCount));
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

            /*uncomment in TE3 to avoid wating cursor infront of dialogs*/
            bool waitCursor = Application.UseWaitCursor;
            Application.UseWaitCursor = false;



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

            //comment this line in TE2
            Application.UseWaitCursor = waitCursor;

        }

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
            string reportFunctionsClassFilePath = String.Format(@"{0}\ReportFunctions\ReportFunctions.cs", baseFolderPath);
            string daxUserDefinedFunctionClassFilePath = String.Format(@"{0}\DaxUserDefinedFunction\DaxUserDefinedFunction.cs", baseFolderPath);

            TE_Scripting.TE_Scripts.CopyMacroFromVSFile(
                macroFilePath, generalFunctionsClassFilePath, reportClassFilePath,reportFunctionsClassFilePath, daxUserDefinedFunctionClassFilePath
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
                int endOfHashrMacroCode = previousCode.IndexOf(Environment.NewLine, Math.Max(hasrLastMacroCode,0));

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
                            hashrFirstMacroCode = previousCode.IndexOf("#r");
                            hasrLastMacroCode = previousCode.LastIndexOf("#r");
                            endOfHashrMacroCode = previousCode.IndexOf(Environment.NewLine, Math.Max(hasrLastMacroCode, 0));
                        }


                    }



                    //remove #r directives from custom class 
                    codeToAppend = codeToAppend.Replace(codeToAppend.Substring(hashrFirstCustomClass, endOfHashrCustomClass - hashrFirstCustomClass), "");

                    int usingFirstMacroCode = Math.Max(previousCode.IndexOf("using"), endOfHashrMacroCode);
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
        
    }
}
