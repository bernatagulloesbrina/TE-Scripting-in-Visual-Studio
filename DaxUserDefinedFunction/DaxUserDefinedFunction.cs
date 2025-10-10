using GeneralFunctions; 
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TabularEditor.TOMWrapper;
namespace DaxUserDefinedFunction
{
    

    public class FunctionExtended
    {
        public string Name { get; set; }
        public string Expression { get; set; }
        public string Description { get; set; }
        public string OutputFormatString { get; set; }
        public string OutputNameTemplate { get; set; }
        public string OutputType { get; set; }
        public string OutputDisplayFolder { get; set; }

        public string OutputDestination { get; set; } 
        public Function OriginalFunction { get; set; }
        public List<FunctionParameter> Parameters { get; set; }
        private static List<FunctionParameter> ExtractParametersFromExpression(string expression)
        {
            // Find the first set of parentheses before the "=>"
            int arrowIndex = expression.IndexOf("=>");
            if (arrowIndex == -1)
                return new List<FunctionParameter>();

            int openParenIndex = expression.LastIndexOf('(', arrowIndex);
            int closeParenIndex = expression.IndexOf(')', openParenIndex);
            if (openParenIndex == -1 || closeParenIndex == -1 || closeParenIndex > arrowIndex)
                return new List<FunctionParameter>();

            string paramSection = expression.Substring(openParenIndex + 1, closeParenIndex - openParenIndex - 1);
            var paramList = new List<FunctionParameter>();
            var paramStrings = paramSection.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var param in paramStrings)
            {
                var trimmed = param.Trim();
                var nameParams = trimmed.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
                //var parts = trimmed.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                var fp = new FunctionParameter();
                fp.Name = nameParams.Length > 0 ? nameParams[0] : param;
                fp.Type =
                    (nameParams.Length == 1) ? "ANYVAL" :
                    (nameParams.Length > 1) ?
                        (nameParams[1].IndexOf("anyVal", StringComparison.OrdinalIgnoreCase) >= 0 ? "ANYVAL" :
                         nameParams[1].IndexOf("Scalar", StringComparison.OrdinalIgnoreCase) >= 0 ? "SCALAR" :
                         nameParams[1].IndexOf("Table", StringComparison.OrdinalIgnoreCase) >= 0 ? "TABLE" :
                         nameParams[1].IndexOf("AnyRef", StringComparison.OrdinalIgnoreCase) >= 0 ? "ANYREF" :
                         "ANYVAL")
                    : "ANYVAL";

                fp.Subtype =
                    (fp.Type == "SCALAR" && nameParams.Length > 1) ?
                        (
                            nameParams[1].IndexOf("variant", StringComparison.OrdinalIgnoreCase) >= 0 ? "VARIANT" :
                            nameParams[1].IndexOf("int64", StringComparison.OrdinalIgnoreCase) >= 0 ? "INT64" :
                            nameParams[1].IndexOf("decimal", StringComparison.OrdinalIgnoreCase) >= 0 ? "DECIMAL" :
                            nameParams[1].IndexOf("double", StringComparison.OrdinalIgnoreCase) >= 0 ? "DOUBLE" :
                            nameParams[1].IndexOf("string", StringComparison.OrdinalIgnoreCase) >= 0 ? "STRING" :
                            nameParams[1].IndexOf("datetime", StringComparison.OrdinalIgnoreCase) >= 0 ? "DATETIME" :
                            nameParams[1].IndexOf("boolean", StringComparison.OrdinalIgnoreCase) >= 0 ? "BOOLEAN" :
                            nameParams[1].IndexOf("numeric", StringComparison.OrdinalIgnoreCase) >= 0 ? "NUMERIC" :
                            null
                        )
                    : null;

                // ParameterMode: check for VAL or EXPR (any casing) in the parameter string
                string paramMode = null;
                if (trimmed.IndexOf("VAL", StringComparison.OrdinalIgnoreCase) >= 0)
                    paramMode = "VAL";
                else if (trimmed.IndexOf("EXPR", StringComparison.OrdinalIgnoreCase) >= 0)
                    paramMode = "EXPR";
                else
                    paramMode = "VAL";
                fp.ParameterMode = paramMode;

                paramList.Add(fp);
            }

            return paramList;
        }

        

        public static FunctionExtended CreateFunctionExtended(Function function)
        {

            FunctionExtended emptyFunction = null as FunctionExtended;
            List<FunctionParameter> Parameters =  ExtractParametersFromExpression (function.Expression);

            string nameTemplateDefault = "";
            string formatStringDefault = "";
            string displayFolderDefault = "";
            string functionNameShort = function.Name;
            string destinationDefault = ""; 

            if(function.Name.IndexOf(".") > 0)
            {
                functionNameShort = function.Name.Substring(function.Name.LastIndexOf(".") + 1);
            }

            if (Parameters.Count == 0) {
                nameTemplateDefault = function.Name;
                formatStringDefault = "";
                displayFolderDefault = "";
                destinationDefault = "";
            }
            else
            {
                nameTemplateDefault = string.Join(" ", Parameters.Select(p => p.Name + "Name"));
                if(function.Name.Contains("Pct"))
                {
                    formatStringDefault = "+0.0%;-0.0%;-";
                }
                else
                {
                    formatStringDefault = Parameters[0].Name + "FormatStringRoot";
                }


                    
                displayFolderDefault = 
                    String.Format(
                        @"{0}DisplayFolder/{1}Name {2}", 
                        Parameters[0].Name, 
                        Parameters[0].Name,
                        functionNameShort);
                
                if (Parameters[0].Name.ToUpper().Contains("TABLE"))
                {
                    destinationDefault = Parameters[0].Name;
                }
                else if (Parameters[0].Name.ToUpper().Contains("MEASURE") || Parameters[0].Name.ToUpper().Contains("COLUMN"))
                {
                    destinationDefault = Parameters[0].Name + "Table";
                }
                else
                {
                    destinationDefault = "Custom";
                }
                

            };
            
            


            string myOutputType = function.GetAnnotation("outputType"); 
            string myNameTemplate = function.GetAnnotation("nameTemplate");
            string myFormatString = function.GetAnnotation("formatString");
            string myDisplayFolder = function.GetAnnotation("displayFolder");
            string myOutputDestination = function.GetAnnotation("outputDestination");

            if (string.IsNullOrEmpty(myOutputType))
            {
                IList<string> selectionTypeOptions = new List<string> { "Table", "Column", "Measure", "None" };
                myOutputType = 
                    Fx.ChooseString(
                        OptionList: selectionTypeOptions, 
                        label: "Choose output type for function" + function.Name, 
                        customWidth: 600);
                if (string.IsNullOrEmpty(myOutputType)) return emptyFunction;
                function.SetAnnotation("outputType", myOutputType);
            }

            if (string.IsNullOrEmpty(myNameTemplate))
            {
                myNameTemplate = Fx.GetNameFromUser(Prompt:"Enter output name template for function " + function.Name,  "Name Template", nameTemplateDefault);
                if (string.IsNullOrEmpty(myNameTemplate)) return emptyFunction;
                function.SetAnnotation("nameTemplate", myNameTemplate);
            }
            if(string.IsNullOrEmpty(myFormatString))
            {
                myFormatString = Fx.GetNameFromUser(Prompt: "Enter output format string for function " + function.Name, "Format String", formatStringDefault);
                if (string.IsNullOrEmpty(myFormatString)) return emptyFunction;
                function.SetAnnotation("formatString", myFormatString);

            }
            if(string.IsNullOrEmpty(myDisplayFolder))
            {
                myDisplayFolder = 
                    Fx.GetNameFromUser(
                        Prompt: "Enter output display folder for function " + function.Name, 
                        Title:"Display Folder", 
                        DefaultResponse: displayFolderDefault);

                if (string.IsNullOrEmpty(myDisplayFolder)) return emptyFunction;
                function.SetAnnotation("displayFolder", myDisplayFolder);
            }

            if (string.IsNullOrEmpty(myOutputDestination))
            {
                if(myOutputType ==  "Table")
                {
                    myOutputDestination = "Model";
                }
                else if(myOutputType == "Column" ||  myOutputType == "Measure")
                {
                    myOutputDestination = 
                        Fx.GetNameFromUser(
                            Prompt: "Enter Destination template for " + function.Name, 
                            Title:"Destination", 
                            DefaultResponse: destinationDefault);

                    if(string.IsNullOrEmpty(myOutputDestination)) return emptyFunction;
                    function.SetAnnotation("outputDestination", destinationDefault);
                }
            }


            var functionExtended = new FunctionExtended
            {
                Name = function.Name,
                Expression = function.Expression,
                Description = function.Description,
                Parameters = Parameters,
                OutputFormatString = myFormatString,
                OutputNameTemplate = myNameTemplate,
                OutputType = myOutputType,
                OutputDisplayFolder = myDisplayFolder,
                OutputDestination = myOutputDestination,
                OriginalFunction = function

            };

            return functionExtended;
        }

        
    }

    public class FunctionParameter
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string Subtype { get; set; }
        public string ParameterMode { get; set; }
    }

    
}
