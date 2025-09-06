using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DaxUserDefinedFunction
{
    

    public class FunctionExtended
    {
        public string Name { get; set; }
        public string Expression { get; set; }
        public string Description { get; set; }
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
        public static FunctionExtended CreateFunctionExtended(string name, string expression, string description)
        {
            var function = new FunctionExtended
            {
                Name = name,
                Expression = expression,
                Description = description,
                Parameters = ExtractParametersFromExpression(expression)
            };
            return function;
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
