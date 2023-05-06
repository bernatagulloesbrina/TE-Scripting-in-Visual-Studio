using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TabularEditor.TOMWrapper;
using TabularEditor.Scripting;
using System.Reflection.Emit;

namespace GeneralFunctions
{

    //copy from the following line up to ****** and remove the // before the closing bracket
    //after the class declaration add all the #r and using statements necessary for the custom class code to run in Tabular Editor
    //these directives will be combined with the ones from the macro when using the CopyMacro script
    public static class Fx
    {
        //#r "Microsoft.VisualBasic"
        //using Microsoft.VisualBasic;


        //in TE2 (at least up to 2.17.2) any method that accesses or modifies the model needs a reference to the model 
        //the following is an example method where you can build extra logic
        public static Table CreateCalcTable(Model model, string tableName, string tableExpression) 
        { 
            return model.AddCalculatedTable(name:tableName,expression:tableExpression);
        }

        //add other methods always as "public static" followed by the data type they will return or void if they do not return anything.

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

        public static void Output(object value, int lineNumber = -1)
        {
            ScriptHelper.Output(value: value, lineNumber: lineNumber);
        }
    }
}
