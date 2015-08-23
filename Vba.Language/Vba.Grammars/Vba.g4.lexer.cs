using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Vba.Grammars
{
    partial class VbaLexer
    {
        private static readonly List<string> ambiguousIdentifiers = new List<string>()
        {
            "Access",
            "Alias",
            "Append",
            "Base",
            "Binary",
            "Class_Initialize",
            "Class_Terminate",
            "CLngLng",
            "Compare",
            "Database",
            "DefLngLng",
            "Error",
            "Explicit",
            "Lib",
            "Line",
            "LongLong",
            "Mid",
            "MidB",
            "Module",
            "Object",
            "Output",
            "Property",
            "PtrSafe",
            "Random",
            "Read",
            "Reset",
            "Step",
            "Text",
            "Width"
        };

        private static readonly List<string> trueKeywords = new List<string>()
        {
            "Abs",
            "AddressOf",
            "And",
            "Any",
            "Array",
            "As",
            "Attribute",
            "Boolean",
            "ByRef",
            "Byte",
            "ByVal",
            "Call",
            "Case",
            "CBool",
            "CByte",
            "CCur",
            "CDate",
            "CDbl",
            "CDec",
            "CDecl",
            "CInt",
            "Circle",
            "CLng",
            "CLngPtr",
            "Close",
            "Const",
            "CSng",
            "CStr",
            "Currency",
            "CVar",
            "CVErr",
            "Date",
            "Debug",
            "Decimal",
            "Declare",
            "DefBool",
            "DefByte",
            "DefCur",
            "DefDate",
            "DefDbl",
            "DefDec",
            "DefInt",
            "DefLng",
            "DefLngPtr",
            "DefObj",
            "DefSng",
            "DefStr",
            "DefVar",
            "Dim",
            "Do",
            "DoEvents",
            "Double",
            "Each",
            "Else",
            "ElseIf",
            "Empty",
            "End",
            "EndIf",
            "Enum",
            "Eqv",
            "Erase",
            "Event",
            "Exit",
            "False",
            "Fix",
            "For",
            "Friend",
            "Function",
            "Get",
            "Global",
            "GoSub",
            "GoTo",
            "If",
            "Imp",
            "Implements",
            "In",
            "Input",
            "InputB",
            "Int",
            "Integer",
            "Is",
            "LBound",
            "Len",
            "LenB",
            "Let",
            "Like",
            "LINEINPUT",
            "Lock",
            "Long",
            "LongPtr",
            "Loop",
            "LSet",
            "Me",
            "Mod",
            "New",
            "Next",
            "Not",
            "Nothing",
            "Null",
            "On",
            "Open",
            "Option",
            "Optional",
            "Or",
            "ParamArray",
            "Preserve",
            "Print",
            "Private",
            "PSet",
            "Public",
            "Put",
            "RaiseEvent",
            "ReDim",
            "Rem",
            "Resume",
            "Return",
            "RSet",
            "Scale",
            "Seek",
            "Select",
            "Set",
            "Sgn",
            "Shared",
            "Single",
            "Spc",
            "Static",
            "Stop",
            "String",
            "Sub",
            "Tab",
            "Then",
            "To",
            "True",
            "Type",
            "TypeOf",
            "UBound",
            "Unlock",
            "Until",
            "Variant",
            "VB_Base",
            "VB_Control",
            "VB_Creatable",
            "VB_Customizable",
            "VB_Description",
            "VB_Exposed",
            "VB_Ext_KEY",
            "VB_GlobalNameSpace",
            "VB_HelpID",
            "VB_Invoke_Func",
            "VB_Invoke_Property",
            "VB_Invoke_PropertyPut",
            "VB_Invoke_PropertyPutRef",
            "VB_MemberFlags",
            "VB_Name",
            "VB_PredeclaredId",
            "VB_ProcData",
            "VB_TemplateDerived",
            "VB_UserMemId",
            "VB_VarDescription",
            "VB_VarHelpID",
            "VB_VarMemberFlags",
            "VB_VarProcData",
            "VB_VarUserMemId",
            "Wend",
            "While",
            "With",
            "WithEvents",
            "Write",
            "Xor"
        };

        public static IEnumerable<string> AmbiguousIdentifiers => ambiguousIdentifiers;

        public static IEnumerable<string> TrueKeywords => trueKeywords;

        // Type member names can include most keywords.  The notable exceptions are 'Me' and 'Rem'.
        public static IEnumerable<string> TypeMemberIdentifiers =>
            AllKeywords
            .Where(k => k != "Me" && k != "Rem")
            .ToList();

        public static IEnumerable<string> AllKeywords =>
            ambiguousIdentifiers
            .Concat(trueKeywords)
            .ToList();

        private static IDictionary<string, int> constants;

        public static int ConvertTokenNameToValue(string tokenName)
        {
            if (constants == null)
            {
                // http://stackoverflow.com/a/10261848/2301065

                constants = new Dictionary<string, int>();

                var type = typeof(VbaParser);
                var fieldInfos = type.GetFields(BindingFlags.Public | BindingFlags.Static);
                var constantInfos = fieldInfos.Where(f => f.IsLiteral && !f.IsInitOnly).ToList();
                foreach (var c in constantInfos)
                {
                    var name = c.Name;
                    var value = (int)c.GetRawConstantValue();
                    constants.Add(name, value);
                }
            }

            if (constants.ContainsKey(tokenName))
            {
                return constants[tokenName];
            }
            throw new ArgumentException("Token name does not exist.");
        }
    }
}
