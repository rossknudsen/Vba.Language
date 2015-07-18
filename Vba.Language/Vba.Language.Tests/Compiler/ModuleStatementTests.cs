using System;
using System.Collections.Generic;
using System.Runtime.Remoting.Messaging;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Vba.Grammars;
using Xunit;

namespace Vba.Language.Tests.Compiler
{
    public class ModuleStatementTests
    {
        private readonly List<string> ambiguousIdentifiers = new List<string>()
        {
            "Alias",
            "Base",
            "Binary",
            "ClassInit",
            "ClassTerm",
            "CLngLng",
            "Compare",
            "Database",
            "DefLngLng",
            "Explicit",
            "Lib",
            "LongLong",
            "Module",
            "Object",
            "Property",
            "Text",
            "VB_Invoke_PropertyPutRefVB_MemberFlags",
            "PtrSafe",
        };

        private readonly List<string> trueKeywords = new List<string>()
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
            
        [Theory]
        [InlineData("Option Compare Text", "(directiveElement (optionCompareDirective Option Compare Text))")]
        [InlineData("Option Compare Binary", "(directiveElement (optionCompareDirective Option Compare Binary))")]
        [InlineData("Option Compare Database", "(directiveElement (optionCompareDirective Option Compare Database))")]
        [InlineData("Option Base 0", "(directiveElement (optionBaseDirective Option Base 0))")]
        [InlineData("Option Base 1", "(directiveElement (optionBaseDirective Option Base 1))")]
        [InlineData("Option Explicit", "(directiveElement (optionExplicitDirective Option Explicit))")]
        [InlineData("Option Private Module", "(directiveElement (optionPrivateDirective Option Private Module))")]
        [InlineData("Implements InterfaceName", "(directiveElement (implementsDirective Implements (classTypeName (definedTypeExpression (simpleNameExpression (name (untypedName (identifier InterfaceName))))))))")]
        [InlineData("DefStr A", "(directiveElement (defDirective (defType DefStr) (letterSpec A)))")]
        [InlineData("DefStr A-Q", "(directiveElement (defDirective (defType DefStr) (letterSpec A - Q)))")]
        public void CanParseDirectiveElement(string source, string expectedTree)
        {
            var parser = VbaCompilerHelper.BuildVbaParser(source);

            var result = parser.directiveElement();

            Assert.Null(result.exception);
            ParseTreeHelper.TreesAreEqual(expectedTree, result.ToStringTree(parser));
        }

        [Fact]
        public void CanParseAmbiguousIdentifierInVariableDeclaration()
        {
            const string variableDeclarationTemplate = "Dim {0} As String";
            const string expectedVariableDeclaration =
                            "(variableDeclaration Dim (variableDclList (variableDcl (untypedVariableDcl (identifier {0}) (asClause (asType As (typeSpec (typeExpression (builtInType (reservedTypeIdentifier String))))))))))";

            CanParseAllAmbiguousIdentifiers(variableDeclarationTemplate, expectedVariableDeclaration, p => p.variableDeclaration());
        }

        [Fact]
        public void CanParseAmbiguousIdentifierInWithEventsDeclaration()
        {
            const string withEventsTemplate = "Private WithEvents {0} As Application";
            const string expectedWithEventsTemplate = "(variableDeclaration Private (variableDclList (witheventsVariableDcl WithEvents (identifier {0}) As (classTypeName (definedTypeExpression (simpleNameExpression (name (untypedName (identifier Application)))))))))";

            CanParseAllAmbiguousIdentifiers(withEventsTemplate, expectedWithEventsTemplate, p => p.variableDeclaration());
        }

        [Fact]
        public void CannotParseInvalidIdentifierInVariableDeclaration()
        {
            const string variableDeclarationTemplate = "Dim {0} As String";
            
            CannotParseAnyTrueKeywords(variableDeclarationTemplate, p => p.variableDeclaration());
        }

        [Fact]
        public void CannotParseInvalidIdentifierInWithEventsVariableDeclaration()
        {
            const string variableDeclarationTemplate = "Dim {0} As String";
            
            CannotParseAnyTrueKeywords(variableDeclarationTemplate, p => p.variableDeclaration());
        }

        [Fact]
        public void CanParseIdentifierInSubDefinition()
        {
            const string subDefinitionTemplate = "Sub {0} \r\nEnd Sub\r\n";
            const string expectedOutputSubDefinitionTemplate = "(subroutineDeclaration Sub (subroutineName (identifier {0})) \\r\\n (procedureBody statementBlock) End Sub \\r\\n)";

            CanParseAllAmbiguousIdentifiers(subDefinitionTemplate, expectedOutputSubDefinitionTemplate, p => p.subroutineDeclaration());
        }

        [Fact]
        public void CannotParseInvalidIdentifierInSubDefinition()
        {
            const string subDeclarationTemplate = "Sub {0} \r\nEnd Sub\r\n";

            CannotParseAnyTrueKeywords(subDeclarationTemplate, p => p.subroutineDeclaration());
        }

        [Fact]
        public void CanParseIdentifierInConstItem()
        {
            const string constDeclarationTemplate = "Private Const {0} As String = \"\"";
            const string expectedOutputConstDeclarationTemplate = "(constDeclaration Private Const (constItemList (constItem (identifier {0}) (asClause (asType As (typeSpec (typeExpression (builtInType (reservedTypeIdentifier String)))))) = (constantExpression (expression (literalExpression \"\"))))))";

            CanParseAllAmbiguousIdentifiers(constDeclarationTemplate, expectedOutputConstDeclarationTemplate, p => p.constDeclaration());
        }

        [Fact]
        public void CannotParseInvalidIdentifierInConstItem()
        {
            const string constDeclarationTemplate = "Private Const {0} As String = \"\"";

            CannotParseAnyTrueKeywords(constDeclarationTemplate, p => p.constDeclaration());
        }

        private void CannotParseAnyTrueKeywords(string sourceTemplate, Func<VbaParser, ParserRuleContext> rule)
        {
            foreach (var id in trueKeywords)
            {
                var source = string.Format(sourceTemplate, id);
                var parser = VbaCompilerHelper.BuildVbaParser(source);

                Assert.Throws<ParseCanceledException>(() => rule(parser));
            }
        }

        private void CanParseAllAmbiguousIdentifiers(string sourceTemplate, string expectedOutputTemplate, Func<VbaParser, ParserRuleContext> rule)
        {
            foreach (var id in ambiguousIdentifiers)
            {
                var source = string.Format(sourceTemplate, id);
                var expectedTree = string.Format(expectedOutputTemplate, id);

                CanParseSource(source, expectedTree, rule);
            }
        }

        private static void CanParseSource(string source, string expectedTree, Func<VbaParser, ParserRuleContext> rule)
        {
            var parser = VbaCompilerHelper.BuildVbaParser(source);

            var result = rule(parser);

            Assert.Null(result.exception);
            ParseTreeHelper.TreesAreEqual(expectedTree, result.ToStringTree(parser));
        }
    }
}
