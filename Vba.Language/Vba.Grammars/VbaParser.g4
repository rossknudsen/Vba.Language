parser grammar VbaParser;

//import Common;

options {
    tokenVocab=VbaLexer;
}

// Specification: [MS-VBAL] VBA Language Specification

// 3.3.5.2 Reserved Identifiers and IDENTIFIER
identifier      
    :   Access
    |   Alias
    |   Append
    |   Base
    |   Binary
    |   Class_Initialize
    |   Class_Terminate
    |   CLngLng
    |   Compare
    |   Database
    |   DefLngLng
    |   Error
    |   Explicit
    |   Lib
    |   Line
    |   LongLong
    |   Mid
    |   MidB
    |   Module
    |   Object
    |   Output
    |   Property
    |   Random
    |   Read
    |   Reset
    |   Step
    |   Text
    |   PtrSafe
    |   Width
    |   ID
    ;

reservedIdentifier
    :   statementKeyword
    |   markerKeyword
    |   operatorIdentifier
    |   specialForm
    |   reservedName
    |   literalIdentifier
    |   Rem
    |   reservedForImplementationUse
    |   futureReserved
    ;
statementKeyword
    :    Call | Case | Close | Const
    |    Declare | defType | Dim |    Do
    |    Else | ElseIf | End | EndIf | Enum | Erase | Event | Exit
    |    For | Friend | Function
    |    Get | Global | GoSub | GoTo
    |    If | Implements | Input
    |    Let | Lock | Loop | LSet
    |    Next
    |    On | Open | Option
    |    Print | Private | Public | Put
    |    RaiseEvent | ReDim | Resume | Return | RSet
    |    Seek | Select | Set | Static | Stop | Sub
    |    Type
    |    Unlock
    |    Wend | While | With | Write
    ;
markerKeyword
    :    Any | As
    |    ByRef | ByVal
    |    Case
    |    Each | Else
    |    In
    |    New
    |    Shared
    |    Until
    |    WithEvents | Write
    |    Optional
    |    ParamArray
    |    Preserve
    |    Spc
    |    Tab
    |    Then
    |    To
    ;
operatorIdentifier
    :    AddressOf
    |    And
    |    Eqv
    |    Imp
    |    Is
    |    Like
    |    New
    |    Mod
    |    Not
    |    Or
    |    TypeOf
    |    Xor
    ;
reservedName
    :    Abs
    |    CBool | CByte | CCur | CDate | CDbl | CDec | CInt | CLng | CLngLng | CLngPtr | CSng | CStr | CVar | CVErr
    |    Date | Debug | DoEvents
    |    Fix
    |    Int
    |    Len | LenB
    //|    Me   The spec says this should be included but UDT Members do not permit this.
    |    PSet
    |    Scale | Sgn | String
    ;
specialForm
    :    Array
    |    Circle
    |    Input
    |    InputB
    |    LBound
    |    Scale
    |    UBound
    ;
reservedTypeIdentifier
    :   Boolean |   Byte
    |   Currency
    |   Date |   Double
    |   Integer
    |   Long |   LongLong |   LongPtr
    |   Single |   String
    |   Variant
    ;
literalIdentifier
    :   booleanLiteralIdentifier
    |   Nothing
    |   Empty
    |   Null
    ;
booleanLiteralIdentifier
    :   True
    |   False
    ;
reservedForImplementationUse
    :    Attribute
    |    LINEINPUT
    |    VB_Base
    |    VB_Control
    |    VB_Creatable
    |    VB_Customizable
    |    VB_Description
    |    VB_Exposed
    |    VB_Ext_KEY
    |    VB_GlobalNameSpace
    |    VB_HelpID
    |    VB_Invoke_Func
    |    VB_Invoke_Property
    |    VB_Invoke_PropertyPut
    |    VB_Invoke_PropertyPutRef
    |    VB_MemberFlags
    |    VB_Name
    |    VB_PredeclaredId
    |    VB_ProcData
    |    VB_TemplateDerived
    |    VB_UserMemId
    |    VB_VarDescription
    |    VB_VarHelpID
    |    VB_VarMemberFlags
    |    VB_VarProcData
    |    VB_VarUserMemId
    ;
futureReserved              :   CDecl | Decimal | DefDec;

// 3.3.5.3 Special Identifier Forms
builtInType
    :   reservedTypeIdentifier
    |   '[' reservedTypeIdentifier ']'
    |   Object
    |   '[' Object ']'
    ;
typedName                   :   identifier typeSuffix;
typeSuffix                  
    :   '%'  // Integer 16 bit signed (default)
    |   '&'  // Long 32 bit signed
    |   '^'  // LongLong 64bit signed
    |   '!'  // Single
    |   '#'  // Double (default)
    |   '@'  // Currency
    |   '$';

// 5.1 Module Body Structure

// We use a single definition for a class and standard module.  Distinction will be made
// during semantic analysis.
module
    :   declarationSection moduleCodeSection
    ;

unrestrictedName            :   name | reservedIdentifier;
name                        :   untypedName | typedName;
untypedName                 :   identifier | ForeignName;

// 5.2 Module Declaration Section Structure
declarationSection
    :   (directiveElement EOS)*;

directiveElement
    :   optionCompareDirective
    |   optionBaseDirective
    |   optionExplicitDirective
    |   optionPrivateDirective 
    |   RemStatement
    |   implementsDirective
    |   defDirective
    |   variableDeclaration
    |   enumDeclaration 
    |   constDeclaration
    |   typeDeclaration 
    |   extProcDeclaration
    ;

// 5.2.1.1-4 Option Directives
optionCompareDirective      :   Option Compare (Binary | Text | Database);
optionBaseDirective         :   Option Base IntegerLiteral;
optionExplicitDirective     :   Option Explicit;
optionPrivateDirective      :   Option Private Module;

// 5.2.2 Implicit Definition Directives
// TODO
defDirective        :   defType letterSpec (',' letterSpec)*;
letterSpec          :   LETTER ('-' LETTER)?;
defType
    :   DefBool
    |   DefByte
    |   DefCur
    |   DefDate
    |   DefDbl
    |   DefInt
    |   DefLng
    |   DefLngLng
    |   DefLngPtr
    |   DefObj
    |   DefSng
    |   DefStr
    |   DefVar
    ;

// 5.2.3.1 Module Variable Declaration Lists
variableDeclaration :   (Global | Public | Private | Dim) Shared? variableDclList;
variableDclList     :   (variableDcl | witheventsVariableDcl) (',' (variableDcl | witheventsVariableDcl))*; 
variableDeclarationList     :   variableDcl (',' variableDcl)*;

// 5.2.3.1.1 Variable Declarations
variableDcl         :   typedVariableDcl | untypedVariableDcl;
typedVariableDcl    :   typedName arrayDim?;
untypedVariableDcl  :   identifier (arrayClause | asClause)?;
arrayClause         :   arrayDim asClause?;
asClause            :   asAutoObject | asType;

// 5.2.3.1.2 WithEvents Variable Declarations
witheventsVariableDcl : WithEvents identifier As classTypeName;
classTypeName       :   definedTypeExpression;

// 5.2.3.1.3 Array Dimensions and Bounds
arrayDim            :   '(' boundsList ')';
boundsList          :   dimSpec (',' dimSpec)*;
dimSpec             :   constantExpression (To constantExpression)?;

// 5.2.3.1.4 Variable Type Declarations
asAutoObject        :   As New classTypeName;
asType              :   As typeSpec;
typeSpec            :   fixedLengthStringSpec | typeExpression;
fixedLengthStringSpec : String '*' stringLength;
stringLength        :   constantName | IntegerLiteral;
constantName        :   simpleNameExpression;

// 5.2.3.2 Const Declarations
constDeclaration    :   (Global | Public | Private)? Const constItemList;
constItemList       :   constItem (',' constItem)*;
constItem           :   identifier asClause? '=' constantExpression;

// 5.2.3.3 User Defined Type Declarations
typeDeclaration         :   (Global | Public | Private)? udtDeclaration;
udtDeclaration          :   Type untypedName EOS
                            udtMemberList EOS
                            End Type;
udtMemberList           :   udtElement (EOS udtElement)*;
udtElement              :   udtMember | RemStatement;
udtMember               :   reservedNameMemberDcl | untypedNameMemberDcl;
untypedNameMemberDcl    :   identifier optionalArrayClause;
reservedNameMemberDcl   :   reservedMemberName asClause;
optionalArrayClause     :   arrayDim? asClause;
reservedMemberName
    :   statementKeyword
    |   markerKeyword
    |   operatorIdentifier
    |   specialForm
    |   reservedName
    |   literalIdentifier
    |   reservedForImplementationUse
    |   futureReserved
    |   reservedTypeIdentifier   // This is not in the specification but appears to be required for UDT member names.
    ;

// 5.2.3.4 Enum Declarations
enumDeclaration         :   (Global | Public | Private)? Enum untypedName EOS
                            memberList EOS
                            End Enum;
memberList              :   enumElement (EOS enumElement)*;  
enumElement             :   enumMember | RemStatement;
enumMember              :   untypedName ('=' constantExpression)?;

// 5.2.3.5 External Procedure Declaration
extProcDeclaration      :   (Public | Private)? externalProcDcl;
externalProcDcl         :   Declare PtrSafe? (externalSub | externalFunction);
externalSub             :   Sub subroutineName libInfo procedureParameters?;
externalFunction        :   Function functionName libInfo procedureParameters? functionType?;
libInfo                 :   Lib StringLiteral (Alias StringLiteral)?;

// 5.2.4.2 Implements Directive
implementsDirective     :   Implements classTypeName;

// 5.2.4.3 Event Declaration
eventDeclaration        :   Public? Event identifier eventParameterList?;
eventParameterList      :   '(' positionalParameters? ')';

// 5.3 Module Code Section Structure
// TODO implement LINE-START / LINE-END
moduleCodeSection       :   moduleCodeElement*;
moduleCodeElement
    :   subroutineDeclaration
    |   functionDeclaration
    |   propGetDeclaration
    |   propLhsDeclaration
    |   RemStatement
    |   implementsDirective
    ;

// 5.3.1 Procedure Declarations
subroutineDeclaration
    :   procedureScope? Static? Sub subroutineName procedureParameters? Static? EOS
        procedureBody
        endLabel? End Sub EOS;

functionDeclaration
    :   procedureScope? Static? Function functionName procedureParameters? functionType? Static? EOS
        procedureBody
        endLabel? End Function EOS;

propGetDeclaration
    :   procedureScope? Static? Property Get functionName procedureParameters? functionType? Static? EOS
        procedureBody
        endLabel? End Property EOS;

propLhsDeclaration
    :   procedureScope? Static? Property (Set | Let) subroutineName procedureParameters? functionType? Static? EOS
        procedureBody
        endLabel? End Property EOS;

endLabel            :   statementLabelDefinition;
//procedureTail       :   TODO

// 5.3.1.1 Procedure Scope
procedureScope
    :   Global 
    |   Public 
    |   Private
    |   Friend
    ;

// 5.3.1.3 Procedure Names
subroutineName  :   identifier | prefixedName;
functionName    :   typedName | identifier | prefixedName;
prefixedName    :   eventHandlerName | implementedName | lifecycleHandlerName;

// 5.3.1.4 Function Type Declarations
functionType    :   As typeExpression arrayDesignator?;
arrayDesignator :   '(' ')';

// 5.3.1.5 Parameter Lists
procedureParameters     :   '(' parameterList? ')';
propertyParameters      :   '(' (parameterList ',')? valueParam ')';

parameterList           
    :   (positionalParameters ',' optionalParameters)
    |   (positionalParameters (',' paramArray)?)
    |   optionalParameters
    |   paramArray
    ;
positionalParameters    :   positionalParam (',' positionalParam)*;
optionalParameters      :   optionalParam (',' optionalParam)*;
valueParam              :   positionalParam;
positionalParam         :   parameterMechanism? paramDcl;
optionalParam           :   optionalPrefix paramDcl defaultValue?;
paramArray              :   ParamArray identifier '(' ')' (As (Variant | '[' Variant ']'))?;
paramDcl                :   untypedNameParamDcl | typedNameDcl;
untypedNameParamDcl     :   identifier parameterType?;
typedNameDcl            :   typedName arrayDesignator?;
optionalPrefix          :   Optional parameterMechanism? | parameterMechanism? Optional;
parameterMechanism      :   ByRef | ByVal;
parameterType           :   arrayDesignator? As (typeExpression | Any);
defaultValue            :   '=' constantExpression;

// 5.3.1.8 Event Handler Declarations
eventHandlerName        :   identifier;

// 5.3.1.9 Implemented Name Declarations
implementedName         :   identifier;

// 5.3.1.10 Lifecycle Handler Declarations
lifecycleHandlerName    :   Class_Initialize | Class_Terminate;

// 5.4 Procedure Bodies and Statements
procedureBody           :   statementBlock;

// 5.4.1 Statement Blocks
statementBlock          :   (blockStatement EOS)*;
blockStatement          
    :   statementLabelDefinition 
    |   statement
    |   RemStatement 
    ;
statement               
    :   controlStatement
    |   dataManipulationStatement
    |   errorHandlingStatement
    |   fileStatement
    ;

// 5.4.1.1 Statement Labels
statementLabelDefinition    :   identifier ':' | lineNumberLabel ':'?;
statementLabel              :   identifier | lineNumberLabel;
statementLabelList          :   statementLabel (',' statementLabel);
lineNumberLabel             :   IntegerLiteral;  // actually can only be a decimal literal.

// 5.4.1.2 Rem Statement
// converted to lexer rule.

// 5.4.2 Control Statements
controlStatement            
    :   ifStatement 
    |   controlStatementExceptMultilineIf
    ;
controlStatementExceptMultilineIf
    :   callStatement
    |   whileStatement
    |   forStatement
    |   exitForStatement
    |   doStatement
    |   exitDoStatement
    |   singleLineIfStatement
    |   selectCaseStatement
    |   stopStatement
    |   gotoStatement
    |   onGotoStatement
    |   gosubStatement
    |   returnStatement
    |   onGosubStatement
    |   forEachStatement
    |   exitSubStatement
    |   exitFunctionStatement
    |   exitPropertyStatement
    |   raiseEventStatement
    |   withStatement
    ;

// 5.4.2.1 Call Statement
callStatement               
    :   Call? (simpleNameExpression | memberAccessExpression | indexExpression | withExpression);

// 5.4.2.2 While Statement
whileStatement              :   While booleanExpression EOS
                                statementBlock Wend;

// 5.4.2.3 For Statement
forStatement                :   simpleForStatement | explicitForStatement;
simpleForStatement          :   forClause EOS statementBlock Next;
explicitForStatement        :   forClause EOS
                                statementBlock
                                (Next | (nestedForStatement ',')) boundVariableExpression;
nestedForStatement          :   explicitForStatement | explicitForEachStatement;
forClause                   :   For boundVariableExpression '=' expression To expression stepClause?;
stepClause                  :   Step expression;

// 5.4.2.4 For Each Statement
forEachStatement            :   simpleForEachStatement | explicitForEachStatement;
simpleForEachStatement      :   forEachClause EOS statementBlock Next;
explicitForEachStatement    :   forEachClause EOS 
                                statementBlock
                                (Next | (nestedForStatement ',')) boundVariableExpression;
forEachClause               :   For Each boundVariableExpression In expression;

// 5.4.2.5 Exit For Statement
exitForStatement            :   Exit For;

// 5.4.2.6 Do Statement
doStatement                 :   Do conditionClause? EOS 
                                statementBlock
                                Loop conditionClause?;
conditionClause             :   (While | Until) booleanExpression;

// 5.4.2.7 Exit Do Statement
exitDoStatement             :   Exit Do;

// 5.4.2.8 If Statement
ifStatement                 :   If booleanExpression Then EOS
                                statementBlock
                                elseIfBlock*
                                elseBlock?
                                End If;
elseIfBlock                 :   ElseIf booleanExpression Then EOS
                                statementBlock;
elseBlock                   :   Else
                                statementBlock;

// 5.4.2.9 Single-line If Statement
singleLineIfStatement       :   ifWithNonEmptyThen | ifWithEmptyThen;
ifWithNonEmptyThen          :   If booleanExpression Then listOrLabel singleLineElseClause?;
ifWithEmptyThen             :   If booleanExpression Then singleLineElseClause;
singleLineElseClause        :   Else listOrLabel?;
listOrLabel                 :   statementLabel (':' sameLineStatement?)?
                            |   ':'? sameLineStatement (':' sameLineStatement?)*;
sameLineStatement           
    :   fileStatement 
    |   errorHandlingStatement
    |   dataManipulationStatement
    |   controlStatementExceptMultilineIf
    ;

// 5.4.2.10 Select Case Statement
selectCaseStatement         :   Select Case expression EOS
                                caseClause*
                                caseElseClause?
                                End Select;
caseClause                  :   Case rangeClause (',' rangeClause)? EOS 
                                statementBlock;
caseElseClause              :   Case Else EOS 
                                statementBlock;
rangeClause                
    :   expression (To expression)?
    |   Is? comparisonOperator expression
    ;
comparisonOperator 
    :   '=' 
    |   '<>'
    |   '<' 
    |   '>' 
    |   '>=' 
    |   '<='
    ;

// 5.4.2.11 Stop Statement
stopStatement               :   Stop;

// 5.4.2.12 GoTo Statement
gotoStatement               :   GoTo statementLabel;

// 5.4.2.13 On GoTo Statement
onGotoStatement             :   On expression GoTo statementLabelList;

// 5.4.2.14 GoSub Statement
gosubStatement              :   GoSub statementLabel;

// 5.4.2.15 Return Statement
returnStatement             :   Return;

// 5.4.2.16 On GoSub Statement
onGosubStatement            :   On expression GoSub statementLabelList;

// 5.4.2.17 Exit Sub Statement
exitSubStatement            :   Exit Sub;

// 5.4.2.18 Exit Function Statement
exitFunctionStatement       :   Exit Function;

// 5.4.2.19 Exit Property Statement
exitPropertyStatement       :   Exit Property;

// 5.4.2.20 RaiseEvent Statement
raiseEventStatement         :   RaiseEvent identifier ('(' eventArgumentList? ')')?;
eventArgumentList           :   expression (',' expression)*;

// 5.4.2.21 With Statement
withStatement               :   With expression EOS 
                                statementBlock 
                                End With;

// 5.4.3 Data Manipulation Statements
dataManipulationStatement
    :   localVariableDeclaration
    |   staticVariableDeclaration
    |   localConstDeclaration
    |   redimStatement
    |   midStatement
    |   rsetStatement
    |   lsetStatement
    |   letStatement
    |   setStatement
    ;

localVariableDeclaration    :   Dim Shared? variableDeclarationList;
staticVariableDeclaration   :   Static variableDeclarationList;
localConstDeclaration       :   constDeclaration;

redimStatement              :   ReDim Preserve? redimDeclarationList;
redimDeclarationList        :   redimVariableDcl (',' redimVariableDcl)*;
redimVariableDcl            :   redimTypedVariableDcl | redimUntypedDcl;
redimTypedVariableDcl       :   typedName dynamicArrayDim;
redimUntypedDcl             :   untypedName dynamicArrayClause;
dynamicArrayDim             :   '(' dynamicBoundsList ')';
dynamicBoundsList           :   dynamicDimSpec (',' dynamicDimSpec)*;
dynamicDimSpec              :   (integerExpression To)? integerExpression;
dynamicArrayClause          :   dynamicArrayDim asClause?;

eraseStatement              :   Erase eraseList;
eraseList                   :   lExpression (',' lExpression)*;

midStatement                
    :   modeSpecifier '(' stringArgument ',' integerExpression (',' integerExpression)? ')' '=' expression;
modeSpecifier               :   Mid '$'? | MidB '$'?;
stringArgument              :   boundVariableExpression;

lsetStatement               :   LSet boundVariableExpression '=' expression;
rsetStatement               :   RSet boundVariableExpression '=' expression;

letStatement                :   Let? lExpression '=' expression;

setStatement                :   Set lExpression '=' expression;
                                
// 5.4.4 Error Handling Statements
errorHandlingStatement      
    :   onErrorStatement
    |   resumeStatement
    |   errorStatement
    ;
onErrorStatement            :   On Error errorbehavior;
errorbehavior               :   (Resume Next) | (GoTo statementLabel);

resumeStatement             :   Resume (Next | statementLabel)?;

errorStatement              :   Error integerExpression;

// 5.4.5 File Statements
fileStatement
    :   openStatement
    |   closeStatement
    |   seekStatement
    |   lockStatement
    |   unlockStatement
    |   lineInputStatement
    |   widthStatement
    |   writeStatement
    |   inputStatement
    |   putStatement
    |   getStatement
    ;

openStatement               
    :   Open pathName modeClause? accessClause? lock As fileNumber lenClause?;
pathName                    :   expression;
modeClause                  :   For (Append | Binary | Input | Output | Random);
accessClause                :   Access (Read Write | Read | Write);
lock                        
    :   Shared 
    |   Lock Read 
    |   Lock Write 
    |   Lock Read Write
    ;
lenClause                   :   Len '=' expression;

fileNumber                  :   markedFileNumber | expression;
markedFileNumber            :   '#' expression;

closeStatement              :   Reset | (Close fileNumberList?);
fileNumberList              :   fileNumber (',' fileNumber)*;
seekStatement               :   Seek fileNumber ',' expression;

lockStatement               :   Lock fileNumber (',' recordRange)?;
recordRange                 :   expression (To expression)?;

unlockStatement             :   Unlock fileNumber (',' recordRange)?;

lineInputStatement          :   Line Input markedFileNumber ',' variableExpression;

widthStatement              :   Width markedFileNumber ',' expression;

printStatement              :   Print markedFileNumber ',' outputList?;

outputList                  :   outputClause (charPosition? outputClause)*;
outputClause                
    :   spcClause 
    |   tabClause 
    |   expression
    ;
charPosition                :   ';' | ',';
spcClause                   :   Spc '(' expression ')';
tabClause                   :   Tab ('(' expression ')')?;

writeStatement              :   Write markedFileNumber ',' outputList?;

inputStatement              :   Input markedFileNumber ',' inputList;
inputList                   :   boundVariableExpression (',' boundVariableExpression)*;

putStatement                :   Put fileNumber ',' expression? ',' expression;

getStatement                :   Get fileNumber ',' expression? ',' variableExpression;

// 5.6 Expressions
expression                  
    :   lExpression       
    |   literalExpression 
    |   parenthesizedExpression
    |   typeOfIsExpression
    |   newExpression
    |   '-' expression
    |   expression '^'<assoc=right> expression
    |   expression '/' expression
    |   expression '\\' expression
    |   expression Mod expression
    |   expression '*' expression
    |   expression '+' expression
    |   expression '-' expression
    |   expression comparisonOperator expression
    |   expression Eqv expression
    |   expression Imp expression
    ;
lExpression
    :   simpleNameExpression
    |   Me //instanceExpression 5.6.11
    |   lExpression '.' unrestrictedName  //memberAccessExpression
    |   lExpression '(' argumentList ')'  //indexExpression
    |   lExpression '!' unrestrictedName  //dictionaryAccessExpression
    |   withExpression
    ;

// 5.6.5 Literal Expressions
literalExpression
    :   IntegerLiteral
    |   FloatLiteral //|   DateLiteral
    |   StringLiteral
    |   literalIdentifier typeSuffix?
    ;

// 5.6.6 Parenthesized Expressions
parenthesizedExpression     :   '(' expression ')';

// 5.6.7 TypeOf Is Expressions
typeOfIsExpression          :   TypeOf expression Is typeExpression;

// 5.6.8 New Expressions
newExpression               :   New typeExpression;

// 5.6.10 Simple Name Expressions
simpleNameExpression        :   name;

// 5.6.12 Member Access Expressions
// TODO no line continuation or whitespace before period.
memberAccessExpression      :   lExpression '.' unrestrictedName; 

// 5.6.13 Index Expressions
indexExpression             :   lExpression '(' argumentList ')';

// 5.6.13.1 Argument Lists
argumentList  
    :   (argumentExpression? ',')* argumentExpression
    |   (argumentExpression? ',')* namedArgumentList
    ;
namedArgumentList           :   namedArgument (',' namedArgument)*;
namedArgument               :   unrestrictedName ':=' argumentExpression;
argumentExpression          :   ByVal? expression | addressOfExpression;

// 5.6.14 Dictionary Access Expressions
// TODO no whitespace/line-continuation between lExpression and bang.
dictionaryAccessExpression  :   lExpression '!' unrestrictedName;

// 5.6.15 With Expressions
withExpression
    :   ('.' | '!') unrestrictedName
    ;

// 5.6.16.1 Constant Expressions
// TODO enforce semantics on this.
constantExpression          :   expression;

// 5.6.16.3 Boolean Expressions
booleanExpression           :   expression;

// 5.6.16.4 Integer Expressions
integerExpression           :   expression;

// 5.6.16.5 Variable Expressions
variableExpression          :   lExpression;

// 5.6.16.6 Bound Variable Expressions
boundVariableExpression     :   lExpression;

// 5.6.16.7 Type Expressions
typeExpression              
    :   builtInType
    |   simpleNameExpression
    |   memberAccessExpression
    ;
definedTypeExpression       : simpleNameExpression | memberAccessExpression;

// 5.6.16.8 AddressOf Expressions
addressOfExpression         :   AddressOf (simpleNameExpression | memberAccessExpression);

