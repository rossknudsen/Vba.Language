grammar Vba;

import Common;

// Specification: [MS-VBAL] VBA Language Specification

// 3.3.5.2 Reserved Identifiers and IDENTIFIER
identifier      
    :   Alias
    |   Base
    |   Binary
    |   ClassInit
    |   ClassTerm
    |   CLngLng
    |   Compare
    |   Database
    |   DefLngLng
    |   Explicit
    |   Lib
    |   LongLong
    |   Module
    |   Object
    |   Property
    |   Text
    |   VB_Invoke_PropertyPutRefVB_MemberFlags
    |   PtrSafe
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
    |    Me
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
    :   boolLiteral
    |   Nothing
    |   Empty
    |   Null
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
    |    VB_Invoke_PropertyPutRefVB_MemberFlags
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
typedName                   :   ID typeSuffix;
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
    :   declarationSection NL moduleCodeSection
    ;

unrestrictedName            :   name | reservedIdentifier;
name                        :   untypedName | typedName;
untypedName                 :   ID; //| foreignName;

// 5.2 Module Declaration Section Structure
declarationSection
    :   (directiveElement EOS)*;

directiveElement
    :   optionCompareDirective
    |   optionBaseDirective
    |   optionExplicitDirective
    |   optionPrivateDirective 
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

// 5.2.4.2 Implements Directive
implementsDirective         :   Implements ID;

// 5.2.3.1 Module Variable Declaration Lists
variableDeclaration :   (Global | Public | Private | Dim) Shared? variableDclList;
variableDclList     :   (variableDcl | witheventsVariableDcl) (',' (variableDcl | witheventsVariableDcl))*; 

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

moduleCodeSection
    :   subroutineDeclaration
    |   functionDeclaration
    |   propGetDeclaration
    |   propLhsDeclaration
    |   implementsDirective
    ;

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
udtElement              :   udtMember;  //| remStatement;
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
    ;

// 5.2.3.4 Enum Declarations
enumDeclaration         :   (Global | Public | Private)? Enum untypedName EOS
                            memberList EOS
                            End Enum;
memberList              :   enumElement (EOS enumElement)*;  
enumElement             :   untypedName ('=' constantExpression)?; // TODO remStatement.

// 5.2.3.5 External Procedure Declaration
extProcDeclaration      :   (Public | Private)? Declare PtrSafe? (Sub | Function) ID 
                            Lib StringLiteral (Alias StringLiteral)? procedureParameters;
externalProcDcl         :   Declare PtrSafe? (externalSub | externalFunction);
externalSub             :   Sub subroutineName libInfo procedureParameters?;
externalFunction        :   Function functionName libInfo procedureParameters? functionType?;
libInfo                 :   Lib StringLiteral (Alias StringLiteral)?;

// 5.2.4.3 Event Declaration
eventDeclaration    :   (Public )? Event ID ('(' ( positionalParameters )? ')')?;

// 5.3.1 Procedure Declarations
subroutineDeclaration
    :   procedureScope? Static? Sub identifier procedureParameters? Static? EOS
        End Sub EOS;

functionDeclaration
    :   procedureScope? Static? Function ID procedureParameters? functionType? Static? EOS
        End Function EOS;

propGetDeclaration
    :   procedureScope? Static? Property Get ID procedureParameters? functionType? Static? EOS
        End Property;

propLhsDeclaration
    :   procedureScope? Static? Property (Set | Let) ID procedureParameters? functionType? Static? EOS
        End Property;

// 5.3.1.1 Procedure Scope
procedureScope
    :   Global 
    |   Public 
    |   Private
    |   Friend
    ;

// 5.3.1.3 Procedure Names
subroutineName  :   ID | prefixedName;
functionName    :   typedName | ID | prefixedName;
prefixedName    :   eventHandlerName | implementedName | lifecycleHandlerName;

// 5.3.1.4 Function Type Declarations
functionType    :   As ID arrayDesignator?;
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
paramArray              :   ParamArray ID '(' ')' (As (Variant | '[' Variant ']'))?;
paramDcl                :   ID arrayDesignator?;
optionalPrefix          :   Optional parameterMechanism? | parameterMechanism? Optional;
parameterMechanism      :   ByRef | ByVal;
parameterType           :   arrayDesignator? As (typeExpression | Any);
defaultValue            :   '=' constantExpression;

// 5.3.1.8 Event Handler Declarations
eventHandlerName        :   ID;

// 5.3.1.9 Implemented Name Declarations
implementedName         :   ID;

// 5.3.1.10 Lifecycle Handler Declarations
lifecycleHandlerName    :   ClassInit | ClassTerm;

// 5.4 Procedure Bodies and Statements
//procedureBody           :   statementBlock;

/*
// 5.6 Expressions
expression              :   valueExpression | lExpression;
valueExpression         
    :   literalExpression 
    |   parenthesizedExpression
    |   typeOfIsExpression
    |   newExpression
    |   operatorExpression
    ;
lExpression
    :   simpleNameExpression
    |   Me //instanceExpression 5.6.11
    |   memberAccessExpression
    |   indexExpression
    |   dictionaryAccessExpression
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

// 5.6.7 TypeOf…Is Expressions
typeOfIsExpression          : TypeOf expression Is typeExpression;

// 5.6.8 New Expressions
newExpression               :   New typeExpression;

// 5.6.9 Operator Expressions
operatorExpression
    :   arithmeticOperator
    |   concatenationOperator
    |   relationalOperator
    |   likeOperator
    |   isOperator
    |   logicalOperator
    ;
*/
// 5.6.10 Simple Name Expressions
simpleNameExpression        :   name;
/*
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
*/

// 5.6.16.1 Constant Expressions
// TODO enforce semantics on this.
constantExpression          :   expression;

// 5.6.16.5 Variable Expressions
//variableExpression          :   lExpression;

// 5.6.16.6 Bound Variable Expressions
//boundVariableExpression     :   lExpression;

// 5.6.16.7 Type Expressions
typeExpression              
    :   builtInType
    //|   simpleNameExpression
    //|   memberAccessExpression
    ;
definedTypeExpression : simpleNameExpression ;//| memberAccessExpression;

// 5.6.16.8 AddressOf Expressions
//addressOfExpression         :   AddressOf (simpleNameExpression | memberAccessExpression);
                                
Abs                     :    'Abs';
AddressOf               :    'AddressOf';
Alias                   :    'Alias';
And                     :    'And';
Any                     :    'Any';
Array                   :    'Array';
As                      :    'As';
Attribute               :    'Attribute';
Base                    :    'Base';
Binary                  :    'Binary';
Boolean                 :    'Boolean';
ByRef                   :    'ByRef';
Byte                    :    'Byte';
ByVal                   :    'ByVal';
Call                    :    'Call';
Case                    :    'Case';
CBool                   :    'CBool';
CByte                   :    'CByte';
CCur                    :    'CCur';
CDate                   :    'CDate';
CDbl                    :    'CDbl';
CDec                    :    'CDec';
CDecl                   :    'CDecl';
CInt                    :    'CInt';
Circle                  :    'Circle';
CLng                    :    'CLng';
CLngLng                 :    'CLngLng';
CLngPtr                 :    'CLngPtr';
ClassInit               :    'Class_Initialize';
ClassTerm               :    'Class_Terminate';
Close                   :    'Close';
Compare                 :    'Compare';
Const                   :    'Const';
CSng                    :    'CSng';
CStr                    :    'CStr';
Currency                :    'Currency';
CVar                    :    'CVar';
CVErr                   :    'CVErr';
Database                :    'Database';
Date                    :    'Date';
Debug                   :    'Debug';
Decimal                 :    'Decimal';
Declare                 :    'Declare';
DefBool                 :    'DefBool';
DefByte                 :    'DefByte';
DefCur                  :    'DefCur';
DefDate                 :    'DefDate';
DefDbl                  :    'DefDbl';
DefInt                  :    'DefInt';
DefLng                  :    'DefLng';
DefLngLng               :    'DefLngLng';
DefLngPtr               :    'DefLngPtr';
DefObj                  :    'DefObj';
DefSng                  :    'DefSng';
DefStr                  :    'DefStr';
DefVar                  :    'DefVar';
DefDec                  :    'DefDec';
Dim                     :    'Dim';
Do                      :    'Do';
DoEvents                :    'DoEvents';
Double                  :    'Double';
Each                    :    'Each';
Else                    :    'Else';
ElseIf                  :    'ElseIf';
Empty                   :    'Empty';
End                     :    'End';
EndIf                   :    'EndIf';
Enum                    :    'Enum';
Eqv                     :    'Eqv';
Erase                   :    'Erase';
Event                   :    'Event';
Exit                    :    'Exit';
Explicit                :    'Explicit';
Fix                     :    'Fix';
For                     :    'For';
Friend                  :    'Friend';
Function                :    'Function';
Get                     :    'Get';
Global                  :    'Global';
GoSub                   :    'GoSub';
GoTo                    :    'GoTo';
If                      :    'If';
Imp                     :    'Imp';
Implements              :    'Implements';
In                      :    'In';
Input                   :    'Input';
InputB                  :    'InputB';
Int                     :    'Int';
Integer                 :    'Integer';
Is                      :    'Is';
LBound                  :    'LBound';
Len                     :    'Len';
LenB                    :    'LenB';
Let                     :    'Let';
Lib                     :    'Lib';
Like                    :    'Like';
LINEINPUT               :    'LINEINPUT';
Lock                    :    'Lock';
Long                    :    'Long';
LongLong                :    'LongLong';
LongPtr                 :    'LongPtr';
Loop                    :    'Loop';
LSet                    :    'LSet';
Me                      :    'Me';
Mod                     :    'Mod';
Module                  :    'Module';
New                     :    'New';
Next                    :    'Next';
Not                     :    'Not';
Nothing                 :    'Nothing';
Null                    :    'Null';
Object                  :    'Object';
On                      :    'On';
Open                    :    'Open';
Option                  :    'Option';
Optional                :    'Optional';
Or                      :    'Or';
ParamArray              :    'ParamArray';
Preserve                :    'Preserve';
Print                   :    'Print';
Private                 :    'Private';
Property                :    'Property';
PSet                    :    'PSet';
PtrSafe                 :    'PtrSafe';
Public                  :    'Public';
Put                     :    'Put';
RaiseEvent              :    'RaiseEvent';
ReDim                   :    'ReDim';
Rem                     :    'Rem';
Resume                  :    'Resume';
Return                  :    'Return';
RSet                    :    'RSet';
Scale                   :    'Scale';
Seek                    :    'Seek';
Select                  :    'Select';
Set                     :    'Set';
Sgn                     :    'Sgn';
Shared                  :    'Shared';
Single                  :    'Single';
Spc                     :    'Spc';
Static                  :    'Static';
Stop                    :    'Stop';
String                  :    'String';
Sub                     :    'Sub';
Tab                     :    'Tab';
Text                    :    'Text';
Then                    :    'Then';
To                      :    'To';
Type                    :    'Type';
TypeOf                  :    'TypeOf';
UBound                  :    'UBound';
Unlock                  :    'Unlock';
Until                   :    'Until';
Variant                 :    'Variant';
VB_Base                 :    'VB_Base';
VB_Control              :    'VB_Control';
VB_Creatable            :    'VB_Creatable';
VB_Customizable         :    'VB_Customizable';
VB_Description          :    'VB_Description';
VB_Exposed              :    'VB_Exposed';
VB_Ext_KEY              :    'VB_Ext_KEY';
VB_GlobalNameSpace      :    'VB_GlobalNameSpace';
VB_HelpID               :    'VB_HelpID';
VB_Invoke_Func          :    'VB_Invoke_Func';
VB_Invoke_Property      :    'VB_Invoke_Property';
VB_Invoke_PropertyPut   :    'VB_Invoke_PropertyPut';
VB_Invoke_PropertyPutRefVB_MemberFlags  :    'VB_Invoke_PropertyPutRefVB_MemberFlags';
VB_Name                 :    'VB_Name';
VB_PredeclaredId        :    'VB_PredeclaredId';
VB_ProcData             :    'VB_ProcData';
VB_TemplateDerived      :    'VB_TemplateDerived';
VB_UserMemId            :    'VB_UserMemId';
VB_VarDescription       :    'VB_VarDescription';
VB_VarHelpID            :    'VB_VarHelpID';
VB_VarMemberFlags       :    'VB_VarMemberFlags';
VB_VarProcData          :    'VB_VarProcData';
VB_VarUserMemId         :    'VB_VarUserMemId';
Wend                    :    'Wend';
While                   :    'While';
With                    :    'With';
WithEvents              :    'WithEvents';
Write                   :    'Write';
Xor                     :    'Xor';

LETTER                  :    [A-Za-z];

// End of statement.
EOS                     :   NL | ':';

WS                      :   [ \t] -> channel(HIDDEN);
