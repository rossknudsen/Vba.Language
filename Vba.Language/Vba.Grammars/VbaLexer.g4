lexer grammar VbaLexer;

// 3.3.2 Number Tokens
IntegerLiteral  
    :   DecimalLiteral
    |   HexIntegerLiteral
    |   OctalIntegerLiteral
    ;

FloatLiteral
    :   Sign? [0-9]+ '.' [0-9]* Exponent Sign? DecimalDigits+ FloatSuffix?
    |   Sign? DecimalDigits Exponent Sign? DecimalDigits+ FloatSuffix?
    |   Sign? [0-9]+ '.' [0-9]* FloatSuffix?
    |   Sign? DecimalDigits FloatSuffix  // No decimal or exponent.  Suffix required.
    ;

fragment DecimalLiteral         :   Sign? DecimalDigits IntegerSuffix?;
fragment HexIntegerLiteral      :   Sign? '&' [Hh] [1-9A-Fa-f] [0-9A-Fa-f]* IntegerSuffix?;
fragment OctalIntegerLiteral    :   Sign? '&' [Oo] [1-7] [0-7]* IntegerSuffix?;
fragment IntegerSuffix
    :   '%'  // Integer 16 bit signed (default)
    |   '&'  // Long 32 bit signed
    |   '^'  // LongLong 64bit signed
    ;
fragment DecimalDigits  :   [0-9]+;
fragment Exponent       :   [DdEe];
fragment Sign           :   '+' | '-';
fragment FloatSuffix
    :   '!'  // Single
    |   '#'  // Double (default)
    |   '@'  // Currency
    ;

// 3.3.3 Date Tokens
// The specification allows for a wider range of date expressions than what the VBE
// supports.  At this stage we are limiting the options to what is supported by the VBE.
DateLiteral             :   '#' DateOrTime '#';
fragment DateOrTime 
    :   (DateValue TimeValue)
    |   DateValue 
    |   TimeValue
    ;
fragment DateValue      :   DateComponent DateSeparator DateComponent DateSeparator DateComponent;
fragment DateComponent  :   DecimalLiteral;// | MonthName;
fragment DateSeparator  :   '/';// | '-' | ',';
fragment MonthName       
    :   EnglishMonthName 
    |   EnglishMonthAbbreviation
    ;
fragment EnglishMonthName
    :   'january' 
    |   'february' 
    |   'march' 
    |   'april' 
    |   'may' 
    |   'june' 
    |   'august' 
    |   'september' 
    |   'october' 
    |   'november' 
    |   'december'
    ;
fragment EnglishMonthAbbreviation 
    :   'jan' 
    |   'feb' 
    |   'mar' 
    |   'apr' 
    |   'jun' 
    |   'jul' 
    |   'aug' 
    |   'sep' 
    |   'oct' 
    |   'nov' 
    |   'dec'
    ;
fragment TimeValue 
    :   (DecimalLiteral TimeSeparator DecimalLiteral TimeSeparator DecimalLiteral AmPm);
    /*|   (DecimalLiteral AmPm)
    ;*/
fragment TimeSeparator   
    :   ':';// | '.';
fragment AmPm 
    :   'am' 
    |   'pm';/* 
    |   'a' 
    |   'p'
    ;*/

// 3.3.4 String Tokens
StringLiteral           :   '"' (~["\r\n] | '""')*  '"';

Abs                     :   'Abs';
Access                  :   'Access';
AddressOf               :   'AddressOf';
Alias                   :   'Alias';
And                     :   'And';
Any                     :   'Any';
Append                  :   'Append';
Array                   :   'Array';
As                      :   'As';
Attribute               :   'Attribute';
Base                    :   'Base';
Binary                  :   'Binary';
Boolean                 :   'Boolean';
ByRef                   :   'ByRef';
Byte                    :   'Byte';
ByVal                   :   'ByVal';
Call                    :   'Call';
Case                    :   'Case';
CBool                   :   'CBool';
CByte                   :   'CByte';
CCur                    :   'CCur';
CDate                   :   'CDate';
CDbl                    :   'CDbl';
CDec                    :   'CDec';
CDecl                   :   'CDecl';
CInt                    :   'CInt';
Circle                  :   'Circle';
CLng                    :   'CLng';
CLngLng                 :   'CLngLng';
CLngPtr                 :   'CLngPtr';
Class_Initialize        :   'Class_Initialize';
Class_Terminate         :   'Class_Terminate';
Close                   :   'Close';
Compare                 :   'Compare';
Const                   :   'Const';
CSng                    :   'CSng';
CStr                    :   'CStr';
Currency                :   'Currency';
CVar                    :   'CVar';
CVErr                   :   'CVErr';
Database                :   'Database';
Date                    :   'Date';
Debug                   :   'Debug';
Decimal                 :   'Decimal';
Declare                 :   'Declare';
DefBool                 :   'DefBool';
DefByte                 :   'DefByte';
DefCur                  :   'DefCur';
DefDate                 :   'DefDate';
DefDbl                  :   'DefDbl';
DefInt                  :   'DefInt';
DefLng                  :   'DefLng';
DefLngLng               :   'DefLngLng';
DefLngPtr               :   'DefLngPtr';
DefObj                  :   'DefObj';
DefSng                  :   'DefSng';
DefStr                  :   'DefStr';
DefVar                  :   'DefVar';
DefDec                  :   'DefDec';
Dim                     :   'Dim';
Do                      :   'Do';
DoEvents                :   'DoEvents';
Double                  :   'Double';
Each                    :   'Each';
Else                    :   'Else';
ElseIf                  :   'ElseIf';
Empty                   :   'Empty';
End                     :   'End';
EndIf                   :   'EndIf';
Enum                    :   'Enum';
Eqv                     :   'Eqv';
Erase                   :   'Erase';
Error                   :   'Error';
Event                   :   'Event';
Exit                    :   'Exit';
Explicit                :   'Explicit';
False                   :   'False';
Fix                     :   'Fix';
For                     :   'For';
Friend                  :   'Friend';
Function                :   'Function';
Get                     :   'Get';
Global                  :   'Global';
GoSub                   :   'GoSub';
GoTo                    :   'GoTo';
If                      :   'If';
Imp                     :   'Imp';
Implements              :   'Implements';
In                      :   'In';
Input                   :   'Input';
InputB                  :   'InputB';
Int                     :   'Int';
Integer                 :   'Integer';
Is                      :   'Is';
LBound                  :   'LBound';
Len                     :   'Len';
LenB                    :   'LenB';
Let                     :   'Let';
Lib                     :   'Lib';
Like                    :   'Like';
Line                    :   'Line';
LINEINPUT               :   'LINEINPUT';
Lock                    :   'Lock';
Long                    :   'Long';
LongLong                :   'LongLong';
LongPtr                 :   'LongPtr';
Loop                    :   'Loop';
LSet                    :   'LSet';
Me                      :   'Me';
Mid                     :   'Mid';
MidB                    :   'MidB';
Mod                     :   'Mod';
Module                  :   'Module';
New                     :   'New';
Next                    :   'Next';
Not                     :   'Not';
Nothing                 :   'Nothing';
Null                    :   'Null';
Object                  :   'Object';
On                      :   'On';
Open                    :   'Open';
Option                  :   'Option';
Optional                :   'Optional';
Or                      :   'Or';
Output                  :   'Output';
ParamArray              :   'ParamArray';
Preserve                :   'Preserve';
Print                   :   'Print';
Private                 :   'Private';
Property                :   'Property';
PSet                    :   'PSet';
PtrSafe                 :   'PtrSafe';
Public                  :   'Public';
Put                     :   'Put';
RaiseEvent              :   'RaiseEvent';
Random                  :   'Random';
Read                    :   'Read';
ReDim                   :   'ReDim';
Rem                     :   'Rem';
Reset                   :   'Reset';
Resume                  :   'Resume';
Return                  :   'Return';
RSet                    :   'RSet';
Scale                   :   'Scale';
Seek                    :   'Seek';
Select                  :   'Select';
Set                     :   'Set';
Sgn                     :   'Sgn';
Shared                  :   'Shared';
Single                  :   'Single';
Spc                     :   'Spc';
Static                  :   'Static';
Stop                    :   'Stop';
Step                    :   'Step';
String                  :   'String';
Sub                     :   'Sub';
Tab                     :   'Tab';
Text                    :   'Text';
Then                    :   'Then';
To                      :   'To';
True                    :   'True';
Type                    :   'Type';
TypeOf                  :   'TypeOf';
UBound                  :   'UBound';
Unlock                  :   'Unlock';
Until                   :   'Until';
Variant                 :   'Variant';
VB_Base                 :   'VB_Base';
VB_Control              :   'VB_Control';
VB_Creatable            :   'VB_Creatable';
VB_Customizable         :   'VB_Customizable';
VB_Description          :   'VB_Description';
VB_Exposed              :   'VB_Exposed';
VB_Ext_KEY              :   'VB_Ext_KEY';
VB_GlobalNameSpace      :   'VB_GlobalNameSpace';
VB_HelpID               :   'VB_HelpID';
VB_Invoke_Func          :   'VB_Invoke_Func';
VB_Invoke_Property      :   'VB_Invoke_Property';
VB_Invoke_PropertyPut   :   'VB_Invoke_PropertyPut';
VB_Invoke_PropertyPutRef:   'VB_Invoke_PropertyPutRef';
VB_MemberFlags          :   'VB_MemberFlags';
VB_Name                 :   'VB_Name';
VB_PredeclaredId        :   'VB_PredeclaredId';
VB_ProcData             :   'VB_ProcData';
VB_TemplateDerived      :   'VB_TemplateDerived';
VB_UserMemId            :   'VB_UserMemId';
VB_VarDescription       :   'VB_VarDescription';
VB_VarHelpID            :   'VB_VarHelpID';
VB_VarMemberFlags       :   'VB_VarMemberFlags';
VB_VarProcData          :   'VB_VarProcData';
VB_VarUserMemId         :   'VB_VarUserMemId';
Wend                    :   'Wend';
While                   :   'While';
Width                   :   'Width';
With                    :   'With';
WithEvents              :   'WithEvents';
Write                   :   'Write';
Xor                     :   'Xor';

// Symbols.
LB                      :   '[';
RB                      :   ']';
LP                      :   '(';
RP                      :   ')';
PERCENT                 :   '%';
AMP                     :   '&';
CARET                   :   '^';
EXCL                    :   '!';
HASH                    :   '#';
AT                      :   '@';
DOLLAR                  :   '$';
COMMA                   :   ',';
DASH                    :   '-';
AST                     :   '*';
EQUALS                  :   '=';
SEMI                    :   ';';
COLON                   :   ':';
LT                      :   '<';
GT                      :   '>';
LTET                    :   '<=';
GTET                    :   '>=';
NE                      :   '<>';
BS                      :   '\\';
FS                      :   '/';
PLUS                    :   '+';
DOT                     :   '.';
COLONEQUAL              :   ':=';

ForeignName             :   '[' (~('\u000D' | '\u000A' | '\u2028' | '\u2029'))+ ']';

LETTER                  :   [A-Za-z];
ID                      :   [A-Za-z] [A-Za-z0-9_]*;
EOS                     :   (NL | ':')+;
NL                      :   '\r'? '\n';
LC                      :   WS+ '_' WS* NL              -> channel(HIDDEN);
WS                      :   [ \t]                       -> channel(HIDDEN);
RemStatement            :   Rem (LC | ~('\r' | '\n')*);
COMMENT                 :   '\'' (LC | ~('\r' | '\n')*) -> channel(HIDDEN);
