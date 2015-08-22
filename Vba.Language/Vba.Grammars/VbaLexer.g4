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

Abs                     :   A B S;
Access                  :   A C C E S S;
AddressOf               :   A D D R E S S O F;
Alias                   :   A L I A S;
And                     :   A N D;
Any                     :   A N Y;
Append                  :   A P P E N D;
Array                   :   A R R A Y;
As                      :   A S;
Attribute               :   A T T R I B U T E;
Base                    :   B A S E;
Binary                  :   B I N A R Y;
Boolean                 :   B O O L E A N;
ByRef                   :   B Y R E F;
Byte                    :   B Y T E;
ByVal                   :   B Y V A L;
Call                    :   C A L L;
Case                    :   C A S E;
CBool                   :   C B O O L;
CByte                   :   C B Y T E;
CCur                    :   C C U R;
CDate                   :   C D A T E;
CDbl                    :   C D B L;
CDec                    :   C D E C;
CDecl                   :   C D E C L;
CInt                    :   C I N T;
Circle                  :   C I R C L E;
CLng                    :   C L N G;
CLngLng                 :   C L N G L N G;
CLngPtr                 :   C L N G P T R;
Class_Initialize        :   C L A S S '_' I N I T I A L I Z E;
Class_Terminate         :   C L A S S '_' T E R M I N A T E;
Close                   :   C L O S E;
Compare                 :   C O M P A R E;
Const                   :   C O N S T;
CSng                    :   C S N G;
CStr                    :   C S T R;
Currency                :   C U R R E N C Y;
CVar                    :   C V A R;
CVErr                   :   C V E R R;
Database                :   D A T A B A S E;
Date                    :   D A T E;
Debug                   :   D E B U G;
Decimal                 :   D E C I M A L;
Declare                 :   D E C L A R E;
DefBool                 :   D E F B O O L;
DefByte                 :   D E F B Y T E;
DefCur                  :   D E F C U R;
DefDate                 :   D E F D A T E;
DefDbl                  :   D E F D B L;
DefInt                  :   D E F I N T;
DefLng                  :   D E F L N G;
DefLngLng               :   D E F L N G L N G;
DefLngPtr               :   D E F L N G P T R;
DefObj                  :   D E F O B J;
DefSng                  :   D E F S N G;
DefStr                  :   D E F S T R;
DefVar                  :   D E F V A R;
DefDec                  :   D E F D E C;
Dim                     :   D I M;
Do                      :   D O;
DoEvents                :   D O E V E N T S;
Double                  :   D O U B L E;
Each                    :   E A C H;
Else                    :   E L S E;
ElseIf                  :   E L S E I F;
Empty                   :   E M P T Y;
End                     :   E N D;
EndIf                   :   E N D I F;
Enum                    :   E N U M;
Eqv                     :   E Q V;
Erase                   :   E R A S E;
Error                   :   E R R O R;
Event                   :   E V E N T;
Exit                    :   E X I T;
Explicit                :   E X P L I C I T;
False                   :   F A L S E;
Fix                     :   F I X;
For                     :   F O R;
Friend                  :   F R I E N D;
Function                :   F U N C T I O N;
Get                     :   G E T;
Global                  :   G L O B A L;
GoSub                   :   G O S U B;
GoTo                    :   G O T O;
If                      :   I F;
Imp                     :   I M P;
Implements              :   I M P L E M E N T S;
In                      :   I N;
Input                   :   I N P U T;
InputB                  :   I N P U T B;
Int                     :   I N T;
Integer                 :   I N T E G E R;
Is                      :   I S;
LBound                  :   L B O U N D;
Len                     :   L E N;
LenB                    :   L E N B;
Let                     :   L E T;
Lib                     :   L I B;
Like                    :   L I K E;
Line                    :   L I N E;
LINEINPUT               :   L I N E I N P U T;
Lock                    :   L O C K;
Long                    :   L O N G;
LongLong                :   L O N G L O N G;
LongPtr                 :   L O N G P T R;
Loop                    :   L O O P;
LSet                    :   L S E T;
Me                      :   M E;
Mid                     :   M I D;
MidB                    :   M I D B;
Mod                     :   M O D;
Module                  :   M O D U L E;
New                     :   N E W;
Next                    :   N E X T;
Not                     :   N O T;
Nothing                 :   N O T H I N G;
Null                    :   N U L L;
Object                  :   O B J E C T;
On                      :   O N;
Open                    :   O P E N;
Option                  :   O P T I O N;
Optional                :   O P T I O N A L;
Or                      :   O R;
Output                  :   O U T P U T;
ParamArray              :   P A R A M A R R A Y;
Preserve                :   P R E S E R V E;
Print                   :   P R I N T;
Private                 :   P R I V A T E;
Property                :   P R O P E R T Y;
PSet                    :   P S E T;
PtrSafe                 :   P T R S A F E;
Public                  :   P U B L I C;
Put                     :   P U T;
RaiseEvent              :   R A I S E E V E N T;
Random                  :   R A N D O M;
Read                    :   R E A D;
ReDim                   :   R E D I M;
Rem                     :   R E M;
Reset                   :   R E S E T;
Resume                  :   R E S U M E;
Return                  :   R E T U R N;
RSet                    :   R S E T;
Scale                   :   S C A L E;
Seek                    :   S E E K;
Select                  :   S E L E C T;
Set                     :   S E T;
Sgn                     :   S G N;
Shared                  :   S H A R E D;
Single                  :   S I N G L E;
Spc                     :   S P C;
Static                  :   S T A T I C;
Stop                    :   S T O P;
Step                    :   S T E P;
String                  :   S T R I N G;
Sub                     :   S U B;
Tab                     :   T A B;
Text                    :   T E X T;
Then                    :   T H E N;
To                      :   T O;
True                    :   T R U E;
Type                    :   T Y P E;
TypeOf                  :   T Y P E O F;
UBound                  :   U B O U N D;
Unlock                  :   U N L O C K;
Until                   :   U N T I L;
Variant                 :   V A R I A N T;
VB_Base                 :   V B '_' B A S E;
VB_Control              :   V B '_' C O N T R O L;
VB_Creatable            :   V B '_' C R E A T A B L E;
VB_Customizable         :   V B '_' C U S T O M I Z A B L E;
VB_Description          :   V B '_' D E S C R I P T I O N;
VB_Exposed              :   V B '_' E X P O S E D;
VB_Ext_KEY              :   V B '_' E X T '_' K E Y;
VB_GlobalNameSpace      :   V B '_' G L O B A L N A M E S P A C E;
VB_HelpID               :   V B '_' H E L P I D;
VB_Invoke_Func          :   V B '_' I N V O K E '_' F U N C;
VB_Invoke_Property      :   V B '_' I N V O K E '_' P R O P E R T Y;
VB_Invoke_PropertyPut   :   V B '_' I N V O K E '_' P R O P E R T Y P U T;
VB_Invoke_PropertyPutRef:   V B '_' I N V O K E '_' P R O P E R T Y P U T R E F;
VB_MemberFlags          :   V B '_' M E M B E R F L A G S;
VB_Name                 :   V B '_' N A M E;
VB_PredeclaredId        :   V B '_' P R E D E C L A R E D I D;
VB_ProcData             :   V B '_' P R O C D A T A;
VB_TemplateDerived      :   V B '_' T E M P L A T E D E R I V E D;
VB_UserMemId            :   V B '_' U S E R M E M I D;
VB_VarDescription       :   V B '_' V A R D E S C R I P T I O N;
VB_VarHelpID            :   V B '_' V A R H E L P I D;
VB_VarMemberFlags       :   V B '_' V A R M E M B E R F L A G S;
VB_VarProcData          :   V B '_' V A R P R O C D A T A;
VB_VarUserMemId         :   V B '_' V A R U S E R M E M I D;
Wend                    :   W E N D;
While                   :   W H I L E;
Width                   :   W I D T H;
With                    :   W I T H;
WithEvents              :   W I T H E V E N T S;
Write                   :   W R I T E;
Xor                     :   X O R;

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

fragment A:('a'|'A');
fragment B:('b'|'B');
fragment C:('c'|'C');
fragment D:('d'|'D');
fragment E:('e'|'E');
fragment F:('f'|'F');
fragment G:('g'|'G');
fragment H:('h'|'H');
fragment I:('i'|'I');
fragment J:('j'|'J');
fragment K:('k'|'K');
fragment L:('l'|'L');
fragment M:('m'|'M');
fragment N:('n'|'N');
fragment O:('o'|'O');
fragment P:('p'|'P');
fragment Q:('q'|'Q');
fragment R:('r'|'R');
fragment S:('s'|'S');
fragment T:('t'|'T');
fragment U:('u'|'U');
fragment V:('v'|'V');
fragment W:('w'|'W');
fragment X:('x'|'X');
fragment Y:('y'|'Y');
fragment Z:('z'|'Z');