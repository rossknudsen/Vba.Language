grammar Preprocessor;

// Parser Rules
headerLine
    :   classHeaderLine
    |   moduleAttribute
    ;

classHeaderLine
    :   WS* Version WS+ FloatLiteral WS+ Class WS*
    |   WS* Begin WS*
    |   WS* MultiUse WS* '=' WS* IntegerLiteral WS* COMMENT?
    |   WS* End WS*
    ;

moduleAttribute
    :	Attribute WS+ VbName WS* '=' WS* ModuleName WS*
    |   Attribute WS+ optionalAttributes WS* '=' WS* boolLiteral WS*
    ;

optionalAttributes
    :	VbExposed
    |   VbGlobalNamespace
    |   VbCreatable 
    |   VbPredeclaredId 
    |   VbCustomizable
    ;

preprocessorStatement
    :	constantDeclaration
    |	ifStatement
    |	elseIfStatement
    |	elseStatement
    |	endIfStatement
    ;

constantDeclaration
    :	'#' Const WS+ ID WS* '=' WS* expression WS*
    ;

ifStatement
    :	'#' If WS+ expression WS+ Then WS*
    ;

elseIfStatement
    :   '#' Else WS+ If WS+ expression WS+ Then WS*
    ;

elseStatement
    :	'#' Else WS*
    ;

endIfStatement
    :	'#' End WS+ If WS*
    ;

expression
    :   '(' WS* expression WS* ')'
    |   arithmeticExpression
    |   expression WS* comparisonOperator WS* expression
    |   stringExpression
    |   eqvExpression
    |   impExpression 
    |   boolExpression
    |   ID
    ;

comparisonOperator
    :   '='
    |   '<'
    |   '>'
    |   '<='
    |   '>='
    |   '<>'
    |   '><'
    ;

impExpression
    :   (boolExpression | arithmeticExpression) WS* Imp WS* (boolExpression | arithmeticExpression)
    ;

eqvExpression
    :   (boolExpression | arithmeticExpression) WS* Eqv WS* (boolExpression | arithmeticExpression)
    ;

stringExpression
    :   stringExpression WS* (AMP | PLUS) WS* stringExpression
    |   stringExpression WS* Like WS* stringExpression
    |   StringLiteral
    ;

arithmeticExpression
    :   MINUS arithmeticExpression
    |   arithmeticExpression WS* CARET<assoc=right> WS* arithmeticExpression
    |   arithmeticExpression WS* FS WS* arithmeticExpression
    |   arithmeticExpression WS* BS WS* arithmeticExpression
    |   arithmeticExpression WS* Mod WS* arithmeticExpression
    |   arithmeticExpression WS* STAR WS* arithmeticExpression
    |   arithmeticExpression WS* PLUS WS* arithmeticExpression
    |   arithmeticExpression WS* MINUS WS* arithmeticExpression
    |   FloatLiteral
    |   IntegerLiteral
    ;

boolExpression 
    :   Not WS+ boolExpression
    |   boolExpression WS+ And WS+ boolExpression 
    |   boolExpression WS+ Or WS+ boolExpression 
    |   boolExpression WS+ Xor WS+ boolExpression 
    |   boolLiteral
    ;

boolLiteral
    :   True
    |   False
    ;

//Lexer Rules

// Literals.
IntegerLiteral  
    :   DecimalLiteral
    |   HexIntegerLiteral
    |   OctalIntegerLiteral
    ;

fragment
DecimalLiteral
    :   NegativeSymbol? DecimalDigits IntegerSuffix?
    ;

fragment
HexIntegerLiteral
    :   NegativeSymbol? '&' [Hh] [1-9A-Fa-f] [0-9A-Fa-f]* IntegerSuffix?
    ;

fragment
OctalIntegerLiteral
    :   NegativeSymbol? '&' [Oo] [1-7] [0-7]* IntegerSuffix?
    ;

fragment
NegativeSymbol
    :   '-'
    ;

fragment
IntegerSuffix
    :   '%'  // Integer 16 bit signed (default)
    |   '&'  // Long 32 bit signed
    |   '^'  // LongLong 64bit signed
    ;

FloatLiteral
    :   '-'? [0-9]+ '.' [0-9]* Exponent FloatSign? DecimalDigits+ FloatSuffix?
    |   '-'? DecimalDigits Exponent FloatSign? DecimalDigits+ FloatSuffix?
    |   '-'? [0-9]+ '.' [0-9]* FloatSuffix?
    |   '-'? DecimalDigits FloatSuffix  // No decimal or exponent.  Suffix required.
    ;

fragment
DecimalDigits
    :   [1-9] [0-9]*
    ;

fragment 
Exponent   
    :   [DdEe]
    ;

fragment
FloatSign
    :   '+'
    |   '-'
    ;

fragment
FloatSuffix
    :   '!'  // Single
    |   '#'  // Double (default)
    |   '@'  // Currency
    ;

ModuleName  :   '"' [A-Za-z] [A-Za-z0-9_]* '"';

StringLiteral   :   '"' (~["\r\n] | '""')*  '"';

// Module Header Keywords
Attribute           :   'Attribute';
VbName              :   'VB_Name';
VbExposed           :   'VB_Exposed';
VbGlobalNamespace   :   'VB_GlobalNameSpace';
VbCreatable         :   'VB_Creatable';
VbPredeclaredId     :   'VB_PredeclaredId';
VbCustomizable      :   'VB_Customizable';

// Keywords
And         :   'And';
Begin       :   'BEGIN';
Class       :   'CLASS';
Const       :   'Const';
Else        :   'Else';
End         :   'End';
Eqv         :   'Eqv';
Empty       :   'Empty';
False       :   'False';
If          :   'If';
Imp         :   'Imp';
Like        :   'Like';
Mod         :   'Mod';
MultiUse    :   'MultiUse';
Not         :   'Not';
Nothing     :   'Nothing';
Null        :   'Null';
Or          :   'Or';
True        :   'True';
Then        :   'Then';
Version     :   'VERSION';
Xor         :   'Xor';

ID  :	[A-Za-z] [A-Za-z0-9_]*;

WS          :   [ \t];
NL          :   '\r'? '\n';
COMMENT     :	'\'' ~('\r' | '\n')*;

// Symbols
AMP         :   '&';
MINUS       :   '-';
PLUS        :   '+';

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