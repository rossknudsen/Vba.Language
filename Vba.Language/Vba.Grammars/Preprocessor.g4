grammar Preprocessor;

// Parser Rules
preprocessorStatement
    :	moduleAttribute
    |	constantDeclaration
    |	ifStatement
    |	elseIfStatement
    |	elseStatement
    |	endIfStatement
    ;

classHeader  // TODO consider making this case insensitive.
    :   WS* 'VERSION' WS+ '1.0' WS+ 'CLASS' WS* NL
        WS* 'BEGIN' WS* NL
        WS* 'MultiUse' WS* '=' WS* '-1' WS* '\'' WS* 'True' WS* NL
        WS* 'END' WS* NL
    ;

moduleAttribute
    :	Attribute WS+ moduleAttributeOption WS* '=' WS* moduleAttributeValue WS*
    ;

moduleAttributeValue
    :   ModuleName
    |   True
    |   False
    ;

moduleAttributeOption
    :	VbName 
    |   VbExposed
    |   VbGlobalNamespace
    |   VbCreatable 
    |   VbPredeclaredId 
    |   VbCustomizable
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
    ;

impExpression
    :   (boolExpression | arithmeticExpression) WS* Imp WS* (boolExpression | arithmeticExpression)
    ;

eqvExpression
    :   (boolExpression | arithmeticExpression) WS* Eqv WS* (boolExpression | arithmeticExpression)
    ;

stringExpression
    :   stringExpression WS* ('&' | '+') WS* stringExpression
    |   stringExpression WS* Like WS* stringExpression
    |   StringLiteral
    ;

arithmeticExpression
    :   '-' arithmeticExpression
    |   arithmeticExpression WS* '^'<assoc=right> WS* arithmeticExpression
    |   arithmeticExpression WS* '/' WS* arithmeticExpression
    |   arithmeticExpression WS* '\\' WS* arithmeticExpression
    |   arithmeticExpression WS* Mod WS* arithmeticExpression
    |   arithmeticExpression WS* '*' WS* arithmeticExpression
    |   arithmeticExpression WS* '+' WS* arithmeticExpression
    |   arithmeticExpression WS* '-' WS* arithmeticExpression
    |   FloatLiteral
    |   IntegerLiteral
    ;

boolExpression 
    :   Not WS+ boolExpression
    |   boolExpression WS+ And WS+ boolExpression 
    |   boolExpression WS+ Or WS+ boolExpression 
    |   boolExpression WS+ Xor WS+ boolExpression 
    |   True
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
Not         :   'Not';
Nothing     :   'Nothing';
Null        :   'Null';
Or          :   'Or';
True        :   'True';
Then        :   'Then';
Xor         :   'Xor';

ID  :	[A-Za-z] [A-Za-z0-9_]*;

WS          :   [ \t];
NL          :   '\r'? '\n';
COMMENT     :	'\'' ~('\r' | '\n')*;

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