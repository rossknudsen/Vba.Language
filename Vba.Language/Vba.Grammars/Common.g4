grammar Common;

// Parser Rules

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
    :   EQ
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

// Lexer Rules

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
    :   [0-9]+
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

StringLiteral   :   '"' (~["\r\n] | '""')*  '"';

// Keywords
And         :   'And';
Const       :   'Const';
Else        :   'Else';
ElseIf      :   'ElseIf';
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

ID			:	[A-Za-z] [A-Za-z0-9_]*;

WS          :   [ \t];
NL          :   '\r'? '\n';
COMMENT     :	'\'' ~('\r' | '\n')*;

// Symbols
AMP         :   '&';
MINUS       :   '-';
PLUS        :   '+';
FS          :   '/';
BS          :   '\\';
CARET       :   '^';
STAR        :   '*';
EQ          :   '=';

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
