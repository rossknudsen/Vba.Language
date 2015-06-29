grammar Vba;

//import Common;

// Parser Rules

compileUnit
	:	EOF
	;

// Lexer Rules

WS
	:	' ' -> channel(HIDDEN)
	;
