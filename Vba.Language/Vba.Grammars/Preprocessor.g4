grammar Preprocessor;

import Common;

// Parser Rules
headerLine
    :   classHeaderLine
    |   moduleAttribute
    ;

classHeaderLine
    :   WS* Version WS+ FloatLiteral WS+ Class WS*
    |   WS* Begin WS*
    |   WS* MultiUse WS* EQ WS* IntegerLiteral WS* COMMENT?
    |   WS* End WS*
    ;

moduleAttribute
    :	Attribute WS+ VbName WS* EQ WS* ModuleName WS*
    |   Attribute WS+ optionalAttributes WS* EQ WS* boolLiteral WS*
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
    :	WS* '#' Const WS+ ID WS* EQ WS* expression WS*
    ;

ifStatement
    :	WS* '#' If WS+ expression WS+ Then WS*
    ;

elseIfStatement
    :   WS* '#' ElseIf WS+ expression WS+ Then WS*
    ;

elseStatement
    :	WS* '#' Else WS*
    ;

endIfStatement
    :	WS* '#' End WS+ If WS*
    ;

//Lexer Rules

ModuleName  :   '"' [A-Za-z] [A-Za-z0-9_]* '"';

// Module Header Keywords
Attribute           :   'Attribute';
VbName              :   'VB_Name';
VbExposed           :   'VB_Exposed';
VbGlobalNamespace   :   'VB_GlobalNameSpace';
VbCreatable         :   'VB_Creatable';
VbPredeclaredId     :   'VB_PredeclaredId';
VbCustomizable      :   'VB_Customizable';

// Keywords
Begin       :   'BEGIN';
Class       :   'CLASS';
MultiUse    :   'MultiUse';
Version     :   'VERSION';
