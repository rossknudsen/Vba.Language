# Vba.Language
A grammar built using ANTLR for Visual Basic for Applications (VBA)

Architecture
------------
This project is intending to create a complete VBA compiler chain.  This means building a precompiler and compiler chain.  The public interface will probably use a Facade class which will assemble the necessary object graph required and provide a simple way to access the compiler functionality.

Current Status
--------------
Currently we are working on the precompiler.  To understand how this works, it is recommended to look at the class diagram showing the interaction of the different objects.  Basically the precompiler is derived from the System.IO.TextReader class.  It takes a reference to the actual source code and reads any header or precompiler statements.  When the ReadXXX and Peek methods are called, it emits the resulting precompiled VBA source code to the caller.

The reason we have done this is because the Antlr toolchain accepts a TextReader object as an input making it straight-forward to create a VBA Antlr grammar which can act directly on the output.

Challenges
----------
Because the VBA compiler will not see all the source input, it will not calculate the line numbers correctly.  The plan to resolve this issue is to put a proxy between the VBA lexer and parser which will be able to query the preprocessor for the correct line number and alter the tokens before the parser receives them.

Recommended Workflow
--------------------
The grammar(s) have been put into a separate project within the solution, mainly to resolve conflicts when using Resharper.  We recommend building Vba.Grammar then unloading the project in Visual Studio.  Vba.Grammar is referenced by dll (Debug) so the rest of the solution will function correctly and Resharper should be happy.

We have also found that using AntlrWorks to edit the grammar(s) outside of Visual Studio quite helpful in this workflow.  You just need to remember to rebuild Vba.Grammar (and unload again it if you loaded it in Visual Studio) after making changes to the grammar.
