Creole Forth for VB6
--------------------

Intro
-----

This is a Forth-like scripting language built in VB6 similar to Creole Forth For Excel. It is intended
to run as a DSL on top of an existing application. 

Methodology
-----------
Primitives are defined as VB methods attached to objects. They are roughly analagous to core words defined 
in assembly language in some Forth compilers. They are then passed to the BuildPrimitive method in the CreoleForthBundle class, which assigns them a name, vocabulary, and integer token value, which is used as an address. 

High-level or colon definitions are assemblages of primitives and previously defined high-level definitions.
They are defined by the colon compiler. 


