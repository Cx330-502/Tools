# C++ Dependency
```mermaid
graph LR

CPP1(lab4.cpp); 
CPP2(Compiler.cpp); 	H2(Compiler.h); 
CPP3(Error.cpp); 	H3(Error.h); 
CPP4(Lexer.cpp); 	H4(Lexer.h); 
CPP5(Parser.cpp); 	H5(Parser.h); CPP1-->H2 ; 
CPP2-->H2 ; CPP2-->H4 ; CPP2-->H5 ; 
H2-->H3 ; 
CPP3-->H3 ; 

CPP4-->H4 ; 
H4-->H2 ; 
CPP5-->H5 ; 
H5-->H4 ; H5-->H2 ; 
