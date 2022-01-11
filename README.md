# HIL converter

## Documentation

**Basic concepts**
* *TC Build* : .xlsx file containing all TCs with functions to substitute 
* *Substitution* : .xlsx file containing the implementation of all the functions
* *TC Run* : .xlsx output file
* *pathfile.json* : input file indicating the path and sheets name of the previous files

**What does the program do?**
1.  Read all substitution files reported in *pathfile.json*
2.  Create an internal dictionary for mapping functions with their implementation
3.  Import *TC Build*
4.  Create a copy if it and save it as *TC Run*
5.  Close *TC Build* : no modification will be performed on original file
6.  Skim throughout the **ONE** sheet of *TC Run* reported in *pathfile.json* in order to find all function and substitute alla functions
7.  Remove all rows containing all manual tests 
8.  Fill Test N column
9.  Fill Test Enable column
10. Fill Step ID Counter

## Installation and Run

No installation required.
Download TC_tool.exe from installer/ folder, copy the 'filepath.json' file next to it and run it.

## Release History

