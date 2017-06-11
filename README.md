# Excel Flat File Converter
A tool written in C# to remove all formulas, database queries, external links, etc. from an Excel file, whilst maintaining formatting.

## Usage
Compile the code to an .exe file, which can be run with the first argument as the Excel file that is to be cloned and converted. The exectuable will generate a copy of the file which will be given a suffix to indicate that it is a 'flat' file. 
If no arguments are given, the executable will take input in order to determine the file to convert.
