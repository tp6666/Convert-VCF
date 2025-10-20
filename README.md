Powershell code to process a VCF (contact) data file and convert its data for output.
Examples (at end of this PS script file) show how to output to console, GridView or a CVS file, as required.
VCF data comes in many formats - this code only caters for some formats / versions, NOT all.
VCF format is documented at https://www.rfc-editor.org/rfc/rfc6350
IMPORTANTLY: this code only extracts data for fields specified in the New-Card data structure.
All other contact fields, are ignored.

Adding data extraction for other data fields, requires writting code to 
 - add the field to the data structure.
 - identify relevant lines for the field and its data (can be multi-line) in the VCF file.
 - extracting the fields data and writting it to the data structure.
