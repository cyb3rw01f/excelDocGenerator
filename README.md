# excelDocGenerator
PowerShell script that will generate an Excel document that contains a macro. 

Author @cyberw01f  
Last updated: 05/14/2020  

## SYNOPSIS 
This file is for security control testing purposes or malware analysis research only. Malicious use of this script is prohibitted.  

## DESCRIPTION
Script generates a Microsoft Excel document containing a macro. The macro vba code can be customzed. Simply replace the informastion in the $code variable secion of the script.  

Look for this section, $code = @" @" Just place your code between the @" and "@ section and execute the script to generate your excel file"  

The PS1 script contains code that generates random data to insure each file generated is forensicly unique.  

## Notes
Macro enabled Excel Document generater requires Microsoft Excel to be installed on the syestem that the scipt will be executed.  

Some anti malware programs detect the excel file when its generated so you may need to supress detection while the script is in use.  

## How to use excelDocGenerator
To generate a excel document with a unique MD5sum simply execute the PS1 script.  
PS C:\Users\cyb3rw01f\excelDocGenerator> .\excelDocGenerator.ps1
