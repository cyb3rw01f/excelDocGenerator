<#                             
Author of excelDocGenerator script @cyberw01f
Last updated 6/24/2018

.SYNOPSIS 
This file is for testing purposes only. Malicious use of this script is prohibitted. 

.DESCRIPTION
Script generates a Microsoft Excel document containing a maco. The macro vba code can be customzed. 

.Notes
##############################################################
Macro enabled Excel Document generater requiers Microsoft Excel
##############################################################
I developed this script to assist testing security controls ability to detect/alert on Word documents containing macros. 

#>

#nl var simply does a carrage return / new line
$nl = [Environment]::NewLine

$logo = @"
	=================================================================
				   _                         ___  _  __ 
			 ___ _   _| |__   ___ _ ____      __/ _ \/ |/ _|
			/ __| | | | '_ \ / _ \ '__\ \ /\ / / | | | | |_ 
		   | (__| |_| | |_) |  __/ |   \ V  V /| |_| | |  _|
			\___|\__, |_.__/ \___|_|    \_/\_/  \___/|_|_|  
				|___/                                      
		
	==================================================================
		,%#((%                                             %%%%%.
		%(%%,#%((%                                       %%%%%,&%%%               
		&(#(    /(&%%                                   &%%&*    #%%&              
		%%%*       .&%%#                               %%%&        #&%%                 
		,%%%          ,&%%                             %%%,          &%%               
		%%%             %%%,                         ,%%%             %%%                          
		%%%              %%%*                       *%%%              &%%                            
		,%%,         /%%#  %%%                       %%&  #%%/         ,%%,                           
		#%%           &%%.  %#%&&%%%%%%%%%%%%%%%&%%&%%%  /%%%           %%(                           
		(%%            %%%  .%%%&%*.       .*#&%&%%%.  %%%            %%(;)                              
		,%%*            %%%  %%%                   %&&  %%%            *%%                               
		%%%            %&%                             &%&            &%%                                  
	   &%%                                                           %%&                                    
	   *%%%                                                         %%%                                      
		%%%                                                         %%%                                       
		 &%%                                                       &%%                                         
		 .%%%                                                     %%%.                                          
		  ,%%%                                                   %%%.                                            
		   %%%                                                   %%%                                            
		 .%%%                                                     %%%.                                            
		,&%%      .%%%%%%%%%%%#                 *&(%%(%(%%%&.      %%%,                                            
		%%%       %%%%%%%%%%%%%&/             /((&(&(%((%%%%%       %%%                                            
	   %%%              &%%&%%%%%&           &%&((((((&              %%%                                            
	  %(&.                &%%%,%%%           %%%,%%%%                 %%%                                           
	  %#(                      %%%           %%%                      %%%                                           
	 #%%                       %%%           %%%                       &%&                                         
	 &%%                       %%%           %%%                       %%%                                         
	 %%(                       %%%           &%%                       %%%                                         
	,%%,                       %%&           %%%                       (%%,                                        
	 %%%%                                                             %%%%                                         
	   &%%.                                                         .%%%                                           
		%%%%               %%%    #%%%%%%%%%#    %%%               %%%%                                            
		 ,%%%              %#%/  %%&/.....(%%%  /%%%              &%%,                                             
		   &%%              %%%   %%%     %%%   %%&              %%%                                               
			&%%              &%%  ,%%&,,,%&%,  %%%              %%%                                                
			 %%%  ,&,         %%%  .&%%%&%&.  &%%         ,&,  &%%                                                     
			 .%%%%%%%%        ,%%%           %%%,        (%%%%%%%.                                                 
			  *(*  .%%%        /%%%%%%%%%%%%%%&.        (%%.  /&(                                                     
					 %%%          #&&&&&&&&          %%%                                                           
					  %%%                             %%%                                                           
					   &%%                           %#%                                                             
					   %%%(  (%%               #((  /%%/                                                            
						%%&%%%%%&             %%%%%%%%%                                                             
						*%%(   %%         %%%%   (%%*                                                            
								#%%%       &%%#                                                                   
								  %%&     %%%                                                                     
								   %%%   %%%                                                                     
									%%% %%%                                                                     
									 %%%%%                                                                     
									 #%%%#                                                                    
									  %%&  
									   . 
								   @cyberw01f
"@

$label = @"  
                   Macro Enabled Excel Document Generator
  Documents generated by this script are to be used only for testing security controls
                       Responsible use only permited
"@

function Get-RandomAlphaNum($len)
{
	$r = "1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
	$tmp = foreach ($i in 1..[int]$len) {$r[(Get-Random -Minimum 1 -Maximum $r.Length)]}
	return [string]::Join('', $tmp)
}

function Create-Excel()
{
###############################################################################
# Place your custom macro vba code between the code=@" and "@ lines.
###############################################################################
$code = @"
'PowerShell encoded command macro.
'Gets ip default gateway addr and Pings it 5 times. Ping results are displayed in a message box.
Private Sub Workbook_Open()
    strCommand = "Powershell -noprofile -windowstyle hidden -e JABnAGEAdABlACAAPQAgACgARwBlAHQALQB3AG0AaQBPAGIAagBlAGMAdAAgAFcAaQBuADMAMgBfAG4AZQB0AHcAbwByAGsAQQBkAGEAcAB0AGUAcgBDAG8AbgBmAGkAZwB1AHIAYQB0AGkAbwBuACAAfAAgAD8AewAkAF8ALgBJAFAARQBuAGEAYgBsAGUAZAB9ACkALgBEAGUAZgBhAHUAbAB0AEkAUABHAGEAdABlAHcAYQB5AA0ACgBwAGkAbgBnACAALQBuACAANQAgACQAZwBhAHQAZQA="
    Set WshShell = CreateObject("WScript.Shell")
    Set WshShellExec = WshShell.Exec(strCommand)
    strOutput = WshShellExec.StdOut.ReadAll
    MsgBox strOutput
End Sub
"@
        # 1. Create new excel application instance
        $time = ([WMI]'').ConvertToDateTime((gwmi win32_operatingsystem).LocalDateTime)
        [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Vbe.Interop") | Out-Null
        #$docName = Read-Host "Enter a name for the document but do not include file extension"
        $docName = Get-RandomAlphaNum 5
		$excel = New-Object -ComObject Excel.Application
        $excel.Visible=$true
        $excel.DisplayAlerts = $true
        $excel.ScreenUpdating = $true
		#$excelVersion = $excel01.Version
        New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($excel.Version)\excel\Security" -Name AccessVBOM -Value 1 -Force | Out-Null
        New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($excel.Version)\excel\Security" -Name VBAWarnings -Value 1 -Force | Out-Null
        $workbook = $excel.Workbooks.Add(1)
        $worksheet=$workbook.WorkSheets.item(1)
        $Data= $workbook.Worksheets.Item(1)
        $Data.Name = 'MySpreadsheet'
        $Data.Cells.Item(1,1) = "This file was auto generated on $time by @cyberw01f's powershell script for the purpose of testing anti-malware controls."
        $Data.Cells.Item(3,1) = "This Excel workbook contains a macro that will auto run a commands when the document is opened."
        $Data.Cells.Item(4,1) = "A message box will be displayed after the script executes with results. This will let you know that the macro ran when the document opened."
        $Data.Cells.Item(7,1) = "Cells 9A, 11A ,13A and 15A contain a set of random alphanumeric characters that are generated at the time of document creation to ensure the file is a unique hash."
        $Data.Cells.Item(9,1) = Get-RandomAlphaNum 100
        $Data.Cells.Item(11,1) = Get-RandomAlphaNum 100
        $Data.Cells.Item(13,1) = Get-RandomAlphaNum 100
        $Data.Cells.Item(15,1) = Get-RandomAlphaNum 100

        # Format, save and quit excel
        $usedRange = $Data.UsedRange                                                                                              
        $usedRange.EntireColumn.AutoFit() | Out-Null
        $xlmodule = $workbook.VBProject.VBComponents.item(1)
        $xlmodule.CodeModule.AddFromString($code)

        #$saveName = ("$($ENV:UserProfile)\Desktop\$docName", [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel8)
        #$workbook.SaveAs($saveName, 52)
        $workbook.SaveAs("$($ENV:UserProfile)\Desktop\$docName", [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel8)

         #Cleanup
        $excel.Workbooks.Close()
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | out-null
        $Excel01 = $Null
        if (ps excel){kill -name excel}
        #####################################################################################
        # Document will be saved to the curent users desktop
        #####################################################################################
        $file = ("$($ENV:UserProfile)\Desktop\$docName.*")
        #$md5sum= Get-FileHash $file -Algorithm MD5
        #$md5sum
        $md5sum = Get-FileHash $file -Algorithm MD5 #| Format-List
        $sha1 = Get-FileHash $file -Algorithm SHA1 #| Format-List
        $sha256 = Get-FileHash $file -Algorithm SHA256 #| Format-List        
        $md5sum
        $sha1
        $sha256
        
}
Write-Host -f Magenta $logo
Write-Host
Write-Host -f Green $label		
Create-Excel
