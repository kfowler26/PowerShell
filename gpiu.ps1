<#
Program: GOS Printer Installation Utility (GPIU)
Author: Kendall Fowler
Date Created: October 2021
Last Update: 02/12/2021
Description: Program to install printer drivers onto clients asset. Code modified from VBScript
#>

#--------------------------------------------
# Global Variables and Functions here
#--------------------------------------------

[bool]$global:validPath

[string]$global:bitType
[string]$global:driver
[string]$global:inf
[string]$global:model
[string]$global:port
[string]$global:printer
[string]$global:type

#File Paths
[string]$global:driver64CSV = "\\******\gpiu$\bin\csv\64bit.csv"
[string]$global:driver32CSV = "\\******\gpiu$\bin\csv\32bit.csv"
[string]$global:printerCSV = "\\******\gpiu$\bin\csv\printers.csv"

#Message Box Strings
[string]$global:msgErrDriver = "The driver is not available to install for the printer requested. Please select an alternate driver and try again. If you require assistance, please call the Service Desk at 306-787-5000."
[string]$global:msgErrPrinter = "The printer tag you have entered cannot be found. Please ensure the correct printer tag is being used. If you continue to recieve this error message, please call the Service Desk at 306-787-5000 to install this printer."
[string]$global:msgErrTag = "Please ensure the printer tag being entered follows proper naming conventions."
[string]$global:msgErr32Bit = "This computer requires manual driver installation, please call the Service Desk at ***-***-**** to install this printer."
[string]$global:msgPrinterSucess = "The printer tag you have entered has successfully been installed. Would you like to set this printer as your default printer?"

#Message Box Titles
[string]$global:msgErrInstall = "Printer Cannot Be Installed"
[string]$global:msgErrInstall32 = "Printer Cannot Be Installed - 32 Bit OS"
[string]$global:msgErrInvalidTag = "Invalid Printer Tag"
[string]$global:msgSuccess = "Printer Installed Successfully!"


$btnInstall_Click = {
	#Changes cursor to wait on click 
	$frmMain.Cursor = 'WaitCursor'
	
	#Checks for 32 bit OS, if OS is 32 bit, client is directed to call SD, program stops.
	if ((Get-WmiObject win32_operatingsystem | Select-Object osarchitecture).osarchitecture -eq "32-bit")
	{
		[System.Windows.Forms.MessageBox]::Show($global:msgErr32Bit, $global:msgErrInstall32, 'OK', 'Error')
		$frmMain.Cursor = 'Default'
		$txtPrinterTag.Clear()
		$txtPrinterTag.Focus()
		$txtPrinterTag.ReadOnly() #Makes it so the user can't enter in a new tag
	}
	else
	{
		#Gets user input from form
		$global:printer = $txtPrinterTag.Text
		
		#Validates user input
		if (($printer -cmatch '\d{6}' -and $printer.Length -eq 6) -or ($printer -imatch 'loaner\d{2}' -and $printer.Length -eq 8) -or ($printer -imatch 'loaner\d{1}' -and $printer.Length -eq 7))
		{
			#Sets driver type from user input
			If ($radBtnPCL5.Checked)
			{
				$global:type = "PCL5"
			}
			ElseIf ($radBtnPCL6.Checked)
			{
				$global:type = "PCL6"
			}
			ElseIf ($radBtnPS.Checked)
			{
				$global:type = "PS"
			}
			ElseIf ($radBtnHPGL2.Checked)
			{
				$global:type = "HPGL2"
			}
			ElseIf ($radBtnFax.Checked)
			{
				$global:type = "Fax"
			}
			
			#Creates port name based on printer tag #no p for loaners according to Rob			
			if (($global:printer -imatch 'loaner\d{2}' -and $global:printer.Length -eq 8) -or ($global:printer -imatch 'loaner\d{1}' -and $global:printer.Length -eq 7))
			{
				$global:port = $global:printer + ".gos.ca"
			}
			else
			{
				$global:port = $global:printer + "p.gos.ca"
			}
			
			
			#Finds printer in GPIU printers csv based on Printer tag, outputs model
			$filePrinter = Import-Csv $global:printerCSV -Header Printer, Model | Where-Object { $_."Printer" -eq $global:printer }
			$global:model = $filePrinter."Model"
			
			#If the printer model is not found on list, throws an error
			if (!$global:model)
			{
				[System.Windows.Forms.MessageBox]::Show($global:msgErrPrinter, $global:msgErrInvalidTag, 'OK', 'Error')
				$frmMain.Cursor = 'Default'
				$txtPrinterTag.Clear()
				$txtPrinterTag.Focus()
			}
			else
			{
				[string]$printerName = $global:printer + " - " + $global:model + " " + $global:type
				
				#Finds driver from information pulled from printer.csv
				$fileDriver = Import-Csv $global:driver64CSV -Header model, driverType, driver, inf | Where-Object { ($_."model" -eq $global:model) -and ($_."driverType" -eq $global:type) }
				
				$global:driver = $fileDriver."driver"
				$global:inf = $fileDriver."inf"
				
				
				if (!$global:driver)
				{
					[System.Windows.Forms.MessageBox]::Show($global:msgErrDriver, $global:msgErrInstall, 'OK', 'Error')
					$frmMain.Cursor = 'Default'
					$txtPrinterTag.Clear()
					$txtPrinterTag.Focus()
				}
				else
				{
					#Installs driver to driver store
					pnputil.exe /a ("\\*****\gpiu$\bin\drivers\64bit\$global:driver\$global:inf")
					
					#Checks if port exists, does not install if its there
					$portExists = Get-Printerport -Name $global:port -ErrorAction SilentlyContinue
					if (-not $portExists)
					{
						Add-PrinterPort -name $global:port
					}
					
					Add-PrinterDriver -Name $global:driver
					
					#Checks if printer exists, does not install if printer exists 
					$printerExists = Get-Printer -Name $printerName -ErrorAction SilentlyContinue
					if (-not $printerExists)
					{
						Add-Printer -Name $printerName -DriverName $global:driver -PortName $global:port
					}
					
					$default = [System.Windows.Forms.MessageBox]::Show($global:msgPrinterSucess, $global:msgSuccess, 'YesNo', 'None')
					
					if ($default -eq [System.Windows.Forms.DialogResult]::Yes)
					{
						(New-Object -ComObject WScript.Network).SetDefaultPrinter($printerName)
					}
					
					$frmMain.Cursor = 'Default'
					$txtPrinterTag.Clear()
					$txtPrinterTag.Focus()
				}
			}
			
		}
		else
		{
			[System.Windows.Forms.MessageBox]::Show($global:msgErrTag, $global:msgErrInvalidTag, 'OK', 'Error')
			$frmMain.Cursor = 'Default'
			$txtPrinterTag.Clear()
			$txtPrinterTag.Focus()
		}
	}
}