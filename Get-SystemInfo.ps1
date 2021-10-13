<#
.SYNOPSIS
  This will quickly grab system information on a system to save time.
.DESCRIPTION
  1. We start by getting the Systems Name, Description, DNS Host Name, Domain, and Hardware configuration. 
  2. We then get the Systems network configuration.
  3. Harddrive space and partitions.
  4. We finish by listing all of the installed programs.
  5. We combine all of the seperate csv files into one master xlsx sheet utilizing the tabs for orginization
  6. Clean up the folder by deleting the seperate csv files leaving just the master xlsx sheet.
  7. Open the master xlsx sheet minimizing muscle fatigue of mouse-clicks.
.PARAMETER <Parameter_Name>
  none
.INPUTS
  none
.OUTPUTS
  Using PowSho environmental variables we create a folder and and a file in "Desktop/audit/ServerAudit.xlsx"
.NOTES
  Version:        1.0
  Author:         Scott B Lichty
  Creation Date:  10/10/2021
  Purpose/Change: Gather quick system information from the system itself
.EXAMPLE
  No Exaples Needed
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
  $dir
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

#Import Modules & Snap-ins

#----------------------------------------------------------[Declarations]----------------------------------------------------------

$var = $env:USERPROFILE + "\Desktop"
$dir = $var + "\logs"
mkdir $dir

#-----------------------------------------------------------[Functions]------------------------------------------------------------

<
Function Get-SystemInfo {
  Param (
    $dir
  )
  Begin {
    Write-Host 'Gathering System Information...'
  }
  Process {
    Try {
      get-wmiobject Win32_ComputerSystem | Select-Object Name, Description, @{Label = "DNS Host Name"; Expression = { $_.DNSHostName } }, Domain, Manufacturer, Model, @{Label = "# Processors"; Expression = { $_.NumberOfProcessors } }, @{Label = "System Type"; Expression = { $_.SystemType } }, @{Label = "Physical Memory"; Expression = { "{0,12:n0} MB" -f ($_.TotalPhysicalMemory / 1mb) } } | Out-File $dir\1_SystemInfo.csv
      ipconfig /all | Out-File $dir\2_Network.csv
      get-wmiobject Win32_LogicalDisk | Select-Object Name, Description, FileSystem, @{Label = "Size"; Expression = { "{0,12:n0} MB" -f ($_.Size / 1mb) } }, @{Label = "Free Space"; Expression = { "{0,12:n0} MB" -f ($_.FreeSpace / 1mb) } }, ProviderName  | Export-Csv $dir\3_HDDConfig.csv -NoTypeInformation
      Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | Export-csv $dir\4_InstalledPrograms.csv -NoTypeInformation
      Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | Export-csv -Append $dir\4_InstalledPrograms.csv -NoTypeInformation

      Function Merge-CSVFiles {
        Param(
          $CSVPath = "$dir", ## Source CSV Folder
          $XLOutput = "$dir\ServerAudit.xlsx" ## Output file name
        )

        Write-Host 'Merging CSV Files...'
        $csvFiles = Get-ChildItem ("$CSVPath\*") -Include *.csv
        $Excel = New-Object -ComObject excel.application 
        $Excel.visible = $false
        $Excel.sheetsInNewWorkbook = $csvFiles.Count
        $workbooks = $excel.Workbooks.Add()
        $CSVSheet = 1

        Foreach ($CSV in $Csvfiles)
        {
          $worksheets = $workbooks.worksheets
          $CSVFullPath = $CSV.FullName
          $SheetName = ($CSV.name -split "\.")[0]
          $worksheet = $worksheets.Item($CSVSheet)
          $worksheet.Name = $SheetName
          $TxtConnector = ("TEXT;" + $CSVFullPath)
          $CellRef = $worksheet.Range("A1")
          $Connector = $worksheet.QueryTables.add($TxtConnector, $CellRef)
          $worksheet.QueryTables.item($Connector.name).TextFileCommaDelimiter = $True
          $worksheet.QueryTables.item($Connector.name).TextFileParseType = 1
          $worksheet.QueryTables.item($Connector.name).Refresh()
          $worksheet.QueryTables.item($Connector.name).delete()
          $worksheet.UsedRange.EntireColumn.AutoFit()
          $CSVSheet++

        }

        $workbooks.SaveAs($XLOutput, 51)
        $workbooks.Saved = $true
        $workbooks.Close()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbooks) | Out-Null
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

      }

      Merge-CSVFiles

      Write-Host 'Cleaning up the audit folder from SPAM csv files...'
      Remove-Item $dir\1_SystemInfo.csv
      Remove-Item $dir\2_Network.csv
      Remove-Item $dir\3_HDDConfig.csv
      Remove-Item $dir\4_InstalledPrograms.csv

      Write-Host 'Opening you master sheet. Enjoy...'
      Invoke-Item $dir\ServerAudit.xlsx
    }
    Catch {
      Write-Host -BackgroundColor Red "Error: $($_.Exception)"
      Break
    }
  }
  End {
    If ($?) {
      Write-Host 'Completed Successfully.'
      Write-Host ' '
    }
  }
}


#-----------------------------------------------------------[Execution]------------------------------------------------------------

Get-SystemInfo