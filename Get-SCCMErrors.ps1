<#
    .SYNOPSIS
    This script pulls all warnings and errors produced within the amount of time specified in hours

    .DESCRIPTION
    This script produces an excel spreadsheet with the errors and warnings produced in a time period from now to a time in the past.  Variable takes a number of elapsed hours, folder path and excel spreadsheet name.

    A spreadsheet with all of the warnings and errors is written to the path specified.

    .EXAMPLE
    Get-SCCMErrors -NumElapsedHours 24, -LogFolderPath "c:\temp\" -ExcelFileName "test.xlsx"
#>

Function Get-SCCMErrors()
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [int]$NumElapsedHours,
        [Parameter(Mandatory=$true)]
        [String]$LogFolderPath,
        [Parameter(Mandatory=$true)]
        [String]$ExcelFileName
        )

    $FolderLength = $LogFolderPath.Length
    If ($LogFolderPath.Substring($FolderLength-1,1) -eq "\")
    {
        $LogFilePath = "$LogFolderPath$ExcelFileName"
    }
    else
    {
        $LogFilePath = "$LogFolderPath\$ExcelFileName"
    }

    $TestFullPath = $true
    $TestFolderPath = $false
    $TestExcel = $false

    #Check that the file name ends in xlsx
    $FileLen = $LogFilePath.Length
    $TestExcel = $LogFilePath.Substring($FileLen-4,4) -eq "xlsx"
    If ($TestExcel -eq $false)
    {
        Write-Output "The filename must end with .xlsx"
    }

    #Check the folder path
    $TestFolderPath = Test-Path -Path $LogFolderPath
    If ($TestFolderPath -eq $false)
    {
        Write-Output "The folder doesn't exist"
    }

    #Check the full path to see if it exists
    $TestFullPath = Test-Path -Path $LogFilePath
    If ($TestFullPath -eq $true)
    {
        Write-Output "The file already exists"
    }

    If (($TestFullPath -eq $true) -or ($TestFolderPath -eq $false) -or ($TestExcel -eq $false))
    {
        Break
    }

    #Create excel COM object
    $excel = New-Object -ComObject excel.application

    #Make Visible
    $excel.Visible = $True

    #Add a workbook
    $workbook = $excel.Workbooks.Add()

    #Connect to first worksheet to rename and make active
    $serverInfoSheet = $workbook.Worksheets.Item(1)
    $serverInfoSheet.Name = 'WarningsAndErrors'
    $serverInfoSheet.Activate() | Out-Null

    $LogFiles = Get-ChildItem C:\Windows\CCM\Logs
    $RowNum = 1

    #For each logfile, grab the contents of the file
    foreach ($LogFile in $LogFiles)
    {
        $LogFile.Name
        $LogContents = Get-Content $LogFile.FullName
        $LineCount = 0

        #Reach each line, looking for Failed or Error
        foreach ($Line in $LogContents)
        {
            $LineCount++

            if (($Line -like "*Failed*") -or ($Line -like "*Error*"))
            {
                #Make sure the line is long enough to contain an error
                If ($Line.Length -gt 7)
                {
                    $TimePresent = ""
                    $DashPresent = ""
                    $CheckLine = $false
                    $TimePresent = $Line.IndexOf("time=")
                    $DashPresent = $Line.Substring(4,1)

                    #Check for each style of date that is produced by the log files
                    If ($TimePresent -ne "-1")
                    {
                        $DateStart = $Line.IndexOf("date=")
                        $Time = $Line.Substring($TimePresent+6,8)
                        $Date = $Line.Substring($DateStart+6,10)
                        $Month = $Date.Substring(0,2)
                        $Day = $Date.Substring(3,2)
                        $Year = $Date.Substring(6,4)
                        $Hour = $Time.Substring(0,2)
                        $Minute = $Time.Substring(3,2)
                        $Second = $Time.Substring(6,2)
                        $CheckLine = $true
                    }
                    elseif ($DashPresent -eq "-")
                    {
                        $Date = $Line.Substring(0,10)
                        $Time = $Line.Substring(11,8)
                        $Month = $Date.Substring(5,2)
                        $Day = $Date.Substring(8,2)
                        $Year = $Date.Substring(0,4)
                        $Hour = $Time.Substring(0,2)
                        $Minute = $Time.Substring(3,2)
                        $Second = $Time.Substring(6,2)
                        $CheckLine = $true
                    }

                    #If an error is found, find the difference between now and when the error was logged
                    If ($CheckLine -eq $true)
                    {
                        $PastDate = get-date -Month $Month -Day $Day -Year $Year -Hour $Hour -Minute $Minute -Second $Second
                        $DateDiff = New-TimeSpan -Start $Date -End (Get-Date)

                        #Write the error to the excel spreadsheet in either orange for a warning or red for an error
                        If ($DateDiff.TotalHours -lt $NumElapsedHours)
                        {
                            If ($Line -like "*Error*")
                            {
                                $serverInfoSheet.Cells.Item($RowNum,1)= $LogFile.Name
                                $serverInfoSheet.Cells.Item($RowNum,2)= $Line
                                $serverInfoSheet.Cells.Item($RowNum,2).font.colorindex = 3
                            }
                            elseif ($Line -like "*Failed*")
                            {
                                $serverInfoSheet.Cells.Item($RowNum,1)= $LogFile.Name
                                $serverInfoSheet.Cells.Item($RowNum,2)= $Line
                                $serverInfoSheet.Cells.Item($RowNum,2).font.colorindex = 12
                            }
                            $RowNum++
                        }
                    }
                }
            }
        }
    }

    #Save the file
    $workbook.SaveAs($LogFilePath)

    #Quit the application
    $excel.Quit()

    #Release COM Object
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel) | Out-Null
}
