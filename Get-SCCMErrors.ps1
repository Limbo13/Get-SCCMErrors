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

$NumElapsedHours = 72
$LogFiles = Get-ChildItem C:\Windows\CCM\Logs
$RowNum = 1

foreach ($LogFile in $LogFiles)
{
    $LogFile.Name
    $LogContents = Get-Content $LogFile.FullName
    $LineCount = 0

    foreach ($Line in $LogContents)
    {
        $LineCount++

        if (($Line -like "*Failed*") -or ($Line -like "*Error*"))
        {
            If ($Line.Length -gt 7)
            {
                $TimePresent = ""
                $DashPresent = ""
                $CheckLine = $false
                $TimePresent = $Line.IndexOf("time=")
                $DashPresent = $Line.Substring(4,1)
                #$TimePresent

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
                else
                {
                    #write-output $Line
                }

                If ($CheckLine -eq $true)
                {
                    $PastDate = get-date -Month $Month -Day $Day -Year $Year -Hour $Hour -Minute $Minute -Second $Second
                    $DateDiff = New-TimeSpan -Start $Date -End (Get-Date)

                    If ($DateDiff.TotalHours -lt $NumElapsedHours)
                    {
                        If ($Line -like "*Error*")
                        {
                            #Write-Host $Line -ForegroundColor "red"
                            $serverInfoSheet.Cells.Item($RowNum,1)= $LogFile.Name
                            $serverInfoSheet.Cells.Item($RowNum,2)= $Line
                            $serverInfoSheet.Cells.Item($RowNum,2).font.colorindex = 3
                        }
                        elseif ($Line -like "*Failed*")
                        {
                            #Write-Host $Line -ForegroundColor "Yellow"
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
$workbook.SaveAs("C:\temp\SCCMErrors.xlsx")

#Quit the application
$excel.Quit()

#Release COM Object
[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel) | Out-Null
