<#
#### requires ps-version 5.1 ####

.DESCRIPTION
A small and simple script to check if a GPO is linked to an OU and/or if a GPO is Enabled
Export to CSV
Requirement: RSAT tools installed on de machine

.NOTES
   Version:        0.1
   Author:         Maarten Schmeitz
   Creation Date:  Tuesday, January 14th 2020, 2:52:32 pm
   File: GetGPOUsage.ps1
   Copyright (c) 2020 Advantive

HISTORY:
Date      	          By	Comments
----------	          ---	----------------------------------------------------------
2020-01-14-14-52	 MSCH	Initial Version

.LINK
   www.advantive.nl

.LICENSE
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the Software), to deal
in the Software without restriction, including without limitation the rights
to use copy, modify, merge, publish, distribute sublicense and /or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED AS IS, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
 
#>


Function Get-SaveFile($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.savefiledialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

$savefile =Get-SaveFile

$gpoGuids = Get-GPO -All | Select-Object @{ n='GUID'; e = {$_.Id.ToString()}} | Select-Object -ExpandProperty GUID
$arrAlias = @();
$gpoTotal = $gpoGuids.count
$i=0

foreach ($gpo in $gpoGuids)
{
    Write-Progress -Activity “Working on GPO” -status “Working on GPO $i from $gpoTotal" -percentComplete ($i / $gpoTotal*100)
   $i++
   
   $gpoXML = Get-GPOReport -GUID $($gpo) -ReportType xml
   $results = [xml]$gpoXML
   $totalSOMName = $results.GPO.LinksTo.SOMName.count

   if ( $totalSOMName -gt 1){
       $loop=0
       while ($loop -lt $totalSOMName){
           $loop++
           $objAlias = New-Object psobject   
           Add-Member -InputObject $objAlias -MemberType NoteProperty -Name "Name" -Value $results.gpo.Name 
           Add-Member -InputObject $objAlias -MemberType NoteProperty -Name "SOMName" -Value $results.GPO.LinksTo.SOMName[$loop]
           Add-Member -InputObject $objAlias -MemberType NoteProperty -Name "SOMPATH" -Value $results.GPO.LinksTo.SOMpath[$loop]
           Add-Member -InputObject $objAlias -MemberType NoteProperty -Name "ENABLED" -Value $results.GPO.LinksTo.enabled[$loop]
           $arrAlias += $objAlias
        }
   } else
   {   $objAlias = New-Object psobject
       Add-Member -InputObject $objAlias -MemberType NoteProperty -Name "Name" -Value $results.gpo.Name 
       Add-Member -InputObject $objAlias -MemberType NoteProperty -Name "SOMName" -Value $results.GPO.LinksTo.SOMName
       Add-Member -InputObject $objAlias -MemberType NoteProperty -Name "SOMPATH" -Value $results.GPO.LinksTo.SOMpath
       Add-Member -InputObject $objAlias -MemberType NoteProperty -Name "ENABLED" -Value $results.GPO.LinksTo.enabled
       $arrAlias += $objAlias
       }
   
 
   
}

$arrAlias | Export-Csv $savefile -Delimiter ';' -NoTypeInformation -Encoding UTF8

