<#
    .SYNOPSIS
    Retrieves information from block lists and outputs to console or grid

    .DESCRIPTION
    Sends an HTTP get to the specified URI and returns
    data from a block list (.txt) file either as console
    or .NET GridView object

    .PARAMETER -FileWithURIs
    Default parameter that expects a path to a .txt file

    .PARAMETER -URI
    Parameter that expects either an 'http://' or 'https://' prefix
    Cannot be used with -FileWithURIs

    .PARAMETER -ViewAsGrid
    Displays output as a .NET GridView object
    Cannot be used with -FileWithURIs

    .EXAMPLE
    .\Get-URIList.ps1 -FileWithURIs C:\Temp\myfile.txt

    .EXAMPLE
    .\Get-URIList.ps1 -URI "https://www.example.com/somelist.txt"

    .EXAMPLE
    .\Get-URIList.ps1 -URI "https://www.example.com/somelist.txt" -ViewAsGrid


#>


[cmdletBinding(
    DefaultParameterSetName='FileWithURIs'
)]
Param(
    [Parameter(ParameterSetName='FileWithURIs',
               Mandatory=$true
    )]
    [ValidateNotNull()]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
            if(-Not ($_ | Test-Path) ){
                throw "File or folder does not exist"
            }
            if(-Not ($_ | Test-Path -PathType Leaf) ){
                throw "The Path argument must be a file. Folder paths are not allowed."
            }
            if($_ -notmatch "(\.txt)"){
                throw "The file specified in the path argument must be of type txt"
            }
            return $true 
        })]
    [String]$FileWithURIs,

    [Parameter(ParameterSetName='SingleURI',
               Mandatory=$True)]
    [ValidateNotNull()]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
        if($_.StartsWith("http://") -eq $false -and $_.StartsWith("https://" -eq $false))
        {
            throw "User specified URI must start with http:// or https://"
        }
        else
        {
            return $true
        }
    })]
    [String]$URI,
    [Switch]$ViewAsGrid
)

BEGIN
{
    Function Check-CustomType()
    {
        if("TrustAllCertsPolicy" -as [type])
        {
            Out-Null
        }
        else
        {
            Set-CustomType
        }
    }

    Function Set-CustomType()
    {

add-type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) {
        return true;
    }
}
"@
        $script:newCertPolicy = [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
    }

    Function Evaluate-URIs()
    {
        if($URI)
        {
            Get-Blocklist -ListURI $URI
        }
        elseif($FileWithURIs)
        {
            $lines = Get-Content $FileWithURIs
            foreach($line in $lines)
            {
                Get-Blocklist -ListURI $line
            }
        }
    }

    Function Create-Table()
    {
        $script:tabName = "ResultTable"

        # Create Table object
        $script:table = New-Object System.Data.DataTable “$script:tabName”

        # Create first column
        $script:col1 = New-Object System.Data.DataColumn SiteName,([string])
        $script:table.columns.add($script:col1)
    }

    Function Add-RowToTable($Value)
    {
        # Create new row
        $newRow = $script:table.NewRow()

        # Add value to row
        $newRow.SiteName = $Value

        # Add row to table
        $script:table.Rows.Add($newRow)
    }

    Function Get-Blocklist($ListURI)
    {
        try
        {
            $query = Invoke-WebRequest -Uri "$ListURI"

            if($ViewAsGrid)
            {
                Create-Table
                $content = @($query.Content.Split("`r`n"))
                
                foreach($entry in $content)
                {
                    Add-RowToTable -Value $entry
                }

                $script:table | Out-GridView -Title "Blocklist for $ListURI"
            }
            else
            {
                Write-Host "`nBlocklist for $ListURI " -ForegroundColor Yellow -NoNewline
                Write-Host "`n`n$($query.Content | Sort -Descending)"
            }  
        }
        catch [Exception]
        {
           Write-Host "`nUnable to connect to resource " -ForegroundColor Yellow -NoNewline
           Write-Host "$ListURI" -ForegroundColor Red
           Write-Host "`nERROR: $($Error[0].Exception)"
        }
    }

    Function Run-Stack()
    {
        Check-CustomType
        Evaluate-URIs
    }
}

PROCESS
{
    Run-Stack
}

END { Write-Host "`nEnd of script" }
