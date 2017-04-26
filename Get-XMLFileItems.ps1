<#
.Synopsis
    Gets XML file to parse and displays results
.DESCRIPTION
    Reads content of an XML file, parses the
    manifest.projects.project.project_directory
    tag and displays concatenated results of 
    project.project_directory and child file.filename
    items
.EXAMPLE
    \Get-XMLFileItems -XMLFilePath "C:\users\myusername\desktop\myxmlfile.xml"
.EXAMPLE
    .\Get-XMLFileItems -XMLFilePath "C:\users\myusername\desktop\myxmlfile.xml" -MultipleDirectories $False
.NOTES
    See URL: http://stackoverflow.com/questions/43052970/reading-xml-file-manifest-from-powershell/43054489?noredirect=1#comment73316234_43054489
#>

[cmdletBinding()]
Param(
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_})]
    [string]$XMLFilePath,

    [Parameter()]
    [bool]$MultipleDirectories = $true
)

BEGIN
{
    Function Load-XMLFile
    {
        try
        {
            Write-Output "Loading XML file...`n"
            [xml]$XMLDocument = Get-Content $XMLFilePath

            Get-XMLDirectoryInfo -File $XMLDocument

        }
        catch [Exception]
        {
            Write-Error -Message "Error loading XML file $XMLFilePath"
        } 
    }

    Function Get-XMLDirectoryInfo($File)
    {
        try
        {
            $directoryCount = @($File.manifest.projects.project).Count
        }
        catch [Exception]
        {
            Write-Error "Unable to count 'manifest.projects.project' in schema"
        }  

        try
        {
            $directoryNames = $File.manifest.projects.project.project_directory
        }
        catch [Exception]
        {
            Write-Error "Unable to find names for items in 'manifest.projects.project.project_directory' in schema"
        }

        Get-FilesInXMLSchema -File $File
    }

    Function Get-FilesInXMLSchema($File)
    {
        for($i = 0 ; $i -lt $directoryCount; $i++)
        {
            if($MultipleDirectories -eq $true)
            {
                $files = $File.manifest.projects.project[$i].files.file

                foreach($item in $files.filename)
                {
                    [array]$results += $directoryNames[$i] + $item
                }
            }

            else
            {
                $files = $XmlDocument.manifest.projects.project.files.file

                foreach($item in $files.filename)
                {
                    [array]$results += $directoryNames + $item
                } 
            }
        }

        Show-Results -InputObject $results
    }

    Function Show-Results($InputObject)
    {
        try
        {
            return $InputObject
        }
        catch [Exception]
        {
            Write-Error -Message "Error displaying results"
        }
    }

    Function Write-Error($Message)
    {
        Write-Output $Message
        $Error[0]
        Write-Output "Exiting script"
        Exit
    }
}

PROCESS
{
    Load-XMLFile
}

END
{}