Function Start-IE()
{
    $script:IE = New-Object -ComObject InternetExplorer.Application
    $script:IE.Visible = $True
}

Function Navigate-IE($URL)
{
    Write-Host "Navigating IE to page " -NoNewline
    Write-Host $URL -ForegroundColor Yellow
    $script:IE.Navigate2($URL)
    #Wait-ForPageLoad
}

Function Wait-ForPageLoad($maxPageLoadWatch=20)
{
    $pageLoadWatcher = 0

    do
    {
        Write-Host "Waiting for page load"
        Write-Host "IE Busy state: " -NoNewline
        Write-Host $script:IE.Document.Busy -ForegroundColor Yellow
        Start-Sleep -Seconds 1
        $pageLoadWatcher++
    }
    until($script:IE.Document.Busy -eq $False -or $pageLoadWatcher -gt $maxPageLoadWatch)

    if($pageLoadWatcher -gt $maxPageLoadWatch)
    {
        Write-Host "Page load watcher timed out"
    }
}

Function Get-IEElementByID($ElementID)
{
    $script:IE.Document.GetElementByID($ElementID)
}

Function Get-IEElementByTagName($ElementTagName)
{
    $script:IE.Document.GetElementByID($ElementTagName)
}

Function Fill-TextField($fieldElement, $textValue)
{
    $textField = Get-IEElementByID -ElementID $fieldElement
    $textField.Value = $textvalue
}

Function Click-Button($ButtonElement)
{
    $pageButton = Get-IEElementByID -ElementID $ButtonElement 
    $pageButton.Click()
}

Function Select-DropdownItem($dropwDownID, $selectionValue)
{
    $drowpDown = Get-IEElementByID -ElementID $dropwDownID
    ($dropDown | where {$_.innerHTML -eq "$selectionValue"}).Selected = $true
}


# Start IE session
$browser = Start-IE
