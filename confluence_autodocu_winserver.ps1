<#
.SYNOPSIS
  Retrieves information about a Windows server, determines changes and uploads it to Confluence
  
.DESCRIPTION
  * Scripts determines whether the page already exists in Confluence
  * If yes, it will update the page and add labels if they do not already exist
        * Changes are determined by comparing the WikiMarkup code saved in the $WikiMarkupFilePath folder
		* Existing labels will not be removed
  * If no, it will create a new page and add labels
#>

### Parameters ###
param (
    # Network path for WikiMarkup comparison file. Must not end with "\"    
    [string]$WikiMarkupFilePath = "\\server\share",

    # Confluence URI
    [string]$Confluence_URI = "https://confluence.domain.tld",

    # Confluence credentials with write access to space of ParentPageID (can be replaced by static credentials)
    [SecureString]$Confluence_Cred = (Get-Credentials),

    # Page with this ID will be parent of new pages
    [string]$Confluence_ParentPageID = "12345678",

    # Labels added to the page (does not replace existing)
    [array]$Confluence_Labels = @("custom_tag","auto_generated"),

    # Will not check for changes in the documents, upload all Observium data to Confluence and create a new page version
    [bool]$ForceUpload = $false
)

### Creating strings for system information ###
$SVR_Hostname = $env:computername.ToLower()
$SVR_Domain = ([System.Net.Dns]::GetHostByName((hostname)).HostName | out-string) -creplace '^[^\.]*\.', '' -replace "`n|`r"
$SVR_IPv4 = Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object {$_.IPAddress} | Select-Object -Expand IPAddress | Where-Object {($_ -match "(192\.168\.)" -or $_ -match "(10\.)" -or $_ -match "172\.16")} | out-string
$SVR_OSName = (Get-WmiObject Win32_OperatingSystem).Caption -replace 'Microsoft'
$SVR_SysRelease = (Get-WmiObject Win32_OperatingSystem).Version
$SVR_CPU = (Get-WmiObject Win32_processor).NumberOfCores | Measure-Object -sum | Select-Object -ExpandProperty Sum
$SVR_Memory = (Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum | ForEach-Object {"{0:N2}" -f ([math]::round(($_.Sum / 1GB),2))}).ToString() + " GB"
$SRV_Virtual = (Get-WmiObject Win32_ComputerSystem).Manufacturer


### Creating WikiMarkup code for server ###
$WikiMarkupCode = '
<h1>Characteristics</h1>
<ac:structured-macro ac:macro-id="062cfc45-5a03-44f7-8ce7-34caf2588e13" ac:name="details" ac:schema-version="1">
  <ac:rich-text-body>
    <table class="wrapped">
      <colgroup>
        <col/>
        <col/>
      </colgroup>
      <tbody>
        <tr>
          <th>Hostname</th>
          <td>'+$SVR_Hostname+'</td>
        </tr>
        <tr>
          <th>Domain</th>
          <td>'+$SVR_Domain+'</td>
        </tr>
        <tr>
          <th>IPv4</th>
          <td>'+$SVR_IPv4+'</td>
        </tr>
        <tr>
          <th>OS</th>
          <td>'+$SVR_OSName+'</td>
        </tr>
        <tr>
          <th>System Release</th>
          <td>'+$SVR_SysRelease+'</td>
        </tr>
        <tr>
          <th>CPU</th>
          <td>'+$SVR_CPU+'</td>
        </tr>
        <tr>
          <th>Memory</th>
          <td>'+$SVR_Memory+'</td>
        </tr>
        <tr>
          <th>Virtual</th>
          <td>'+$SRV_Virtual+'</td>
        </tr>
      </tbody>
    </table>
  </ac:rich-text-body>
</ac:structured-macro>
'

### Creating WikiMarkup code files and determining changes ###

# Building file paths
$WikiMarkupFileNEW = $WikiMarkupFilePath + "\" + $SVR_Hostname + "\WikiMarkupFileNEW.txt"
$WikiMarkupFileLIVE = $WikiMarkupFilePath + "\" + $SVR_Hostname + "\WikiMarkupFileLIVE.txt"

# Create WikiMarkup files if they do not exist yet
If (-not (Test-Path $WikiMarkupFileNEW))
{
    New-Item $WikiMarkupFileNEW -Type file -Value "foo" -Force
}

If (-not (Test-Path $WikiMarkupFileLIVE))
{
    New-Item $WikiMarkupFileLIVE -Type file -Value "bar" -Force
}

# Pipe out WikiMarkup code to file
$WikiMarkupCode | Out-File $WikiMarkupFileNEW

# Determine changes
$Changes = $false
If (Compare-Object -ReferenceObject (Get-Content $WikiMarkupFileNEW) -DifferenceObject (Get-Content $WikiMarkupFileLIVE))
{
    $Changes = $true
}


### Uploading to Confluence ###

If ($Changes)
{
  # Making connection to Confluence
  Import-Module ConfluencePS
  Set-ConfluenceInfo -BaseURI $Confluence_URI -Credential $Confluence_Cred

  # Querying all Confluence pages that are children of the given parent page
  $ConfluenceChildPages = Get-ConfluenceChildPage -PageID $Confluence_ParentPageID

  # Checking if page already exists and if yes, storing page ID
  $ConfluencePageID = ($ConfluenceChildPages | Where-Object {$_.Title -eq $SVR_Hostname}).ID

  # Creating / updating page
  If ($ConfluencePageID)
  {
    # Updating existing page
    Set-ConfluencePage -PageID $ConfluencePageID -Body $WikiMarkupCode
  }
  else
  {
    # Creating new page and storing page ID
    $ConfluencePageID = (New-ConfluencePage -Title $SVR_Hostname -ParentID $Confluence_ParentPageID -Body $WikiMarkupCode).ID
  }

  If ($ConfluencePageID)
  {
    # Adding labels - if one does not exist it will be added
    $CurrentLabels = (Get-ConfluenceLabel -PageID $ConfluencePageID).Labels.Name
    ForEach ($Label in $Confluence_Labels)
    {
        If ($CurrentLabels -notcontains $Label)
        {
          Add-ConfluenceLabel -PageID $ConfluencePageID -Label $Label
        }
    }

    # Renaming WikiMarkup files
    Remove-Item $WikiMarkupFileLIVE -Force
    Rename-Item $WikiMarkupFileNEW $WikiMarkupFileLIVE -Force
  }
  else
  {
    Write-Error "Upload to Confluence failed"
  }
}
