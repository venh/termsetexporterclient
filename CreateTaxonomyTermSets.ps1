# =+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+
#   Script: CreateTaxonomyTermSets.ps1
#   Description: Power Shell script to create term sets and terms in SharePoint Taxonomy Term Store using an XML file as input.
#   Author: Venkatesh Subramanian
#   Created Date: 26-March-2013
# =+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+

param (
     [string]$centralAdminUrl = "Enter the Central Admin site URL",
    [string]$termstoreName = "Enter the name of the Managed Metadata Service Instance",
    [string]$filename = "Path to the XML File containing the TermSets and Terms and their GUIDs", 
    [string]$logfile = "Path to the Log File Name (.txt)", 
    [string]$groupName = "Enter the Term Group Name", 
    [string]$owner = "Enter the owner's user ID (for Term Group, Term Sets and Terms) "
)

$ErrorActionPreference = 'Stop'

function LogMessage ([string]$message)
{
    Try
    {
       Write-Host $message
       Out-File -filepath $logfile -InputObject $message -Append
    } 
    catch [System.Exception]
    {
        write-host $_.Exception.ToString() -foregroundcolor Red
        break
    }
}

function GroupExists([Microsoft.SharePoint.Taxonomy.GroupCollection]$grps,[string]$grpName)
{
    Try
    {
        $grpExists = $false
        foreach($tg in $grps){
            if($tg.Name -eq $grpName){
                LogMessage "Group ""$grpName"" exists!" 		
                $grpExists = $true
                break
        }
        }
        Return $grpExists
    }
    catch [System.Exception]
    {
        write-host $_.Exception.ToString() -foregroundcolor Red
        break
    }
}

function TermSetExists([Microsoft.SharePoint.Taxonomy.TermSetCollection]$tSets,[string]$tsName)
{
    Try
    {
        $tsExists = $false
        foreach($ts in $tSets){
            if($ts.Name -eq $tsName){    
                LogMessage "Term Set ""$tsName"" exists!" 	
                $tsExists = $true
                break
        }    
        }
        Return $tsExists
    }
    catch [System.Exception]
    {
        write-host $_.Exception.ToString() -foregroundcolor Red
        break
    }
}

function TermExists([Microsoft.SharePoint.Taxonomy.TermCollection]$termColln,[string]$tName)
{
    Try
    {
        $tExists = $false
        foreach($t in $termColln){
            if($t.Name -eq $tName){       
                LogMessage "Term ""$tName"" exists!" 
                $tExists = $true
                break
        }    
        }
        Return $tExists
    }
    catch [System.Exception]
    {
        write-host $_.Exception.ToString() -foregroundcolor Red
        break
    }
}

function CreateChildTerms([System.Object]$cEles,[Microsoft.SharePoint.Taxonomy.Term]$t,[string]$tNm)
{
    Try
    {
      foreach($cEle in $cEles)
      {
       $cEleName = $cEle.Attribute("Name").Value
       $cEleID = $cEle.Attribute("ID").Value
       $cEleExists = $false
       $cTerms = $t.Terms
       $cEleExists = TermExists $cTerms $tNm
        if($cEleExists -eq $false)
        {
         $cTerm = $t.CreateTerm($cEleName, 1033, [System.Guid]($cEleID))
         $cTerm.Owner = $owner
         $termStore.CommitAll()
         LogMessage "Created Term ""$cEleName"" in Term ""$tNm"""
        }
         $cTermEles = $cEle.Elements() | Where-Object {$_.NodeType -eq [System.Xml.XmlNodeType]::Element -and $_.Name -eq "Term"}    
         if($cTermEles.Count -gt 0)
         {
           CreateChildTerms $cTermEles $cTerm $cEleName
         }
      }
    }
    catch [System.Exception]
    {
        write-host $_.Exception.ToString() -foregroundcolor Red
        break
    }
}

# Load SharePoint PowerShell Snapin
Write-Host " Verifying if SharePoint PowerShell Snapin Loaded.." -ForegroundColor White
$snapin = Get-PSSnapin | Where-Object { $_.Name -eq 'Microsoft.SharePoint.Powershell'}
if($snapin -eq $null){
Write-Host " Loading SharePoint PowerShell Snapin..." -ForegroundColor Gray
Add-PSSnapin "Microsoft.SharePoint.Powershell"
}
Write-Host " Loaded SharePoint PowerShell Snapin..." -ForegroundColor Gray

# Validating Arguments
if (($siteurl -eq "") -or ($filename -eq "")) {
	Write-Host "Site URL and Filename are required" -foregroundcolor Red
	break }

    if ($logfile -eq ""){$logfile="c:\temp\TermSetLog.txt"}

# Read TermSets XML file using XDocument   
[Reflection.Assembly]::LoadWithPartialName("System.Xml.Linq") | Out-Null
$xDoc = [System.Xml.Linq.XDocument]::Load($filename)

Try { 
    #establish session and identify Term Store (same for all groups)
    $site = Get-SPSite $centralAdminUrl
    $session = new-object Microsoft.SharePoint.Taxonomy.TaxonomySession($site)
    $termstore = $session.TermStores[$termstoreName]
    
    get-date | Out-File -filepath $logfile -append
    LogMessage "Importing Data from File: ""$filename"""
    
    $termGroups = $termStore.Groups    
    $groupExists = $false    
    
    # Check if Term Group Exists or Not
   $groupExists = GroupExists $termGroups $groupName
   
    # Create Group and Add Contributor
    if($groupExists -eq $false){
    	$tsGroup = $termStore.CreateGroup($groupName)
        $tsGroup.AddContributor($owner)
        $termStore.CommitAll() 
		LogMessage "Created Group ""$groupName"""
    }
    
    # Create Term Sets
    $tsElements = $xDoc.Root.Descendants() | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element -and $_.Name -eq "TermSet"}
    $termSets = $tsGroup.TermSets
    foreach($tsEle in $tsElements){
        $tsEleName = $tsEle.Attribute("Name").Value
        $tsEleID = $tsEle.Attribute("ID").Value
        $tsExists = $false
        $tsExists = TermSetExists $termSets $tsEleName
        if($tsExists -eq $false)
        {
         $termSet = $tsGroup.CreateTermSet($tsEleName, [System.Guid]($tsEleID), 1033)
         $termSet.Owner = $owner
         $termStore.CommitAll()                  
         LogMessage "Created TermSet ""$tsEleName"" in Group ""$groupName"""
        }
        $termSet = $tsGroup.TermSets[$tsEleName]
        $terms = $termSet.Terms
        $tsElement =  $xDoc.Root.Descendants() | Where-Object {$_.NodeType -eq [System.Xml.XmlNodeType]::Element -and $_.Name -eq "TermSet" -and $_.Attribute("Name").Value -eq $tsEleName}
        $termElements = $tsElement.Elements() | Where-Object {$_.NodeType -eq [System.Xml.XmlNodeType]::Element -and $_.Name -eq "Term"}
        foreach($tEle in $termElements){
        $tEleName = $tEle.Attribute("Name").Value
        $tEleID = $tEle.Attribute("ID").Value
        $tExists = $false
        $tExists = TermExists $terms $tEleName
        if($tExists -eq $false)
        {
         $term = $termSet.CreateTerm($tEleName, 1033, [System.Guid]($tEleID))
         $term.Owner = $owner
         $termStore.CommitAll()
         LogMessage "Created Term ""$tEleName"" in Term Set ""$tsEleName"""
        }
        $childTermEles = $tEle.Elements() | Where-Object {$_.NodeType -eq [System.Xml.XmlNodeType]::Element -and $_.Name -eq "Term"}   
         if($childTermEles.Count -gt 0)
         {
           CreateChildTerms $childTermEles $term $tEleName
         }
        }       
    }
    Write-Host "Completed Successfully.."
   
} catch [System.Exception]{
    write-host $_.Exception.ToString() -foregroundcolor Red
    break
}

$site.dispose()
