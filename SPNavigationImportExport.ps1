function Import-SPNavigation {
}

function Export-SPNavigation {
    param(
    [Parameter(Mandatory = $true)]
    [String]$SiteUrl,
    [Parameter(Mandatory = $true)]
    [String]$OutputXmlPath
    )
    
    $pubAsm = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Publishing");
    $IsSP2010 = $pubAsm.FullName -match "Version=14"
    
    $site = Get-SPSite -Identity $SiteUrl

    if ($IsSP2010) {
    
        $oSiteNav = Export-SPNavigationStructured -WebUrl $site.RootWeb.Url
    } else {

    }
    $site.Dispose()
    $oSiteNav | Export-Clixml D:\Scripts\PowerShell\G2G.SharePoint.H2O\Navigation.xml
}

function Export-SPNavigationManaged {
}

function Export-SPNavigationStructured {
    param(
    [Parameter(Mandatory = $true)]
    [string]$WebUrl)
    $navObj = New-Object SharePointNavigation
    $web = Get-SPWeb -Identity $WebUrl
    $navObj.SiteUrl = $web.Url
    
    #$web.Navigation.QuickLaunch
    if($web.Webs.Count -gt 1) {
        $navObj.Children = @()
        $web.Webs | % {
            $navObj.Children += Export-SPNavigationStructured -WebUrl $_.Url
        }
    }
    $web.Dispose()
    
    return $navObj
}

if (-not(Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)) {
    Add-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
}

<#
Add-Type -IgnoreWarnings -TypeDefinition @"
using System;
public class NavigationNode
{
    public string Title;
    public string Url;
    public NavigationNode[] Children;
    public bool Hidden;
}

public class SharePointNavigation
{
    public string SiteUrl;
    public NavigationNode[] TopNavigation;
    public NavigationNode[] QuickLaunch;
    public bool UseShared;
    public SharePointNavigation[] Children;
}
"@
#>

function Export-SPNavigationNode {
    $topnavNode = New-Object PSObject
    $topnavNode | Add-Member -MemberType NoteProperty -Name "Title" -Value "Home"
    $topnavNode | Add-Member -MemberType NoteProperty -Name "Url" -Value "/"
    $topnavNode | Add-Member -MemberType NoteProperty -Name "Hidden"-Value $false

    $sub1navNode = New-Object PSObject
    $sub1navNode | Add-Member -MemberType NoteProperty -Name "Title" -Value "Sub One"
    $sub1navNode | Add-Member -MemberType NoteProperty -Name "Url" -Value "/subone"
    $sub1navNode | Add-Member -MemberType NoteProperty -Name "Hidden"-Value $false

    $sub2navNode = New-Object PSObject
    $sub2navNode | Add-Member -MemberType NoteProperty -Name "Title" -Value "Sub Two"
    $sub2navNode | Add-Member -MemberType NoteProperty -Name "Url" -Value "/subtwo"
    $sub2navNode | Add-Member -MemberType NoteProperty -Name "Hidden"-Value $true

    $topnavNode | Add-Member -MemberType NoteProperty -Name "Children" -Value @($sub1navNode, $sub2navNode)

    [xml]$navXml = "<SPNavigation/>"
    $navXml
}

function New-NavigationNode {
    param(
    [Parameter(Mandatory = $true)]
    [string]$Title,
    [Parameter(Mandatory = $true)]
    [string]$Url
    )
    $navigationNode = New-Object PSObject
    $navigationNode | Add-Member -MemberType NoteProperty -Name "Title" -Value $Title
    $navigationNode | Add-Member -MemberType NoteProperty -Name "Url" -Value $Url
    #$navigationNode | Add-Member -MemberType NoteProperty -Name "Children" -Value @()
    return $navigationNode
}

$pubAsm = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Publishing");

Get-SPWeb -Identity "http://qa.livgov.com" | % {
    $web = $_
    $navSettings = New-Object -TypeName Microsoft.SharePoint.Publishing.Navigation.WebNavigationSettings -ArgumentList $web
    $site = $web.Site
    $taxSession = Get-SPTaxonomySession -Site $site
    $termStore = $taxSession.DefaultSiteCollectionTermStore
#    $globalTermSet = $termStore.GetTermSet($navSettings.GlobalNavigation.TermSetId)
    $navGlobalTermSet = [Microsoft.SharePoint.Publishing.Navigation.TaxonomyNavigation]::GetTermSetForWeb($web, "GlobalNavigationTaxonomyProvider", $false)
    # use $editableNavSet = $navGlobalTermSet.GetAsEditable($taxSession) # to get an editable version
}