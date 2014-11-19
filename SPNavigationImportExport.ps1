function New-TermsFromNavNodes {
	param(
	[Parameter(Mandatory = $true)]
	[PSObject]$NavigationNodes,
	[Parameter(Mandatory = $true)]
	[Microsoft.SharePoint.Publishing.Navigation.NavigationTerm]$ParentTerm
	)
	$NavigationNodes | % {
		$NavigationNode = $_
		$newTerm = $ParentTerm.CreateTerm($NavigationNode.Title, [Microsoft.SharePoint.Publishing.Navigation.NavigationLinkType]::SimpleLink)
		if($NavigationNode.NavigationNodes.Count -gt 0) {
			New-TermsFromNavNodes -NavigationNodes $NavigationNode.NavigationNodes -ParentTerm $newTerm
		}
	}
}
function New-NavigationTermSet {
	param(
	[Parameter(Mandatory = $true)]
	[PSObject]$NavigationSetNav,
	[Parameter(Mandatory = $true)]
	[string]$WebUrl,
	[Parameter(Mandatory = $true)]
	[string]$TermSetType
	)
	
	$web = Get-SPWeb -Identity $WebUrl
	
	$taxSession = Get-SPTaxonomySession -Site $web.Site
	$termStore = $taxSession.DefaultSiteCollectionTermStore
	$siteCollectionGroup = $termStore.GetSiteCollectionGroup($web.Site)
	
	$termSet = $siteCollectionGroup.CreateTermSet($NavigationSetNav.Title.Value)
	
	if($TermSetType -eq "Current") {
		$navProviderName = [Microsoft.SharePoint.Publishing.Navigation.StandardNavigationProviderNames]::CurrentNavigationTaxonomyProvider
	} else {
		$navProviderName = [Microsoft.SharePoint.Publishing.Navigation.StandardNavigationProviderNames]::GlobalNavigationTaxonomyProvider
	}
	
	$navTermSet = [Microsoft.SharePoint.Publishing.Navigation.NavigationTermSet]::GetAsResolvedByWeb($termSet, $web, $navProviderName)
	$navTermSet.IsNavigationTermSet = $true
	
	$NavigationSetNav.NavigationNodes | % {
		$NavigationNode = $_
		$NavigationTerm = $NavigationTermSet.CreateTerm($NavigationNode.Title, [Microsoft.SharePoint.Publishing.Navigation.NavigationLinkType]::SimpleLink)
		if($NavigationNode.NavigationNodes.Count -gt 0) {
			New-TermsFromNavNodes -NavigationNodes $NavigationNode.NavigationNodes -ParentTerm $NavigationTerm
		}
	}
	
	$termStore.CommitAll()
	
	$web.Dispose()
	
	return $navTermSet
}

function Add-WebNavigation {
	param(
	[Parameter(Mandatory = $true)]
	[string]$SiteUrl,
	[Parameter(Mandatory = $true)]
	[PSObject]$NavigationSet
	)
	
	$s = Get-SPSite -Identity $SiteUrl
	$w = $s.AllWebs | ? { $_.ServerRelativeUrl -eq $NavigationSet.Url }
	
	$navSettings = New-Object -TypeName Microsoft.SharePoint.Publishing.Navigation.WebNavigationSettings -ArgumentList $w
	$taxSession = Get-SPTaxonomySession -Site $s
	# we might need to create a DefaultSiteCollectionTermStore first
	$termStore = $taxSession.DefaultSiteCollectionTermStore
	# we might need toi create SiteCollectionGroup if it doesn't exist
	$termGroup = $termStore.GetSiteCollectionGroup($s)
	
	# reset navigation
	$navSettings.ResetToDefaults();
	if($NavigationSet.GlobalNavigation) {
		# web should use global navigation
		# create a new termstore
		$globalTermSet = New-NavigationTermSet -WebUrl $w.Url -NavigationSetNav $NavigationSet.GlobalNavigation -TermSetType "Global"
		$navSettings.GlobalNavigation.Source = [Microsoft.SharePoint.Publishing.Navigation.StandardNavigationSource]::TaxonomyProvider
		$navSettings.GlobalNavigation.TermStoreId = $termStore.Id
		$navSettings.GlobalNavigation.TermSetId = $globalTermSet.Id
	}
	
	if($NavigationSet.CurrentNavigation) {
		# web should use current navigation
		$currentTermSet = New-NavigationTermSet -WebUrl $w.Url -NavigationSetNav $NavigationSet.CurrentNavigation -TermSetType "Current"
		$navSettings.CurrentNavigation.Source = [Microsoft.SharePoint.Publishing.Navigation.StandardNavigationSource]::TaxonomyProvider
		$navSettings.CurrentNavigation.TermStoreId = $termStore.Id
		$navSettings.CurrentNavigation.TermSetId = $currentTermSet.Id
	}
	
	$navSettings.Update($taxSession)
	[Microsoft.SharePoint.Publishing.Navigation.TaxonomyNavigation]::FlushSiteFromCache($w)
	
	
	
	#$navGlobalTermSet = [Microsoft.SharePoint.Publishing.Navigation.TaxonomyNavigation]::GetTermSetForWeb($w, "GlobalNavigationTaxonomyProvider", $false)
	#$navCurrentTermSet = [Microsoft.SharePoint.Publishing.Navigation.TaxonomyNavigation]::GetTermSetForWeb($w, "CurrentNavigationTaxonomyProvider", $false)	
	# use $editableNavSet = $navGlobalTermSet.GetAsEditable($taxSession) # to get an editable version
	
	$w.Dispose()
	$s.Dispose()
}

function Import-SPNavigation {
	param(
	[Parameter(Mandatory = $true)]
	[string]$SiteUrl,
	[Parameter(Mandatory = $true)]
	[string]$InputXmlPath
	)
	
	$pubAsm = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Publishing");
	
	$NavigationSets = Import-Clixml -LiteralPath $InputXmlPath -Verbose
	$NavigationSets | % {
		$NavigationSet = $_
		Add-WebNavigation -SiteUrl $SiteUrl -NavigationSet $NavigationSet
	}
}

function Export-SPNavigation {
    param(
    [Parameter(Mandatory = $true)]
    [String]$SiteUrl,
    [Parameter(Mandatory = $true)]
    [String]$OutputXmlPath
    )
    
    $pubAsm = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Publishing");

	$NavTermSets = @()

	Get-SPWeb -Site $SiteUrl -Limit All | % {
	    $web = $_
	    $navSettings = New-Object -TypeName Microsoft.SharePoint.Publishing.Navigation.WebNavigationSettings -ArgumentList $web
	    $site = $web.Site
	    $taxSession = Get-SPTaxonomySession -Site $site
	    $termStore = $taxSession.DefaultSiteCollectionTermStore
	#    $globalTermSet = $termStore.GetTermSet($navSettings.GlobalNavigation.TermSetId)
	    $navGlobalTermSet = [Microsoft.SharePoint.Publishing.Navigation.TaxonomyNavigation]::GetTermSetForWeb($web, "GlobalNavigationTaxonomyProvider", $false)
		$navCurrentTermSet = [Microsoft.SharePoint.Publishing.Navigation.TaxonomyNavigation]::GetTermSetForWeb($web, "CurrentNavigationTaxonomyProvider", $false)	
		$globalNav = Get-NavTermSetFromTaxonomy -NavigationTermSet $navGlobalTermSet	
		$currentNav = Get-NavTermSetFromTaxonomy -NavigationTermSet $navCurrentTermSet
		$webNavigation = New-Object PSObject
		$webNavigation | Add-Member -MemberType NoteProperty -Name "Url" -Value $web.ServerRelativeUrl
		$webNavigation | Add-Member -MemberType NoteProperty -Name "GlobalNavigation" -Value $globalNav
		$webNavigation | Add-Member -MemberType NoteProperty -Name "CurrentNavigation" -Value $currentNav
	    # use $editableNavSet = $navGlobalTermSet.GetAsEditable($taxSession) # to get an editable version
		$NavTermSets += $webNavigation	
		#ConvertTo-Xml -InputObject $webNavigation -Depth 8 -As String
	}
	Export-Clixml -Depth 9 -InputObject $NavTermSets -Path $OutputXmlPath
#	ConvertTo-Xml -InputObject $NavTermSets -Depth 8 -As String
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

function Get-NavNodesFromTerms {
	param(
		[Parameter(Mandatory = $true)]
    	$Terms
	)
	
	$nodes = @()
	
	$Terms | % {
		$Term = $_
		$navNode = New-Object PSObject
		$navNode | Add-Member -MemberType NoteProperty -Name "Title" -Value $Term.Title
		$navNode | Add-Member -MemberType NoteProperty -Name "Url" -Value $Term.SimpleLinkUrl
		$navNode | Add-Member -MemberType NoteProperty -Name "NavigationNodes" -Value (Get-NavNodesFromTerms -Terms $Term.Terms)
		$nodes += $navNode
	}
	return $nodes
}

function Get-NavTermSetFromTaxonomy {
	param(
		[Parameter(Mandatory = $false)]
    	[Microsoft.SharePoint.Publishing.Navigation.NavigationTermSet]$NavigationTermSet
	)
	
	if($NavigationTermSet -eq $null) {
		return $null
	} else {
		$navTermSet = New-Object PSObject
		$navTermSet | Add-Member -MemberType NoteProperty -Name "Id" -Value $NavigationTermSet.Id
		$navTermSet | Add-Member -MemberType NoteProperty -Name "Title" -Value $NavigationTermSet.Title
		$navTermSet | Add-Member -MemberType NoteProperty -Name "NavigationNodes" -Value (Get-NavNodesFromTerms -Terms $NavigationTermSet.Terms)
		
		return $navTermSet
	}
}


#Export-SPNavigation -SiteUrl "http://qa.livgov.com" -OutputXmlPath 'D:\Nav Backups\qa.livgov.com.xml'
Import-SPNavigation -SiteUrl "http://qa2.livgov.com" -InputXmlPath 'D:\Nav Backups\qa.livgov.com.xml'
