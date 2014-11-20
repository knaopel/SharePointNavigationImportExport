#region Import Methods

function New-SimpleLinkFromNavigationNode  {
	param(
	[Parameter(Mandatory = $true)]
	[PSObject]$NavigationNode,
	[Parameter(Mandatory = $true)]
	$ParentObject,
	[Parameter(Mandatory = $false)]
	[Switch]$IsGlobal
)
	
	Write-Host "$($tabs)New-SimpleLinkFromNavigationNode for `"$($NavigationNode.Title)`" to `"$($NavigationNode.Url)`""
	[Microsoft.SharePoint.Publishing.Navigation.NavigationTerm]$childTerm = $ParentObject.CreateTerm($NavigationNode.Title, [Microsoft.SharePoint.Publishing.Navigation.NavigationLinkType]::SimpleLink)
	$childTerm.SimpleLinkUrl = $NavigationNode.Url
	$childTerm.ExcludeFromCurrentNavigation = $IsGlobal
	$childTerm.ExcludeFromGlobalNavigation = (-not $IsGlobal)
	if ($NavigationNode.NavigationNodes.Count -gt 0) {
		$tabs += "`t"
		$NavigationNode.NavigationNodes | % {
			New-SimpleLinkFromNavigationNode -ParentObject $childTerm -NavigationNode $_ -IsGlobal:$IsGlobal
		}
	}
}

function Get-UniqueTermSetName {
	param(
	[Parameter(Mandatory = $true)]
	[string]$NavNodeTitle,
	[Parameter(Mandatory = $true)]
	[Microsoft.SharePoint.Taxonomy.Group]$TaxonomyGroup
	)
	
	$str = "$NavNodeTitle Navigation"
	$termSetName = $str
	$setNames = @()
	$TaxonomyGroup.TermSets | Select Name | % {
		$setNames += $_.Name
	}
	
	for ($i = 2; $i -lt 1000 -and $setNames -contains $termSetName; $i++) {
		$oArray = $str, $i.ToString()
		$termSetName = "{0} {1}" -f $oArray
	}
	
	return $termSetName
}

function New-NavigationTermSet {
	param(
	[Parameter(Mandatory = $true)]
	[PSObject]$NavigationSetNav,
	[Parameter(Mandatory = $true)]
	[string]$WebUrl,
	[Parameter(Mandatory = $true)]
	[string]$TermSetType,
	[Parameter(Mandatory = $false)]
	[Guid]$ExistingTermSetId
	)
	Write-Host "New-NavigationTermSet named $($NavigationSetNav.Title) for $WebUrl"
	$web = Get-SPWeb -Identity $WebUrl
	
	$taxSession = Get-SPTaxonomySession -Site $web.Site
	$termStore = $taxSession.DefaultSiteCollectionTermStore
	$siteCollectionGroup = $termStore.GetSiteCollectionGroup($web.Site)
	
	
	if ($ExistingTermSetId) {
		$termSet = $siteCollectionGroup.TermSets.Item($ExistingTermSetId)
	} else {
		$termSet = $siteCollectionGroup.CreateTermSet((Get-UniqueTermSetName -NavNodeTitle $NavigationSetNav.Title -TaxonomyGroup $siteCollectionGroup))
	}
	
	if($TermSetType -eq "Current") {
		$navProviderName = [Microsoft.SharePoint.Publishing.Navigation.StandardNavigationProviderNames]::CurrentNavigationTaxonomyProvider
	} else {
		$navProviderName = [Microsoft.SharePoint.Publishing.Navigation.StandardNavigationProviderNames]::GlobalNavigationTaxonomyProvider
	}
	
	$navTermSet = [Microsoft.SharePoint.Publishing.Navigation.NavigationTermSet]::GetAsResolvedByWeb($termSet, $web, $navProviderName)
	$editNavTermSet = $navTermSet.GetAsEditable($taxSession)
	$editNavTermSet.IsNavigationTermSet = $true
	
	$termStore.CommitAll() # do we need this one?
	
	if ($NavigationSetNav.NavigationNodes.Count -gt 0) {
		$tabs = "`t"
		$NavigationSetNav.NavigationNodes | % {
			$NavigationNode = $_
			New-SimpleLinkFromNavigationNode -ParentObject $editNavTermSet -NavigationNode $NavigationNode -IsGlobal:($TermSetType -eq "Global")
#			Ensure-RootSimpleLink -WebUrl $web.Url -NavigationNode $NavigationNode -IsCurrent:($TermSetType -eq "Current")
		}
	}
	
	$termStore.CommitAll() # do we need this one?
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
	"Add-WebNavigation for Web: $($NavigationSet.Url) on $SiteUrl"
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
	"`tResetting Navigation settings"
	
	$globalTermSet = $null

	if($NavigationSet.GlobalNavigation) {
		# web should use global navigation
		$globalTermSet = New-NavigationTermSet -WebUrl $w.Url -NavigationSetNav $NavigationSet.GlobalNavigation -TermSetType "Global"
		$navSettings.GlobalNavigation.Source = [Microsoft.SharePoint.Publishing.Navigation.StandardNavigationSource]::TaxonomyProvider
		$navSettings.GlobalNavigation.TermStoreId = $termStore.Id
		$navSettings.GlobalNavigation.TermSetId = $globalTermSet.Id
	}
	
	if($NavigationSet.CurrentNavigation) {
		# web should use current navigation
		if ($globalTermSet) {
			$currentTermSet = New-NavigationTermSet -WebUrl $w.Url -NavigationSetNav $NavigationSet.CurrentNavigation -TermSetType "Current" -ExistingTermSetId $globalTermSet.Id
		} else {
			$currentTermSet = New-NavigationTermSet -WebUrl $w.Url -NavigationSetNav $NavigationSet.CurrentNavigation -TermSetType "Current"
		}
		$navSettings.CurrentNavigation.Source = [Microsoft.SharePoint.Publishing.Navigation.StandardNavigationSource]::TaxonomyProvider
		$navSettings.CurrentNavigation.TermStoreId = $termStore.Id
		$navSettings.CurrentNavigation.TermSetId = $currentTermSet.Id
	}
	$termStore.CommitAll()
	$navSettings.Update($taxSession)
	[Microsoft.SharePoint.Publishing.Navigation.TaxonomyNavigation]::FlushSiteFromCache($w.Site)
	
	$w.Dispose()
	$s.Dispose()
}

function Clear-TermSets {
	param(
	[Parameter(Mandatory = $true)]
	[string]$SiteUrl
	)
	
	Write-Host "Now deleting all term sets from $SiteUrl"
	$tsSite = Get-SPSite -Identity $SiteUrl
	$txSession = Get-SPTaxonomySession -Site $tsSite
	$trmStore = $txSession.DefaultSiteCollectionTermStore
	$trmGroup = $trmStore.GetSiteCollectionGroup($tsSite)
	
	$trmGroup.TermSets | Select Id, Name | % {
		$trmStore.GetTermSet($_.Id).Delete()
	}

	$trmStore.CommitAll()
	
	[Microsoft.SharePoint.Publishing.Navigation.TaxonomyNavigation]::FlushSiteFromCache($tsSite)
	
	$tsSite.Dispose()
}

function Import-SPNavigation {
	param(
	[Parameter(Mandatory = $true)]
	[string]$SiteUrl,
	[Parameter(Mandatory = $true)]
	[string]$InputXmlPath
	)
	
	$pubAsm = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Publishing");
	"Importing Navigation from $InputXmlPath to $SiteUrl."
	
	$NavigationSets = Import-Clixml -LiteralPath $InputXmlPath -Verbose
	if ($NavigationSets.Count -gt 0) {
		Clear-TermSets -SiteUrl $SiteUrl
		$NavigationSets | % {
			$NavigationSet = $_
			Add-WebNavigation -SiteUrl $SiteUrl -NavigationSet $NavigationSet
		}
	}
}
#endregion

#region Export Methods
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

function Get-NavNodesFromTerms {
	param(
		[Parameter(Mandatory = $true)]
    	$Terms
	)
	
	$nodes = @()
	
	$Terms | % {
		$Term = $_
		$navNode = New-Object PSObject
		$navNode | Add-Member -MemberType NoteProperty -Name "Title" -Value $Term.Title.ToString()
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
		$navTermSet | Add-Member -MemberType NoteProperty -Name "Title" -Value $NavigationTermSet.Title.ToString()
		$navTermSet | Add-Member -MemberType NoteProperty -Name "NavigationNodes" -Value (Get-NavNodesFromTerms -Terms $NavigationTermSet.Terms)
		
		return $navTermSet
	}
}

#endregion