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
function New-WebNavigation {
    param(
    [Parameter(Mandatory = $true)]
    [string]$Url,
    [Parameter(Mandatory = $false)]
    [PSObject]$GlobalNavigation,
    [Parameter(Mandatory = $false)]
    [PSObject]$CurrentNavigation
    )
    
    $webNavigation = New-Object PSObject
    $webNavigation | Add-Member -MemberType NoteProperty -Name "Url" -Value $Url
    $webNavigation | Add-Member -MemberType NoteProperty -Name "GlobalNavigation" -Value $GlobalNavigation
    $webNavigation | Add-Member -MemberType NoteProperty -Name "CurrentNavigation" -Value $CurrentNavigation
    
    return $webNavigation
}

function New-NavigationSet {
	param(
    [Parameter(Mandatory = $true)]
    [guid]$Id,
    [Parameter(Mandatory = $true)]
    [String]$Title
    )
    
	$navSet = New-Object PSObject
	$navSet | Add-Member -MemberType NoteProperty -Name "Id" -Value $Id
	$navSet | Add-Member -MemberType NoteProperty -Name "Title" -Value $Title
	$navSet | Add-Member -MemberType NoteProperty -Name "NavigationNodes" -Value @()
	
	return $navSet
}

function New-NavigationNode {
	param(
    [Parameter(Mandatory = $true)]
    [string]$Title,
    [Parameter(Mandatory = $false)]
    [string]$Url
	)
	
	$navNode = New-Object PSObject
	$navNode | Add-Member -MemberType NoteProperty -Name "Title" -Value $Title
	$navNode | Add-Member -MemberType NoteProperty -Name "Url" -Value $Url
	$navNode | Add-Member -MemberType NoteProperty -Name "NavigationNodes" -Value @()
	
	return $navNode
}

function Get-NodeVisiblity {
	param(
    [Parameter(Mandatory = $true)]
    [Microsoft.SharePoint.Publishing.PublishingWeb]$PublishingWeb,
    [Parameter(Mandatory = $true)]
    [Microsoft.SharePoint.Navigation.SPNavigationNode]$SPNavNode,
    [Parameter()]
    [Switch]$IsGlobal
    )
	
	$IsVisible = $true

	if ($IsGlobal) {
		$IncludeSubSites = $PublishingWeb.Navigation.GlobalIncludeSubSites
		$IncludePages = $PublishingWeb.Navigation.GlobalIncludePages
	} else {
		$IncludeSubSites = $PublishingWeb.Navigation.CurrentIncludeSubSites
		$IncludePages = $PublishingWeb.Navigation.CurrentIncludePages
	}
	
	if ($SPNavNode.Properties["NodeType"] -ne $null -and (-not [string]::IsNullOrEmpty($SPNavNode.Properties["NodeType"].ToString()))) {
    	$type = [Microsoft.SharePoint.Publishing.NodeTypes]$SPNavNode.Properties["NodeType"]
    }
	
	switch ($type) {
		"Area" {
			if ($IncludeSubSites) {
				# check the sub site and see if it is included in navigation
				try {
					$name = $SPNavNode.Url.Trim('/')
					if ($name.Length -ne 0 -and $name.IndexOf("/") -gt 0) {
						$name = $name.Substring($name.LastIndexOf("/") + 1)
					}
				
					try {
						$web = $PublishingWeb.Web.Webs[$name]
					} catch [System.ArgumentException] {
					}
				
					if ($web -ne $null -and $web.Exists -and $web.ServerRelativeUrl.ToLower() -eq $SPNavNode.Url.ToLower() -and [Microsoft.SharePoint.Publishing.PublishingWeb]::IsPublishingWeb($web)) {
						$tempPubWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($web)
						if (-not ($IsGlobal -and $tempPubWeb.IncludeInGlobalNavigation) -or $tempPubWeb.IncludeInCurrentNavigation) {
							$IsVisible = $false
						}
					}
				} finally {
					if ($web -ne $null) {
						$web.Dispose()
					}
				}
			} else {
				# don't show sub sites
				$IsVisible = $IncludeSubSites
			}
		}
		"Page" {
			if ($IncludePages) {
				#check the page to see if it is included in navigation
				try {
					$page = $PublishingWeb.GetPublishingPages()[$SPNavNode.Url]
				} catch [System.ArgumentException] {
				}
			
				if ($page -ne $null) {
					if (-not ($IsGlobal -and $page.IncludeInGlobalNavigation) -or $page.IncludeInCurrentNavigation) {
						$IsVisible = $false
					}
				}
			} else {
				# this page is not shown
				$IsVisible = $IncludePages
			}
		}
		default {
		}
	}
	return $IsVisible
}

function Enumerate-Collection {
    param(
    [Parameter(Mandatory = $true)]
    [Microsoft.SharePoint.Publishing.PublishingWeb]$PublishingWeb,
    [Parameter(Mandatory = $true)]
    [Microsoft.SharePoint.Navigation.SPNavigationNodeCollection]$NavNodes,
    [Parameter(Mandatory = $true)]
    [PSObject]$ParentNode,
    [Parameter()]
    [Switch]$IsGlobal
    )
    
    if ($NavNodes -eq $null -or $NavNodes.Count -eq 0) {
        return
    }
    
    $NavNodes | % {
		if ((Get-NodeVisiblity -SPNavNode $_ -PublishingWeb $PublishingWeb -IsGlobal:$IsGlobal)) {
			$childNode = New-NavigationNode -Title $_.Title -Url $_.Url
			$ParentNode.NavigationNodes += $childNode
			Write-Host "$tabs`"$($_.Title)`" - Url: $($_.Url)"
			if ($_.Children.Count -gt 0) {	
				Enumerate-Collection -PublishingWeb $PublishingWeb -NavNodes $_.Children -ParentNode $childNode
			}
		}
    }
}

function Get-NavNodesFromPortal {
	param(
	[Parameter(Mandatory = $true)]
	[Microsoft.SharePoint.Navigation.SPNavigationNodeCollection]$NavigationNodes,
	[Parameter()]
	[switch]$IsGlobal
	)
	
	$nodes = @()
	$w = $NavigationNodes.Navigation.Web
	if ([Microsoft.SharePoint.Publishing.PublishingWeb]::IsPublishingWeb($w)) {
		$pubWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($w)
		$NavigationNodes | % {
			if ((Get-NodeVisiblity -SPNavNode $_ -PublishingWeb $pubWeb -IsGlobal:$IsGlobal)) {
				$navNode = New-NavigationNode -Title $_.Title -Url $_.Url
				$navNode.NavigationNodes = Get-NavNodesFromPortal -NavigationNodes $_.Children
				$nodes += $navNode
			}
		}
	}
	
	$w.Dispose()
	
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

function Get-NavTermSetFromPortal {
	param(
	[Parameter()]
	[Microsoft.SharePoint.Publishing.PublishingWeb]$PublishingWeb,
	[Parameter(Mandatory = $true)]
	[Microsoft.SharePoint.Navigation.SPNavigationNodeCollection]$NavigationNodes,
	[Parameter()]
	[Switch]$IsGlobal
	)
	
#	if ($IsGlobal) {
#		$NavigationNodes = $PublishingWeb.Navigation.GlobalNavigationNodes
#	} else {
#		$NavigationNodes = $PublishingWeb.Navigation.CurrentNavigationNodes
#	}
	
	$navSet = New-NavigationSet -Id ([Guid]::NewGuid()) -Title $NavigationNodes.Navigation.Web.Title
	$navSet.NavigationNodes = Get-NavNodesFromPortal -NavigationNodes $NavigationNodes -IsGlobal:$IsGlobal
#	Enumerate-Collection -PublishingWeb $PublishingWeb -NavNodes $NavigationNodes -ParentNode $navSet.NavigationNodes
	return $navSet
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

    if ($pubAsm.FullName -match "14") {
		Get-SPWeb -Site $SiteUrl -Limit ALL | % {
			if ([Microsoft.SharePoint.Publishing.PublishingWeb]::IsPublishingWeb($_)) {
				$pubWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($_)
				$globalNav = $null
				$currentNav = $null
				if (-not $pubWeb.Navigation.InheritGlobal) {
					$globalNav = Get-NavTermSetFromPortal -PublishingWeb $pubWeb -NavigationNodes $pubWeb.Navigation.GlobalNavigationNodes -IsGlobal
				}
				
				if (-not $pubWeb.Navigation.InheritCurrent) {
					$currentNav = Get-NavTermSetFromPortal -PublishingWeb $pubWeb -NavigationNodes $pubWeb.Navigation.CurrentNavigationNodes
				}
				$webNavigation = New-WebNavigation -Url $_.ServerRelativeUrl -GlobalNavigation $globalNav -CurrentNavigation $currentNav
				$NavTermSets += $webNavigation
			}
		}
    } else {
		Get-SPWeb -Site $SiteUrl -Limit All | % {
    	    $web = $_
    	    $navSettings = New-Object -TypeName Microsoft.SharePoint.Publishing.Navigation.WebNavigationSettings -ArgumentList $web
    	    $site = $web.Site
    	    $taxSession = Get-SPTaxonomySession -Site $site
    	    $termStore = $taxSession.DefaultSiteCollectionTermStore
    	
    	    $navGlobalTermSet = [Microsoft.SharePoint.Publishing.Navigation.TaxonomyNavigation]::GetTermSetForWeb($web, "GlobalNavigationTaxonomyProvider", $false)
    		$navCurrentTermSet = [Microsoft.SharePoint.Publishing.Navigation.TaxonomyNavigation]::GetTermSetForWeb($web, "CurrentNavigationTaxonomyProvider", $false)	
    		$globalNav = Get-NavTermSetFromTaxonomy -NavigationTermSet $navGlobalTermSet	
    		$currentNav = Get-NavTermSetFromTaxonomy -NavigationTermSet $navCurrentTermSet
    		$webNavigation = New-Object PSObject
    		$webNavigation | Add-Member -MemberType NoteProperty -Name "Url" -Value $web.ServerRelativeUrl
    		$webNavigation | Add-Member -MemberType NoteProperty -Name "GlobalNavigation" -Value $globalNav
    		$webNavigation | Add-Member -MemberType NoteProperty -Name "CurrentNavigation" -Value $currentNav
    	
    		$NavTermSets += $webNavigation
    	}
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

#endregion

Export-SPNavigation -SiteUrl "http://www.oakgov.com" -OutputXmlPath D:\NavBackups\www.oakgov.com.xml
#Import-SPNavigation -SiteUrl "http://qa2.livgov.com" -InputXmlPath 'D:\Nav Backups\qa.livgov.com.xml'
#Create-SimpleLink -WebUrl http://qa2.livgov.com -NavigationName "My Link" -NavigationUrl "/Pages/default.aspx"
#$tabs = ""
#$w = Get-SPWeb -Identity http://www.oakgov.com/aviation
#$pw = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($w)
#$pn = New-NavigationNode -Title "Parent Node" -Url ""
#Enumerate-Collection -PublishingWeb $pw -NavNodes $pw.Navigation.CurrentNavigationNodes -ParentNode $pn
#$w.Dispose()