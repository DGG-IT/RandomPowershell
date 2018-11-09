#CONSTANT VARIABLES
$DomainsToSteerPath = ''
$AppsToSteerPath = ''
$CriticalDomainsPath = ''
$prepend = "*."
$append = "/"


#END CONSTANT



#HELPER FUNCTIONS

function Update-DomainsToSteer {
    
param(
        [Parameter(Mandatory=$false,position=0)][Boolean] $NoCritical=$false
    )

    #Load the User-Defined List of Apps to steer
    $saSteeredApps = Import-CSV $AppsToSteerPath

    #Use Helper function to build hashmap of All domains belonging to apps
    $admAppDomainMap = Map-AppsToDomains

    
    #Version the current lists
    if (Test-Path -Path $DomainsToSteerPath) {
        Rename-item $DomainsToSteerPath "$DomainsToSteerPath.$(Get-Date -Format yyyy-MM-dd-HH-mm-ss)"
    }

    if (Test-Path -Path $CriticalDomainsPath) {
        Rename-item $CriticalDomainsPath "$CriticalDomainsPath.$(Get-Date -Format yyyy-MM-dd-HH-mm-ss)"
    }

    foreach ($item in $saSteeredApps) {
        Write-Verbose "using $item"
        
        #Write out just the critical apps to a new file for use in Custom Palo Alto URL Category
        if ($item.Critical -eq "Yes") {
            if ($item.Application -eq "Custom") {
                $prepend, $append -join $item.Domain | Out-File -Append -NoClobber $CriticalDomainsPath
                $item.Domain + "/" | Out-File -Append -NoClobber $CriticalDomainsPath
            }
            else {
                #Add the wildcard *. beginning, and trailing /
                $admAppDomainMap[$item.Application] | % {$prepend, $append -join $_} | Out-File -Append -NoClobber $CriticalDomainsPath
                #Also just output the domain with trailing /
                $admAppDomainMap[$item.Application] | % {$_ + "/"} | Out-File -Append -NoClobber $CriticalDomainsPath
            }
        }
        
        #For each App to steer Show an error if the app is not in Netskope's list, and not a Custom app
        if (($item.Application -ne "Custom") -and !($admAppDomainMap.Contains($item.Application))) {
            Write-Error "Error, App <$($item.Application)> is not in the current sfdr_domains.csv list. Other apps have been matched"
        }

        #Add Domains for Custom Apps as long as NoCritical is false
        elseif ( ($item.Application -eq "Custom") -and !($NoCritical) ) {
            Write-Verbose "Custom Domain will be steered with Domain: $($item.Domain)"
            $item.Domain | Out-File -Append -NoClobber $DomainsToSteerPath
        }

        #Add Domains for Custom Apps as long as NoCritical is false
        elseif ( ($item.Application -eq "Custom") -and ($NoCritical) ) {
            
            if ($item.Critical -eq "Yes") {
                Write-Verbose "Because NoCritical=True, not steering Critical Custom Domain: $($item.Domain)"
            }
            else {
                Write-Verbose "Despite NoCritical=True, still steering non-critical Custom Domain: $($item.Domain)"
                $item.Domain | Out-File -Append -NoClobber $DomainsToSteerPath
            }
        }

        #Domains
        elseif ( ($NoCritical) -and ($item.Critical -eq "Yes") ) {
            
            #Tell us why we're not steering the app
            Write-Verbose ("Because NoCritical=True App <$($item.Application)> will not be steered with Domain(s): " + $($admAppDomainMap[$item.Application]))

        }

        #Domains
        else {
        
            Write-Verbose ("App <$($item.Application)> will be steered with Domain(s):" + $($admAppDomainMap[$item.Application]))

            #Write the domains to a list for future use
            $admAppDomainMap[$item.Application] | Out-File -Append -NoClobber $DomainsToSteerPath
        }
    }
}

#This function is to load the current list of ALL SFDR domains as downloaded from GoSkope GUI
#Eventually This function should be replaced with an API call to get the current SFDR domains, but Netskope has not implemented this API endpoint yet
function Map-DomainsToApps {

    #Get the current sfdr_domains export to keep domains for managed apps up to date
    $dampDomainAppMapPath = ''

    #Create Hash Tables for getting App from Domain and Domain from App
    $damMapDomainAppMap = @{}
    $admAppDomainMap = @{}

    $daDomainApp = import-csv $dampDomainAppMapPath 
    foreach ($item in $daDomainApp) {
        $damDomainAppMap[$item.Domain] = @($item.Application)
    }

    #Both maps will be returned in the below order, cast one to $null if not needed
    return $damDomainAppMap
}

#This function is to load the current list of ALL SFDR domains as downloaded from GoSkope GUI
#Eventually This function should be replaced with an API call to get the current SFDR domains, but Netskope has not implemented this API endpoint yet
function Map-AppsToDomains {

    #Get the current sfdr_domains export to keep domains for managed apps up to date
    $dampDomainAppMapPath = ''

    #Create Hash Tables for getting App from Domain and Domain from App
    $admAppDomainMap = @{}

    $daDomainApp = import-csv $dampDomainAppMapPath 
    foreach ($item in $daDomainApp) {
        $admAppDomainMap[$item.Application] += @($item.Domain)
    }

    #Both maps will be returned in the below order, cast one to $null if not needed
    return $admAppDomainMap
}

#This function is to load the current list of ALL SFDR categories as downloaded from the one-time list generated by Netskope Engineering
#Eventually This function should be replaced with an API call to get the current category mappings, but Netskope has not implemented this API endpoint yet
function Map-AppsToCategories {

    $acmpAppCategoryMapPath = ''
    $acmAppCategoryMap = @{}
    $camCategoryAppMap = @{}

    $acAppCategory = import-csv $acmpAppCategoryMapPath 
    foreach ($item in $acAppCategory) {
        $acmAppCategoryMap[$item.app_name] = @($item.category_name)
    }

    #Both maps will be returned in the below order, cast one to $null if not needed
    return $acmAppCategoryMap
}

#This function is to load the current list of ALL SFDR categories as downloaded from the one-time list generated by Netskope Engineering
#Eventually This function should be replaced with an API call to get the current category mappings, but Netskope has not implemented this API endpoint yet
function Map-CategoriesToApps {

    $acmpAppCategoryMapPath = ''
    $camCategoryAppMap = @{}

    $acAppCategory = import-csv $acmpAppCategoryMapPath
    foreach ($item in $acAppCategory) {
        $camCategoryAppMap[$item.category_name] += @($item.app_name)
    }

    #Both maps will be returned in the below order, cast one to $null if not needed
    return $camCategoryAppMap
}
