########################################################################################
#
# Site policy unlock and delete for migrated sites
# 
# 1. Ensure Username below has access via group membership to MSOTenantContext 
#    by submitting CR [o365admin/changerequest.aspx?id=SPOD-13-215]
# 2. Set variables below for your environment
# 3. Dot source this script (. .\MigratedSiteRemoval.ps1), enter password
# 4. To delete or unlock a site see examples
# 
# Examples:
# 
# UnlockSitePolicy http://siteurl.com
# 
# DeleteMigratedSite http://siteurl.com
# 
########################################################################################

# Set these variables for your environment
# User must have access to MSOTenantContext site
$Username = "Domain\username"
# MSOTenantContext site for D
$MSOUrl =  "https://www.site.com/MSOTenantContext" 


########################################################################################
# Setup Types
########################################################################################
try { 
    Add-Type -Path "Microsoft.Online.SharePoint.Client.Tenant.dll"
    Add-Type -Path "Microsoft.SharePoint.Client.dll"
    Add-Type -Path "Microsoft.SharePoint.Client.Runtime.dll"

    # Custom version of JDP commands with additional feature to open site policy (OpenSiteClosedByPolicy)
    Add-Type -path "JDP.Transformation.HttpCommands.dll"
}
catch {
    Write-Error "Failed to load required libraries, check that all files are copied." 
    Write-Error $_.Exception.Message
}

########################################################################################
# Setup Security Context
########################################################################################
try { 

    $Password = Read-Host -Prompt "Enter your password: " -AsSecureString
    $Creds = New-Object System.Net.NetworkCredential($UserName, $Password)
    $Creds.UserName = $Username.Split('\')[1]
    $Creds.Domain = $Username.Split('\')[0]
    $Global:Context = New-Object Microsoft.SharePoint.Client.ClientContext($MSOUrl)
    $context.Credentials = $Creds
}
catch {
    Write-Error "Failed to generate client context." 
    Write-Error $_.Exception.Message 
}


########################################################################################
# Unlock site that was locked using a policy
########################################################################################
Function UnlockSitePolicy($url) {
    Write-Output "Unlocking site: $url"
    try  {
        # pass the site to the custom call and execute
        # set auth by navigating to site
        $tempRequest = Invoke-WebRequest $url
        $openSite = New-Object JDP.Transformation.HttpCommands.OpenSiteClosedByPolicy -argumentlist $url
        $openSite.Execute()
        Write-Output "Site unlocked: $url"
    }
    catch {
        Write-Error "Failed to unlock site: $url" 
        Write-Error $_.Exception.Message
    }
}


########################################################################################
# Delete site 
########################################################################################
Function DeleteMigratedSite($url) {
    try { 
    Write-Output "Removing site: $url"
    # Remove the site
    $tenant = New-Object Microsoft.Online.SharePoint.TenantAdministration.Tenant($context)
    $Removal = $tenant.RemoveSite($url)
    $context.Load($Removal)
    $context.ExecuteQuery()

    }
    catch {
      Write-Error "Failed to delete site: $url" 
      Write-Error $_.Exception.Message
    }
}


