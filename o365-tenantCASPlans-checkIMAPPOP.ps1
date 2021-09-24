# Script assumes you have already run Connect-MsolService
# Distrubted with the MIT license. 
# Copyright 2021 Dentaur Pty Ltd
$host.ui.RawUI.ForegroundColor = "White" 
Write-Host "Check and disable IMAP and POP Status for new mailboxes "
Write-Host "on all Email enabled domains in all Tenants"
Write-Host "-------------------------------------------------------"
$customers = Get-MsolPartnerContract
$role = Get-MsolRole | Where-Object {$_.name -contains "Company Administrator"}
if (Get-module -Name ExchangeOnlineManagement) {
    #Write-Host "is already imported"
}
else {
    Import-Module ExchangeOnlineManagement 
}
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
[console]::Writeline("{0,-60}{1,-10}{2,-10}","Company Name", "IMAP", "POP")
[console]::Writeline("{0,-60}{1,-10}{2,-10}","------------", "----", "---")
foreach($customer in $customers) {    
    try {
        $company = Get-MsolCompanyInformation -TenantId $customer.tenantid
        $domains = Get-MsolDomain -TenantId $customer.tenantid         
    }
    catch {
        # If the above two cmdlets fail, it is likely because we no longer have delagated access. 
        # We don't want to spoil the output and this can be checked and reported on in other scripts. 
    }
    [console]::Writeline("$($company.displayname)")
    foreach ($domain in $domains) {  
        $host.ui.RawUI.ForegroundColor = "Green"   
        $domainCapabilities = Out-String -InputObject $domain.Capabilities        
        if ($domainCapabilities.Contains("Email")){ 
            try {           
                $output = Connect-ExchangeOnline -UserPrincipalName andrew@dentaur.com -DelegatedOrganization $domain.Name -ShowBanner:$false -ShowProgress:$false -EnableErrorReporting:$false
            }
            catch {
                $host.ui.RawUI.ForegroundColor = "White"                
                [console]::Writeline("{0,-50}{1,-20}",$($domain.Name), "No Exchange Service")
                continue
            }            
            $imap = "Disabled"
            $pop = "Disabled"
            if (Get-CASMailboxPlan -Filter {ImapEnabled -eq "true"}) {
                $host.ui.RawUI.ForegroundColor = "Red"
                $imap = "Enabled"                
            } 
            if (Get-CASMailboxPlan -Filter {PopEnabled -eq "true" }) {
                $host.ui.RawUI.ForegroundColor = "Red"
                $pop = "Enabled"                
            }
            $action = " "             
            if ($imap -eq "Enabled" -Or$imap -eq "Enabled")
            { 
                try {                                       
                    Get-CASMailboxPlan -Filter {ImapEnabled -eq "true" -or PopEnabled -eq "true" } | Set-CASMailboxPlan -ImapEnabled $false -PopEnabled $false            
                    $action = "Fixed"
                }
                catch {
                if (Get-CASMailboxPlan -Filter {ImapEnabled -eq "true"}) {
                    $action = "FAILED to fix"
                }
                if (Get-CASMailboxPlan -Filter {PopEnabled -eq "true" }) {                
                    $action = "FAILED to fix"
                }
                }
            }            
            [console]::Writeline("{0,-60}{1,-10}{2,-10}",$($domain.Name),$imap,"$($pop) $($action)")
            Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue            
        }
    }
    $host.ui.RawUI.ForegroundColor = "White"    
}