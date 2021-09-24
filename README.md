# o365-tenantCASPlans-checkIMAPPOP
Check and Disable IMAP and POP in CAS plans

This script runs across all the domains that you have Delegated Partner access to. 
The purpose is to disable the IMAP and POP settings for all new mailboxes created after the script is run. 

The script assumes you have already opened a powershell prompt and run Connect-MsolService
