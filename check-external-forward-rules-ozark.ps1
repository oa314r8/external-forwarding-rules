#0a314r8
#check
#set up the modules
try {
    if(Get-InstalledModule 'PartnerCenter' -ErrorAction SilentlyContinue){Update-Module 'PartnerCenter'}else{Install-Module 'PartnerCenter' -scope CurrentUser}
    if(Get-InstalledModule 'ExchangeOnlineManagement' -ErrorAction SilentlyContinue){Update-Module 'ExchangeOnlineManagement'}else{Install-Module 'ExchangeOnlineManagement' -scope CurrentUser}
    Import-Module -Name PartnerCenter,ExchangeOnlineManagement
}
catch {
    <#Do this if a terminating exception happens#>
    throw $_
}

#get all tenants listing
Connect-PartnerCenter
$partner_customer=Get-PartnerCustomer

#Need to loop through mailboxes
function GetAllRules{
    param(
        [Object[]]$partner
    )
    Connect-ExchangeOnline -DelegatedOrganization $partner.Domain
    $domains = Get-AcceptedDomain
    $mailboxes = Get-Mailbox -ResultSize Unlimited |?{$_.Name -notlike 'Discover*'}
    
    foreach ($mailbox in $mailboxes) {
        $forwardingRules = $null
        Write-Host "Checking rules for $($mailbox.displayname) - $($mailbox.primarysmtpaddress)" -foregroundColor Green
        $rules = get-inboxrule -Mailbox $mailbox.primarysmtpaddress
        
        $forwardingRules = $rules | Where-Object {$_.forwardto -or $_.forwardasattachmentto}
    
        foreach ($rule in $forwardingRules) {
            $recipients = @()
            $recipients = $rule.ForwardTo | Where-Object {$_ -match "SMTP"}
            $recipients += $rule.ForwardAsAttachmentTo | Where-Object {$_ -match "SMTP"}

            $externalRecipients = @()
            foreach ($recipient in $recipients) {
                $email = ($recipient -split "SMTP:")[1].Trim("]")
                $domain = ($email -split "@")[1]
    
                if ($domains.DomainName -notcontains $domain) {
                    $externalRecipients += $email
                }    
            }
    
            if ($externalRecipients) {
                $extRecString = $externalRecipients -join ", "
                Write-Host "$($rule.Name) forwards to $extRecString" -ForegroundColor Yellow
    
                $ruleHash = $null
                $ruleHash = [ordered]@{
                    PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                    DisplayName        = $mailbox.DisplayName
                    RuleId             = $rule.Identity
                    RuleName           = $rule.Name
                    RuleDescription    = $rule.Description
                    ExternalRecipients = $extRecString
                }
                $ruleObject = New-Object PSObject -Property $ruleHash
                $today=Get-Date -format "yyyy-MM-dd"
                $ruleObject | Export-Csv "C:\temp\externalrules-$today.csv" -NoTypeInformation -Append
            }
        }
    }
    Disconnect-ExchangeOnline -Confirm:$false
    $runanother=Read-Host "Run another Inbox Audit? [y/N])"
    if($runanother.ToLower() -ne "y"){
        Break
    }
}

#dynamic list with prompt to select single tenant?
$selected_customer=""
while (($selected_customer -ne "q") -and ($runanother -ne "y")){
    $partner = @()
    Clear-Host
    For ($i=0; $i -lt $partner_customer.Count; $i++)  {
        Write-Host "$($i): $($partner_customer[$i].Name)"
    }
    [int]$selected_customer = Read-Host "Select Partner number`nor press q to quit "
    $partner = $partner_customer[$selected_customer]
    Clear-Host
    Write-Host $partner.Name "chosen"
    GetAllRules($partner)
    Write-Host "Any external rules identified have been placed in c:\temp\externalrules-$today.csv."
    Write-Host -Foregroundcolor orange "Multiple runs for the same Partner will result in duplicate records."
    

}

