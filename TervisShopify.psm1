#Requires -Modules ShopifyPowerShell,TervisPowershellJobs,TervisPasswordstatePowerShell

Set-GetShopifyCredentialScriptBlock -ScriptBlock {
    Get-PasswordstatePassword -AsCredential -ID 5729
}