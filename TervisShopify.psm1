#Requires -Modules ShopifyPowerShell,TervisPowershellJobs,TervisPasswordstatePowerShell

function Set-TervisShopifyCredential {
    $Credential = Get-PasswordstatePassword -ID 5729 | Select-Object UserName, Password
    Set-ShopifyCredential -Credential $Credential
}
function Set-TervisShopifyEnvironment {
    Set-TervisShopifyCredential
}