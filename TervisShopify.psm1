#Requires -Modules ShopifyPowerShell,TervisPowershellJobs,TervisPasswordstatePowerShell

function Set-TervisShopifyCredential {
    $Credential = Get-TervisPasswordstatePassword -Guid "4acc9b2a-080f-4f58-8cbd-843bcbc6d4ab" -AsCredential
    Set-ShopifyCredential -Credential $Credential
}
function Set-TervisShopifyEnvironment {
    Set-TervisShopifyCredential
}