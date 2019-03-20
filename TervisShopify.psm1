#Requires -Modules ShopifyPowerShell,TervisPowershellJobs,TervisPasswordstatePowerShell

Set-TervisShopifyEnvironment

function Set-TervisShopifyCredential {
    $Credential = Get-TervisPasswordstatePassword -Guid "4acc9b2a-080f-4f58-8cbd-843bcbc6d4ab" -AsCredential
    Set-ShopifyCredential -Credential $Credential
}
function Set-TervisShopifyEnvironment {
    Set-TervisShopifyCredential
}

function New-TervisShopifyProduct {
    param (
        [Parameter(Mandatory)]$Title,
        [Parameter(Mandatory)]$Description,
        [Parameter(Mandatory)]$EBSItemNumber,
        [Parameter(Mandatory)]$UPC,
        [Parameter(Mandatory)]$Price,
        $InventoryQuantity = 0,
        $PublishedScope = "global",
        $ShopName = "ospreystoredev"
    )

    $Params = @{
        ShopName = $ShopName
        Title = $Title
        Body_HTML = $Description
        SKU = $EBSItemNumber
        Barcode = $UPC
        Price = $Price
        Inventory_Quantity = $InventoryQuantity
        Published_Scope = $PublishedScope
    }

    New-ShopifyRestProduct @Params
}