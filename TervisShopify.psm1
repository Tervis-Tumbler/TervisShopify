#Requires -Modules ShopifyPowerShell,TervisPowershellJobs,TervisPasswordstatePowerShell

Set-TervisShopifyEnvironment

function Set-TervisShopifyEnvironment {
    param (
        [Parameter(Mandatory)]
        [ValidateSet("DEV","PRD")]$Environment
    )
    $GUID = @{
        DEV = "4acc9b2a-080f-4f58-8cbd-843bcbc6d4ab"
        PRD = "a8957c55-9337-4b94-9469-81b06328a9f6"
    }
    $Credential = Get-TervisPasswordstatePassword -Guid $GUID[$Environment] -AsCredential
    Set-ShopifyCredential -Credential $Credential
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

function New-TervisShopifyImage {
    # Does not work currently
    param (
        [Parameter(Mandatory)]$ImageUrl,
        [Parameter(Mandatory)]$ProductId
    )
    
    $Body = @{
        image = @{
            src = $ImageUrl
        }
    } | ConvertTo-Json -Compress

    Invoke-ShopifyRestAPIFunction -HttpMethod POST -ShopName ospreystoredev -Resource Products -Subresource "$ProductId/images" -Body $Body
}
