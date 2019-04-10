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

function Update-TervisShopifyItemToBePOSReady {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$Product,
        [Parameter(Mandatory)]$ShopName
    )
    begin {
        $Locations = Get-ShopifyRestLocations -ShopName $ShopName
    }
    process {
        $InventoryItemId = $Product.Variants.Inventory_Item_ID
        $ProductVariantId = $Product.Variants.ID

        foreach ($LocationId in $Locations.id) {
            Invoke-ShopifyInventoryActivate -InventoryItemId $InventoryItemId -LocationId $LocationId -ShopName $ShopName | Out-Null
        }
        Set-ShopifyRestProductChannel -ShopName $ShopName -Products $Product -Channel global | Out-Null
        Set-ShopifyProductVariantInventoryPolicy -ProductVariantId $ProductVariantId -InventoryPolicy "CONTINUE" -ShopName $ShopName | Out-Null
    }
}

function Get-TervisShopifyProductInventoryLocations {
    param (
        [Parameter(Mandatory)]$ProductId,
        [Parameter(Mandatory)]$ShopName
    )

    $Query = @"
        query GetProductLocations {
            productVariants(first:1, query:"product_id:$ProductId") {
                edges {
                    node {
                        product {
                            title
                        }
                        inventoryItem {
                            ...locationInfo
                        }
                    }
                }
            }
        }

        fragment locationInfo on InventoryItem {
            inventoryLevels(first: 4) {
                edges {
                    node {
                        location {
                            name
                        }
                    }
                }
            }
        }
"@
    
    $Response = Invoke-ShopifyAPIFunction -ShopName $ShopName -Body $Query
    return [PSCustomObject]@{
        ProductId = $ProductId
        ProductTitle = $Response.data.productVariants.edges.node.product.title
        Locations = $Response.data.productVariants.edges.node.inventoryItem.inventoryLevels.edges.node.location.name
    }
}
