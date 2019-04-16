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
        [Parameter(Mandatory)]$ProductId,
        [Parameter(Mandatory)]$ShopName
    )
    
    $Body = @{
        image = @{
            src = $ImageUrl
        }
    } | ConvertTo-Json -Compress

    Invoke-ShopifyRestAPIFunction -HttpMethod POST -ShopName $ShopName -Resource Products -Subresource "$ProductId/images" -Body $Body
}

function Update-TervisShopifyItemToBePOSReady {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$Product,
        [Parameter(Mandatory)]$ShopName,
        $OutputPath
    )
    begin {
        $Locations = Get-ShopifyRestLocations -ShopName $ShopName

        # Logging 
        $InventoryLevel = @()
        $SalesChannel = @()
        $InventoryPolicy = @()
    }
    process {
        $InventoryItemId = $Product.Variants.Inventory_Item_ID
        $ProductVariantId = $Product.Variants.ID

        foreach ($LocationId in $Locations.id) {
            $InventoryLevel += Invoke-ShopifyInventoryActivate -InventoryItemId $InventoryItemId -LocationId $LocationId -ShopName $ShopName
        }
        $SalesChannel += Set-ShopifyRestProductChannel -ShopName $ShopName -Products $Product -Channel global
        $InventoryPolicy += Set-ShopifyProductVariantInventoryPolicy -ProductVariantId $ProductVariantId -InventoryPolicy "CONTINUE" -ShopName $ShopName
    }
    end {
        if ($OutputPath) {
            $DateStamp = Get-Date -Format "yyyyMMdd-hhmmss"
            $InventoryLevel | ConvertTo-Json -Depth 15 -Compress | Out-File -FilePath "$OutputPath/$DateStamp`_InventoryLevel.json"
            $SalesChannel | ConvertTo-Json -Depth 15 -Compress | Out-File -FilePath "$OutputPath/$DateStamp`_SalesChannel.json"
            $InventoryPolicy | ConvertTo-Json -Depth 15 -Compress | Out-File -FilePath "$OutputPath/$DateStamp`_InventoryPolicy.json"
        }
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

function Get-TervisShopifyProductsAtLocation {
    param (
        [Parameter(Mandatory)]$LocationName,
        [Parameter(Mandatory)]$ShopName
    )
    $Products = @()
    $CurrentCursor = ""
    
    do {
        $Query = @"
            query LocationStuff {
                locations(first: 1, query: "name:$LocationName") {
                    edges {
                        node {
                            inventoryLevels(first: 245 $(if ($CurrentCursor) {", after:`"$CurrentCursor`""} )) {
                                edges {
                                    node {
                                        item {
                                            variant {
                                                product {
                                                    title
                                                }
                                            }
                                        }
                                    }
                                    cursor
                                }
                                pageInfo {
                                    hasNextPage
                                }
                            }
                        }
                    }
                }
            }    
"@   
        $Response = Invoke-ShopifyAPIFunction -ShopName $ShopName -Body $Query
        $CurrentCursor = $Response.data.locations.edges.node.inventoryLevels.edges | Select-Object -Last 1 -ExpandProperty cursor
        $Products += $Response.data.locations.edges.node.inventoryLevels.edges | ForEach-Object {$_.node.item.variant.product.title}
    } while ($Response.data.locations.edges.node.inventoryLevels.pageInfo.hasNextPage)
    return $Products
}
