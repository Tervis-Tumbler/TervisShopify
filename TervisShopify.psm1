#Requires -Modules ShopifyPowerShell,TervisPowershellJobs,TervisPasswordstatePowerShell

# Set-TervisShopifyEnvironment

function Set-TervisShopifyEnvironment {
    param (
        [Parameter(Mandatory)]
        [ValidateSet("Delta","Epsilon","Production")]$Environment
    )
    $GUID = @{
        Delta = "a66d6cd9-a055-46be-ae5b-9e29a6832811"
        Epsilon = "a66d6cd9-a055-46be-ae5b-9e29a6832811"
        Production = "37d9d606-4d1b-49ae-8f89-c0d06c421345"
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

    # $Params = @{
    #     ShopName = $ShopName
    #     Title = $Title
    #     Body_HTML = $Description
    #     SKU = $EBSItemNumber
    #     Barcode = $UPC
    #     Price = $Price
    #     Inventory_Quantity = $InventoryQuantity
    #     Published_Scope = $PublishedScope
    # }

    # New-ShopifyRestProduct @Params

    $Body = [PSCustomObject]@{
        product = @{
            title = $Title
            # body_html = $Body_HTML
            published_scope = $Published_Scope
            variants = @(
                @{
                    price = $Price
                    sku = $SKU
                    barcode = $Barcode
                    inventory_quantity = $Inventory_Quantity
                }
            )
        }
    } | ConvertTo-Json -Compress -Depth 3

    Invoke-ShopifyRestAPIFunction -HttpMethod Post -Resource Products -ShopName $ShopName -Body $Body


}

function Update-TervisShopifyItemToBePOSReady {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$Product,
        [Parameter(Mandatory)]$ShopName,
        $OutputPath
    )
    begin {
        $Locations = Get-ShopifyRestLocations -ShopName $ShopName
        Set-TervisEBSEnvironment -Name Production
        # Logging 
        $InventoryLevel = @()
        $SalesChannel = @()
        $InventoryPolicy = @()
        $Image = @()
    }
    process {
        $InventoryItemId = $Product.Variants.Inventory_Item_ID
        $ProductVariantId = $Product.Variants.ID

        # foreach ($LocationId in $Locations.id) {
        #     $InventoryLevel += Invoke-ShopifyInventoryActivate -InventoryItemId $InventoryItemId -LocationId $LocationId -ShopName $ShopName
        # }
        if ($Product.published_scope -ne "global") {
            $SalesChannel += Set-ShopifyRestProductChannel -ShopName $ShopName -Products $Product -Channel global
        }
        if ($Product.variants.inventory_policy -ne "continue") {
            $InventoryPolicy += Set-ShopifyProductVariantInventoryPolicy -ProductVariantId $ProductVariantId -InventoryPolicy "CONTINUE" -ShopName $ShopName
        }
        if (-not $Product.Image) {
            $Image = $Product | Add-TervisShopifyImageToProduct -ShopName $ShopName
        }
    }
    end {
        if ($OutputPath) {
            $DateStamp = Get-Date -Format "yyyyMMdd-hhmmss"
            $InventoryLevel | ConvertTo-Json -Depth 15 -Compress | Out-File -FilePath "$OutputPath/$DateStamp`_InventoryLevel.json"
            $SalesChannel | ConvertTo-Json -Depth 15 -Compress | Out-File -FilePath "$OutputPath/$DateStamp`_SalesChannel.json"
            $InventoryPolicy | ConvertTo-Json -Depth 15 -Compress | Out-File -FilePath "$OutputPath/$DateStamp`_InventoryPolicy.json"
            $Image | ConvertTo-Json -Depth 15 -Compress | Out-File -FilePath "$OutputPath/$DateStamp`_Image.json"
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

function Add-TervisShopifyImageToProduct {
    param (
        [Parameter(Mandatory)]$ShopName,
        [Parameter(Mandatory,ValueFromPipeline)]$Product
    )
    process {
        $ProductId = $Product.variants.sku
        try {
            $Scene7Url = Invoke-EBSSQL -SQLCommand "SELECT SGURL FROM xxtrvs.XXTRVS_DW_CATALOG_INTF WHERE Product_Id = '$ProductId'" -ErrorAction Stop | Select-Object -ExpandProperty SGURL
        } catch {
            Write-Warning "Could not retrieve image URL from EBS"
        }
        if ($Scene7Url) {            
            $ImageUrl = "https://images.tervis.com/is/image/$Scene7Url"
            New-ShopifyImageByURL -ImageUrl $ImageUrl -ProductId $Product.Id -ShopName $ShopName
        } else {
            # Write-Warning "No image URL for EBS item number $ProductId"
        }
    }
}

function ConvertTo-IndexedArray {
    param (
        [Parameter(Mandatory)]$InputObject,
        [Parameter(Mandatory)]$NumberedPropertyToIndex
    )
    # Get max value to determine size of array
    $MaxValue = $InputObject.$NumberedPropertyToIndex | 
        ForEach-Object {
            try {
                [int]::Parse($_)
            } catch {
                $NonIntWarning = $true
            }
        } |
        Measure-Object -Maximum | 
        Select-Object -ExpandProperty Maximum
    if (($MaxValue + 1) -gt [int]::MaxValue) {
        throw "Largest value is larger than the maximum array length."
    }
    if ($NonIntWarning) {
        Write-Warning "There are values that cannot be converted to an integer."
    }

    # Add values to array by prop's number
    $Array = [System.Object[]]::new($MaxValue + 1)
    $InputObject | ForEach-Object {
        try {
            $Index = [int]::Parse($_.$NumberedPropertyToIndex)
            $Array[$Index] = $_
        } catch {}
    }
    $Array
}

function Get-NonIntegerValues {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$InputObject
    )
    process {
        try {
            [int]::Parse($InputObject) | Out-Null
        }
        catch {
            $InputObject
        }
    }
}

function Find-ObjectsWithNonIntegerValuesInProperty {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$InputObject,
        [Parameter(Mandatory)]$Property
    )
    process {
        try {
            [int]::Parse($InputObject.$Property) | Out-Null
        } catch {
            $InputObject
        }
    }
}

function Get-TervisShopifyPricesFromRMSSQL {
    param (
        [Parameter(Mandatory)]$Server,
        [Parameter(Mandatory)]$Database
    )
    $query = @"
        SELECT * FROM (
            SELECT Alias.Alias AS EBSItemNumber 
            , Item.ItemLookupCode AS ItemUPC
                , Item.Description
                , Item.Price
                , Item.DBTimeStamp
                , ROW_NUMBER() OVER (PARTITION BY Alias.Alias ORDER BY Item.DBTimeStamp DESC) AS RN
            FROM Item, Alias WITH (NOLOCK)
            WHERE Item.ID = Alias.ItemID
            AND Item.Inactive = 0
            AND Item.ItemLookupCode != Alias.Alias
        )  AS WINDOW
        WHERE RN = 1
        ORDER BY EBSItemNumber
"@

    Invoke-MSSQL @PSBoundParameters -SQLCommand $query -ConvertFromDataRow
}

function Get-TervisShopifyDataFromEBS {
    param (
        $Environment = "Production"
    )
    $Query = @"
        SELECT
            items.SEGMENT1 AS ITEM_NUMBER,
            items.INVENTORY_ITEM_ID,
            items.DESCRIPTION,
            xref.CROSS_REFERENCE AS UPC,
            cat.SGURL AS IMG_URL,
            items.LAST_UPDATE_DATE
        FROM mtl_system_items_b items
        LEFT JOIN apps.mtl_cross_references xref ON items.INVENTORY_ITEM_ID = xref.INVENTORY_ITEM_ID
        LEFT JOIN XXTRVS.XXTRVS_DW_CATALOG_INTF cat ON items.SEGMENT1 = cat.PRODUCT_ID
        WHERE xref.CROSS_REFERENCE_TYPE = 'UPC' 
        AND items.ORGANIZATION_ID = 85
        AND items.INVENTORY_ITEM_STATUS_CODE IN ('Active','DTCDeplete')
"@
    Set-TervisEBSEnvironment -Name $Environment
    Invoke-EBSSQL -SQLCommand $Query
}

function New-TervisShopifyInitialUploadData {
    [CmdletBinding()] param ()
    Write-Verbose "Getting SQL data"
    $RMSAccess = Get-TervisPasswordstatePassword -Guid "000108ef-95f8-4232-a62d-97d8c69e0b9f" -AsCredential
    $RMSData = Get-TervisShopifyPricesFromRMSSQL -Server $RMSAccess.Username -Database $RMSAccess.GetNetworkCredential().Password
    $EBSData = Get-TervisShopifyDataFromEBS

    Write-Verbose "Indexing SQL data"
    $IndexedRMSData_1 = ConvertTo-IndexedArray -InputObject $RMSData -NumberedPropertyToIndex "EBSItemNumber"
    $IndexedRMSData_2 = $RMSData | Find-ObjectsWithNonIntegerValuesInProperty -Property EBSItemNumber | ConvertTo-IndexedHashtable -PropertyToIndex EBSItemNumber
    
    Write-Verbose "Generating Shopify data"
    $EBSData | ForEach-Object {
        $Price = if ($IndexedRMSData_1[$_.Item_Number]) {
                $IndexedRMSData_1[$_.Item_Number].Price
            } else {$IndexedRMSData_2["$($_.Item_Number)"].Price}
        [PSCustomObject]@{
            EBSItemNumber = $_.Item_Number
            EBSInventoryItemId = $_.Inventory_Item_Id
            Description = $_.Description
            Price = $Price
            UPC = $_.UPC
            ImageURL = if ($_.Img_Url) {"https://images.tervis.com/is/image/" + $_.Img_Url} else {""}
        }
    }
}

function ConvertTo-ShopifyHandle {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$String
    )
    process {
        $String -replace "[^\w|^\d]+","-"
    }
}

function Export-TervisShopifyCSV {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$ShopifyUploadData,
        [Parameter(Mandatory)]$DirectoryPath,
        $ItemsPerCSV
    )
    begin {
        $CSVData = [System.Collections.ArrayList]::new()
    }
    process {
        $Price = try {
            $ShopifyUploadData.Price.ToString("#.##")
        } catch {"0.00"}
        $CSVData.Add([PSCustomObject]@{
            "Handle" = $ShopifyUploadData.Description | ConvertTo-ShopifyHandle
            "Title" = $ShopifyUploadData.Description
            "Body (HTML)" = $ShopifyUploadData.Description
            "Vendor" = "Tervis"
            "Type" = ""
            "Tags" = ""
            "Published" = "TRUE"
            "Option1 Name" = "Title"
            "Option1 Value" = "Default Title"
            "Option2 Name" = ""
            "Option2 Value" = ""
            "Option3 Name" = ""
            "Option3 Value" = ""
            "Variant SKU" = $ShopifyUploadData.EBSItemNumber
            "Variant Grams" = "0"
            "Variant Inventory Tracker" = "shopify"
            "Variant Inventory Policy" = "continue"
            "Variant Fulfillment Service" = "manual"
            "Variant Price" = $Price
            "Variant Compare At Price" = ""
            "Variant Requires Shipping" = ""
            "Variant Taxable" = "TRUE"
            "Variant Barcode" = $ShopifyUploadData.UPC
            "Image Src" = $ShopifyUploadData.ImageURL
            "Image Position" = ""
            "Image Alt Text" = ""
            "Gift Card" = ""
            "Google Shopping / MPN" = ""
            "Google Shopping / Age Group" = ""
            "Google Shopping / Gender" = ""
            "Google Shopping / Google Product Category" = ""
            "SEO Title" = ""
            "SEO Description" = ""
            "Google Shopping / AdWords Grouping" = ""
            "Google Shopping / AdWords Labels" = ""
            "Google Shopping / Condition" = ""
            "Google Shopping / Custom Product" = ""
            "Google Shopping / Custom Label 0" = ""
            "Google Shopping / Custom Label 1" = ""
            "Google Shopping / Custom Label 2" = ""
            "Google Shopping / Custom Label 3" = ""
            "Google Shopping / Custom Label 4" = ""
            "Variant Image" = ""
            "Variant Weight Unit" = "lb"
            "Cost per item" = ""
        }) | Out-Null
    }
    end {
        if ($ItemsPerCSV) {
            for ($i = 0; $i -lt $CSVData.Count; $i += $ItemsPerCSV) {
                $CSVData | Select-Object -Skip $i -First $ItemsPerCSV | Export-Csv -Path "$DirectoryPath\ShopifyUpload_Start_$i.csv" -Encoding UTF8 -NoTypeInformation -Force
            } 
        } else {
            $CSVData | Export-Csv -Path "$DirectoryPath\ShopifyUpload.csv" -Encoding UTF8 -NoTypeInformation -Force
        }
    }
}

function Invoke-TervisShopifyContinuousUpdate {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]$ShopName
    )

    $Products = Get-ShopifyRestProductsAll -ShopName $ShopName

    do {
        $CurrentCount = $Products.Length
        Write-Verbose "Current count: $CurrentCount"
        $Products | Update-TervisShopifyItemToBePOSReady -ShopName $ShopName
        $Products = Get-ShopifyRestProductsAll -ShopName $ShopName
        $NextRoundCount = $Products.Length
        Write-Verbose "Next round count: $NextRoundCount"
    } while (
        $NextRoundCount -gt $CurrentCount
    )
}

function Get-TervisShopifyLocationDefinition {
    param (
        [Parameter(Mandatory,ParameterSetName="City")]$City,
        [Parameter(Mandatory,ParameterSetName="Name")]$Name
    )
    
    $ModulePath = if ($PSScriptRoot) {
        $PSScriptRoot
    } else {
        (Get-Module -ListAvailable TervisShopify).ModuleBase
    }
    . $ModulePath\LocationDefinition.ps1
    if ($City) {
        $LocationDefinition | Where-Object City -EQ $City
    } elseif ($Name) {
        $LocationDefinition | Where-Object Description -EQ $Name
    }
}

function Get-TervisShopifyOrdersForImport {
    param (
        [Parameter(Mandatory)]$ShopName
    )

    $Orders = Get-ShopifyOrders -ShopName $ShopName -QueryString "NOT tag:ImportedToEBS"
    $Orders | ForEach-Object {
        $LocationDefinition = Get-TervisShopifyLocationDefinition -Name $_.physicalLocation.name
        $OrderId = $_.id | Get-ShopifyIdFromShopifyGid
        $EBSDocumentReference = "$($LocationDefinition.Subinventory)-$OrderId"
        $_ | Add-Member -MemberType NoteProperty -Name EBSDocumentReference -Value $EBSDocumentReference -Force
        $_ | Add-Member -MemberType NoteProperty -Name StoreCustomerNumber -Value $LocationDefinition.CustomerNumber -Force
    }
    return $Orders
}

function Get-TervisShopifyOrdersWithRefundPending {
    param (
        [Parameter(Mandatory)]$ShopName
    )

    Get-ShopifyOrders -ShopName $ShopName -QueryString "tag:RefundPendingUploadToEBS"
}

function Get-TervisShopifyActiveLocations {
    param (
        [Parameter(Mandatory)]$ShopName
    )
    
    $ShopifyLocations = Get-ShopifyLocation -ShopName $ShopName -LocationName * | Where-Object isActive -EQ $true
    $LocationDefinitions = Import-Csv -Path $PSScriptRoot\LocationDefinition.csv
    $ShopifyLocations | foreach {
        $Definition = $LocationDefinitions | Where-Object Description -EQ $_.name
        $_ | Add-Member -MemberType NoteProperty -Force -Name Subinventory -Value $Definition.Subinventory
        $_ | Add-Member -MemberType NoteProperty -Force -Name RMSStoreNumber -Value $Definition.RMSStoreNumber
        $_ | Add-Member -MemberType NoteProperty -Force -Name StoreCustomerNumber -Value $Definition.CustomerNumber
    }
    $ShopifyLocations
}
