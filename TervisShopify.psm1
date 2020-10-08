#Requires -Modules ShopifyPowerShell,TervisPowershellJobs,TervisPasswordstatePowerShell

function Invoke-TervisShopifyModuleImport {
    param (
        [ValidateSet("Delta","Epsilon","Production")]$Environment = "Delta"
    )
    Import-Module -Global -Force -Name ShopifyPowerShell,TervisShopify,TervisShopifyPowerShellApplication
    Set-TervisShopifyEnvironment -Environment $Environment
    Set-TervisEBSEnvironment -Name $Environment
}

function Set-TervisShopifyEnvironment {
    param (
        [Parameter(Mandatory)]
        [ValidateSet("Delta","Epsilon","Production")]$Environment
    )
    $GUID = @{
        Delta = "a66d6cd9-a055-46be-ae5b-9e29a6832811"
        Epsilon = "c1ad053e-6f3e-410d-81f6-b2754b974db4"
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
        [Parameter(ParameterSetName="City")]$City,
        [Parameter(ParameterSetName="Name")]$Name
    )
    
    $ModulePath = if ($PSScriptRoot) {
        $PSScriptRoot
    } else {
        (Get-Module -ListAvailable TervisShopify).ModuleBase
    }
    $LocationDefinition = Import-Csv -Path $ModulePath\LocationDefinition.csv
    if ($City) {
        $LocationDefinition | Where-Object City -EQ $City
    } elseif ($Name) {
        $LocationDefinition | Where-Object Description -EQ $Name
    } else {
        $Online = $LocationDefinition | Where-Object Subinventory -EQ "FL1"
        $Online.Description = "Online"
        $Online
    }
}

function Get-TervisShopifyOrdersForImport {
    param (
        [Parameter(Mandatory)]$ShopName,
        [Parameter(ValueFromPipeline)]$Orders
    )
    if (-not $Orders) {
        $Orders = Get-ShopifyOrders -ShopName $ShopName -QueryString "NOT tag:ImportedToEBS NOT tag:IgnoreImport" #Omit exchanges
    }
    $Orders | ForEach-Object {
        $LocationDefinition = Get-TervisShopifyLocationDefinition -Name $_.physicalLocation.name
        $IsOnlineOrder = if (-not $_.physicalLocation) { $true } else { $false }
        $OrderId = $_.id | Get-ShopifyIdFromShopifyGid
        $EBSDocumentReference = "$($LocationDefinition.Subinventory)-$OrderId"
        $CustomAttributes = $_ | Convert-TervisShopifyCustomAttributesToObject
        $_ | Add-Member -MemberType NoteProperty -Name EBSDocumentReference -Value $EBSDocumentReference -Force
        $_ | Add-Member -MemberType NoteProperty -Name StoreCustomerNumber -Value $LocationDefinition.CustomerNumber -Force
        $_ | Add-Member -MemberType NoteProperty -Name Subinventory -Value $LocationDefinition.Subinventory -Force
        $_ | Add-Member -MemberType NoteProperty -Name ReceiptMethodId -Value $LocationDefinition.ReceiptMethodId -Force
        $_ | Add-Member -MemberType NoteProperty -Name CustomAttributes -Value $CustomAttributes -Force
        $_ | Add-Member -MemberType NoteProperty -Name IsOnlineOrder -Value $IsOnlineOrder -Force
        $_ | Select-TervisShopifyOrderPersonalizationLines | Add-TervisShopifyOrderPersonalizationSKU
        $_ | Set-TervisShopifyOrderPersonalizedItemNumber 
        $_ | Add-TervisShopifyCartDiscountAsLineItem
    }
    return $Orders
}

function Add-TervisShopifyCartDiscountAsLineItem {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$Order
    )
    process {
        if (-not $Order.cartDiscountAmountSet) { return }
        $DiscountSku = "1373422"
        # $DiscountSku = "1371644" # Uncomment for SIT until next refresh after 07-2020
        $DiscountName = $Order.discountCode
        $DiscountAmount = [decimal]$Order.cartDiscountAmountSet.shopMoney.amount * -1
        $Order.lineItems.edges += [PSCustomObject]@{
            node = [PSCustomObject]@{
                name = $DiscountName
                sku = $DiscountSku
                quantity = 1
                originalUnitPriceSet = [PSCustomObject]@{
                    shopMoney = [PSCustomObject]@{
                        amount = $DiscountAmount
                    }
                }
                discountedUnitPriceSet = [PSCustomObject]@{
                    shopMoney = [PSCustomObject]@{
                        amount = $DiscountAmount
                    }
                }
                taxLines = @(
                    [PSCustomObject]@{
                        priceSet = [PSCustomObject]@{
                            shopMoney = [PSCustomObject]@{
                                amount = 0
                            }
                        }
                    }
                )
            }
        }
    }
}

function Get-TervisShopifyOrdersWithRefundPending {
    param (
        [Parameter(Mandatory)]$ShopName,
        $Orders
    )
    
    if (-not $Orders) {
        $Orders = Get-ShopifyOrders -ShopName $ShopName -QueryString "tag:RefundPendingImportToEBS NOT tag:IgnoreImport" #Omit exhanges
    }
    $Refunds = @()
    foreach ($Order in $Orders) {
        $LocationDefinition = Get-TervisShopifyLocationDefinition -Name $Order.physicalLocation.name
        $AllRefundsForOrder = Get-ShopifyRefunds -ShopName $ShopName -OrderGID $Order.id
        $OrderId = $Order.id | Get-ShopifyIdFromShopifyGid
        [array]$RefundIDs = $Order | Get-TervisShopifyRefundIdsFromOrderTags
        
        foreach ($RefundID in $RefundIDs) {
            $Refund = $AllRefundsForOrder | Where-Object id -Match $RefundID
            $EBSDocumentReference = "$($LocationDefinition.Subinventory)-$OrderId-$RefundID"
            $Refund | Add-Member -MemberType NoteProperty -Name EBSDocumentReference -Value $EBSDocumentReference -Force
            $Refund | Add-Member -MemberType NoteProperty -Name StoreCustomerNumber -Value $LocationDefinition.CustomerNumber -Force
            $Refund | Add-Member -MemberType NoteProperty -Name Subinventory -Value $LocationDefinition.Subinventory -Force
            $Refund | Add-Member -MemberType NoteProperty -Name RefundID -Value $RefundID
            $Refund | Add-Member -MemberType NoteProperty -Name RefundTag -Value "Refund_$RefundID"
            $Refund | Add-Member -MemberType NoteProperty -Name Order -Value $Order
            $Refund | Select-TervisShopifyOrderPersonalizationLines | Add-TervisShopifyOrderPersonalizationSKU

            $Refunds += $Refund
        }
    }
    return $Refunds
}

function Get-TervisShopifyRefundIdsFromOrderTags {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$Tags
    )
    process {
        $Tags | 
            Where-Object {$_ -match "Refund_"} |
            ForEach-Object {$_.split("_")[1]}
    }
}

function Get-TervisShopifyActiveLocations {
    param (
        [Parameter(Mandatory)]$ShopName
    )
    $ModulePath = if ($PSScriptRoot) {
        $PSScriptRoot
    } else {
        (Get-Module -ListAvailable TervisShopify).ModuleBase
    }
    $ShopifyLocations = Get-ShopifyLocation -ShopName $ShopName -LocationName * | Where-Object isActive -EQ $true
    $LocationDefinitions = Import-Csv -Path $ModulePath\LocationDefinition.csv
    $ShopifyLocations | foreach {
        $Definition = $LocationDefinitions | Where-Object Description -EQ $_.name
        $_ | Add-Member -MemberType NoteProperty -Force -Name Subinventory -Value $Definition.Subinventory
        $_ | Add-Member -MemberType NoteProperty -Force -Name RMSStoreNumber -Value $Definition.RMSStoreNumber
        $_ | Add-Member -MemberType NoteProperty -Force -Name StoreCustomerNumber -Value $Definition.CustomerNumber
    }
    $ShopifyLocations
}

function Invoke-TervisShopifyRefundPendingTagCleanup {
    param (
        [Parameter(Mandatory)]$ShopName
    )
    $Orders = Get-ShopifyOrders -ShopName $ShopName -QueryString "tag:RefundPendingImportToEBS NOT tag:IgnoreImport"
    foreach ($Order in $Orders) {
        if (-not ($Order.tags -match "Refund_")) {
            Write-Warning "Removing RefundPending tag from Shopify order #$($Order.legacyResourceId)."
            $Order | Set-ShopifyOrderTag -ShopName $ShopName -RemoveTag "RefundPendingImportToEBS" | Out-Null
        }
    }
}

function Get-TervisShopifyExchangesForImport {
    param (
        [Parameter(Mandatory)]$ShopName
    )
    # $Orders = Get-ShopifyOrders -ShopName $ShopName -QueryString "NOT tag:ImportedToEBS" | 
    $Orders = Get-ShopifyOrders -ShopName $ShopName -QueryString "tag:XTest" | 
        Where-Object {$_.events.edges.node.message -match "completed an exchange"}

    foreach ($Order in $Orders) {
        $ExchangeCompletedOrderIDs = $Order | Get-TervisShopifyCompletedExchangeOrderID
        
        # Assuming one refund on exchange for now
        $ExchangeCompletedOrders = foreach ($ID in $ExchangeCompletedOrderIDs) {
            Get-ShopifyOrder -ShopName $ShopName -OrderId $ID | Get-TervisShopifyOrdersForImport -ShopName $ShopName
        }
        $Refunds = Get-TervisShopifyOrdersWithRefundPending -ShopName $ShopName -Orders $Order
        
        $ConvertedOrderHeader = $ExchangeCompletedOrders | Convert-TervisShopifyOrderToEBSOrderLineHeader
        [array]$ConvertedOrderLines = $ExchangeCompletedOrders | Convert-TervisShopifyOrderToEBSOrderLines
        $ConvertedOrderLines += $Refunds | Convert-TervisShopifyRefundToEBSOrderLines
        # if total payment is positive, create payment
        # Need to correctly calculate the total payment (New items minus credit from sale)
        $ConvertedOrderPayment = $ExchangeCompletedOrders | Convert-TervisShopifyPaymentsToEBSPayment -ShopName $ShopName
        [array]$Subqueries = $ConvertedOrderHeader | New-EBSOrderLineHeaderSubquery
        $Subqueries += $ConvertedOrderLines | New-EBSOrderLineSubquery
        $Subqueries += $ConvertedOrderPayment | New-EBSOrderLinePaymentSubquery

        Invoke-EBSSubqueryInsert -ShowQuery -Subquery $Subqueries

        <#
        Next things to work on:
        - Correct line item number (count starts over on refunds)
        - Refunds need to have correct orig sys doc ref (match exchange order, not original order+refund id)
        - Correctly handle refund ID tags on original order
        - Do not attempt to overwrite existing exchange order in EBS
        #>
    }
}

function Get-TervisShopifyCompletedExchangeOrderID {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$Order
    )
    process {
        $CompletedOrderEventMessages = $Order.events.edges.node.message -match "completed an exchange"
        $IDs = foreach ($Message in $CompletedOrderEventMessages) {
            ($Message -split 'orders/' | Select-Object -First 1 -Skip 1) -split '">' | Select-Object -First 1
        }
        return $IDs
    }
}

function Get-TervisShopifySuperCollectionName {
    param (
        [Parameter(Mandatory)]$Collection
    )

    $ModulePath = if ($PSScriptRoot) {
        $PSScriptRoot
    } else {
        (Get-Module -ListAvailable TervisShopify).ModuleBase
    }
    if (-not $Script:CollectionDefinition) {
        $Script:CollectionDefinition = Import-Csv -Path $ModulePath\OnlineCollections.csv
    }
    return $Script:CollectionDefinition | 
        Where-Object ChildCollection -Like $Collection | 
        Select-Object -ExpandProperty ParentCollection -Unique
}

function Find-TervisShopifyEBSOrderNumberAndOrigSysDocumentRef {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$SearchTerm,
        [Parameter(Mandatory)][ValidateSet("order_number","orig_sys_document_ref")]$Column
    )
    begin {
        $BaseQuery = "select order_number, orig_sys_document_ref from apps.oe_order_headers_all "
    }
    process {
        if ($SearchTerm -match "%") {
             $Operator = "LIKE"
        } else {
            $Operator = "="
        }

        $Query = $BaseQuery + "where $Column $Operator '$SearchTerm'"
        Invoke-EBSSQL -SQLCommand $Query
    }
}

function Get-TervisShopifyEBSOrderNumberFromShopifyOrderID {
    param (
        $OrderID
    )
    Find-TervisShopifyEBSOrderNumberAndOrigSysDocumentRef -Column orig_sys_document_ref -SearchTerm "%$OrderID%"
}

function Invoke-TervisShopifyReprocessBTO {
    param (
        $OrderID
    )
    $ShopifyOrder = Get-ShopifyOrder -ShopName tervisstore -OrderId $OrderID
    $Order = Get-TervisShopifyOrdersForImport -ShopName tervisstore -Orders $ShopifyOrder
    $IsBTO = $Order | Test-TervisShopifyBuildToOrder
    if ($IsBTO) {
        $OrderBTO = $Order | ConvertTo-TervisShopifyOrderBTO
        if (-not (Test-TervisShopifyEBSOrderExists -Order $OrderBTO)) {
            $OrderObject = $OrderBTO | New-TervisShopifyBuildToOrderObject
            $EBSQueryBTO = $OrderObject | Convert-TervisShopifyOrderObjectToEBSQuery
            $text = $OrderObject | ConvertTo-JsonEx
            Read-Host "$text`n`nContinue?"
            Invoke-EBSSQL -SQLCommand $EBSQueryBTO
        } else {
            Write-Warning "BTO already in EBS"
        }
    } else {
        Write-Warning "No BTO detected"
    }

}