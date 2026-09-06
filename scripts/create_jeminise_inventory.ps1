param(
  [string]$ShopDomain = "jeminise.com",
  [string]$Collection = "all",
  [string]$RunId = ""
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

if ([string]::IsNullOrWhiteSpace($RunId)) {
  $RunId = Get-Date -Format "yyyyMMdd_HHmmss"
}

$baseUrl = "https://$ShopDomain"
$runDir = Join-Path "seo_runs" (Join-Path $ShopDomain $RunId)
$evidenceDir = Join-Path $runDir "evidence"
$productsEvidenceDir = Join-Path $evidenceDir "products"
$imagesEvidenceDir = Join-Path $evidenceDir "images"
$resultsDir = Join-Path "resutls" (Join-Path $ShopDomain $RunId)
$batchesDir = Join-Path $resultsDir "batches"

New-Item -ItemType Directory -Force -Path $runDir, $productsEvidenceDir, $imagesEvidenceDir, $batchesDir | Out-Null

$rawProducts = @()
$page = 1
$limit = 250

while ($true) {
  $url = "$baseUrl/collections/$Collection/products.json?limit=$limit&page=$page"
  Write-Host "Fetching $url"
  $response = Invoke-WebRequest -Uri $url -UseBasicParsing
  $jsonPath = Join-Path $runDir ("products_page_{0}.json" -f $page)
  Set-Content -Path $jsonPath -Value $response.Content -Encoding UTF8

  $payload = $response.Content | ConvertFrom-Json
  if ($null -eq $payload.products -or $payload.products.Count -eq 0) {
    break
  }

  $rawProducts += $payload.products
  if ($payload.products.Count -lt $limit) {
    break
  }

  $page += 1
}

$deduped = $rawProducts | Group-Object id | ForEach-Object { $_.Group[0] }
$now = (Get-Date).ToString("o")
$inventory = @()
$position = 1

foreach ($product in $deduped) {
  $variantPrices = @($product.variants | ForEach-Object { [decimal]$_.price })
  $availableVariants = @($product.variants | Where-Object { $_.available -eq $true })
  $firstImage = $null
  if ($product.images -and $product.images.Count -gt 0) {
    $firstImage = $product.images[0].src
  }

  $inventory += [pscustomobject]@{
    inventory_position = $position
    shop_domain = $ShopDomain
    product_key = "$ShopDomain+$($product.handle)"
    product_id = [string]$product.id
    product_gid = ""
    Handle = $product.handle
    product_url = "$baseUrl/products/$($product.handle)"
    collection_url = "$baseUrl/collections/$Collection/products/$($product.handle)"
    canonical_url = ""
    title_current = $product.title
    vendor = $product.vendor
    product_type = $product.product_type
    tags = ($product.tags -join ", ")
    published_at = $product.published_at
    created_at = $product.created_at
    updated_at = $product.updated_at
    variant_count = @($product.variants).Count
    image_count = @($product.images).Count
    first_image_url = $firstImage
    min_price = if ($variantPrices.Count -gt 0) { ($variantPrices | Measure-Object -Minimum).Minimum.ToString("0.00") } else { "" }
    max_price = if ($variantPrices.Count -gt 0) { ($variantPrices | Measure-Object -Maximum).Maximum.ToString("0.00") } else { "" }
    available = if ($availableVariants.Count -gt 0) { "TRUE" } else { "FALSE" }
    source = "collections/$Collection/products.json"
    discovery_status = "DISCOVERED"
    processing_status = "DISCOVERED"
    notes = ""
    discovered_at = $now
  }

  $position += 1
}

$inventoryPath = Join-Path $runDir "inventory.csv"
$inventoryJsonPath = Join-Path $runDir "inventory.json"
$rawCombinedPath = Join-Path $runDir "products_all_raw.json"
$progressPath = Join-Path $runDir "progress.json"

$inventory | Export-Csv -Path $inventoryPath -NoTypeInformation -Encoding UTF8
$inventory | ConvertTo-Json -Depth 8 | Set-Content -Path $inventoryJsonPath -Encoding UTF8
$deduped | ConvertTo-Json -Depth 30 | Set-Content -Path $rawCombinedPath -Encoding UTF8

$typeCounts = $inventory | Group-Object product_type | Sort-Object Count -Descending | ForEach-Object {
  [pscustomobject]@{
    product_type = if ([string]::IsNullOrWhiteSpace($_.Name)) { "UNKNOWN" } else { $_.Name }
    count = $_.Count
  }
}

$batchSize = 10
$firstBatch = @($inventory | Select-Object -First $batchSize)

$progress = [ordered]@{
  schema_version = "inventory_v1"
  prompt_version = "2.4"
  run_id = $RunId
  shop_domain = $ShopDomain
  scope = "collections/$Collection"
  market = "UNKNOWN"
  seo_language = "English"
  inventory_ref = $inventoryPath
  inventory_json_ref = $inventoryJsonPath
  raw_products_ref = $rawCombinedPath
  product_count_discovered = $inventory.Count
  expected_collection_count_observed = 338
  count_match_observed_collection = ($inventory.Count -eq 338)
  product_type_counts = $typeCounts
  batch_size = $batchSize
  batch_id = "batch_001"
  batch_product_keys = @($firstBatch | ForEach-Object { $_.product_key })
  batch_status = "INVENTORY_CREATED"
  awaiting_confirmation = $false
  continuation_confirmation_ref = ""
  current_product_key = ""
  current_stage = "INVENTORY_CREATED"
  images_completed = @()
  last_saved_at = $now
  results_dir = $resultsDir
  artifact_paths = [ordered]@{
    inventory_csv = $inventoryPath
    inventory_json = $inventoryJsonPath
    raw_products_json = $rawCombinedPath
    progress_json = $progressPath
    batch_results_dir = $batchesDir
  }
}

$progress | ConvertTo-Json -Depth 10 | Set-Content -Path $progressPath -Encoding UTF8

Write-Host ""
Write-Host "Inventory created"
Write-Host "Run ID: $RunId"
Write-Host "Products discovered: $($inventory.Count)"
Write-Host "Inventory CSV: $inventoryPath"
Write-Host "Progress JSON: $progressPath"
Write-Host "Product type counts:"
$typeCounts | Format-Table -AutoSize
