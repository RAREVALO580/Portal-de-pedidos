# Script de Actualización Automática del Portal de Traslados
# Lee ME2W.xlsx y ME80FN.xlsx directamente y regenera portal_data.js y dashboard_data.json

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "   MACROPAY - ETL SAP MM PORTAL DE TRASLADOS" -ForegroundColor Yellow
Write-Host "==========================================================" -ForegroundColor Cyan

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$me2wPath = Join-Path $baseDir "ME2W.xlsx"
$me80fnPath = Join-Path $baseDir "ME80FN.xlsx"
$jsonPath = Join-Path $baseDir "dashboard_data.json"
$jsPath = Join-Path $baseDir "portal_data.js"

if (-not (Test-Path $me2wPath) -or -not (Test-Path $me80fnPath)) {
    Write-Host "ERROR: Asegúrate de que ME2W.xlsx y ME80FN.xlsx se encuentren en este directorio." -ForegroundColor Red
    Pause
    exit 1
}

Write-Host "Iniciando lectura ultrarrápida de SAP Excel en memoria..." -ForegroundColor Green
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # 1. ME2W
    $wb1 = $excel.Workbooks.Open($me2wPath)
    $ws1 = $wb1.Sheets.Item(1)
    $data1 = $ws1.UsedRange.Value2
    $rows1 = $ws1.UsedRange.Rows.Count
    $wb1.Close($false)
    Write-Host "[OK] ME2W cargado: $rows1 filas" -ForegroundColor Gray

    # 2. ME80FN
    $wb2 = $excel.Workbooks.Open($me80fnPath)
    $ws2 = $wb2.Sheets.Item(1)
    $data2 = $ws2.UsedRange.Value2
    $rows2 = $ws2.UsedRange.Rows.Count
    $wb2.Close($false)
    Write-Host "[OK] ME80FN cargado: $rows2 filas" -ForegroundColor Gray

    # 3. Directorio comercial (Columna C: DIV, Columna H: SUCURSAL)
    $dirComPath = Join-Path $baseDir "Directorio comercial.xlsx"
    $storeInfo = @{}
    if (Test-Path $dirComPath) {
        $tempCopy = "$env:TEMP\Directorio_copy.xlsx"
        Copy-Item $dirComPath $tempCopy -Force
        $wb3 = $excel.Workbooks.Open($tempCopy, [Type]::Missing, $true)
        $ws3 = $wb3.Sheets.Item(1)
        $data3 = $ws3.UsedRange.Value2
        $rows3 = $ws3.UsedRange.Rows.Count
        $wb3.Close($false)
        for ($r = 2; $r -le $rows3; $r++) {
            $c = "$($data3[$r, 1])".Trim()
            $div = "$($data3[$r, 3])".Trim()
            $nom = "$($data3[$r, 8])".Trim()
            if ($c) {
                $info = @{
                    division = if ($div) { $div } else { "Sin División" }
                    sucursal_nombre = $nom
                }
                $storeInfo[$c] = $info
                $cTrim = $c.TrimStart('0')
                if ($cTrim -ne $c) { $storeInfo[$cTrim] = $info }
            }
        }
        Write-Host "[OK] Directorio comercial cargado: $($storeInfo.Keys.Count) tiendas con División" -ForegroundColor Gray
    }

    Write-Host "Indexando movimientos contables e historial de pedidos (EKBE)..." -ForegroundColor Yellow
    $posMovements = @{}
    $orderMovements = @{}

    for ($r = 2; $r -le $rows2; $r++) {
        $ebeln = "$($data2[$r, 1])".Trim()
        if (-not ($ebeln.StartsWith("73") -or $ebeln.StartsWith("67") -or $ebeln.StartsWith("78"))) { continue }
        
        $ebelp = "$($data2[$r, 2])".Trim()
        $bwart = "$($data2[$r, 6])".Trim()
        $budatVal = $data2[$r, 7]
        $mengeVal = $data2[$r, 8]
        $belnr = "$($data2[$r, 4])".Trim()

        if ($bwart -eq "" -or $mengeVal -eq $null) { continue }

        $budatStr = ""
        $budatDate = $null
        if ($budatVal -is [Double] -or $budatVal -is [Int32]) {
            $budatDate = [DateTime]::FromOADate($budatVal)
            $budatStr = $budatDate.ToString("yyyy-MM-dd")
        } elseif ($budatVal -is [String]) {
            $budatStr = $budatVal
        }

        $mengeNum = [Math]::Abs([Convert]::ToDouble($mengeVal))

        $key = "$ebeln-$ebelp"
        if (-not $posMovements.ContainsKey($key)) {
            $posMovements[$key] = [System.Collections.Generic.List[Object]]::new()
        }
        $posMovements[$key].Add(@{
            BWART = $bwart
            BUDAT = $budatStr
            Date = $budatDate
            MENGE = $mengeNum
            BELNR = $belnr
        })

        if (-not $orderMovements.ContainsKey($ebeln)) {
            $orderMovements[$ebeln] = [System.Collections.Generic.List[Object]]::new()
        }
        $orderMovements[$ebeln].Add(@{
            BWART = $bwart
            BUDAT = $budatStr
            Date = $budatDate
            MENGE = $mengeNum
            EBELP = $ebelp
            BELNR = $belnr
        })
    }

    Write-Host "Consolidando pedidos de traslado y aplicando reglas de negocio..." -ForegroundColor Yellow
    $orders = [System.Collections.Specialized.OrderedDictionary]::new()

    for ($r = 2; $r -le $rows1; $r++) {
        $ebeln = "$($data1[$r, 1])".Trim()
        if (-not ($ebeln.StartsWith("73") -or $ebeln.StartsWith("67") -or $ebeln.StartsWith("78"))) { continue }
        
        $loekz = "$($data1[$r, 13])".Trim()
        if ($loekz -ne "") { continue }

        $reswk = "$($data1[$r, 2])".Trim()
        $ebelp = "$($data1[$r, 3])".Trim()
        $aedatVal = $data1[$r, 9]
        $matnr = "$($data1[$r, 10])".Trim()
        $maktx = "$($data1[$r, 11])".Trim()
        $werks = "$($data1[$r, 14])".Trim()
        $mengeVal = $data1[$r, 16]
        $cantPedida = if ($mengeVal -ne $null) { [Convert]::ToDouble($mengeVal) } else { 0 }

        $aedatStr = ""
        $aedatDate = $null
        if ($aedatVal -is [Double] -or $aedatVal -is [Int32]) {
            $aedatDate = [DateTime]::FromOADate($aedatVal)
            $aedatStr = $aedatDate.ToString("yyyy-MM-dd")
        } elseif ($aedatVal -is [String]) {
            $aedatStr = $aedatVal
        }

        if (-not $orders.Contains($ebeln)) {
            $cat = if ($ebeln.StartsWith("73")) { "🏍️ Motos" } elseif ($ebeln.StartsWith("67")) { "🪖 Cascos" } else { "🚲 Bicicletas" }
            $orders[$ebeln] = @{
                ebeln = $ebeln
                reswk = $reswk
                werks = $werks
                aedat = $aedatStr
                aedatDate = $aedatDate
                categoria = $cat
                items = [System.Collections.Generic.List[Object]]::new()
            }
        }

        $posKey = "$ebeln-$ebelp"
        $posDesp = 0
        $posRec = 0
        $posSalidaDate = $null
        $posEntradaDate = $null
        $posDocSalida = ""
        $posDocEntrada = ""

        if ($posMovements.ContainsKey($posKey)) {
            foreach ($m in $posMovements[$posKey]) {
                if ($m.BWART -in @("351", "641")) {
                    $posDesp += $m.MENGE
                    if ($m.Date -and (-not $posSalidaDate -or $m.Date -gt $posSalidaDate)) { 
                        $posSalidaDate = $m.Date 
                        $posDocSalida = $m.BELNR
                    }
                } elseif ($m.BWART -in @("352", "642")) {
                    $posDesp -= $m.MENGE
                } elseif ($m.BWART -eq "101") {
                    $posRec += $m.MENGE
                    if ($m.Date -and (-not $posEntradaDate -or $m.Date -gt $posEntradaDate)) { 
                        $posEntradaDate = $m.Date 
                        $posDocEntrada = $m.BELNR
                    }
                } elseif ($m.BWART -in @("102", "104")) {
                    $posRec -= $m.MENGE
                }
            }
        }

        $posTransito = [Math]::Max(0, ($posDesp - $posRec))

        $orders[$ebeln].items.Add(@{
            ebelp = $ebelp
            matnr = $matnr
            maktx = $maktx
            meins = "PZA"
            cant_pedida = $cantPedida
            cant_despachada = $posDesp
            cant_recibida = $posRec
            cant_en_transito = $posTransito
            doc_salida = $posDocSalida
            doc_entrada = $posDocEntrada
            fecha_salida = if ($posSalidaDate) { $posSalidaDate.ToString("yyyy-MM-dd") } else { "" }
            fecha_entrada = if ($posEntradaDate) { $posEntradaDate.ToString("yyyy-MM-dd") } else { "" }
        })
    }

    $finalOrders = [System.Collections.Generic.List[Object]]::new()
    
    foreach ($ebeln in $orders.Keys) {
        $o = $orders[$ebeln]
        $totPed = 0
        $totDesp = 0
        $totRec = 0
        $totTrans = 0
        $skus = @{}
        $maxSalidaDate = $null
        $maxEntradaDate = $null
        $minBudat = $null

        foreach ($it in $o.items) {
            $totPed += $it.cant_pedida
            $totDesp += $it.cant_despachada
            $totRec += $it.cant_recibida
            $totTrans += $it.cant_en_transito
            $skus[$it.matnr] = $true
        }

        if ($orderMovements.ContainsKey($ebeln)) {
            foreach ($m in $orderMovements[$ebeln]) {
                if ($m.Date) {
                    if (-not $minBudat -or $m.Date -lt $minBudat) { $minBudat = $m.Date }
                    if ($m.BWART -in @("351", "641")) {
                        if (-not $maxSalidaDate -or $m.Date -gt $maxSalidaDate) { $maxSalidaDate = $m.Date }
                    }
                    if ($m.BWART -eq "101") {
                        if (-not $maxEntradaDate -or $m.Date -gt $maxEntradaDate) { $maxEntradaDate = $m.Date }
                    }
                }
            }
        }

        $status = ""
        if ($totDesp -gt $totRec) {
            $status = "🚚 En Ruta a Tienda"
        } elseif ($totRec -gt 0 -and ($totRec -ge $totPed -or $totRec -ge $totDesp)) {
            $status = "✅ Entregado en Sucursal"
        } elseif ($totDesp -eq 0 -and $totRec -eq 0) {
            $status = "🟡 En Preparación CEDI"
        } else {
            $status = "⚠️ Discrepancia"
        }

        $wStr = "$($o.werks)".Trim()
        $wTrim = $wStr.TrimStart('0')
        $sInfo = if ($storeInfo.ContainsKey($wStr)) { $storeInfo[$wStr] } elseif ($storeInfo.ContainsKey($wTrim)) { $storeInfo[$wTrim] } else { $null }
        $divVal = if ($sInfo) { $sInfo.division } else { "Sin División" }
        $nomVal = if ($sInfo) { $sInfo.sucursal_nombre } else { "Centro $wStr" }

        $finalOrders.Add(@{
            ebeln = $ebeln
            reswk = $o.reswk
            werks = $o.werks
            division = $divVal
            sucursal_nombre = $nomVal
            aedat = $o.aedat
            categoria = $o.categoria
            total_modelos = $skus.Keys.Count
            cant_pedida = $totPed
            cant_despachada = $totDesp
            cant_recibida = $totRec
            cant_en_transito = $totTrans
            estatus = $status
            fecha_pedido = $o.aedat
            fecha_liberacion = if ($minBudat) { $minBudat.ToString("yyyy-MM-dd") } else { "" }
            fecha_salida = if ($maxSalidaDate) { $maxSalidaDate.ToString("yyyy-MM-dd") } else { "" }
            fecha_entrega = if ($maxEntradaDate) { $maxEntradaDate.ToString("yyyy-MM-dd") } else { "" }
            items = $o.items
        })
    }

    $json = $finalOrders | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($jsonPath, $json, [System.Text.Encoding]::UTF8)
    [System.IO.File]::WriteAllText($jsPath, "window.PORTAL_DATA = " + $json + ";", [System.Text.Encoding]::UTF8)

    # 4. MB52 (Inventario de CEDIS SAP)
    $mb52Path = Join-Path $baseDir "MB52.xlsx"
    $invJsonPath = Join-Path $baseDir "inventario_data.json"
    $invJsPath = Join-Path $baseDir "inventario_data.js"

    if (Test-Path $mb52Path) {
        Write-Host "Procesando inventario de CEDIS SAP (MB52.xlsx)..." -ForegroundColor Yellow
        $tempMb52 = "$env:TEMP\MB52_copy.xlsx"
        Copy-Item $mb52Path $tempMb52 -Force
        $wb4 = $excel.Workbooks.Open($tempMb52, [Type]::Missing, $true)
        $ws4 = $wb4.Sheets.Item(1)
        $data4 = $ws4.UsedRange.Value2
        $rows4 = $ws4.UsedRange.Rows.Count
        $wb4.Close($false)
        Write-Host "[OK] MB52 cargado: $rows4 filas de inventario SAP" -ForegroundColor Gray

        $cediDivMap = @{
            "0028" = "División 10"; "0029" = "División 10"; "0030" = "División 10";
            "0025" = "División 20"; "0027" = "División 20";
            "0011" = "División 30"; "0026" = "División 30"; "0021" = "División 30";
            "0014" = "División 30"; "0013" = "División 30"; "0017" = "División 30"; "0024" = "División 30"
        }

        $invGrouped = [System.Collections.Specialized.OrderedDictionary]::new()

        for ($r = 2; $r -le $rows4; $r++) {
            $cedi = "$($data4.GetValue($r, 1))".Trim()
            $cediName = "$($data4.GetValue($r, 3))".Trim()
            $mat = "$($data4.GetValue($r, 4))".Trim()
            $txt = "$($data4.GetValue($r, 5))".Trim()
            $sapCat = "$($data4.GetValue($r, 6))".Trim()
            $val = $data4.GetValue($r, 7)
            $stock = if ($val -ne $null) { [Convert]::ToDouble($val) } else { 0 }
            
            if (-not $mat -or -not $txt) { continue }
            
            $key = "$cedi-$mat"
            if ($invGrouped.Contains($key)) {
                $invGrouped[$key].stock += $stock
            } else {
                $t = $txt.ToUpper()
                
                $cat = 'STREET'
                if ($t -match 'ADV|TRAIL|DAKOTA|SNAKE|RUBAK') { $cat = 'ADVENTURE' }
                elseif ($t -match 'CHOPPER|DARK EAGLE') { $cat = 'CHOPPER' }
                elseif ($t -match 'CUATRIMOTO|ATV') { $cat = 'CUATRIMOTO' }
                elseif ($t -match 'HAWK GT|R250Z|R300Z|RAD HYPER|GTR|SPORT') { $cat = 'DERPORTIVA' }
                elseif ($t -match 'STARK|HAWK|DIRT|CROSS') { $cat = 'DOBLE PROPOSITO' }
                elseif ($t -match 'COVALT|MOTOCARRO') { $cat = 'MOTOCARRO' }
                elseif ($t -match 'SCOOTER|FLUX|RABBIT|METRO|FREEDOM') { $cat = 'SCOOTER' }
                elseif ($t -match 'ATRIO|HUNTER|DYNAMIK|REVENGE|U2|U5') { $cat = 'SEMIAUTOMATICA' }
                elseif ($t -match 'CASTORO|SAETA|SB |THUNDER|PULSAR|NOVAK|SYRIUS|THYRON|BEAST|BOXTER|HAMMER|INDY') { $cat = 'STREET' }
                elseif ($t -match 'JAMBO|MAGMA|BOXER|BM 150') { $cat = 'TRABAJO' }
                elseif ($sapCat -eq 'BICI' -or $t -match 'BICI') { $cat = 'BICICLETAS' }
                elseif ($sapCat -eq 'CASC' -or $t -match 'CASCO') { $cat = 'CASCOS' }
                
                $prov = 'OTRO'
                if ($t -match 'BDS|BODESA') { $prov = 'BODESA' }
                elseif ($t -match 'DINAMO') { $prov = 'DINAMO' }
                elseif ($t -match 'CARABELA|ISLO|XEVERUS') { $prov = 'MOTO ROAD' }
                elseif ($t -match 'VELOCI') { $prov = 'VELOCI MOTORS' }
                elseif ($t -match 'BAJAJ') { $prov = 'BAJAJ' }
                elseif ($t -match 'YADEA|KIWO|MACROBIKE') { $prov = 'PROVEEDOR BICI' }
                elseif ($t -match 'XTR|XECURO') { $prov = 'PROVEEDOR CASCOS' }

                $div = if ($cediDivMap.ContainsKey($cedi)) { $cediDivMap[$cedi] } else { "División 10" }

                $invGrouped[$key] = [PSCustomObject]@{
                    material = $mat
                    texto = $txt
                    categoria = $cat
                    sap_categoria = $sapCat
                    proveedor = $prov
                    stock = $stock
                    division = $div
                    cedi = $cedi
                    cedi_nombre = $cediName
                }
            }
        }

        $invJson = $invGrouped.Values | ConvertTo-Json -Depth 3
        [System.IO.File]::WriteAllText($invJsonPath, $invJson, [System.Text.Encoding]::UTF8)
        [System.IO.File]::WriteAllText($invJsPath, "window.INVENTARIO_DATA = " + $invJson + ";", [System.Text.Encoding]::UTF8)
        Write-Host "[OK] 'inventario_data.js' generado con $($invGrouped.Count) materiales de MB52" -ForegroundColor Green
    }

    $sw.Stop()
    Write-Host "`n[OK] PROCESO COMPLETADO EN $($sw.Elapsed.TotalSeconds.ToString('F1')) SEGUNDOS" -ForegroundColor Green
    Write-Host "Total de pedidos procesados: $($finalOrders.Count)" -ForegroundColor Cyan
    Write-Host "Los archivos 'portal_data.js' y 'dashboard_data.json' han sido actualizados con éxito." -ForegroundColor Green

} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
