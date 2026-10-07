<?php
/**
 * Import and Export barcode database via Excel (XLSX, XLS, CSV)
 * Supports:
 *  1. action=stats: Database metrics
 *  2. action=import_places: Direct 1C Nomenclature & Storage Locations (Код, Код Гарда, Номенклатура, Привязанная ячейка)
 *  3. action=import_barcodes: Direct 1C Barcodes & Packaging Register (Штрихкод, Номенклатура, Производитель, Единица измерения, Количество)
 *  4. action=import: Full combined Excel import (legacy merge/replace)
 *  5. action=export: Full Excel database export
 */
require_once __DIR__ . '/../vendor/autoload.php';
require_once __DIR__ . '/../db.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;

// Increase time and memory limits for large 1C spreadsheets
@set_time_limit(300);
@ini_set('memory_limit', '512M');

$action = $_GET['action'] ?? ($_POST['action'] ?? 'stats');

try {
    $pdo = getDB();

    if ($action === 'stats') {
        $productCount = $pdo->query("SELECT COUNT(*) FROM products")->fetchColumn();
        $withPlaceCount = $pdo->query("SELECT COUNT(*) FROM products WHERE place IS NOT NULL AND TRIM(place) != ''")->fetchColumn();
        $barcodeCount = $pdo->query("SELECT COUNT(*) FROM product_barcodes")->fetchColumn();
        $manufacturerCount = $pdo->query("SELECT COUNT(DISTINCT manufacturer) FROM products WHERE manufacturer IS NOT NULL AND manufacturer != ''")->fetchColumn();
        
        jsonResponse([
            'success' => true,
            'stats' => [
                'products' => (int)$productCount,
                'products_with_place' => (int)$withPlaceCount,
                'barcodes' => (int)$barcodeCount,
                'manufacturers' => (int)$manufacturerCount
            ]
        ]);
    }

    if ($action === 'backup_db') {
        $dbFile = getDataPath() . '/warehouse.db';
        if (!file_exists($dbFile)) {
            jsonError('Файл базы данных не найден', 404);
        }
        $filename = 'warehouse_backup_' . date('Y-m-d_His') . '.db';
        header('Content-Type: application/x-sqlite3');
        header('Content-Disposition: attachment; filename="' . $filename . '"');
        header('Content-Length: ' . filesize($dbFile));
        header('Cache-Control: no-cache, no-store, must-revalidate');
        readfile($dbFile);
        exit;
    }

    if ($action === 'clear_database') {
        if (($_POST['confirm'] ?? '') !== 'yes') {
            jsonError('Подтверждение очистки не получено');
        }
        $pdo->beginTransaction();
        $pdo->exec("DELETE FROM product_barcodes;");
        $pdo->exec("DELETE FROM products;");
        $pdo->commit();
        jsonResponse([
            'success' => true,
            'message' => 'База товаров и штрихкодов успешно очищена'
        ]);
    }

    if ($action === 'export') {
        // Fetch all products and their associated barcodes
        $stmt = $pdo->query("SELECT * FROM products ORDER BY manufacturer ASC, name ASC");
        $products = $stmt->fetchAll();

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('Штрихкоды');

        // Headers
        $headers = [
            'A1' => 'код(наш)',
            'B1' => 'код(поставщика)',
            'C1' => 'наименование',
            'D1' => 'штрихкод_бигбокс',
            'E1' => 'количество_бигбокс',
            'F1' => 'штрихкод_минибокс',
            'G1' => 'количество_минибокс',
            'H1' => 'штрихкод_штучный',
            'I1' => 'количество_штучный',
            'J1' => 'Место',
            'K1' => 'Производитель'
        ];

        foreach ($headers as $cell => $val) {
            $sheet->setCellValue($cell, $val);
            $sheet->getStyle($cell)->getFont()->setBold(true);
        }

        $rowNum = 2;
        $barcodeStmt = $pdo->prepare("SELECT * FROM product_barcodes WHERE product_id = ? ORDER BY quantity DESC");

        foreach ($products as $prod) {
            $barcodeStmt->execute([$prod['id']]);
            $barcodes = $barcodeStmt->fetchAll();

            $bigboxBarcode = '';
            $bigboxQty = '';
            $miniboxBarcodes = [];
            $pieceBarcode = '';
            $pieceQty = '';

            foreach ($barcodes as $bc) {
                $type = mb_strtolower(trim($bc['packaging_type'] ?? ''));
                if (str_contains($type, 'биг') || str_contains($type, 'big')) {
                    if (!$bigboxBarcode) {
                        $bigboxBarcode = $bc['barcode'];
                        $bigboxQty = $bc['quantity'];
                    }
                } elseif (str_contains($type, 'штуч') || str_contains($type, 'piece') || $bc['quantity'] == 1) {
                    if (!$pieceBarcode) {
                        $pieceBarcode = $bc['barcode'];
                        $pieceQty = $bc['quantity'];
                    }
                } else {
                    $miniboxBarcodes[] = $bc;
                }
            }

            // Primary row
            $firstMini = array_shift($miniboxBarcodes);
            $sheet->setCellValueExplicit("A$rowNum", $prod['our_code'] ?? '', \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
            $sheet->setCellValueExplicit("B$rowNum", $prod['supplier_code'] ?? '', \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
            $sheet->setCellValue("C$rowNum", $prod['name'] ?? '');
            $sheet->setCellValueExplicit("D$rowNum", $bigboxBarcode, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
            $sheet->setCellValue("E$rowNum", $bigboxQty);
            $sheet->setCellValueExplicit("F$rowNum", $firstMini['barcode'] ?? '', \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
            $sheet->setCellValue("G$rowNum", $firstMini['quantity'] ?? '');
            $sheet->setCellValueExplicit("H$rowNum", $pieceBarcode, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
            $sheet->setCellValue("I$rowNum", $pieceQty);
            $sheet->setCellValue("J$rowNum", $prod['place'] ?? '');
            $sheet->setCellValue("K$rowNum", $prod['manufacturer'] ?? '');
            $rowNum++;

            // If there are additional miniboxes / barcodes for this product, output extra rows
            while (!empty($miniboxBarcodes)) {
                $extraMini = array_shift($miniboxBarcodes);
                $sheet->setCellValueExplicit("A$rowNum", $prod['our_code'] ?? '', \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
                $sheet->setCellValueExplicit("B$rowNum", $prod['supplier_code'] ?? '', \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
                $sheet->setCellValue("C$rowNum", $prod['name'] ?? '');
                $sheet->setCellValueExplicit("F$rowNum", $extraMini['barcode'] ?? '', \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
                $sheet->setCellValue("G$rowNum", $extraMini['quantity'] ?? '');
                $sheet->setCellValue("J$rowNum", $prod['place'] ?? '');
                $sheet->setCellValue("K$rowNum", $prod['manufacturer'] ?? '');
                $rowNum++;
            }
        }

        // Auto-size columns
        foreach (range('A', 'K') as $col) {
            $sheet->getColumnDimension($col)->setAutoSize(true);
        }

        $filename = 'warehouse_backup_' . date('Y-m-d_His') . '.xlsx';
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . $filename . '"');
        header('Cache-Control: max-age=0');

        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('php://output');
        exit;
    }

    /**
     * Helper to load uploaded file rows safely
     */
    function loadUploadedRows() {
        $fileKey = isset($_FILES['file']) ? 'file' : (isset($_FILES['excel_file']) ? 'excel_file' : null);
        if (!$fileKey || $_FILES[$fileKey]['error'] !== UPLOAD_ERR_OK) {
            jsonError('Файл не был загружен или произошла ошибка при передаче');
        }

        $tmpFile = $_FILES[$fileKey]['tmp_name'];
        $reader = IOFactory::createReaderForFile($tmpFile);
        if ($reader instanceof \PhpOffice\PhpSpreadsheet\Reader\Csv) {
            $reader->setInputEncoding('UTF-8');
            $reader->setDelimiter(',');
        }
        $reader->setReadDataOnly(true);
        $spreadsheet = $reader->load($tmpFile);
        $sheet = $spreadsheet->getActiveSheet();

        $rows = $sheet->toArray(null, true, true, true);
        if (count($rows) < 2) {
            jsonError('Таблица Excel пуста или содержит только одну строку');
        }
        return $rows;
    }

    // =========================================================================
    // ACTION: import_places (Товары и Адреса ячеек из 1С)
    // Columns: Код | Код Гарда | Номенклатура | Привязанная ячейка
    // =========================================================================
    if ($action === 'import_places') {
        $rows = loadUploadedRows();

        // 1. Detect columns by header
        $headerRowIdx = null;
        $colMap = [
            'our_code'      => null,
            'supplier_code' => null,
            'name'          => null,
            'place'         => null
        ];

        foreach ($rows as $rIdx => $row) {
            if ($rIdx > 15) break;
            $hasName = false;
            $hasPlace = false;
            $tempMap = ['our_code' => null, 'supplier_code' => null, 'name' => null, 'place' => null];

            foreach ($row as $col => $val) {
                $v = mb_strtolower(trim((string)$val));
                if (in_array($v, ['наш код', 'наш_код', 'код(наш)', 'код 1с']) || ($v === 'код' && $tempMap['our_code'] === null)) {
                    $tempMap['our_code'] = $col;
                } elseif (str_contains($v, 'гарда') || str_contains($v, 'поставщик') || str_contains($v, 'артикул')) {
                    $tempMap['supplier_code'] = $col;
                } elseif (str_contains($v, 'номенклатура') || str_contains($v, 'наименование') || str_contains($v, 'товар') || $v === 'name') {
                    $tempMap['name'] = $col;
                    $hasName = true;
                } elseif (str_contains($v, 'ячейк') || str_contains($v, 'место') || str_contains($v, 'адрес') || str_contains($v, 'place')) {
                    $tempMap['place'] = $col;
                    $hasPlace = true;
                }
            }

            if ($hasName || $hasPlace) {
                $headerRowIdx = $rIdx;
                $colMap = $tempMap;
                break;
            }
        }

        // Fallback column positions if header text was atypical
        if ($headerRowIdx === null) $headerRowIdx = 1;
        if ($colMap['our_code'] === null)      $colMap['our_code']      = 'A';
        if ($colMap['supplier_code'] === null) $colMap['supplier_code'] = 'B';
        if ($colMap['name'] === null)          $colMap['name']          = 'C';
        if ($colMap['place'] === null)         $colMap['place']         = 'D';

        // 2. Pre-cache existing products into memory for lightning fast updates
        $mapByOurCode = [];
        $mapByName = [];
        $mapBySupplierCode = [];

        $cachedProds = $pdo->query("SELECT id, our_code, supplier_code, name, place FROM products")->fetchAll();
        foreach ($cachedProds as $cp) {
            $pid = (int)$cp['id'];
            if (!empty($cp['our_code'])) {
                $mapByOurCode[trim((string)$cp['our_code'])][] = $pid;
            }
            if (!empty($cp['supplier_code'])) {
                $mapBySupplierCode[trim((string)$cp['supplier_code'])][] = $pid;
            }
            $cleanName = mb_strtolower(trim($cp['name']));
            if ($cleanName !== '') {
                $mapByName[$cleanName][] = $pid;
            }
        }

        $pdo->beginTransaction();

        $updateStmt = $pdo->prepare("
            UPDATE products 
            SET place = ?, 
                our_code = COALESCE(NULLIF(?, ''), our_code), 
                supplier_code = COALESCE(NULLIF(?, ''), supplier_code), 
                name = ?, 
                updated_at = CURRENT_TIMESTAMP 
            WHERE id = ?
        ");

        $insertStmt = $pdo->prepare("
            INSERT INTO products (our_code, supplier_code, name, place) 
            VALUES (?, ?, ?, ?)
        ");

        $totalProcessed = 0;
        $addedCount = 0;
        $updatedCount = 0;
        $totalRows = count($rows);

        for ($i = $headerRowIdx + 1; $i <= $totalRows; $i++) {
            $row = $rows[$i] ?? null;
            if (!$row) continue;

            $ourCode      = trim((string)($row[$colMap['our_code']] ?? ''));
            $supplierCode = trim((string)($row[$colMap['supplier_code']] ?? ''));
            $name         = trim((string)($row[$colMap['name']] ?? ''));
            $place        = trim((string)($row[$colMap['place']] ?? ''));

            if (!$name && !$ourCode && !$supplierCode) {
                continue; // Empty row
            }

            // Skip accidental repeated headers
            $nameLower = mb_strtolower($name);
            if (in_array($nameLower, ['номенклатура', 'наименование', 'товар', 'name', 'код'])) {
                continue;
            }

            if (!$name) {
                $name = 'Товар ' . ($ourCode ?: $supplierCode);
                $nameLower = mb_strtolower($name);
            }

            $totalProcessed++;

            // Match priority:
            // 1. our_code (exact)
            // 2. name (exact case-insensitive)
            // 3. supplier_code (exact)
            $matchedIds = [];
            if ($ourCode !== '' && isset($mapByOurCode[$ourCode])) {
                $matchedIds = $mapByOurCode[$ourCode];
            } elseif (isset($mapByName[$nameLower])) {
                $matchedIds = $mapByName[$nameLower];
            } elseif ($supplierCode !== '' && isset($mapBySupplierCode[$supplierCode])) {
                $matchedIds = $mapBySupplierCode[$supplierCode];
            }

            if (!empty($matchedIds)) {
                // Update all matching entries (typically 1)
                foreach ($matchedIds as $mid) {
                    $updateStmt->execute([
                        $place !== '' ? $place : null,
                        $ourCode !== '' ? $ourCode : null,
                        $supplierCode !== '' ? $supplierCode : null,
                        $name,
                        $mid
                    ]);
                    $updatedCount++;
                }
            } else {
                // Insert new product record
                $insertStmt->execute([
                    $ourCode !== '' ? $ourCode : null,
                    $supplierCode !== '' ? $supplierCode : null,
                    $name,
                    $place !== '' ? $place : null
                ]);
                $newId = (int)$pdo->lastInsertId();
                $mapByName[$nameLower][] = $newId;
                if ($ourCode !== '') {
                    $mapByOurCode[$ourCode][] = $newId;
                }
                if ($supplierCode !== '') {
                    $mapBySupplierCode[$supplierCode][] = $newId;
                }
                $addedCount++;
            }
        }

        $pdo->commit();

        jsonResponse([
            'success' => true,
            'message' => 'Справочник ячеек и товаров успешно обработан',
            'stats' => [
                'total_rows'       => $totalProcessed,
                'added_products'   => $addedCount,
                'updated_products' => $updatedCount
            ]
        ]);
    }

    // =========================================================================
    // ACTION: import_barcodes (Реестр штрихкодов и упаковок из 1С)
    // Columns: Штрихкод | Номенклатура | Производитель | Единица измерения | Количество
    // =========================================================================
    if ($action === 'import_barcodes') {
        $rows = loadUploadedRows();

        // 1. Detect columns by header
        $headerRowIdx = null;
        $colMap = [
            'barcode'        => null,
            'name'           => null,
            'manufacturer'   => null,
            'packaging_type' => null,
            'quantity'       => null
        ];

        foreach ($rows as $rIdx => $row) {
            if ($rIdx > 15) break;
            $hasBarcode = false;
            $hasName = false;
            $tempMap = [
                'barcode'        => null,
                'name'           => null,
                'manufacturer'   => null,
                'packaging_type' => null,
                'quantity'       => null
            ];

            foreach ($row as $col => $val) {
                $v = mb_strtolower(trim((string)$val));
                if (str_contains($v, 'штрих') || $v === 'barcode' || $v === 'шк') {
                    $tempMap['barcode'] = $col;
                    $hasBarcode = true;
                } elseif (str_contains($v, 'номенклатура') || str_contains($v, 'наименование') || str_contains($v, 'товар') || $v === 'name') {
                    $tempMap['name'] = $col;
                    $hasName = true;
                } elseif (str_contains($v, 'производ') || str_contains($v, 'бренд') || str_contains($v, 'марка') || $v === 'manufacturer') {
                    $tempMap['manufacturer'] = $col;
                } elseif (str_contains($v, 'единиц') || str_contains($v, 'упаков') || $v === 'unit' || str_contains($v, 'вид')) {
                    $tempMap['packaging_type'] = $col;
                } elseif (str_contains($v, 'колич') || str_contains($v, 'кратн') || str_contains($v, 'коэф') || str_contains($v, 'кол-во') || $v === 'qty') {
                    $tempMap['quantity'] = $col;
                }
            }

            if ($hasBarcode && $hasName) {
                $headerRowIdx = $rIdx;
                $colMap = $tempMap;
                break;
            }
        }

        // Fallback column positions
        if ($headerRowIdx === null) $headerRowIdx = 1;
        if ($colMap['barcode'] === null)        $colMap['barcode']        = 'A';
        if ($colMap['name'] === null)           $colMap['name']           = 'B';
        if ($colMap['manufacturer'] === null)   $colMap['manufacturer']   = 'C';
        if ($colMap['packaging_type'] === null) $colMap['packaging_type'] = 'D';
        if ($colMap['quantity'] === null)       $colMap['quantity']       = 'E';

        // 2. Pre-cache products by name for instant matching
        $mapByName = [];
        $cachedProds = $pdo->query("SELECT id, name, manufacturer FROM products")->fetchAll();
        foreach ($cachedProds as $cp) {
            $cleanName = mb_strtolower(trim($cp['name']));
            if ($cleanName !== '') {
                $mapByName[$cleanName][] = [
                    'id' => (int)$cp['id'],
                    'manufacturer' => $cp['manufacturer'] ?? ''
                ];
            }
        }

        $pdo->beginTransaction();

        $insertProductStmt = $pdo->prepare("
            INSERT INTO products (name, manufacturer) 
            VALUES (?, ?)
        ");

        $updateMfrStmt = $pdo->prepare("
            UPDATE products 
            SET manufacturer = ?, updated_at = CURRENT_TIMESTAMP 
            WHERE id = ?
        ");

        $upsertBcStmt = $pdo->prepare("
            INSERT INTO product_barcodes (product_id, barcode, packaging_type, quantity)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(product_id, barcode, quantity) 
            DO UPDATE SET packaging_type = excluded.packaging_type
        ");

        $totalProcessed = 0;
        $addedBarcodes = 0;
        $updatedBarcodes = 0;
        $newProductsCreated = 0;
        $totalRows = count($rows);

        for ($i = $headerRowIdx + 1; $i <= $totalRows; $i++) {
            $row = $rows[$i] ?? null;
            if (!$row) continue;

            $rawBc = trim((string)($row[$colMap['barcode']] ?? ''));
            // Clean barcode of whitespace and non-alphanumeric chars
            $barcode = preg_replace('/[^0-9A-Za-z]/', '', $rawBc);
            $name    = trim((string)($row[$colMap['name']] ?? ''));
            $mfr     = trim((string)($row[$colMap['manufacturer']] ?? ''));
            $unit    = trim((string)($row[$colMap['packaging_type']] ?? ''));
            $qty     = (int)($row[$colMap['quantity']] ?? 1);

            if (!$barcode || !$name) {
                continue;
            }

            // Skip repeated headers
            $bcLower = mb_strtolower($barcode);
            if (in_array($bcLower, ['штрихкод', 'штрих-код', 'barcode', 'шк'])) {
                continue;
            }

            if ($qty <= 0) $qty = 1;
            if (!$unit) {
                $unit = ($qty > 1) ? 'Упаковка' : 'Штука';
            }

            $totalProcessed++;
            $cleanName = mb_strtolower($name);

            // Match product by name
            $matchedPids = [];
            if (isset($mapByName[$cleanName])) {
                foreach ($mapByName[$cleanName] as &$prodRef) {
                    $matchedPids[] = $prodRef['id'];
                    // Update manufacturer if it was missing
                    if ($mfr !== '' && empty($prodRef['manufacturer'])) {
                        $updateMfrStmt->execute([$mfr, $prodRef['id']]);
                        $prodRef['manufacturer'] = $mfr;
                    }
                }
            } else {
                // Product does not exist yet: create it so barcode is never lost
                $insertProductStmt->execute([$name, $mfr !== '' ? $mfr : null]);
                $newPid = (int)$pdo->lastInsertId();
                $mapByName[$cleanName][] = [
                    'id' => $newPid,
                    'manufacturer' => $mfr
                ];
                $matchedPids[] = $newPid;
                $newProductsCreated++;
            }

            // Insert or update barcode for each matched product ID
            foreach ($matchedPids as $pid) {
                $upsertBcStmt->execute([$pid, $barcode, $unit, $qty]);
                if ($upsertBcStmt->rowCount() === 1) {
                    $addedBarcodes++;
                } else {
                    $updatedBarcodes++;
                }
            }
        }

        $pdo->commit();

        jsonResponse([
            'success' => true,
            'message' => 'Реестр штрихкодов успешно загружен',
            'stats' => [
                'total_rows'       => $totalProcessed,
                'added_barcodes'   => $addedBarcodes,
                'updated_barcodes' => $updatedBarcodes,
                'new_products'     => $newProductsCreated
            ]
        ]);
    }

    // =========================================================================
    // ACTION: import (Полный объединенный Excel-файл / Legacy)
    // =========================================================================
    if ($action === 'import') {
        $rows = loadUploadedRows();
        $mode = $_POST['mode'] ?? 'merge'; // 'merge' or 'replace'

        $pdo->beginTransaction();

        if ($mode === 'replace') {
            $pdo->exec("DELETE FROM product_barcodes;");
            $pdo->exec("DELETE FROM products;");
        }

        $findProductStmt = $pdo->prepare("SELECT id, place, our_code, name, manufacturer FROM products WHERE supplier_code = ? OR our_code = ?");
        $insertProductStmt = $pdo->prepare("INSERT INTO products (our_code, supplier_code, name, place, manufacturer) VALUES (?, ?, ?, ?, ?)");
        $updateProductStmt = $pdo->prepare("UPDATE products SET our_code = COALESCE(NULLIF(?, ''), our_code), name = ?, place = COALESCE(NULLIF(?, ''), place), manufacturer = COALESCE(NULLIF(?, ''), manufacturer), updated_at = CURRENT_TIMESTAMP WHERE id = ?");
        
        $upsertBarcodeStmt = $pdo->prepare("
            INSERT INTO product_barcodes (product_id, barcode, packaging_type, quantity) 
            VALUES (?, ?, ?, ?)
            ON CONFLICT(product_id, barcode, quantity) 
            DO UPDATE SET packaging_type = excluded.packaging_type
        ");

        $addedProducts = 0;
        $updatedProducts = 0;
        $addedBarcodes = 0;

        for ($i = 2; $i <= count($rows); $i++) {
            $row = $rows[$i] ?? null;
            if (!$row) continue;

            $ourCode      = trim((string)($row['A'] ?? ''));
            $supplierCode = trim((string)($row['B'] ?? ''));
            $name         = trim((string)($row['C'] ?? ''));

            if (!$supplierCode && !$ourCode && !$name) {
                continue;
            }

            $scLower = mb_strtolower($supplierCode);
            if (in_array($scLower, ['код', 'code', 'код(поставщика)', 'код поставщика', 'артикул'])) {
                continue;
            }

            if (!$supplierCode) $supplierCode = $ourCode;
            if (!$name) $name = 'Товар ' . ($supplierCode ?: $ourCode);

            $bigboxBarcode  = preg_replace('/[^0-9A-Za-z]/', '', trim((string)($row['D'] ?? '')));
            $bigboxQty      = (int)($row['E'] ?? 0);
            $miniboxBarcode = preg_replace('/[^0-9A-Za-z]/', '', trim((string)($row['F'] ?? '')));
            $miniboxQty     = (int)($row['G'] ?? 0);
            $pieceBarcode   = preg_replace('/[^0-9A-Za-z]/', '', trim((string)($row['H'] ?? '')));
            $pieceQty       = (int)($row['I'] ?? 0);
            $place          = trim((string)($row['J'] ?? ''));
            $manufacturer   = trim((string)($row['K'] ?? ''));

            // Check if product already exists
            $findProductStmt->execute([$supplierCode, $ourCode ?: $supplierCode]);
            $existing = $findProductStmt->fetch();

            if ($existing) {
                $productId = $existing['id'];
                $updateProductStmt->execute([$ourCode ?: null, $name, $place ?: null, $manufacturer ?: null, $productId]);
                $updatedProducts++;
            } else {
                $insertProductStmt->execute([$ourCode ?: null, $supplierCode ?: null, $name, $place ?: null, $manufacturer ?: null]);
                $productId = $pdo->lastInsertId();
                $addedProducts++;
            }

            if ($bigboxBarcode && $bigboxQty > 0) {
                $upsertBarcodeStmt->execute([$productId, $bigboxBarcode, 'Бигбокс', $bigboxQty]);
                $addedBarcodes++;
            }

            if ($miniboxBarcode && $miniboxQty > 0) {
                $upsertBarcodeStmt->execute([$productId, $miniboxBarcode, 'Минибокс', $miniboxQty]);
                $addedBarcodes++;
            }

            if ($pieceBarcode) {
                $pQty = $pieceQty > 0 ? $pieceQty : 1;
                $upsertBarcodeStmt->execute([$productId, $pieceBarcode, 'Штучный', $pQty]);
                $addedBarcodes++;
            }
        }

        $pdo->commit();

        jsonResponse([
            'success' => true,
            'message' => 'База данных успешно обновлена из Excel',
            'stats' => [
                'added_products'   => $addedProducts,
                'updated_products' => $updatedProducts,
                'added_barcodes'   => $addedBarcodes
            ]
        ]);
    }

    jsonError('Неизвестное действие (action)');

} catch (Exception $e) {
    if (isset($pdo) && $pdo->inTransaction()) {
        $pdo->rollBack();
    }
    jsonError($e->getMessage(), 500);
}
