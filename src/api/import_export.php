<?php
/**
 * Import and Export barcode database via Excel (XLSX)
 */
require_once __DIR__ . '/../vendor/autoload.php';
require_once __DIR__ . '/../db.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

$action = $_GET['action'] ?? ($_POST['action'] ?? 'stats');

try {
    $pdo = getDB();

    if ($action === 'stats') {
        $productCount = $pdo->query("SELECT COUNT(*) FROM products")->fetchColumn();
        $barcodeCount = $pdo->query("SELECT COUNT(*) FROM product_barcodes")->fetchColumn();
        $manufacturerCount = $pdo->query("SELECT COUNT(DISTINCT manufacturer) FROM products WHERE manufacturer IS NOT NULL AND manufacturer != ''")->fetchColumn();
        
        jsonResponse([
            'success' => true,
            'stats' => [
                'products' => (int)$productCount,
                'barcodes' => (int)$barcodeCount,
                'manufacturers' => (int)$manufacturerCount
            ]
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
            $miniboxBarcodes = []; // Can have multiple miniboxes
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
                    // Minibox or other intermediate pack
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

            // If there are additional miniboxes / barcodes for this product, output extra rows with same supplier_code
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

        $filename = 'barcodeDB_export_' . date('Y-m-d_His') . '.xlsx';
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . $filename . '"');
        header('Cache-Control: max-age=0');

        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('php://output');
        exit;
    }

    if ($action === 'import') {
        if (!isset($_FILES['excel_file']) || $_FILES['excel_file']['error'] !== UPLOAD_ERR_OK) {
            jsonError('Файл не был загружен или произошла ошибка при передаче');
        }

        $tmpFile = $_FILES['excel_file']['tmp_name'];
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
            jsonError('Таблица Excel пуста или содержит только заголовки');
        }

        $mode = $_POST['mode'] ?? 'merge'; // 'merge' or 'replace'

        $pdo->beginTransaction();

        if ($mode === 'replace') {
            $pdo->exec("DELETE FROM product_barcodes;");
            $pdo->exec("DELETE FROM products;");
        }

        $findProductStmt = $pdo->prepare("SELECT id, place, our_code, name, manufacturer FROM products WHERE supplier_code = ?");
        $insertProductStmt = $pdo->prepare("INSERT INTO products (our_code, supplier_code, name, place, manufacturer) VALUES (?, ?, ?, ?, ?)");
        $updateProductStmt = $pdo->prepare("UPDATE products SET our_code = COALESCE(NULLIF(?, ''), our_code), name = ?, place = COALESCE(NULLIF(?, ''), place), manufacturer = COALESCE(NULLIF(?, ''), manufacturer), updated_at = CURRENT_TIMESTAMP WHERE id = ?");
        
        $findBarcodeStmt = $pdo->prepare("SELECT id FROM product_barcodes WHERE product_id = ? AND barcode = ? AND quantity = ?");
        $insertBarcodeStmt = $pdo->prepare("INSERT OR IGNORE INTO product_barcodes (product_id, barcode, packaging_type, quantity) VALUES (?, ?, ?, ?)");

        $addedProducts = 0;
        $updatedProducts = 0;
        $addedBarcodes = 0;

        // Skip header (row 1)
        for ($i = 2; $i <= count($rows); $i++) {
            $row = $rows[$i] ?? null;
            if (!$row) continue;

            $ourCode      = trim((string)($row['A'] ?? ''));
            $supplierCode = trim((string)($row['B'] ?? ''));
            $name         = trim((string)($row['C'] ?? ''));

            if (!$supplierCode && !$ourCode && !$name) {
                continue; // empty row
            }

            // Skip header if it happens to be read as data
            $scLower = mb_strtolower($supplierCode);
            if (in_array($scLower, ['код', 'code', 'код(поставщика)', 'код поставщика', 'артикул'])) {
                continue;
            }

            // We require at least a supplier code or generate one from our code
            if (!$supplierCode) {
                $supplierCode = $ourCode;
            }
            if (!$name) {
                $name = 'Товар ' . $supplierCode;
            }

            $bigboxBarcode  = trim((string)($row['D'] ?? ''));
            $bigboxQty      = (int)($row['E'] ?? 0);
            $miniboxBarcode = trim((string)($row['F'] ?? ''));
            $miniboxQty     = (int)($row['G'] ?? 0);
            $pieceBarcode   = trim((string)($row['H'] ?? ''));
            $pieceQty       = (int)($row['I'] ?? 0);
            $place          = trim((string)($row['J'] ?? ''));
            $manufacturer   = trim((string)($row['K'] ?? ''));

            // Check if product already exists
            $findProductStmt->execute([$supplierCode]);
            $existing = $findProductStmt->fetch();

            if ($existing) {
                $productId = $existing['id'];
                $updateProductStmt->execute([$ourCode, $name, $place, $manufacturer, $productId]);
                $updatedProducts++;
            } else {
                $insertProductStmt->execute([$ourCode, $supplierCode, $name, $place, $manufacturer]);
                $productId = $pdo->lastInsertId();
                $addedProducts++;
            }

            // Insert Bigbox barcode if present
            if ($bigboxBarcode && $bigboxQty > 0) {
                $findBarcodeStmt->execute([$productId, $bigboxBarcode, $bigboxQty]);
                if (!$findBarcodeStmt->fetch()) {
                    $insertBarcodeStmt->execute([$productId, $bigboxBarcode, 'Бигбокс', $bigboxQty]);
                    $addedBarcodes++;
                }
            }

            // Insert Minibox barcode if present
            if ($miniboxBarcode && $miniboxQty > 0) {
                $findBarcodeStmt->execute([$productId, $miniboxBarcode, $miniboxQty]);
                if (!$findBarcodeStmt->fetch()) {
                    $insertBarcodeStmt->execute([$productId, $miniboxBarcode, 'Минибокс', $miniboxQty]);
                    $addedBarcodes++;
                }
            }

            // Insert Piece barcode if present
            if ($pieceBarcode) {
                $pQty = $pieceQty > 0 ? $pieceQty : 1;
                $findBarcodeStmt->execute([$productId, $pieceBarcode, $pQty]);
                if (!$findBarcodeStmt->fetch()) {
                    $insertBarcodeStmt->execute([$productId, $pieceBarcode, 'Штучный', $pQty]);
                    $addedBarcodes++;
                }
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
