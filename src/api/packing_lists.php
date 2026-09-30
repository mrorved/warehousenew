<?php
/**
 * Packing lists management: List, Upload, View, Complete, Reset (re-accept), Delete, Export Report
 */
require_once __DIR__ . '/../vendor/autoload.php';
require_once __DIR__ . '/../db.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Fill;

$action = $_GET['action'] ?? ($_POST['action'] ?? 'list');

try {
    $pdo = getDB();
    $dataPath = getDataPath();

    if ($action === 'list') {
        $stmt = $pdo->query("SELECT * FROM packing_lists ORDER BY created_at DESC");
        $lists = $stmt->fetchAll();

        jsonResponse([
            'success' => true,
            'lists' => $lists
        ]);
    }

    if ($action === 'get_items') {
        $listId = (int)($_GET['id'] ?? 0);
        if (!$listId) {
            jsonError('ID упаковочного листа не указан');
        }

        $listStmt = $pdo->prepare("SELECT * FROM packing_lists WHERE id = ?");
        $listStmt->execute([$listId]);
        $list = $listStmt->fetch();
        if (!$list) {
            jsonError('Упаковочный лист не найден', 404);
        }

        // Get items joined with products to fetch current place and our_code
        $itemsStmt = $pdo->prepare("
            SELECT 
                pi.id,
                pi.num,
                pi.code,
                pi.name,
                pi.expected_qty,
                pi.accepted_qty,
                pi.notes,
                p.our_code,
                COALESCE(p.place, '') as place,
                COALESCE(p.manufacturer, '') as manufacturer
            FROM packing_items pi
            LEFT JOIN products p ON (pi.code = p.supplier_code OR pi.code = p.our_code OR pi.name = p.name)
            WHERE pi.packing_list_id = ?
              AND (pi.expected_qty > 0 OR pi.accepted_qty > 0)
            ORDER BY CAST(pi.num AS INTEGER) ASC, pi.id ASC
        ");
        $itemsStmt->execute([$listId]);
        $items = $itemsStmt->fetchAll();

        // Also fetch all barcodes for fast local client lookup
        $barcodesStmt = $pdo->query("
            SELECT 
                pb.barcode,
                pb.quantity,
                pb.packaging_type,
                p.supplier_code,
                p.our_code,
                p.name,
                p.place
            FROM product_barcodes pb
            JOIN products p ON pb.product_id = p.id
        ");
        $allBarcodes = $barcodesStmt->fetchAll();

        jsonResponse([
            'success' => true,
            'list' => $list,
            'items' => $items,
            'barcodes' => $allBarcodes
        ]);
    }

    if ($action === 'upload') {
        if (!isset($_FILES['file']) || $_FILES['file']['error'] !== UPLOAD_ERR_OK) {
            jsonError('Файл не был загружен или поврежден');
        }

        $origName = $_FILES['file']['name'];
        $ext = strtolower(pathinfo($origName, PATHINFO_EXTENSION));

        if (!in_array($ext, ['xlsx', 'xls', 'csv'])) {
            jsonError('Поддерживаются только файлы Excel (.xlsx, .xls) и CSV');
        }

        // Safe unique file name
        $safeName = date('Ymd_His') . '_' . preg_replace('/[^a-zA-Z0-9_\-\.]/u', '_', $origName);
        $targetDir = $dataPath . '/packing_lists';
        $targetPath = $targetDir . '/' . $safeName;

        if (!move_uploaded_file($_FILES['file']['tmp_name'], $targetPath)) {
            jsonError('Не удалось сохранить файл на сервере');
        }

        // Parse with PhpSpreadsheet
        $reader = IOFactory::createReaderForFile($targetPath);
        if ($reader instanceof \PhpOffice\PhpSpreadsheet\Reader\Csv) {
            $reader->setInputEncoding('UTF-8');
            $reader->setDelimiter(',');
        }
        $reader->setReadDataOnly(true);
        $spreadsheet = $reader->load($targetPath);
        $sheet = $spreadsheet->getActiveSheet();
        $rows = $sheet->toArray(null, true, true, true);

        if (count($rows) < 2) {
            @unlink($targetPath);
            jsonError('Файл пуст');
        }

        // Dynamically find header row in the first 20 rows
        $headerRowIdx = null;
        $colMap = [
            'num'  => null,
            'code' => null,
            'name' => null,
            'qty'  => null
        ];

        foreach ($rows as $rIdx => $row) {
            if ($rIdx > 20) break;
            $hasCode = false;
            $hasName = false;
            $tempMap = ['num' => null, 'code' => null, 'name' => null, 'qty' => null];

            foreach ($row as $col => $val) {
                $v = mb_strtolower(trim((string)$val));
                if (str_contains($v, 'код') || str_contains($v, 'code') || str_contains($v, 'артикул')) {
                    if ($tempMap['code'] === null) {
                        $tempMap['code'] = $col;
                        $hasCode = true;
                    }
                } elseif (str_contains($v, 'назв') || str_contains($v, 'наим') || str_contains($v, 'name') || str_contains($v, 'товар')) {
                    if ($tempMap['name'] === null) {
                        $tempMap['name'] = $col;
                        $hasName = true;
                    }
                } elseif (str_contains($v, 'кол') || str_contains($v, 'qty') || str_contains($v, 'заявл')) {
                    if ($tempMap['qty'] === null) {
                        $tempMap['qty'] = $col;
                    }
                } elseif (str_contains($v, '№') || str_contains($v, 'номер') || $v === 'n' || $v === 'no') {
                    if ($tempMap['num'] === null) {
                        $tempMap['num'] = $col;
                    }
                }
            }

            if ($hasCode && ($hasName || $tempMap['qty'] !== null)) {
                $headerRowIdx = $rIdx;
                $colMap = $tempMap;
                break;
            }
        }

        // Fallbacks if not found by name
        if ($headerRowIdx === null) {
            $headerRowIdx = 1;
        }
        if ($colMap['code'] === null) $colMap['code'] = 'B';
        if ($colMap['name'] === null) $colMap['name'] = 'C';
        if ($colMap['qty'] === null)  $colMap['qty']  = 'D';
        if ($colMap['num'] === null)  $colMap['num']  = 'A';

        // Aggregate items by code (starting strictly after header row)
        $aggregated = [];
        $totalRows = count($rows);
        for ($i = $headerRowIdx + 1; $i <= $totalRows; $i++) {
            $row = $rows[$i] ?? null;
            if (!$row) continue;

            $code = trim((string)($row[$colMap['code']] ?? ''));
            if (!$code) continue;

            $codeLower = mb_strtolower($code);
            // Ignore header text or metadata in item rows
            if (in_array($codeLower, ['код', 'code', 'код товара', 'код(поставщика)', 'код(наш)', 'артикул', '№', 'номер'])) {
                continue;
            }
            if (str_contains($codeLower, 'дата формирования') || str_contains($codeLower, 'итого')) {
                continue;
            }

            $name = trim((string)($row[$colMap['name']] ?? ''));
            $nameLower = mb_strtolower($name);
            if (in_array($nameLower, ['название', 'наименование', 'товар', 'name', 'наименование товара'])) {
                continue;
            }

            $num = trim((string)($row[$colMap['num']] ?? ''));
            $qty = (int)($row[$colMap['qty']] ?? 0);

            if (isset($aggregated[$code])) {
                $aggregated[$code]['expected_qty'] += $qty;
            } else {
                $aggregated[$code] = [
                    'num' => $num ?: (count($aggregated) + 1),
                    'code' => $code,
                    'name' => $name ?: 'Позиция ' . $code,
                    'expected_qty' => $qty,
                    'notes' => ''
                ];
            }
        }

        if (empty($aggregated)) {
            @unlink($targetPath);
            jsonError('Не удалось найти строки товаров в файле');
        }

        $totalExpected = array_sum(array_column($aggregated, 'expected_qty'));

        $pdo->beginTransaction();

        $listStmt = $pdo->prepare("
            INSERT INTO packing_lists (filename, original_name, status, total_items, total_expected_qty, total_accepted_qty)
            VALUES (?, ?, 'new', ?, ?, 0)
        ");
        $listStmt->execute([$safeName, $origName, count($aggregated), $totalExpected]);
        $listId = $pdo->lastInsertId();

        $itemStmt = $pdo->prepare("
            INSERT INTO packing_items (packing_list_id, num, code, name, expected_qty, accepted_qty)
            VALUES (?, ?, ?, ?, ?, 0)
        ");

        foreach ($aggregated as $item) {
            $itemStmt->execute([$listId, $item['num'], $item['code'], $item['name'], $item['expected_qty']]);
        }

        $pdo->commit();

        jsonResponse([
            'success' => true,
            'message' => 'Упаковочный лист успешно загружен',
            'list_id' => $listId,
            'items_count' => count($aggregated),
            'total_qty' => $totalExpected
        ]);
    }

    if ($action === 'complete') {
        $listId = (int)($_POST['list_id'] ?? 0);
        if (!$listId) jsonError('ID листа не указан');

        $pdo->beginTransaction();

        // Calculate totals
        $totals = $pdo->query("SELECT SUM(accepted_qty) as total_accepted FROM packing_items WHERE packing_list_id = $listId")->fetch();
        $totalAccepted = (int)($totals['total_accepted'] ?? 0);

        $stmt = $pdo->prepare("
            UPDATE packing_lists 
            SET status = 'completed', total_accepted_qty = ?, completed_at = CURRENT_TIMESTAMP, updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        ");
        $stmt->execute([$totalAccepted, $listId]);

        // Move file to archive directory
        $listInfo = $pdo->query("SELECT filename FROM packing_lists WHERE id = $listId")->fetch();
        if ($listInfo) {
            $source = $dataPath . '/packing_lists/' . $listInfo['filename'];
            $archive = $dataPath . '/archive/' . $listInfo['filename'];
            if (file_exists($source)) {
                @rename($source, $archive);
            }
        }

        $pdo->commit();

        jsonResponse([
            'success' => true,
            'message' => 'Приёмка успешно завершена и перемещена в архив'
        ]);
    }

    if ($action === 'reset') { // «Перепринять»
        $listId = (int)($_POST['list_id'] ?? 0);
        if (!$listId) jsonError('ID листа не указан');

        $pdo->beginTransaction();

        // 1. Delete all unplanned / excess items added during scanning (not in original sheet)
        $pdo->prepare("DELETE FROM packing_items WHERE packing_list_id = ? AND (expected_qty = 0 OR num = '+')")->execute([$listId]);

        // 2. Reset accepted_qty to 0 for all original sheet items
        $pdo->prepare("UPDATE packing_items SET accepted_qty = 0 WHERE packing_list_id = ?")->execute([$listId]);

        // 3. Recalculate original items count
        $origCount = (int)$pdo->query("SELECT COUNT(*) FROM packing_items WHERE packing_list_id = $listId")->fetchColumn();

        // 4. Reset sheet status and counters
        $pdo->prepare("
            UPDATE packing_lists 
            SET status = 'new', 
                total_accepted_qty = 0, 
                total_items = ?,
                completed_at = NULL, 
                updated_at = CURRENT_TIMESTAMP 
            WHERE id = ?
        ")->execute([$origCount, $listId]);

        // 5. Clear scan logs
        $pdo->prepare("DELETE FROM scan_events WHERE packing_list_id = ?")->execute([$listId]);

        // Move file back from archive if it was archived
        $listInfo = $pdo->query("SELECT filename FROM packing_lists WHERE id = $listId")->fetch();
        if ($listInfo) {
            $archive = $dataPath . '/archive/' . $listInfo['filename'];
            $source = $dataPath . '/packing_lists/' . $listInfo['filename'];
            if (file_exists($archive) && !file_exists($source)) {
                @rename($archive, $source);
            }
        }

        $pdo->commit();

        jsonResponse([
            'success' => true,
            'message' => 'Приёмка сброшена. Внеплановые позиции удалены.'
        ]);
    }

    if ($action === 'delete') {
        $listId = (int)($_POST['list_id'] ?? 0);
        if (!$listId) jsonError('ID листа не указан');

        $listInfo = $pdo->query("SELECT filename FROM packing_lists WHERE id = $listId")->fetch();
        if ($listInfo) {
            @unlink($dataPath . '/packing_lists/' . $listInfo['filename']);
            @unlink($dataPath . '/archive/' . $listInfo['filename']);
        }

        $pdo->prepare("DELETE FROM packing_lists WHERE id = ?")->execute([$listId]);

        jsonResponse([
            'success' => true,
            'message' => 'Упаковочный лист удален'
        ]);
    }

    if ($action === 'export_report') { // Акт расхождений Excel
        $listId = (int)($_GET['list_id'] ?? 0);
        if (!$listId) jsonError('ID листа не указан');

        $listStmt = $pdo->prepare("SELECT * FROM packing_lists WHERE id = ?");
        $listStmt->execute([$listId]);
        $list = $listStmt->fetch();
        if (!$list) jsonError('Лист не найден');

        $itemsStmt = $pdo->prepare("
            SELECT pi.*, p.place, p.our_code 
            FROM packing_items pi
            LEFT JOIN products p ON (pi.code = p.supplier_code OR pi.code = p.our_code OR pi.name = p.name)
            WHERE pi.packing_list_id = ?
            ORDER BY CAST(pi.num AS INTEGER) ASC
        ");
        $itemsStmt->execute([$listId]);
        $items = $itemsStmt->fetchAll();

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('Акт приемки');

        // Header info
        $sheet->setCellValue('A1', 'АКТ ПРИЕМКИ ТОВАРА');
        $sheet->setCellValue('A2', 'Файл: ' . $list['original_name']);
        $sheet->setCellValue('A3', 'Дата: ' . ($list['completed_at'] ?? date('Y-m-d H:i:s')));
        $sheet->getStyle('A1')->getFont()->setBold(true)->setSize(14);

        $headers = ['A5' => '№', 'B5' => 'Наш код', 'C5' => 'Код поставщика', 'D5' => 'Наименование', 'E5' => 'Место', 'F5' => 'Заявлено', 'G5' => 'Принято', 'H5' => 'Расхождение', 'I5' => 'Статус'];
        foreach ($headers as $cell => $val) {
            $sheet->setCellValue($cell, $val);
            $sheet->getStyle($cell)->getFont()->setBold(true);
        }

        $rowNum = 6;
        foreach ($items as $item) {
            $diff = $item['accepted_qty'] - $item['expected_qty'];
            $statusText = 'Совпадает';
            $color = 'E8F5E9'; // light green

            if ($item['expected_qty'] == 0 && $item['accepted_qty'] > 0) {
                $statusText = 'Сверх плана (лишняя позиция)';
                $color = 'FFF3E0'; // light orange
            } elseif ($diff < 0) {
                $statusText = 'Недостача (' . $diff . ')';
                $color = 'FFEBEE'; // light red
            } elseif ($diff > 0) {
                $statusText = 'Излишек (+' . $diff . ')';
                $color = 'FFF8E1'; // light yellow
            }

            $sheet->setCellValue("A$rowNum", $item['num']);
            $sheet->setCellValueExplicit("B$rowNum", $item['our_code'] ?? '', \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
            $sheet->setCellValueExplicit("C$rowNum", $item['code'], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
            $sheet->setCellValue("D$rowNum", $item['name']);
            $sheet->setCellValue("E$rowNum", $item['place'] ?? '');
            $sheet->setCellValue("F$rowNum", $item['expected_qty']);
            $sheet->setCellValue("G$rowNum", $item['accepted_qty']);
            $sheet->setCellValue("H$rowNum", $diff);
            $sheet->setCellValue("I$rowNum", $statusText);

            $sheet->getStyle("A$rowNum:I$rowNum")->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB($color);
            $rowNum++;
        }

        foreach (range('A', 'I') as $col) {
            $sheet->getColumnDimension($col)->setAutoSize(true);
        }

        $filename = 'Акт_приемки_' . preg_replace('/[^a-zA-Z0-9_\-]/u', '_', $list['original_name']) . '.xlsx';
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . $filename . '"');
        header('Cache-Control: max-age=0');

        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('php://output');
        exit;
    }

    jsonError('Неизвестное действие (action)');

} catch (Exception $e) {
    if (isset($pdo) && $pdo->inTransaction()) {
        $pdo->rollBack();
    }
    jsonError($e->getMessage(), 500);
}
