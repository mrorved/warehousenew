<?php
/**
 * Packing state synchronization API (Offline-first Outbox sync handler)
 */
require_once __DIR__ . '/../db.php';

$rawInput = file_get_contents('php://input');
$input = json_decode($rawInput, true);

if (!$input) {
    jsonError('Неверный формат JSON запроса');
}

$action = $input['action'] ?? 'sync';
$pdo = getDB();

try {
    if ($action === 'sync') {
        $listId = (int)($input['list_id'] ?? 0);
        if (!$listId) {
            jsonError('list_id обязателен для синхронизации');
        }

        $itemsMap = $input['items'] ?? []; // Map of code => accepted_qty
        $scanEvents = $input['events'] ?? []; // Array of { code, barcode, qty, client_time }

        $pdo->beginTransaction();

        // 1. Update quantities in packing_items
        $updateStmt = $pdo->prepare("
            UPDATE packing_items 
            SET accepted_qty = ? 
            WHERE packing_list_id = ? AND code = ?
        ");

        // Prepare insert statement for items that were scanned but not in original packing list (excess/unplanned)
        $insertUnplannedStmt = $pdo->prepare("
            INSERT INTO packing_items (packing_list_id, num, code, name, expected_qty, accepted_qty)
            VALUES (?, '+', ?, ?, 0, ?)
            ON CONFLICT(packing_list_id, code) DO UPDATE SET accepted_qty = excluded.accepted_qty
        ");

        $findProductStmt = $pdo->prepare("SELECT name FROM products WHERE supplier_code = ?");

        foreach ($itemsMap as $code => $qty) {
            $code = (string)$code;
            $qty = max(0, (int)$qty);

            if ($qty === 0) {
                // If it is an unplanned item with 0 accepted qty, delete it from the list
                $pdo->prepare("DELETE FROM packing_items WHERE packing_list_id = ? AND code = ? AND (expected_qty = 0 OR num = '+')")
                    ->execute([$listId, $code]);
            }

            $updateStmt->execute([$qty, $listId, $code]);
            if ($updateStmt->rowCount() === 0 && $qty > 0) {
                // Not found in original sheet, find product name if in DB
                $findProductStmt->execute([$code]);
                $p = $findProductStmt->fetch();
                $name = $p ? $p['name'] : 'Товар ' . $code;
                $insertUnplannedStmt->execute([$listId, $code, $name, $qty]);
            }
        }

        // 2. Insert scan events log
        if (!empty($scanEvents)) {
            $logStmt = $pdo->prepare("
                INSERT INTO scan_events (packing_list_id, item_code, barcode, qty, client_time)
                VALUES (?, ?, ?, ?, ?)
            ");
            foreach ($scanEvents as $ev) {
                $logStmt->execute([
                    $listId,
                    (string)($ev['code'] ?? ''),
                    (string)($ev['barcode'] ?? ''),
                    (int)($ev['qty'] ?? 1),
                    (string)($ev['client_time'] ?? '')
                ]);
            }
        }

        // 3. Recalculate totals and set status to 'in_progress' if currently 'new'
        $totalAccepted = (int)$pdo->query("SELECT SUM(accepted_qty) FROM packing_items WHERE packing_list_id = $listId")->fetchColumn();
        
        $pdo->prepare("
            UPDATE packing_lists 
            SET total_accepted_qty = ?, 
                status = CASE WHEN status = 'new' THEN 'in_progress' ELSE status END,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        ")->execute([$totalAccepted, $listId]);

        $pdo->commit();

        jsonResponse([
            'success' => true,
            'message' => 'Синхронизировано',
            'server_time' => time(),
            'total_accepted' => $totalAccepted
        ]);
    }

    if ($action === 'set_item_qty') {
        $listId = (int)($input['list_id'] ?? 0);
        $code = trim((string)($input['code'] ?? ''));
        $qty = max(0, (int)($input['accepted_qty'] ?? 0));

        if (!$listId || !$code) {
            jsonError('Параметры list_id и code обязательны');
        }

        $pdo->beginTransaction();

        $stmt = $pdo->prepare("UPDATE packing_items SET accepted_qty = ? WHERE packing_list_id = ? AND code = ?");
        $stmt->execute([$qty, $listId, $code]);

        $totalAccepted = (int)$pdo->query("SELECT SUM(accepted_qty) FROM packing_items WHERE packing_list_id = $listId")->fetchColumn();
        
        $pdo->prepare("
            UPDATE packing_lists 
            SET total_accepted_qty = ?, 
                status = CASE WHEN status = 'new' THEN 'in_progress' ELSE status END,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        ")->execute([$totalAccepted, $listId]);

        $pdo->commit();

        jsonResponse([
            'success' => true,
            'accepted_qty' => $qty,
            'total_accepted' => $totalAccepted
        ]);
    }

    jsonError('Неизвестное действие');

} catch (Exception $e) {
    if (isset($pdo) && $pdo->inTransaction()) {
        $pdo->rollBack();
    }
    jsonError($e->getMessage(), 500);
}
