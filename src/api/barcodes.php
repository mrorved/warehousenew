<?php
/**
 * Barcode database management and offline sync cache
 */
require_once __DIR__ . '/../db.php';

$action = $_GET['action'] ?? ($_POST['action'] ?? 'all');
$pdo = getDB();

try {
    if ($action === 'all') {
        // Full barcode list for client-side offline cache
        $stmt = $pdo->query("
            SELECT 
                pb.id as barcode_id,
                pb.barcode,
                pb.packaging_type,
                pb.quantity as box_qty,
                p.id as product_id,
                p.our_code,
                p.supplier_code,
                p.name,
                COALESCE(p.place, '') as place,
                COALESCE(p.manufacturer, '') as manufacturer
            FROM product_barcodes pb
            JOIN products p ON pb.product_id = p.id
        ");
        $all = $stmt->fetchAll();

        jsonResponse([
            'success' => true,
            'count' => count($all),
            'barcodes' => $all
        ]);
    }

    if ($action === 'manufacturers') {
        $stmt = $pdo->query("
            SELECT DISTINCT manufacturer 
            FROM products 
            WHERE manufacturer IS NOT NULL AND manufacturer != '' 
            ORDER BY manufacturer ASC
        ");
        $manufacturers = $stmt->fetchAll(PDO::FETCH_COLUMN);

        jsonResponse([
            'success' => true,
            'manufacturers' => $manufacturers
        ]);
    }

    if ($action === 'by_manufacturer') {
        $manufacturer = trim($_GET['manufacturer'] ?? '');
        $stmt = $pdo->prepare("
            SELECT p.* 
            FROM products p 
            WHERE p.manufacturer = ? 
            ORDER BY p.name ASC
        ");
        $stmt->execute([$manufacturer]);
        $products = $stmt->fetchAll();

        $barcodeStmt = $pdo->prepare("SELECT * FROM product_barcodes WHERE product_id = ? ORDER BY quantity DESC");
        foreach ($products as &$prod) {
            $barcodeStmt->execute([$prod['id']]);
            $prod['barcodes'] = $barcodeStmt->fetchAll();
        }

        jsonResponse([
            'success' => true,
            'products' => $products
        ]);
    }

    if ($action === 'update_product') {
        $input = json_decode(file_get_contents('php://input'), true);
        $productId = (int)($input['product_id'] ?? 0);
        $place = trim($input['place'] ?? '');
        $ourCode = trim($input['our_code'] ?? '');
        $name = trim($input['name'] ?? '');

        if (!$productId) jsonError('product_id обязателен');

        $stmt = $pdo->prepare("
            UPDATE products 
            SET place = ?, our_code = ?, name = ?, updated_at = CURRENT_TIMESTAMP 
            WHERE id = ?
        ");
        $stmt->execute([$place, $ourCode, $name, $productId]);

        jsonResponse(['success' => true, 'message' => 'Товар обновлен']);
    }

    if ($action === 'add_barcode') {
        $input = json_decode(file_get_contents('php://input'), true);
        $productId = (int)($input['product_id'] ?? 0);
        $barcode = trim($input['barcode'] ?? '');
        $type = trim($input['packaging_type'] ?? 'Минибокс');
        $qty = max(1, (int)($input['quantity'] ?? 1));

        if (!$productId || !$barcode) {
            jsonError('product_id и barcode обязательны');
        }

        $stmt = $pdo->prepare("
            INSERT INTO product_barcodes (product_id, barcode, packaging_type, quantity)
            VALUES (?, ?, ?, ?)
        ");
        $stmt->execute([$productId, $barcode, $type, $qty]);

        jsonResponse(['success' => true, 'message' => 'Штрихкод добавлен', 'id' => $pdo->lastInsertId()]);
    }

    if ($action === 'delete_barcode') {
        $input = json_decode(file_get_contents('php://input'), true);
        $barcodeId = (int)($input['barcode_id'] ?? 0);

        if (!$barcodeId) jsonError('barcode_id обязателен');

        $pdo->prepare("DELETE FROM product_barcodes WHERE id = ?")->execute([$barcodeId]);

        jsonResponse(['success' => true, 'message' => 'Штрихкод удален']);
    }

    jsonError('Неизвестное действие');

} catch (Exception $e) {
    jsonError($e->getMessage(), 500);
}
