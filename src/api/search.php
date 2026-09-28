<?php
/**
 * Fast product and location search by barcode, code, or name
 */
require_once __DIR__ . '/../db.php';

$query = trim($_GET['q'] ?? '');
if (str_ends_with($query, '+ET')) {
    $query = trim(substr($query, 0, -3));
}
if (!$query) {
    jsonResponse(['success' => true, 'results' => []]);
}

$pdo = getDB();

try {
    $normQuery = ltrim($query, '0');
    $results = [];

    // 1. Direct barcode match
    $bcStmt = $pdo->prepare("
        SELECT 
            p.id,
            p.our_code,
            p.supplier_code,
            p.name,
            p.place,
            p.manufacturer,
            pb.barcode as matched_barcode,
            pb.packaging_type,
            pb.quantity as matched_qty
        FROM product_barcodes pb
        JOIN products p ON pb.product_id = p.id
        WHERE pb.barcode = ?
    ");
    $bcStmt->execute([$query]);
    $bcMatches = $bcStmt->fetchAll();

    if (!empty($bcMatches)) {
        // Collect all barcodes for each matched product
        $prodBarcodeStmt = $pdo->prepare("SELECT barcode, packaging_type, quantity FROM product_barcodes WHERE product_id = ? ORDER BY quantity DESC");
        foreach ($bcMatches as $m) {
            $prodBarcodeStmt->execute([$m['id']]);
            $m['all_barcodes'] = $prodBarcodeStmt->fetchAll();
            $results[] = $m;
        }
    } else {
        // 2. Search by our_code or supplier_code
        $codeStmt = $pdo->prepare("
            SELECT 
                p.id,
                p.our_code,
                p.supplier_code,
                p.name,
                p.place,
                p.manufacturer
            FROM products p
            WHERE p.supplier_code = ? 
               OR p.our_code = ? 
               OR LTRIM(p.supplier_code, '0') = ? 
               OR LTRIM(p.our_code, '0') = ?
            LIMIT 10
        ");
        $codeStmt->execute([$query, $query, $normQuery, $normQuery]);
        $codeMatches = $codeStmt->fetchAll();

        $prodBarcodeStmt = $pdo->prepare("SELECT barcode, packaging_type, quantity FROM product_barcodes WHERE product_id = ? ORDER BY quantity DESC");
        foreach ($codeMatches as $m) {
            $prodBarcodeStmt->execute([$m['id']]);
            $m['all_barcodes'] = $prodBarcodeStmt->fetchAll();
            $m['matched_barcode'] = null;
            $m['matched_qty'] = null;
            $results[] = $m;
        }

        // 3. Fallback: search by name
        if (empty($results) && mb_strlen($query) >= 3) {
            $nameStmt = $pdo->prepare("
                SELECT 
                    p.id,
                    p.our_code,
                    p.supplier_code,
                    p.name,
                    p.place,
                    p.manufacturer
                FROM products p
                WHERE p.name LIKE ?
                LIMIT 15
            ");
            $nameStmt->execute(['%' . $query . '%']);
            $nameMatches = $nameStmt->fetchAll();
            foreach ($nameMatches as $m) {
                $prodBarcodeStmt->execute([$m['id']]);
                $m['all_barcodes'] = $prodBarcodeStmt->fetchAll();
                $results[] = $m;
            }
        }
    }

    jsonResponse([
        'success' => true,
        'query' => $query,
        'results' => $results
    ]);

} catch (Exception $e) {
    jsonError($e->getMessage(), 500);
}
