<?php
/**
 * Database connection and initialization helper for SQLite
 */

// Define storage root path
function getDataPath() {
    $envPath = getenv('DATA_PATH');
    if ($envPath) {
        $path = rtrim($envPath, '/\\');
    } else {
        $path = __DIR__ . '/data';
    }
    
    // Ensure standard directories exist
    if (!is_dir($path)) {
        @mkdir($path, 0775, true);
    }
    if (!is_dir($path . '/packing_lists')) {
        @mkdir($path . '/packing_lists', 0775, true);
    }
    if (!is_dir($path . '/archive')) {
        @mkdir($path . '/archive', 0775, true);
    }
    
    return $path;
}

function getDB() {
    static $pdo = null;
    if ($pdo !== null) {
        return $pdo;
    }

    $dataPath = getDataPath();
    $dbFile = $dataPath . '/warehouse.db';
    $isNew = !file_exists($dbFile);

    try {
        $pdo = new PDO("sqlite:" . $dbFile, null, null, [
            PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
            PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
            PDO::ATTR_TIMEOUT => 15
        ]);

        // Enable SQLite WAL mode (Write-Ahead Logging) for high concurrency and performance
        $pdo->exec("PRAGMA journal_mode = WAL;");
        $pdo->exec("PRAGMA synchronous = NORMAL;");
        $pdo->exec("PRAGMA foreign_keys = ON;");

        // Run schema initialization if tables do not exist
        initSchema($pdo);

        return $pdo;
    } catch (Exception $e) {
        http_response_code(500);
        echo json_encode(['success' => false, 'error' => 'Database connection error: ' . $e->getMessage()], JSON_UNESCAPED_UNICODE);
        exit;
    }
}

function initSchema(PDO $pdo) {
    // 1. Products table
    $pdo->exec("CREATE TABLE IF NOT EXISTS products (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        our_code TEXT,
        supplier_code TEXT NOT NULL UNIQUE,
        name TEXT NOT NULL,
        place TEXT,
        manufacturer TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
    );");

    // 2. Barcodes table (1 product -> many packaging types / barcodes)
    $pdo->exec("CREATE TABLE IF NOT EXISTS product_barcodes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
        barcode TEXT NOT NULL,
        packaging_type TEXT DEFAULT 'Штучный',
        quantity INTEGER NOT NULL DEFAULT 1,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(product_id, barcode, quantity)
    );");

    // 3. Packing Lists table
    $pdo->exec("CREATE TABLE IF NOT EXISTS packing_lists (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        filename TEXT UNIQUE NOT NULL,
        original_name TEXT NOT NULL,
        status TEXT DEFAULT 'new', -- 'new', 'in_progress', 'completed'
        total_items INTEGER DEFAULT 0,
        total_expected_qty INTEGER DEFAULT 0,
        total_accepted_qty INTEGER DEFAULT 0,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        completed_at DATETIME
    );");

    // 4. Packing items (lines within a packing list)
    $pdo->exec("CREATE TABLE IF NOT EXISTS packing_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        packing_list_id INTEGER NOT NULL REFERENCES packing_lists(id) ON DELETE CASCADE,
        num TEXT,
        code TEXT NOT NULL,
        name TEXT NOT NULL,
        expected_qty INTEGER NOT NULL DEFAULT 0,
        accepted_qty INTEGER NOT NULL DEFAULT 0,
        notes TEXT,
        UNIQUE(packing_list_id, code)
    );");

    // 5. Scan events log (for audit and sync tracking)
    $pdo->exec("CREATE TABLE IF NOT EXISTS scan_events (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        packing_list_id INTEGER NOT NULL REFERENCES packing_lists(id) ON DELETE CASCADE,
        item_code TEXT NOT NULL,
        barcode TEXT,
        qty INTEGER NOT NULL,
        client_time TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    );");

    // Indexes for fast lookups
    $pdo->exec("CREATE INDEX IF NOT EXISTS idx_products_supplier_code ON products(supplier_code);");
    $pdo->exec("CREATE INDEX IF NOT EXISTS idx_products_our_code ON products(our_code);");
    $pdo->exec("CREATE INDEX IF NOT EXISTS idx_products_manufacturer ON products(manufacturer);");
    $pdo->exec("CREATE INDEX IF NOT EXISTS idx_barcodes_barcode ON product_barcodes(barcode);");
    $pdo->exec("CREATE INDEX IF NOT EXISTS idx_barcodes_product_id ON product_barcodes(product_id);");
    $pdo->exec("CREATE INDEX IF NOT EXISTS idx_packing_items_list ON packing_items(packing_list_id);");
    $pdo->exec("CREATE INDEX IF NOT EXISTS idx_packing_items_code ON packing_items(packing_list_id, code);");
}

function jsonResponse($data, $status = 200) {
    http_response_code($status);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode($data, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);
    exit;
}

function jsonError($message, $status = 400) {
    jsonResponse(['success' => false, 'error' => $message], $status);
}
