<?php
require_once __DIR__ . '/db.php';

try {
    $stmt = $pdo->prepare('SELECT data FROM app_state WHERE id = 1');
    $stmt->execute();
    $row = $stmt->fetch();

    if ($row && $row['data']) {
        $decoded = json_decode($row['data'], true);
        echo json_encode(['success' => true, 'state' => $decoded], JSON_UNESCAPED_UNICODE);
    } else {
        // No data saved yet — front-end will fall back to its built-in defaults.
        echo json_encode(['success' => true, 'state' => null], JSON_UNESCAPED_UNICODE);
    }
} catch (PDOException $e) {
    http_response_code(500);
    echo json_encode(['success' => false, 'error' => $e->getMessage()]);
}
