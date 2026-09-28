<?php
require_once __DIR__ . '/db.php';

$raw     = file_get_contents('php://input');
$decoded = json_decode($raw, true);

if ($decoded === null && trim($raw) !== 'null') {
    http_response_code(400);
    echo json_encode(['success' => false, 'error' => 'Invalid JSON payload']);
    exit;
}

try {
    $json = json_encode($decoded, JSON_UNESCAPED_UNICODE);

    $stmt = $pdo->prepare(
        'INSERT INTO app_state (id, data, updated_at)
         VALUES (1, :data, NOW())
         ON DUPLICATE KEY UPDATE data = :data2, updated_at = NOW()'
    );
    $stmt->execute(['data' => $json, 'data2' => $json]);

    echo json_encode(['success' => true]);
} catch (PDOException $e) {
    http_response_code(500);
    echo json_encode(['success' => false, 'error' => $e->getMessage()]);
}
