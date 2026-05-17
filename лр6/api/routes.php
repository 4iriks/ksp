<?php
declare(strict_types=1);

function sendNotFound(): void
{
    http_response_code(404);
    echo json_encode([
        'status' => 'error',
        'message' => 'Endpoint не найден',
    ], JSON_UNESCAPED_UNICODE);
}

function handleRoutes(UserController $controller): void
{
    $method = $_SERVER['REQUEST_METHOD'];
    $path = parse_url($_SERVER['REQUEST_URI'], PHP_URL_PATH) ?? '/';
    $path = preg_replace('#^/api#', '', $path);
    $path = trim($path, '/');
    $parts = $path === '' ? [] : explode('/', $path);

    if (($parts[0] ?? '') !== 'v1') {
        $controller->info();
        return;
    }

    $resource = $parts[1] ?? '';
    $id = isset($parts[2]) ? (int)$parts[2] : null;

    if ($method === 'POST' && $resource === 'register') {
        $controller->register();
        return;
    }

    if ($method === 'POST' && $resource === 'login') {
        $controller->login();
        return;
    }

    if ($method === 'GET' && $resource === 'users' && $id === null) {
        $controller->index();
        return;
    }

    if ($method === 'GET' && $resource === 'users' && $id !== null) {
        $controller->show($id);
        return;
    }

    if (($method === 'PUT' || $method === 'PATCH') && $resource === 'users' && $id !== null) {
        $controller->updatePassword($id);
        return;
    }

    if ($method === 'DELETE' && $resource === 'users' && $id !== null) {
        $controller->delete($id);
        return;
    }

    sendNotFound();
}

