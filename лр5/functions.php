<?php
session_start();

function getUsers(): array
{
    return require __DIR__ . '/users.php';
}

function e(string $value): string
{
    return htmlspecialchars($value, ENT_QUOTES, 'UTF-8');
}

function isLoggedIn(): bool
{
    return isset($_SESSION['user_id']);
}

function currentUserName(): string
{
    return $_SESSION['user_name'] ?? 'Пользователь';
}

function currentUserLogin(): string
{
    return $_SESSION['login'] ?? '';
}

function requireAuth(): void
{
    if (!isLoggedIn()) {
        header('Location: index.php?error=auth_required');
        exit;
    }
}

function logAuth(string $login, string $action, string $info = ''): void
{
    $dir = __DIR__ . '/logs';
    $file = $dir . '/auth.log';

    if (!is_dir($dir)) {
        mkdir($dir, 0777, true);
    }

    $time = date('Y-m-d H:i:s');
    $ip = $_SERVER['REMOTE_ADDR'] ?? 'unknown';
    $loginValue = $login !== '' ? $login : 'empty';

    $line = $time
        . ' | ip=' . $ip
        . ' | login=' . $loginValue
        . ' | action=' . $action;

    if ($info !== '') {
        $line .= ' | info=' . $info;
    }

    $line .= PHP_EOL;

    file_put_contents($file, $line, FILE_APPEND);
}
