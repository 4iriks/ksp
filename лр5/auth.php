<?php
require_once __DIR__ . '/functions.php';

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    header('Location: index.php?error=method');
    exit;
}

$login = trim($_POST['login'] ?? '');
$password = trim($_POST['password'] ?? '');

if ($login === '' || $password === '') {
    logAuth($login, 'FAIL_LOGIN', 'empty_fields');
    header('Location: index.php?error=empty');
    exit;
}

if (strlen($login) < 3 || strlen($password) < 6) {
    logAuth($login, 'FAIL_LOGIN', 'invalid_format');
    header('Location: index.php?error=invalid_format');
    exit;
}

$users = getUsers();

if (!isset($users[$login])) {
    logAuth($login, 'FAIL_LOGIN', 'user_not_found');
    header('Location: index.php?error=user_not_found');
    exit;
}

$user = $users[$login];

if (!password_verify($password, $user['password_hash'])) {
    logAuth($login, 'FAIL_LOGIN', 'wrong_password');
    header('Location: index.php?error=wrong_password');
    exit;
}

session_regenerate_id(true);

$_SESSION['user_id'] = $user['id'];
$_SESSION['login'] = $login;
$_SESSION['user_name'] = $user['name'];

logAuth($login, 'SUCCESS_LOGIN');

header('Location: profile.php');
exit;
