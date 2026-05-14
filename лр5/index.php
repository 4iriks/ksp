<?php
require_once __DIR__ . '/functions.php';

if (isLoggedIn()) {
    header('Location: profile.php');
    exit;
}

$errors = [
    'empty' => 'Введите логин и пароль.',
    'invalid_format' => 'Логин или пароль введены некорректно.',
    'user_not_found' => 'Пользователь с таким логином не найден.',
    'wrong_password' => 'Неверный пароль.',
    'auth_required' => 'Для доступа к странице необходимо войти.',
    'method' => 'Форма должна быть отправлена методом POST.',
];

$error = $_GET['error'] ?? '';
$message = $errors[$error] ?? '';
$isLogout = isset($_GET['logout']);
?>
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <title>Авторизация — Каталог смартфонов</title>
    <link rel="stylesheet" type="text/css" href="style.css">
</head>
<body>

<ul class="menu">
    <li><a href="index.php">Вход</a></li>
</ul>

<hr>

<h1>Авторизация</h1>

<div class="auth-card">
    <h2>Вход в каталог</h2>
    <p>Введите логин и пароль, чтобы открыть защищённые страницы каталога смартфонов.</p>

    <?php if ($message !== ''): ?>
        <p class="message message-error"><?= e($message) ?></p>
    <?php endif; ?>

    <?php if ($isLogout): ?>
        <p class="message message-success">Вы успешно вышли из системы.</p>
    <?php endif; ?>

    <form class="auth-form" action="auth.php" method="post">
        <label for="login">Логин</label>
        <input type="text" id="login" name="login" required minlength="3" maxlength="30" autocomplete="username">

        <label for="password">Пароль</label>
        <input type="password" id="password" name="password" required minlength="6" autocomplete="current-password">

        <button class="btn btn-add" type="submit">Войти</button>
    </form>

    <p class="auth-hint">Тестовые данные: <strong>admin / password</strong> или <strong>user / password</strong>.</p>
</div>

<hr>

<p class="footer"><small>&copy; 2026 Каталог смартфонов. Все права защищены.</small></p>

</body>
</html>
