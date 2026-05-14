<?php
require_once __DIR__ . '/functions.php';
requireAuth();
?>
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <title>Личный кабинет — Каталог смартфонов</title>
    <link rel="stylesheet" type="text/css" href="style.css">
</head>
<body>

<ul class="menu">
    <li><a href="profile.php">Личный кабинет</a></li>
    <li><a href="catalog.php">Каталог</a></li>
    <li><a class="logout-link" href="logout.php">Выход</a></li>
</ul>

<p class="user-panel">Вы вошли как: <strong><?= e(currentUserName()) ?></strong> (<?= e(currentUserLogin()) ?>)</p>

<hr>

<h1>Личный кабинет</h1>

<div class="profile-box">
    <h2>Добро пожаловать!</h2>
    <p>Авторизация прошла успешно. Теперь вам доступны защищённые страницы каталога смартфонов.</p>

    <h3 class="section-heading">Что реализовано в практической работе</h3>
    <ul class="info-list">
        <li>форма авторизации на PHP;</li>
        <li>проверка логина и пароля через password_verify;</li>
        <li>сессия пользователя после успешного входа;</li>
        <li>защита страниц каталога от неавторизованных пользователей;</li>
        <li>выход из системы через уничтожение сессии;</li>
        <li>запись событий авторизации в файл logs/auth.log.</li>
    </ul>

    <p><a class="btn btn-add profile-link" href="catalog.php">Перейти в каталог</a></p>
</div>

<hr>

<p class="footer"><small>&copy; 2026 Каталог смартфонов. Все права защищены.</small></p>

</body>
</html>
