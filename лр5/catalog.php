<?php
require_once __DIR__ . '/functions.php';
requireAuth();
?>
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <title>Каталог — Каталог смартфонов</title>
    <link rel="stylesheet" type="text/css" href="style.css">
</head>
<body>

<ul class="menu">
    <li><a href="profile.php">Личный кабинет</a></li>
    <li><a href="catalog.php">Каталог</a></li>
    <li><a class="logout-link" href="logout.php">Выход</a></li>
</ul>

<p class="user-panel">Вы вошли как: <strong><?= e(currentUserName()) ?></strong></p>

<hr>

<h2>Каталог</h2>

<div class="filter-bar">
    <label for="category-filter">Фильтр по бренду:</label>
    <select id="category-filter">
        <option value="all">Все</option>
        <option value="apple">Apple</option>
        <option value="samsung">Samsung</option>
        <option value="google">Google</option>
    </select>
</div>

<div class="product-card" data-category="apple">
    <img src="images/iphone 15 pro.jpg" width="150" height="150" alt="iPhone 15 Pro">
    <div class="product-name"><a href="item_iphone.php">iPhone 15 Pro</a></div>
    <div class="product-price">129 990 руб.</div>
    <button class="btn btn-add" data-name="iPhone 15 Pro" data-price="129990">Добавить в корзину</button>
</div>

<div class="product-card" data-category="samsung">
    <img src="images/samsung Galaxy s24.jpg" width="150" height="150" alt="Samsung Galaxy S24 Ultra">
    <div class="product-name"><a href="item_samsung.php">Samsung Galaxy S24 Ultra</a></div>
    <div class="product-price">109 990 руб.</div>
    <button class="btn btn-add" data-name="Samsung Galaxy S24 Ultra" data-price="109990">Добавить в корзину</button>
</div>

<div class="product-card" data-category="google">
    <img src="images/Google Pixel 8 proo.jpg" width="150" height="150" alt="Google Pixel 8 Pro">
    <div class="product-name"><a href="item_pixel.php">Google Pixel 8 Pro</a></div>
    <div class="product-price">84 990 руб.</div>
    <button class="btn btn-add" data-name="Google Pixel 8 Pro" data-price="84990">Добавить в корзину</button>
</div>

<div class="cart-section">
    <h3>Корзина</h3>
    <div id="cart-list"></div>
    <div id="cart-total" class="cart-total">Итого: 0 руб.</div>
    <div class="cart-buttons">
        <button class="btn btn-pay" id="btn-pay">Оплатить</button>
        <button class="btn btn-clear" id="btn-clear">Очистить корзину</button>
    </div>
</div>

<hr>

<p class="footer"><small>&copy; 2026 Каталог смартфонов. Все права защищены.</small></p>

<script src="script.js"></script>
</body>
</html>
