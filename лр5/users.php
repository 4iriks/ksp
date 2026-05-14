<?php
// Массив пользователей для практической работы.
// Хэши заранее созданы функцией password_hash(..., PASSWORD_DEFAULT).
return [
    'admin' => [
        'id' => 1,
        'name' => 'Администратор',
        'password_hash' => '$2y$10$92IXUNpkjO0rOQ5byMi.Ye4oKoEa3Ro9llC/.og/at2.uheWG/igi',
    ],
    'user' => [
        'id' => 2,
        'name' => 'Пользователь',
        'password_hash' => '$2y$10$92IXUNpkjO0rOQ5byMi.Ye4oKoEa3Ro9llC/.og/at2.uheWG/igi',
    ],
];
