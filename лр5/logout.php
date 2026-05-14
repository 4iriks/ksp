<?php
require_once __DIR__ . '/functions.php';

if (isLoggedIn()) {
    logAuth(currentUserLogin(), 'LOGOUT');
}

session_unset();
session_destroy();

header('Location: index.php?logout=1');
exit;
