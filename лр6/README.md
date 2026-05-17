# Практическая работа 6

## Разработка REST API на PHP

Проект реализует базовый REST API для работы с пользователями.
Данные пользователей хранятся в файле `data/users.json`.
Ответы сервера возвращаются в формате JSON.

## Запуск через Docker

```powershell
docker run -d --name lr6-php -p 8081:80 -v C:\tmp\lr6:/var/www/html php:8.4-apache
docker exec lr6-php a2enmod rewrite
docker restart lr6-php
```

Адрес API:

```text
http://127.0.0.1:8081/api/v1
```

## Endpoint'ы

```text
POST   /api/v1/register
POST   /api/v1/login
GET    /api/v1/users
GET    /api/v1/users/{id}
PUT    /api/v1/users/{id}
PATCH  /api/v1/users/{id}
DELETE /api/v1/users/{id}
```

## Примеры запросов cURL

Регистрация:

```bash
curl -X POST http://127.0.0.1:8081/api/v1/register \
  -H "Content-Type: application/json" \
  -d "{\"name\":\"Nikita\",\"email\":\"nikita@test.com\",\"password\":\"123456\"}"
```

Авторизация:

```bash
curl -X POST http://127.0.0.1:8081/api/v1/login \
  -H "Content-Type: application/json" \
  -d "{\"email\":\"admin@test.com\",\"password\":\"password\"}"
```

Получение списка пользователей:

```bash
curl http://127.0.0.1:8081/api/v1/users
```

Получение одного пользователя:

```bash
curl http://127.0.0.1:8081/api/v1/users/1
```

Изменение пароля:

```bash
curl -X PATCH http://127.0.0.1:8081/api/v1/users/1 \
  -H "Content-Type: application/json" \
  -d "{\"new_password\":\"654321\"}"
```

Удаление пользователя:

```bash
curl -X DELETE http://127.0.0.1:8081/api/v1/users/2
```
