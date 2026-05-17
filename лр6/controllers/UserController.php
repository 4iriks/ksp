<?php
declare(strict_types=1);

class UserController
{
    private User $users;

    public function __construct(User $users)
    {
        $this->users = $users;
    }

    public function info(): void
    {
        $this->json([
            'status' => 'success',
            'message' => 'REST API пользователей работает',
            'endpoints' => [
                'POST /api/v1/register',
                'POST /api/v1/login',
                'GET /api/v1/users',
                'GET /api/v1/users/{id}',
                'PUT /api/v1/users/{id}',
                'PATCH /api/v1/users/{id}',
                'DELETE /api/v1/users/{id}',
            ],
        ]);
    }

    public function register(): void
    {
        $data = $this->input();
        $name = trim((string)($data['name'] ?? ''));
        $email = trim((string)($data['email'] ?? ''));
        $password = (string)($data['password'] ?? '');

        if ($name === '' || $email === '' || $password === '') {
            $this->jsonError('Заполните name, email и password', 400);
            return;
        }

        if (!filter_var($email, FILTER_VALIDATE_EMAIL)) {
            $this->jsonError('Email указан некорректно', 400);
            return;
        }

        if (strlen($password) < 6) {
            $this->jsonError('Пароль должен быть не короче 6 символов', 400);
            return;
        }

        if ($this->users->findByEmail($email) !== null) {
            $this->jsonError('Пользователь с таким email уже существует', 409);
            return;
        }

        $user = $this->users->create($name, $email, $password);

        $this->json([
            'status' => 'success',
            'message' => 'Пользователь зарегистрирован',
            'user' => $user,
        ], 201);
    }

    public function login(): void
    {
        $data = $this->input();
        $email = trim((string)($data['email'] ?? ''));
        $password = (string)($data['password'] ?? '');

        if ($email === '' || $password === '') {
            $this->jsonError('Введите email и password', 400);
            return;
        }

        $user = $this->users->findByEmail($email, true);

        if ($user === null) {
            $this->jsonError('Пользователь не найден', 404);
            return;
        }

        if (!password_verify($password, $user['password_hash'])) {
            $this->jsonError('Неверный пароль', 401);
            return;
        }

        unset($user['password_hash']);

        $this->json([
            'status' => 'success',
            'message' => 'Авторизация выполнена успешно',
            'user' => $user,
        ]);
    }

    public function index(): void
    {
        $this->json([
            'status' => 'success',
            'users' => $this->users->all(),
        ]);
    }

    public function show(int $id): void
    {
        $user = $this->users->find($id);

        if ($user === null) {
            $this->jsonError('Пользователь не найден', 404);
            return;
        }

        $this->json([
            'status' => 'success',
            'user' => $user,
        ]);
    }

    public function updatePassword(int $id): void
    {
        $data = $this->input();
        $password = (string)($data['password'] ?? $data['new_password'] ?? '');

        if ($password === '') {
            $this->jsonError('Введите новый пароль в поле password или new_password', 400);
            return;
        }

        if (strlen($password) < 6) {
            $this->jsonError('Пароль должен быть не короче 6 символов', 400);
            return;
        }

        $user = $this->users->updatePassword($id, $password);

        if ($user === null) {
            $this->jsonError('Пользователь не найден', 404);
            return;
        }

        $this->json([
            'status' => 'success',
            'message' => 'Пароль пользователя изменён',
            'user' => $user,
        ]);
    }

    public function delete(int $id): void
    {
        $deleted = $this->users->delete($id);

        if (!$deleted) {
            $this->jsonError('Пользователь не найден', 404);
            return;
        }

        $this->json([
            'status' => 'success',
            'message' => 'Пользователь удалён',
        ]);
    }

    private function input(): array
    {
        $raw = file_get_contents('php://input');

        if ($raw === false || trim($raw) === '') {
            return [];
        }

        $data = json_decode($raw, true);

        return is_array($data) ? $data : [];
    }

    private function json(array $data, int $code = 200): void
    {
        http_response_code($code);
        echo json_encode($data, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);
    }

    private function jsonError(string $message, int $code): void
    {
        $this->json([
            'status' => 'error',
            'message' => $message,
        ], $code);
    }
}

