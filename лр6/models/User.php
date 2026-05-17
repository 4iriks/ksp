<?php
declare(strict_types=1);

class User
{
    private string $file;

    public function __construct(string $file)
    {
        $this->file = $file;

        if (!is_dir(dirname($this->file))) {
            mkdir(dirname($this->file), 0777, true);
        }

        if (!file_exists($this->file)) {
            file_put_contents($this->file, json_encode([], JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT));
        }
    }

    public function all(bool $withPassword = false): array
    {
        $users = $this->read();

        if ($withPassword) {
            return $users;
        }

        return array_map([$this, 'withoutPassword'], $users);
    }

    public function find(int $id, bool $withPassword = false): ?array
    {
        foreach ($this->read() as $user) {
            if ((int)$user['id'] === $id) {
                return $withPassword ? $user : $this->withoutPassword($user);
            }
        }

        return null;
    }

    public function findByEmail(string $email, bool $withPassword = false): ?array
    {
        foreach ($this->read() as $user) {
            if (strtolower($user['email']) === strtolower($email)) {
                return $withPassword ? $user : $this->withoutPassword($user);
            }
        }

        return null;
    }

    public function create(string $name, string $email, string $password): array
    {
        $users = $this->read();
        $ids = array_column($users, 'id');
        $nextId = empty($ids) ? 1 : max($ids) + 1;

        $user = [
            'id' => $nextId,
            'name' => $name,
            'email' => $email,
            'password_hash' => password_hash($password, PASSWORD_DEFAULT),
        ];

        $users[] = $user;
        $this->write($users);

        return $this->withoutPassword($user);
    }

    public function updatePassword(int $id, string $password): ?array
    {
        $users = $this->read();

        foreach ($users as $index => $user) {
            if ((int)$user['id'] === $id) {
                $users[$index]['password_hash'] = password_hash($password, PASSWORD_DEFAULT);
                $this->write($users);

                return $this->withoutPassword($users[$index]);
            }
        }

        return null;
    }

    public function delete(int $id): bool
    {
        $users = $this->read();
        $filtered = array_values(array_filter($users, fn(array $user): bool => (int)$user['id'] !== $id));

        if (count($users) === count($filtered)) {
            return false;
        }

        $this->write($filtered);

        return true;
    }

    private function read(): array
    {
        $json = file_get_contents($this->file);
        $data = json_decode($json ?: '[]', true);

        return is_array($data) ? $data : [];
    }

    private function write(array $users): void
    {
        file_put_contents($this->file, json_encode($users, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT), LOCK_EX);
    }

    private function withoutPassword(array $user): array
    {
        unset($user['password_hash']);

        return $user;
    }
}
