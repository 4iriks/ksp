import React, { useEffect, useState } from "react";
import UserCard from "./UserCard.jsx";

function UsersList() {
  const [users, setUsers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  useEffect(() => {
    const loadUsers = async () => {
      try {
        setLoading(true);
        setError("");

        await new Promise((resolve) => setTimeout(resolve, 700));

        const response = await fetch("/api/users.json");

        if (!response.ok) {
          throw new Error("Ошибка ответа сервера");
        }

        const data = await response.json();
        setUsers(data);
      } catch {
        setError("Ошибка загрузки данных");
      } finally {
        setLoading(false);
      }
    };

    loadUsers();
  }, []);

  return (
    <article className="data-panel">
      <div className="panel-title">
        <h2>Пользователи из API</h2>
        <span>GET /api/users.json</span>
      </div>

      {loading && <p className="state-message">Загрузка...</p>}
      {error && <p className="state-message error">{error}</p>}

      {!loading && !error && (
        <ul className="data-list">
          {users.map((user) => (
            <UserCard key={user.id} user={user} />
          ))}
        </ul>
      )}
    </article>
  );
}

export default UsersList;
