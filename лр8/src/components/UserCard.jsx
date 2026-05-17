import React from "react";

function UserCard({ user }) {
  return (
    <li className="data-item">
      <strong>{user.name}</strong>
      <span>Email: {user.email}</span>
      <span>Город: {user.address.city}</span>
    </li>
  );
}

export default UserCard;
