import React, { useEffect, useState } from "react";

function ProfileCard() {
  const [avatar, setAvatar] = useState("");

  const handleAvatarChange = (event) => {
    const file = event.target.files[0];

    if (!file) {
      return;
    }

    setAvatar(URL.createObjectURL(file));
  };

  useEffect(() => {
    return () => {
      if (avatar) {
        URL.revokeObjectURL(avatar);
      }
    };
  }, [avatar]);

  return (
    <article className="profile-card">
      <h1>Моя визитка</h1>

      <div className="avatar-preview">
        {avatar ? <img src={avatar} alt="Аватар студента" /> : <span>ВГ</span>}
      </div>

      <label className="avatar-upload">
        Загрузить аватарку
        <input type="file" accept="image/*" onChange={handleAvatarChange} />
      </label>

      <h2>Горбачёв Вадим Александрович</h2>

      <p>
        <strong>Специальность:</strong> Информатика и вычислительная техника
      </p>
      <p>
        <strong>Группа:</strong> БИВТ-24-2
      </p>

      <ul className="profile-notes">
        <li>Учебный проект по основам React.</li>
        <li>Компонент показывает данные студента и выбранную аватарку.</li>
      </ul>
    </article>
  );
}

export default ProfileCard;
