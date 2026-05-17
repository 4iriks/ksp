import React, { useEffect, useState } from "react";

const student = {
  name: "Горбачёв В. А.",
  specialty: "Информатика и вычислительная техника",
  group: "БИВТ-24-2",
  description:
    "Студент, изучающий клиент-серверные приложения, веб-разработку и основы React. В этой работе создан отдельный компонент визитки с возможностью загрузки изображения для аватарки.",
  skills: ["HTML", "CSS", "JavaScript", "React"],
};

function ProfileCard() {
  const [avatar, setAvatar] = useState("");

  const handleAvatarChange = (event) => {
    const file = event.target.files[0];

    if (!file) {
      return;
    }

    const imageUrl = URL.createObjectURL(file);
    setAvatar(imageUrl);
  };

  useEffect(() => {
    return () => {
      if (avatar) {
        URL.revokeObjectURL(avatar);
      }
    };
  }, [avatar]);

  return (
    <section className="profile-card">
      <div className="avatar-block">
        <div className="avatar-preview">
          {avatar ? (
            <img src={avatar} alt="Аватар студента" />
          ) : (
            <span>ВГ</span>
          )}
        </div>

        <label className="avatar-upload">
          Загрузить аватарку
          <input type="file" accept="image/*" onChange={handleAvatarChange} />
        </label>
      </div>

      <div className="profile-info">
        <h1>Моя визитка</h1>
        <h2>{student.name}</h2>

        <p>
          <strong>Специальность:</strong> {student.specialty}
        </p>
        <p>
          <strong>Группа:</strong> {student.group}
        </p>
        <p>{student.description}</p>

        <div className="skills">
          <h3>Навыки</h3>
          <ul>
            {student.skills.map((skill) => (
              <li key={skill}>{skill}</li>
            ))}
          </ul>
        </div>
      </div>
    </section>
  );
}

export default ProfileCard;
