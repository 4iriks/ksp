import React from "react";
import ProfileCard from "./components/ProfileCard.jsx";
import UsersList from "./components/UsersList.jsx";

function App() {
  return (
    <main className="app">
      <section className="work-area">
        <ProfileCard />
        <UsersList />
      </section>
    </main>
  );
}

export default App;
