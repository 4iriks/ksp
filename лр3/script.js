// Глобальный массив корзины
let cart = [];

// Массив товаров каталога с ценами
const products = [
    { id: 1, name: "iPhone 15 Pro", price: 129990, category: "apple" },
    { id: 2, name: "Samsung Galaxy S24 Ultra", price: 109990, category: "samsung" },
    { id: 3, name: "Google Pixel 8 Pro", price: 84990, category: "google" }
];

// Стрелочная функция: подсчёт общей суммы корзины
const calculateTotal = () => {
    let total = 0;
    cart.forEach(item => total += item.price);
    return total;
};

// Стрелочная функция: добавление товара в корзину
const addToCart = (product) => {
    cart.push(product);
    renderCart();
};

// Стрелочная функция: удаление товара из корзины по индексу
const removeFromCart = (index) => {
    cart.splice(index, 1);
    renderCart();
};

// Стрелочная функция: очистка корзины
const clearCart = () => {
    cart = [];
    renderCart();
};

// Стрелочная функция: оплата
const pay = () => {
    if (cart.length === 0) {
        alert("Корзина пуста!");
        return;
    }
    alert("Оплата прошла успешно! Сумма: " + calculateTotal() + " руб.");
    clearCart();
};

// Стрелочная функция: отрисовка корзины
const renderCart = () => {
    const cartList = document.querySelector("#cart-list");
    const cartTotal = document.querySelector("#cart-total");

    if (!cartList || !cartTotal) return;

    // Очистка текущего списка
    cartList.innerHTML = "";

    // Отрисовка каждого товара
    cart.forEach((item, index) => {
        const div = document.createElement("div");
        div.className = "cart-item";

        const span = document.createElement("span");
        span.textContent = item.name + " — " + item.price + " руб.";

        const btn = document.createElement("button");
        btn.className = "btn btn-remove";
        btn.textContent = "Удалить";
        btn.addEventListener("click", () => removeFromCart(index));

        div.appendChild(span);
        div.appendChild(btn);
        cartList.appendChild(div);
    });

    // Обновление суммы
    const total = calculateTotal();
    cartTotal.textContent = "Итого: " + total + " руб.";
};

// Стрелочная функция: фильтрация товаров по категории
const filterProducts = (category) => {
    const cards = document.querySelectorAll(".product-card");
    cards.forEach(card => {
        const cardCategory = card.dataset.category;
        if (category === "all" || cardCategory === category) {
            card.style.display = "inline-block";
        } else {
            card.style.display = "none";
        }
    });
};

// Инициализация после загрузки DOM
document.addEventListener("DOMContentLoaded", () => {
    // Обработчик фильтра
    const filterSelect = document.querySelector("#category-filter");
    if (filterSelect) {
        filterSelect.addEventListener("change", () => {
            filterProducts(filterSelect.value);
        });
    }

    // Обработчики кнопок "Добавить в корзину"
    const addButtons = document.querySelectorAll(".btn-add");
    addButtons.forEach(button => {
        button.addEventListener("click", () => {
            const name = button.dataset.name;
            const price = Number(button.dataset.price);
            addToCart({ name, price });
        });
    });

    // Обработчик кнопки "Оплатить"
    const payButton = document.querySelector("#btn-pay");
    if (payButton) {
        payButton.addEventListener("click", () => pay());
    }

    // Обработчик кнопки "Очистить корзину"
    const clearButton = document.querySelector("#btn-clear");
    if (clearButton) {
        clearButton.addEventListener("click", () => clearCart());
    }
});
