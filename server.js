const express = require("express");
const mysql = require("mysql2");
const multer = require("multer");
const cors = require("cors");
const path = require("path");
const fs = require("fs");
const bcrypt = require("bcrypt");
const jwt = require("jsonwebtoken");

const SECRET_KEY = "20072004iska";

const app = express();
const PORT = 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Настройка статики
app.use(express.static(__dirname)); // Для index.html в корне
app.use("/public", express.static(path.join(__dirname, "public"))); // Для файлов в public
app.use("/uploads", express.static(path.join(__dirname, "uploads"))); // Для картинок

// Multer setup для загрузки файлов
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const dir = "uploads/";
    if (!fs.existsSync(dir)) fs.mkdirSync(dir); // Создаем папку, если нет
    cb(null, dir);
  },
  filename: (req, file, cb) => {
    const uniqueName = Date.now() + "-" + file.originalname;
    cb(null, uniqueName);
  },
});
const upload = multer({ storage });

// MySQL connection
const db = mysql.createConnection({
  host: "localhost",
  user: "root",
  password: "20072004iska",
  database: "lavandabloom",
});

db.connect((err) => {
  if (err) {
    console.error("Ошибка подключения к БД:", err);
    return;
  }
  console.log("Connected to MySQL");
});

// Получить все товары
app.get("/products", (req, res) => {
  db.query("SELECT * FROM products", (err, results) => {
    if (err) return res.status(500).json({ error: err });
    res.json(results);
  });
});

// ---------- REVIEWS ----------
// Возвращает отзывы для конкретного товара
app.get("/reviews/:productId", (req, res) => {
  const { productId } = req.params;
  db.query(
    "SELECT r.rating, r.comment, r.created_at, u.name FROM reviews r JOIN users u ON r.user_id = u.id WHERE r.product_id = ? ORDER BY r.created_at DESC",
    [productId],
    (err, results) => {
      if (err) return res.status(500).json({ error: err });
      res.json(results);
    },
  );
});

// Добавление нового отзыва (только для авторизованных пользователей)
app.post("/reviews", (req, res) => {
  const token = req.headers.authorization?.split(" ")[1];
  if (!token) return res.status(401).json({ error: "Не авторизован" });

  try {
    const decoded = jwt.verify(token, SECRET_KEY);
    const userId = decoded.id;
    const { productId, rating, comment } = req.body;

    if (!productId || !rating) {
      return res.status(400).json({ error: "Продукт и оценка обязательны" });
    }

    db.query(
      "INSERT INTO reviews (product_id, user_id, rating, comment, created_at) VALUES (?, ?, ?, ?, NOW())",
      [productId, userId, rating, comment || null],
      (err, results) => {
        if (err) return res.status(500).json({ error: err });
        res.json({ message: "Отзыв добавлен" });
      },
    );
  } catch (err) {
    res.status(401).json({ error: "Неверный токен" });
  }
});

// Добавить товар
app.post("/add-product", upload.single("image"), (req, res) => {
  const { title, description, price } = req.body;
  const image = req.file ? req.file.filename : null;

  if (!title || !description || !price || !image) {
    return res
      .status(400)
      .json({ error: "Заполните все поля и загрузите фото!" });
  }

  db.query(
    "INSERT INTO products (title, description, price, image_url) VALUES (?, ?, ?, ?)",
    [title, description, price, image],
    (err, results) => {
      if (err) return res.status(500).json({ error: err });
      res.json({ message: "Товар добавлен", id: results.insertId });
    },
  );
});

// Получить отдельный товар (используется на странице отзывов)
app.get("/product/:id", (req, res) => {
  const { id } = req.params;
  db.query("SELECT * FROM products WHERE id = ?", [id], (err, results) => {
    if (err) return res.status(500).json({ error: err });
    if (!results[0]) return res.status(404).json({ error: "Товар не найден" });
    res.json(results[0]);
  });
});

// Удалить товар
app.delete("/delete-product/:id", (req, res) => {
  const { id } = req.params;
  db.query(
    "SELECT image_url FROM products WHERE id = ?",
    [id],
    (err, results) => {
      if (results && results[0]) {
        const imagePath = path.join(__dirname, "uploads", results[0].image_url);
        fs.unlink(imagePath, (err) => {
          if (err) console.log("Файл не найден");
        });
      }
      db.query("DELETE FROM products WHERE id = ?", [id], (err2) => {
        if (err2) return res.status(500).json({ error: err2 });
        res.json({ message: "Товар удален" });
      });
    },
  );
});

app.post("/register", async (req, res) => {
  const { name, email, password } = req.body;

  if (!name || !email || !password)
    return res.status(400).json({ error: "Все поля обязательны" });

  try {
    const hashedPassword = await bcrypt.hash(password, 10);

    db.query(
      "INSERT INTO users (name, email, password) VALUES (?, ?, ?)",
      [name, email, hashedPassword],
      (err) => {
        if (err) return res.status(500).json({ error: "Email уже существует" });
        res.json({ message: "Регистрация успешна" });
      },
    );
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post("/login", (req, res) => {
  const { email, password } = req.body;

  db.query(
    "SELECT * FROM users WHERE email = ?",
    [email],
    async (err, results) => {
      if (err) return res.status(500).json({ error: err });
      if (!results[0])
        return res.status(400).json({ error: "Пользователь не найден" });

      const user = results[0];
      const match = await bcrypt.compare(password, user.password);

      if (!match) return res.status(400).json({ error: "Неверный пароль" });

      const token = jwt.sign(
        { id: user.id, name: user.name, role: user.role },
        SECRET_KEY,
        { expiresIn: "2h" },
      );

      res.json({ message: "Успешный вход", token });
    },
  );
});

app.listen(PORT, () =>
  console.log(`Server running on http://localhost:${PORT}`),
);

app.post("/checkout", async (req, res) => {
    const token = req.headers.authorization?.split(" ")[1];
    if (!token) return res.status(401).json({ error: "Не авторизован" });

    try {
        const decoded = jwt.verify(token, SECRET_KEY);
        const userId = decoded.id;
        const cartItems = req.body.cart;
        const delivery = req.body.delivery; // { method, address }

        if (!cartItems || cartItems.length === 0)
            return res.status(400).json({ error: "Корзина пуста" });

        if (!delivery || !delivery.method)
            return res.status(400).json({ error: "Выберите способ доставки" });

        let total = 0;

        // Подсчёт суммы
        for (const item of cartItems) {
            total += item.price * (item.quantity || 1);

            // Сохраняем в cart_items
            await new Promise((resolve, reject) => {
                db.query(
                    "INSERT INTO cart_items (user_id, product_id, quantity) VALUES (?, ?, ?)",
                    [userId, item.id, item.quantity || 1],
                    (err) => (err ? reject(err) : resolve())
                );
            });
        }

        // Создаём заказ (с доставкой)
        db.query(
            "INSERT INTO orders (user_id, total_price, delivery_method, delivery_address) VALUES (?, ?, ?, ?)",
            [userId, total, delivery.method, delivery.address || null],
            (err, result) => {
                if (err) return res.status(500).json({ error: err });
                res.json({ message: "Заказ оформлен", total });
            }
        );
    } catch (err) {
        res.status(401).json({ error: "Неверный токен" });
    }
});
