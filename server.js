const express = require("express");
const mysql = require("mysql2");
const multer = require("multer");
const cors = require("cors");
const path = require("path");
const fs = require("fs");
const bcrypt = require("bcrypt");
const jwt = require("jsonwebtoken");
const ExcelJS = require("exceljs");

const SECRET_KEY = "20072004iska";

const app = express();
const PORT = 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Настройка статики
app.use(express.static(__dirname));
app.use("/public", express.static(path.join(__dirname, "public")));
app.use("/uploads", express.static(path.join(__dirname, "uploads")));

// Multer setup
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const dir = "uploads/";
    if (!fs.existsSync(dir)) fs.mkdirSync(dir);
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
    return res.status(400).json({ error: "Заполните все поля и загрузите фото!" });
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
  db.query("SELECT image_url FROM products WHERE id = ?", [id], (err, results) => {
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
  });
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

  db.query("SELECT * FROM users WHERE email = ?", [email], async (err, results) => {
    if (err) return res.status(500).json({ error: err });
    if (!results[0]) return res.status(400).json({ error: "Пользователь не найден" });

    const user = results[0];
    const match = await bcrypt.compare(password, user.password);

    if (!match) return res.status(400).json({ error: "Неверный пароль" });

    const token = jwt.sign(
      { id: user.id, name: user.name, role: user.role },
      SECRET_KEY,
      { expiresIn: "2h" },
    );

    res.json({ message: "Успешный вход", token });
  });
});

app.post("/checkout", async (req, res) => {
  const token = req.headers.authorization?.split(" ")[1];
  if (!token) return res.status(401).json({ error: "Не авторизован" });

  try {
    const decoded = jwt.verify(token, SECRET_KEY);
    const userId = decoded.id;
    const cartItems = req.body.cart;
    const delivery = req.body.delivery;

    if (!cartItems || cartItems.length === 0)
      return res.status(400).json({ error: "Корзина пуста" });

    if (!delivery || !delivery.method)
      return res.status(400).json({ error: "Выберите способ доставки" });

    let total = 0;

    for (const item of cartItems) {
      total += item.price * (item.quantity || 1);

      await new Promise((resolve, reject) => {
        db.query(
          "INSERT INTO cart_items (user_id, product_id, quantity) VALUES (?, ?, ?)",
          [userId, item.id, item.quantity || 1],
          (err) => (err ? reject(err) : resolve()),
        );
      });
    }

    db.query(
      "INSERT INTO orders (user_id, total_price, delivery_method, delivery_address) VALUES (?, ?, ?, ?)",
      [userId, total, delivery.method, delivery.address || null],
      (err, result) => {
        if (err) return res.status(500).json({ error: err });
        res.json({ message: "Заказ оформлен", total });
      },
    );
  } catch (err) {
    res.status(401).json({ error: "Неверный токен" });
  }
});

// ============================================================
//  ФИНАНСОВЫЙ ОТЧЁТ — скачать как Excel
// ============================================================
app.get("/financial-report", async (req, res) => {
  try {
    // 1. Статистика по товарам (сколько штук куплено и выручка)
    const productStats = await new Promise((resolve, reject) => {
      db.query(
        `SELECT 
           p.id,
           p.title,
           p.price,
           COALESCE(SUM(ci.quantity), 0) AS total_qty,
           COALESCE(SUM(ci.quantity * p.price), 0) AS total_revenue
         FROM products p
         LEFT JOIN cart_items ci ON ci.product_id = p.id
         GROUP BY p.id, p.title, p.price
         ORDER BY total_revenue DESC`,
        (err, rows) => (err ? reject(err) : resolve(rows)),
      );
    });

    // 2. Заказы по дням
    const ordersByDay = await new Promise((resolve, reject) => {
      db.query(
        `SELECT 
           DATE(created_at) AS day,
           COUNT(*) AS orders_count,
           SUM(total_price) AS day_revenue
         FROM orders
         GROUP BY DATE(created_at)
         ORDER BY day DESC`,
        (err, rows) => (err ? reject(err) : resolve(rows)),
      );
    });

    // 3. Заказы по способу доставки
    const byDelivery = await new Promise((resolve, reject) => {
      db.query(
        `SELECT 
           delivery_method,
           COUNT(*) AS cnt,
           SUM(total_price) AS revenue
         FROM orders
         GROUP BY delivery_method`,
        (err, rows) => (err ? reject(err) : resolve(rows)),
      );
    });

    // 4. Топ покупатели
    const topUsers = await new Promise((resolve, reject) => {
      db.query(
        `SELECT 
           u.name,
           u.email,
           COUNT(o.id) AS orders_count,
           SUM(o.total_price) AS total_spent
         FROM users u
         JOIN orders o ON o.user_id = u.id
         GROUP BY u.id, u.name, u.email
         ORDER BY total_spent DESC
         LIMIT 20`,
        (err, rows) => (err ? reject(err) : resolve(rows)),
      );
    });

    // ---- Создаём Excel ----
    const wb = new ExcelJS.Workbook();
    wb.creator = "LavandaBloom";
    wb.created = new Date();

    // ----- ЛИСТ 1: Товары -----
    const wsProducts = wb.addWorksheet("Продажи по товарам");

    // Заголовок
    wsProducts.mergeCells("A1:E1");
    const titleCell1 = wsProducts.getCell("A1");
    titleCell1.value = "LavandaBloom — Продажи по товарам";
    titleCell1.font = { name: "Arial", bold: true, size: 14, color: { argb: "FF6A1B9A" } };
    titleCell1.alignment = { horizontal: "center" };
    titleCell1.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF3E5F5" } };

    wsProducts.addRow([]);

    const hdrProducts = wsProducts.addRow(["№", "Название товара", "Цена (₸)", "Кол-во продаж", "Выручка (₸)"]);
    hdrProducts.eachCell((cell) => {
      cell.font = { name: "Arial", bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF8E24AA" } };
      cell.alignment = { horizontal: "center" };
      cell.border = {
        top: { style: "thin" }, bottom: { style: "thin" },
        left: { style: "thin" }, right: { style: "thin" },
      };
    });

    productStats.forEach((p, i) => {
      const row = wsProducts.addRow([i + 1, p.title, p.price, p.total_qty, p.total_revenue]);
      const isEven = i % 2 === 0;
      row.eachCell((cell) => {
        cell.font = { name: "Arial" };
        cell.fill = {
          type: "pattern", pattern: "solid",
          fgColor: { argb: isEven ? "FFFDF8FF" : "FFFCE4EC" },
        };
        cell.border = {
          top: { style: "thin", color: { argb: "FFE1BEE7" } },
          bottom: { style: "thin", color: { argb: "FFE1BEE7" } },
          left: { style: "thin", color: { argb: "FFE1BEE7" } },
          right: { style: "thin", color: { argb: "FFE1BEE7" } },
        };
      });
      // Числа вправо
      row.getCell(3).alignment = { horizontal: "right" };
      row.getCell(4).alignment = { horizontal: "center" };
      row.getCell(5).alignment = { horizontal: "right" };
    });

    // Итоговая строка
    const totalRow = productStats.length + 3; // данные с row 3, плюс 2 строки заголовков
    const dataStart = 4;
    const dataEnd = dataStart + productStats.length - 1;
    const summaryRow = wsProducts.addRow([
      "", "ИТОГО", "",
      `=SUM(D${dataStart}:D${dataEnd})`,
      `=SUM(E${dataStart}:E${dataEnd})`,
    ]);
    summaryRow.eachCell((cell) => {
      cell.font = { name: "Arial", bold: true, color: { argb: "FF4A148C" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFCE93D8" } };
      cell.border = {
        top: { style: "medium", color: { argb: "FF8E24AA" } },
        bottom: { style: "medium", color: { argb: "FF8E24AA" } },
        left: { style: "thin", color: { argb: "FFE1BEE7" } },
        right: { style: "thin", color: { argb: "FFE1BEE7" } },
      };
    });

    wsProducts.columns = [
      { key: "num", width: 5 },
      { key: "title", width: 35 },
      { key: "price", width: 14 },
      { key: "qty", width: 16 },
      { key: "rev", width: 18 },
    ];

    // ----- ЛИСТ 2: Заказы по дням -----
    const wsDaily = wb.addWorksheet("Заказы по дням");

    wsDaily.mergeCells("A1:C1");
    const titleCell2 = wsDaily.getCell("A1");
    titleCell2.value = "LavandaBloom — Заказы по дням";
    titleCell2.font = { name: "Arial", bold: true, size: 14, color: { argb: "FF6A1B9A" } };
    titleCell2.alignment = { horizontal: "center" };
    titleCell2.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF3E5F5" } };

    wsDaily.addRow([]);

    const hdrDaily = wsDaily.addRow(["Дата", "Кол-во заказов", "Выручка за день (₸)"]);
    hdrDaily.eachCell((cell) => {
      cell.font = { name: "Arial", bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF8E24AA" } };
      cell.alignment = { horizontal: "center" };
      cell.border = {
        top: { style: "thin" }, bottom: { style: "thin" },
        left: { style: "thin" }, right: { style: "thin" },
      };
    });

    ordersByDay.forEach((d, i) => {
      const dayStr = d.day ? new Date(d.day).toLocaleDateString("ru-RU") : "—";
      const row = wsDaily.addRow([dayStr, d.orders_count, d.day_revenue]);
      const isEven = i % 2 === 0;
      row.eachCell((cell) => {
        cell.font = { name: "Arial" };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: isEven ? "FFFDF8FF" : "FFFCE4EC" } };
        cell.border = {
          top: { style: "thin", color: { argb: "FFE1BEE7" } },
          bottom: { style: "thin", color: { argb: "FFE1BEE7" } },
          left: { style: "thin", color: { argb: "FFE1BEE7" } },
          right: { style: "thin", color: { argb: "FFE1BEE7" } },
        };
      });
      row.getCell(2).alignment = { horizontal: "center" };
      row.getCell(3).alignment = { horizontal: "right" };
    });

    wsDaily.columns = [
      { key: "day", width: 16 },
      { key: "cnt", width: 18 },
      { key: "rev", width: 22 },
    ];

    // ----- ЛИСТ 3: По способу доставки -----
    const wsDelivery = wb.addWorksheet("По доставке");

    wsDelivery.mergeCells("A1:C1");
    const titleCell3 = wsDelivery.getCell("A1");
    titleCell3.value = "LavandaBloom — Заказы по способу доставки";
    titleCell3.font = { name: "Arial", bold: true, size: 14, color: { argb: "FF6A1B9A" } };
    titleCell3.alignment = { horizontal: "center" };
    titleCell3.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF3E5F5" } };

    wsDelivery.addRow([]);

    const hdrDel = wsDelivery.addRow(["Способ доставки", "Кол-во заказов", "Выручка (₸)"]);
    hdrDel.eachCell((cell) => {
      cell.font = { name: "Arial", bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF8E24AA" } };
      cell.alignment = { horizontal: "center" };
      cell.border = {
        top: { style: "thin" }, bottom: { style: "thin" },
        left: { style: "thin" }, right: { style: "thin" },
      };
    });

    const deliveryLabels = { pickup: "Самовывоз", courier: "Курьер", postal: "Почта" };

    byDelivery.forEach((d, i) => {
      const row = wsDelivery.addRow([
        deliveryLabels[d.delivery_method] || d.delivery_method,
        d.cnt,
        d.revenue,
      ]);
      const isEven = i % 2 === 0;
      row.eachCell((cell) => {
        cell.font = { name: "Arial" };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: isEven ? "FFFDF8FF" : "FFFCE4EC" } };
        cell.border = {
          top: { style: "thin", color: { argb: "FFE1BEE7" } },
          bottom: { style: "thin", color: { argb: "FFE1BEE7" } },
          left: { style: "thin", color: { argb: "FFE1BEE7" } },
          right: { style: "thin", color: { argb: "FFE1BEE7" } },
        };
      });
      row.getCell(2).alignment = { horizontal: "center" };
      row.getCell(3).alignment = { horizontal: "right" };
    });

    wsDelivery.columns = [
      { key: "method", width: 20 },
      { key: "cnt", width: 18 },
      { key: "rev", width: 18 },
    ];

    // ----- ЛИСТ 4: Топ покупатели -----
    const wsUsers = wb.addWorksheet("Топ покупатели");

    wsUsers.mergeCells("A1:D1");
    const titleCell4 = wsUsers.getCell("A1");
    titleCell4.value = "LavandaBloom — Топ покупатели";
    titleCell4.font = { name: "Arial", bold: true, size: 14, color: { argb: "FF6A1B9A" } };
    titleCell4.alignment = { horizontal: "center" };
    titleCell4.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF3E5F5" } };

    wsUsers.addRow([]);

    const hdrUsers = wsUsers.addRow(["Имя", "Email", "Кол-во заказов", "Сумма покупок (₸)"]);
    hdrUsers.eachCell((cell) => {
      cell.font = { name: "Arial", bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF8E24AA" } };
      cell.alignment = { horizontal: "center" };
      cell.border = {
        top: { style: "thin" }, bottom: { style: "thin" },
        left: { style: "thin" }, right: { style: "thin" },
      };
    });

    topUsers.forEach((u, i) => {
      const row = wsUsers.addRow([u.name, u.email, u.orders_count, u.total_spent]);
      const isEven = i % 2 === 0;
      row.eachCell((cell) => {
        cell.font = { name: "Arial" };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: isEven ? "FFFDF8FF" : "FFFCE4EC" } };
        cell.border = {
          top: { style: "thin", color: { argb: "FFE1BEE7" } },
          bottom: { style: "thin", color: { argb: "FFE1BEE7" } },
          left: { style: "thin", color: { argb: "FFE1BEE7" } },
          right: { style: "thin", color: { argb: "FFE1BEE7" } },
        };
      });
      row.getCell(3).alignment = { horizontal: "center" };
      row.getCell(4).alignment = { horizontal: "right" };
    });

    wsUsers.columns = [
      { key: "name", width: 24 },
      { key: "email", width: 30 },
      { key: "cnt", width: 16 },
      { key: "spent", width: 20 },
    ];

    // Отдаём файл
    const dateStr = new Date().toISOString().slice(0, 10);
    const filename = `LavandaBloom_report_${dateStr}.xlsx`;

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);

    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error("Ошибка генерации отчёта:", err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));