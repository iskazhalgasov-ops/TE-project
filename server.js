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

// ─── Склад: Астана, Жиембет жырау 2 ───────────────────────────────────────
const WAREHOUSE_LAT = 51.1605;
const WAREHOUSE_LNG = 71.4704;
const DELIVERY_BASE_PRICE = 500;    // тенге — минимум за доставку
const DELIVERY_PRICE_PER_KM = 80;  // тенге за каждый км

function haversineKm(lat1, lng1, lat2, lng2) {
  const R = 6371;
  const toRad = x => x * Math.PI / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a = Math.sin(dLat/2)**2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng/2)**2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function calcDeliveryPrice(distKm) {
  return Math.round(DELIVERY_BASE_PRICE + distKm * DELIVERY_PRICE_PER_KM);
}

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(__dirname));
app.use("/public", express.static(path.join(__dirname, "public")));
app.use("/uploads", express.static(path.join(__dirname, "uploads")));

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const dir = "uploads/";
    if (!fs.existsSync(dir)) fs.mkdirSync(dir);
    cb(null, dir);
  },
  filename: (req, file, cb) => { cb(null, Date.now() + "-" + file.originalname); },
});
const upload = multer({ storage });

const db = mysql.createConnection({
  host: "localhost", user: "root", password: "20072004iska", database: "lavandabloom",
});
db.connect(err => {
  if (err) { console.error("Ошибка подключения к БД:", err); return; }
  console.log("Connected to MySQL");
});

// ── Расчёт цены доставки по координатам ──────────────────────────────────
app.post("/delivery-price", (req, res) => {
  const { lat, lng } = req.body;
  if (lat == null || lng == null) return res.status(400).json({ error: "Нет координат" });
  const dist = haversineKm(WAREHOUSE_LAT, WAREHOUSE_LNG, parseFloat(lat), parseFloat(lng));
  res.json({ distance: Math.round(dist * 10) / 10, price: calcDeliveryPrice(dist) });
});

// ── PRODUCTS ──────────────────────────────────────────────────────────────
app.get("/products", (req, res) => {
  db.query("SELECT * FROM products", (err, r) => {
    if (err) return res.status(500).json({ error: err });
    res.json(r);
  });
});

app.get("/product/:id", (req, res) => {
  db.query("SELECT * FROM products WHERE id = ?", [req.params.id], (err, r) => {
    if (err) return res.status(500).json({ error: err });
    if (!r[0]) return res.status(404).json({ error: "Товар не найден" });
    res.json(r[0]);
  });
});

app.post("/add-product", upload.single("image"), (req, res) => {
  const { title, description, price } = req.body;
  const image = req.file ? req.file.filename : null;
  if (!title || !description || !price || !image)
    return res.status(400).json({ error: "Заполните все поля и загрузите фото!" });
  db.query(
    "INSERT INTO products (title, description, price, image_url) VALUES (?, ?, ?, ?)",
    [title, description, price, image],
    (err, r) => {
      if (err) return res.status(500).json({ error: err });
      res.json({ message: "Товар добавлен", id: r.insertId });
    }
  );
});

app.delete("/delete-product/:id", (req, res) => {
  db.query("SELECT image_url FROM products WHERE id = ?", [req.params.id], (err, r) => {
    if (r && r[0]) fs.unlink(path.join(__dirname, "uploads", r[0].image_url), () => {});
    db.query("DELETE FROM products WHERE id = ?", [req.params.id], (err2) => {
      if (err2) return res.status(500).json({ error: err2 });
      res.json({ message: "Товар удален" });
    });
  });
});

// ── REVIEWS ───────────────────────────────────────────────────────────────
// Получить отзывы (возвращаем id и user_id для кнопки удаления)
app.get("/reviews/:productId", (req, res) => {
  db.query(
    `SELECT r.id, r.rating, r.comment, r.created_at, r.user_id, u.name
     FROM reviews r JOIN users u ON r.user_id = u.id
     WHERE r.product_id = ? ORDER BY r.created_at DESC`,
    [req.params.productId],
    (err, r) => {
      if (err) return res.status(500).json({ error: err });
      res.json(r);
    }
  );
});

// Добавить отзыв
app.post("/reviews", (req, res) => {
  const token = req.headers.authorization?.split(" ")[1];
  if (!token) return res.status(401).json({ error: "Не авторизован" });
  try {
    const decoded = jwt.verify(token, SECRET_KEY);
    const { productId, rating, comment } = req.body;
    if (!productId || !rating) return res.status(400).json({ error: "Продукт и оценка обязательны" });
    db.query(
      "INSERT INTO reviews (product_id, user_id, rating, comment, created_at) VALUES (?, ?, ?, ?, NOW())",
      [productId, decoded.id, rating, comment || null],
      (err) => {
        if (err) return res.status(500).json({ error: err });
        res.json({ message: "Отзыв добавлен" });
      }
    );
  } catch { res.status(401).json({ error: "Неверный токен" }); }
});

// Удалить отзыв — только свой или admin
app.delete("/reviews/:id", (req, res) => {
  const token = req.headers.authorization?.split(" ")[1];
  if (!token) return res.status(401).json({ error: "Не авторизован" });
  try {
    const decoded = jwt.verify(token, SECRET_KEY);
    db.query("SELECT user_id FROM reviews WHERE id = ?", [req.params.id], (err, r) => {
      if (err || !r[0]) return res.status(404).json({ error: "Отзыв не найден" });
      if (r[0].user_id !== decoded.id && decoded.role !== "admin")
        return res.status(403).json({ error: "Нет доступа" });
      db.query("DELETE FROM reviews WHERE id = ?", [req.params.id], (err2) => {
        if (err2) return res.status(500).json({ error: err2 });
        res.json({ message: "Отзыв удалён" });
      });
    });
  } catch { res.status(401).json({ error: "Неверный токен" }); }
});

// ── AUTH ──────────────────────────────────────────────────────────────────
app.post("/register", async (req, res) => {
  const { name, email, password } = req.body;
  if (!name || !email || !password) return res.status(400).json({ error: "Все поля обязательны" });
  try {
    const hash = await bcrypt.hash(password, 10);
    db.query("INSERT INTO users (name, email, password) VALUES (?, ?, ?)", [name, email, hash],
      (err) => {
        if (err) return res.status(500).json({ error: "Email уже существует" });
        res.json({ message: "Регистрация успешна" });
      }
    );
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post("/login", (req, res) => {
  const { email, password } = req.body;
  db.query("SELECT * FROM users WHERE email = ?", [email], async (err, r) => {
    if (err) return res.status(500).json({ error: err });
    if (!r[0]) return res.status(400).json({ error: "Пользователь не найден" });
    const match = await bcrypt.compare(password, r[0].password);
    if (!match) return res.status(400).json({ error: "Неверный пароль" });
    const token = jwt.sign({ id: r[0].id, name: r[0].name, role: r[0].role }, SECRET_KEY, { expiresIn: "2h" });
    res.json({ message: "Успешный вход", token });
  });
});

// ── CHECKOUT ──────────────────────────────────────────────────────────────
app.post("/checkout", async (req, res) => {
  const token = req.headers.authorization?.split(" ")[1];
  if (!token) return res.status(401).json({ error: "Не авторизован" });
  try {
    const decoded = jwt.verify(token, SECRET_KEY);
    const { cart: cartItems, delivery } = req.body;
    if (!cartItems?.length) return res.status(400).json({ error: "Корзина пуста" });
    if (!delivery?.method) return res.status(400).json({ error: "Выберите способ доставки" });

    let productsTotal = 0;
    for (const item of cartItems) {
      productsTotal += item.price * (item.quantity || 1);
      await new Promise((resolve, reject) => {
        db.query(
          "INSERT INTO cart_items (user_id, product_id, quantity) VALUES (?, ?, ?)",
          [decoded.id, item.id, item.quantity || 1],
          err => err ? reject(err) : resolve()
        );
      });
    }

    // Считаем цену доставки по координатам (переданным с фронта)
    let deliveryPrice = 0;
    if (delivery.method === "courier") {
      if (delivery.lat != null && delivery.lng != null) {
        const dist = haversineKm(WAREHOUSE_LAT, WAREHOUSE_LNG, parseFloat(delivery.lat), parseFloat(delivery.lng));
        deliveryPrice = calcDeliveryPrice(dist);
      } else if (delivery.deliveryPrice) {
        deliveryPrice = Number(delivery.deliveryPrice);
      }
    }

    const grandTotal = productsTotal + deliveryPrice;

    db.query(
      "INSERT INTO orders (user_id, total_price, delivery_method, delivery_address, delivery_price) VALUES (?, ?, ?, ?, ?)",
      [decoded.id, grandTotal, delivery.method, delivery.address || null, deliveryPrice],
      (err) => {
        if (err) {
          // Fallback если колонки delivery_price нет
          db.query(
            "INSERT INTO orders (user_id, total_price, delivery_method, delivery_address) VALUES (?, ?, ?, ?)",
            [decoded.id, grandTotal, delivery.method, delivery.address || null],
            (err2) => {
              if (err2) return res.status(500).json({ error: err2 });
              res.json({ message: "Заказ оформлен", total: grandTotal, deliveryPrice });
            }
          );
          return;
        }
        res.json({ message: "Заказ оформлен", total: grandTotal, deliveryPrice });
      }
    );
  } catch { res.status(401).json({ error: "Неверный токен" }); }
});

// ── ФИНАНСОВЫЙ ОТЧЁТ ──────────────────────────────────────────────────────
app.get("/financial-report", async (req, res) => {
  const q = (sql, params = []) =>
    new Promise((resolve, reject) =>
      db.query(sql, params, (err, rows) => err ? reject(err) : resolve(rows))
    );

  try {
    const productStats = await q(
      `SELECT p.id, p.title, p.price,
         COALESCE(SUM(ci.quantity),0) AS total_qty,
         COALESCE(SUM(ci.quantity * p.price),0) AS total_revenue
       FROM products p LEFT JOIN cart_items ci ON ci.product_id = p.id
       GROUP BY p.id, p.title, p.price ORDER BY total_revenue DESC`
    );

    const ordersByDay = await q(
      `SELECT DATE(created_at) AS day, COUNT(*) AS cnt, SUM(total_price) AS revenue
       FROM orders GROUP BY DATE(created_at) ORDER BY day DESC`
    );

    const byDelivery = await q(
      `SELECT delivery_method, COUNT(*) AS cnt, SUM(total_price) AS revenue,
         SUM(COALESCE(delivery_price,0)) AS delivery_sum,
         AVG(NULLIF(delivery_price,0)) AS avg_delivery,
         MAX(COALESCE(delivery_price,0)) AS max_delivery
       FROM orders GROUP BY delivery_method`
    ).catch(() => q(
      `SELECT delivery_method, COUNT(*) AS cnt, SUM(total_price) AS revenue FROM orders GROUP BY delivery_method`
    ));

    const orderDetails = await q(
      `SELECT o.id, u.name AS buyer, u.email, o.total_price, o.delivery_method,
         o.delivery_address, COALESCE(o.delivery_price,0) AS delivery_price, o.created_at
       FROM orders o JOIN users u ON o.user_id = u.id ORDER BY o.created_at DESC LIMIT 100`
    ).catch(() => q(
      `SELECT o.id, u.name AS buyer, u.email, o.total_price, o.delivery_method,
         o.delivery_address, o.created_at
       FROM orders o JOIN users u ON o.user_id = u.id ORDER BY o.created_at DESC LIMIT 100`
    ));

    const topUsers = await q(
      `SELECT u.name, u.email, COUNT(o.id) AS cnt, SUM(o.total_price) AS spent
       FROM users u JOIN orders o ON o.user_id = u.id
       GROUP BY u.id, u.name, u.email ORDER BY spent DESC LIMIT 20`
    );

    // ─── Стили ────────────────────────────────────────────────────────────
    const PURPLE   = "FF8E24AA";
    const DPURPLE  = "FF6A1B9A";
    const LIGHT    = "FFF3E5F5";
    const WHITE    = "FFFFFFFF";
    const SUMMARY  = "FFCE93D8";
    const ROW_EVEN = "FFFDF8FF";
    const ROW_ODD  = "FFFCE4EC";

    const setTitle = (ws, merge, text) => {
      ws.mergeCells(merge);
      const c = ws.getCell(merge.split(":")[0]);
      c.value = text;
      c.font = { name:"Arial", bold:true, size:13, color:{argb:DPURPLE} };
      c.alignment = { horizontal:"center", vertical:"middle" };
      c.fill = { type:"pattern", pattern:"solid", fgColor:{argb:LIGHT} };
    };

    const setHeader = row => {
      row.eachCell(c => {
        c.font = { name:"Arial", bold:true, color:{argb:WHITE} };
        c.fill = { type:"pattern", pattern:"solid", fgColor:{argb:PURPLE} };
        c.alignment = { horizontal:"center", vertical:"middle", wrapText:true };
        c.border = { top:{style:"thin"}, bottom:{style:"thin"}, left:{style:"thin"}, right:{style:"thin"} };
      });
      row.height = 28;
    };

    const setDataRow = (row, i) => {
      row.eachCell(c => {
        c.font = { name:"Arial" };
        c.fill = { type:"pattern", pattern:"solid", fgColor:{argb: i%2===0 ? ROW_EVEN : ROW_ODD} };
        c.border = {
          top:{style:"thin",color:{argb:"FFE1BEE7"}}, bottom:{style:"thin",color:{argb:"FFE1BEE7"}},
          left:{style:"thin",color:{argb:"FFE1BEE7"}}, right:{style:"thin",color:{argb:"FFE1BEE7"}}
        };
      });
    };

    const setSumRow = row => {
      row.eachCell(c => {
        c.font = { name:"Arial", bold:true, color:{argb:DPURPLE} };
        c.fill = { type:"pattern", pattern:"solid", fgColor:{argb:SUMMARY} };
        c.border = { top:{style:"medium",color:{argb:PURPLE}}, bottom:{style:"medium",color:{argb:PURPLE}}, left:{style:"thin"}, right:{style:"thin"} };
      });
    };

    const dLabel = { pickup:"Самовывоз", courier:"Курьер" };
    const wb = new ExcelJS.Workbook();
    wb.creator = "LavandaBloom";
    wb.created = new Date();

    // ─── Предварительные расчёты ─────────────────────────────────────────
    const totalOrdersAll   = byDelivery.reduce((s,d) => s + Number(d.cnt), 0);
    const totalRevenueAll  = byDelivery.reduce((s,d) => s + Number(d.revenue), 0);
    const totalDeliveryAll = byDelivery.reduce((s,d) => s + (Number(d.delivery_sum)||0), 0);
    const totalGoodsRev    = totalRevenueAll - totalDeliveryAll;
    const avgCheck         = totalOrdersAll > 0 ? Math.round(totalRevenueAll / totalOrdersAll) : 0;
    const courierRow       = byDelivery.find(d => d.delivery_method === "courier");
    const courierOrders    = courierRow ? Number(courierRow.cnt) : 0;
    const deliveryPct      = totalOrdersAll > 0 ? Math.round(courierOrders / totalOrdersAll * 100) : 0;
    const topProduct       = productStats[0];

    // ─── Лист 0: Сводка ──────────────────────────────────────────────────
    const ws0 = wb.addWorksheet("Сводка");

    ws0.mergeCells("A1:D1");
    const titleCell0 = ws0.getCell("A1");
    titleCell0.value = "LavandaBloom — Финансовый отчёт";
    titleCell0.font = { name:"Arial", bold:true, size:16, color:{argb:DPURPLE} };
    titleCell0.alignment = { horizontal:"center", vertical:"middle" };
    titleCell0.fill = { type:"pattern", pattern:"solid", fgColor:{argb:LIGHT} };
    ws0.getRow(1).height = 36;

    ws0.mergeCells("A2:D2");
    const dateCell0 = ws0.getCell("A2");
    dateCell0.value = "Сформирован: " + new Date().toLocaleString("ru-RU");
    dateCell0.font = { name:"Arial", size:10, italic:true, color:{argb:"FF9C27B0"} };
    dateCell0.alignment = { horizontal:"center" };

    ws0.addRow([]);

    // Заголовки KPI-таблицы
    const kpiHdr = ws0.addRow(["Показатель", "Значение", "Пояснение", ""]);
    kpiHdr.eachCell(c => {
      c.font = { name:"Arial", bold:true, color:{argb:WHITE} };
      c.fill = { type:"pattern", pattern:"solid", fgColor:{argb:PURPLE} };
      c.alignment = { horizontal:"center", vertical:"middle" };
      c.border = { top:{style:"thin"}, bottom:{style:"thin"}, left:{style:"thin"}, right:{style:"thin"} };
    });
    kpiHdr.height = 24;

    const kpiData = [
      ["\uD83D\uDCB0 Общая выручка",          totalRevenueAll,   "тенге (товары + доставка)"],
      ["\uD83D\uDED2 Выручка от товаров",       totalGoodsRev,     "тенге (без учёта доставки)"],
      ["\uD83D\uDE9A Выручка от доставки",      totalDeliveryAll,  "тенге (плата за курьера)"],
      ["\uD83D\uDCE6 Всего заказов",            totalOrdersAll,    "шт."],
      ["\uD83D\uDEF5 Заказов с курьером",       courierOrders,     "шт. (" + deliveryPct + "% от всех)"],
      ["\uD83D\uDCB3 Средний чек",              avgCheck,          "тенге на один заказ"],
      ["\uD83D\uDCCA Доля доставки в выручке", totalRevenueAll > 0 ? Math.round(totalDeliveryAll / totalRevenueAll * 100) : 0, "% от общей выручки"],
      ["\uD83C\uDFC6 Топ товар",               topProduct ? topProduct.title : "—", topProduct ? Number(topProduct.total_revenue) + " \u20B8" : ""],
    ];

    kpiData.forEach((kpi, i) => {
      const row = ws0.addRow([kpi[0], kpi[1], kpi[2], ""]);
      const bg = i % 2 === 0 ? ROW_EVEN : ROW_ODD;
      row.eachCell(c => {
        c.fill = { type:"pattern", pattern:"solid", fgColor:{argb:bg} };
        c.border = {
          top:{style:"thin",color:{argb:"FFE1BEE7"}}, bottom:{style:"thin",color:{argb:"FFE1BEE7"}},
          left:{style:"thin",color:{argb:"FFE1BEE7"}}, right:{style:"thin",color:{argb:"FFE1BEE7"}}
        };
      });
      row.getCell(1).font = { name:"Arial", size:11, color:{argb:DPURPLE} };
      if (typeof kpi[1] === "number") {
        row.getCell(2).numFmt = "#,##0";
        row.getCell(2).alignment = { horizontal:"right" };
        row.getCell(2).font = { name:"Arial", size:12, bold:true, color:{argb:PURPLE} };
      } else {
        row.getCell(2).font = { name:"Arial", size:11, bold:true, color:{argb:PURPLE} };
      }
      row.getCell(3).font = { name:"Arial", size:10, italic:true, color:{argb:"FF9C27B0"} };
      row.height = 22;
    });

    ws0.addRow([]);

    // Мини-таблица разбивки по доставке
    const delSubHdr = ws0.addRow(["Разбивка по способу доставки", "", "", ""]);
    ws0.mergeCells(`A${delSubHdr.number}:D${delSubHdr.number}`);
    delSubHdr.getCell(1).font = { name:"Arial", bold:true, size:12, color:{argb:DPURPLE} };
    delSubHdr.getCell(1).fill = { type:"pattern", pattern:"solid", fgColor:{argb:LIGHT} };
    delSubHdr.height = 22;

    const delHdr2 = ws0.addRow(["Способ доставки", "Заказов", "Выручка от товаров (\u20B8)", "Сумма доставки (\u20B8)"]);
    delHdr2.eachCell(c => {
      c.font = { name:"Arial", bold:true, color:{argb:WHITE} };
      c.fill = { type:"pattern", pattern:"solid", fgColor:{argb:PURPLE} };
      c.alignment = { horizontal:"center" };
      c.border = { top:{style:"thin"}, bottom:{style:"thin"}, left:{style:"thin"}, right:{style:"thin"} };
    });

    byDelivery.forEach((d, i) => {
      const ds2 = Number(d.delivery_sum) || 0;
      const rv2 = Number(d.revenue) || 0;
      const row = ws0.addRow([dLabel[d.delivery_method] || d.delivery_method, Number(d.cnt), rv2 - ds2, ds2]);
      const bg = i % 2 === 0 ? ROW_EVEN : ROW_ODD;
      row.eachCell(c => {
        c.fill = { type:"pattern", pattern:"solid", fgColor:{argb:bg} };
        c.border = {
          top:{style:"thin",color:{argb:"FFE1BEE7"}}, bottom:{style:"thin",color:{argb:"FFE1BEE7"}},
          left:{style:"thin",color:{argb:"FFE1BEE7"}}, right:{style:"thin",color:{argb:"FFE1BEE7"}}
        };
      });
      row.getCell(2).alignment = {horizontal:"center"};
      row.getCell(3).numFmt = "#,##0";
      row.getCell(4).numFmt = "#,##0";
    });

    const delSum0 = ws0.addRow(["ИТОГО", totalOrdersAll, totalGoodsRev, totalDeliveryAll]);
    setSumRow(delSum0);
    delSum0.getCell(2).alignment = {horizontal:"center"};
    delSum0.getCell(3).numFmt = "#,##0";
    delSum0.getCell(4).numFmt = "#,##0";

    ws0.columns = [{width:36},{width:18},{width:28},{width:24}];

    // ─── Лист 1: Товары ──────────────────────────────────────────────────
    const ws1 = wb.addWorksheet("Товары");
    setTitle(ws1, "A1:E1", "LavandaBloom — Продажи по товарам");
    ws1.addRow([]);
    setHeader(ws1.addRow(["№", "Название товара", "Цена (₸)", "Продано шт.", "Выручка (₸)"]));
    const r1s = 4;
    productStats.forEach((p, i) => {
      const row = ws1.addRow([i+1, p.title, Number(p.price), Number(p.total_qty), Number(p.total_revenue)]);
      setDataRow(row, i);
      row.getCell(3).numFmt = "#,##0"; row.getCell(4).alignment = {horizontal:"center"}; row.getCell(5).numFmt = "#,##0";
    });
    const r1e = r1s + productStats.length - 1;
    // Итого — реальные числа, не формулы (формулы не пересчитываются без Excel)
    const totalQty = productStats.reduce((s, p) => s + Number(p.total_qty), 0);
    const totalRev = productStats.reduce((s, p) => s + Number(p.total_revenue), 0);
    const s1 = ws1.addRow(["","ИТОГО","", totalQty, totalRev]);
    setSumRow(s1); s1.getCell(4).alignment={horizontal:"center"}; s1.getCell(5).numFmt = "#,##0"; s1.getCell(4).numFmt = "#,##0";
    ws1.columns = [{width:5},{width:36},{width:14},{width:14},{width:18}];

    // ─── Лист 2: По дням ─────────────────────────────────────────────────
    const ws2 = wb.addWorksheet("По дням");
    setTitle(ws2, "A1:C1", "LavandaBloom — Заказы по дням");
    ws2.addRow([]);
    setHeader(ws2.addRow(["Дата", "Кол-во заказов", "Выручка (₸)"]));
    ordersByDay.forEach((d, i) => {
      const row = ws2.addRow([
        d.day ? new Date(d.day).toLocaleDateString("ru-RU") : "—", Number(d.cnt), Number(d.revenue)
      ]);
      setDataRow(row, i);
      row.getCell(2).alignment = {horizontal:"center"}; row.getCell(3).numFmt = "#,##0";
    });
    ws2.columns = [{width:16},{width:18},{width:20}];

    // ─── Лист 3: Анализ доставки ─────────────────────────────────────────
    const ws3 = wb.addWorksheet("Доставка");
    setTitle(ws3, "A1:F1", "LavandaBloom — Анализ доставки  |  Склад: Астана, Жиембет жырау 2");
    ws3.addRow([]);
    setHeader(ws3.addRow([
      "Способ доставки", "Заказов", "Выручка от товаров (₸)",
      "Сумма за доставку (₸)", "Ср. цена доставки (₸)", "Макс. цена доставки (₸)"
    ]));
    byDelivery.forEach((d, i) => {
      const dsum = Number(d.delivery_sum) || 0;
      const rev  = Number(d.revenue) || 0;
      const cnt  = Number(d.cnt) || 0;
      const row = ws3.addRow([
        dLabel[d.delivery_method] || d.delivery_method,
        cnt,
        rev - dsum,
        dsum > 0 ? dsum : 0,
        d.avg_delivery ? Math.round(Number(d.avg_delivery)) : 0,
        d.max_delivery ? Math.round(Number(d.max_delivery)) : 0
      ]);
      setDataRow(row, i);
      row.getCell(2).alignment = {horizontal:"center"};
      [3,4,5,6].forEach(n => { row.getCell(n).numFmt="#,##0"; });
    });
    // Итоговая строка по доставке
    const totalOrders = byDelivery.reduce((s,d) => s + Number(d.cnt), 0);
    const totalRevenue = byDelivery.reduce((s,d) => s + Number(d.revenue), 0);
    const totalDeliverySum = byDelivery.reduce((s,d) => s + (Number(d.delivery_sum)||0), 0);
    const sumRow3 = ws3.addRow(["ИТОГО", totalOrders, totalRevenue - totalDeliverySum, totalDeliverySum, "", ""]);
    setSumRow(sumRow3);
    sumRow3.getCell(2).alignment={horizontal:"center"};
    [3,4].forEach(n => sumRow3.getCell(n).numFmt="#,##0");
    ws3.addRow([]);
    const noteRow = ws3.addRow([`Тариф: ${DELIVERY_BASE_PRICE} ₸ минимум + ${DELIVERY_PRICE_PER_KM} ₸/км от склада`]);
    noteRow.getCell(1).font = {name:"Arial", italic:true, color:{argb:DPURPLE}};
    ws3.mergeCells(`A${noteRow.number}:F${noteRow.number}`);
    ws3.columns = [{width:18},{width:12},{width:26},{width:24},{width:26},{width:26}];

    // ─── Лист 4: Детали заказов ──────────────────────────────────────────
    const ws4 = wb.addWorksheet("Детали заказов");
    setTitle(ws4, "A1:H1", "LavandaBloom — Детали заказов с ценой доставки");
    ws4.addRow([]);
    setHeader(ws4.addRow([
      "№", "Покупатель", "Email", "Товары (₸)",
      "Доставка", "Цена доставки (₸)", "Адрес доставки", "Дата"
    ]));
    orderDetails.forEach((o, i) => {
      const dp = Number(o.delivery_price) || 0;
      const tp = Number(o.total_price) || 0;
      const row = ws4.addRow([
        o.id, o.buyer, o.email,
        tp - dp,
        dLabel[o.delivery_method] || "—",
        dp,
        o.delivery_address || "Самовывоз",
        o.created_at ? new Date(o.created_at).toLocaleString("ru-RU") : "—"
      ]);
      setDataRow(row, i);
      row.getCell(4).numFmt = "#,##0";
      row.getCell(6).numFmt = "#,##0";
    });
    ws4.columns = [{width:6},{width:20},{width:28},{width:16},{width:14},{width:22},{width:42},{width:20}];

    // ─── Лист 5: Топ покупатели ──────────────────────────────────────────
    const ws5 = wb.addWorksheet("Топ покупатели");
    setTitle(ws5, "A1:D1", "LavandaBloom — Топ покупатели");
    ws5.addRow([]);
    setHeader(ws5.addRow(["Имя", "Email", "Заказов", "Потрачено (₸)"]));
    topUsers.forEach((u, i) => {
      const row = ws5.addRow([u.name, u.email, Number(u.cnt), Number(u.spent)]);
      setDataRow(row, i); row.getCell(3).alignment={horizontal:"center"}; row.getCell(4).numFmt="#,##0";
    });
    ws5.columns = [{width:24},{width:30},{width:12},{width:20}];

    const dateStr = new Date().toISOString().slice(0,10);
    res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition",`attachment; filename="LavandaBloom_report_${dateStr}.xlsx"`);
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error("Ошибка отчёта:", err);
    res.status(500).json({ error: err.message });
  }
});


// ── ДЕТАЛЬНЫЙ ОТЧЁТ ПО ПОЛЬЗОВАТЕЛЯМ ─────────────────────────────────────
app.get("/user-report", async (req, res) => {
  const q = (sql, params = []) =>
    new Promise((resolve, reject) =>
      db.query(sql, params, (err, rows) => err ? reject(err) : resolve(rows))
    );

  try {
    const users = await q(
      `SELECT u.id, u.name, u.email,
         COUNT(DISTINCT o.id) AS orders_cnt,
         COALESCE(SUM(o.total_price), 0) AS total_spent,
         COALESCE(SUM(COALESCE(o.delivery_price,0)),0) AS delivery_spent,
         COALESCE(SUM(o.total_price) - SUM(COALESCE(o.delivery_price,0)), 0) AS goods_spent,
         MIN(o.created_at) AS first_order,
         MAX(o.created_at) AS last_order,
         GROUP_CONCAT(DISTINCT o.delivery_method ORDER BY o.delivery_method SEPARATOR ', ') AS delivery_methods
       FROM users u
       LEFT JOIN orders o ON o.user_id = u.id
       GROUP BY u.id, u.name, u.email
       ORDER BY total_spent DESC`
    );

    const orderRows = await q(
      `SELECT o.id AS order_id, o.user_id, o.total_price,
         COALESCE(o.delivery_price,0) AS delivery_price,
         o.delivery_method, o.delivery_address, o.created_at,
         p.title AS product_title, p.price AS product_price, ci.quantity
       FROM orders o
       LEFT JOIN cart_items ci ON ci.user_id = o.user_id
       LEFT JOIN products p ON p.id = ci.product_id
       ORDER BY o.user_id, o.created_at DESC, o.id`
    );

    // Группируем товары по заказу
    const itemsByOrder = {};
    const seenOrderUser = {};
    orderRows.forEach(r => {
      const key = r.order_id;
      if (!itemsByOrder[key]) itemsByOrder[key] = { order: r, items: [] };
      if (r.product_title) itemsByOrder[key].items.push(r);
    });

    // Заказы по пользователю (дедупликация)
    const ordByUser = {};
    orderRows.forEach(r => {
      if (!ordByUser[r.user_id]) ordByUser[r.user_id] = [];
      const already = ordByUser[r.user_id].find(o => o.order_id === r.order_id);
      if (!already) ordByUser[r.user_id].push(r);
    });

    const PURPLE  = "FF8E24AA";
    const DPURPLE = "FF6A1B9A";
    const LIGHT   = "FFF3E5F5";
    const WHITE   = "FFFFFFFF";
    const SUMMARY = "FFCE93D8";
    const ROW_EVEN = "FFFDF8FF";
    const ROW_ODD  = "FFFCE4EC";
    const dLabel  = { pickup:"Самовывоз", courier:"Курьер" };

    const setHeader = row => {
      row.eachCell(c => {
        c.font = { name:"Arial", bold:true, color:{argb:WHITE} };
        c.fill = { type:"pattern", pattern:"solid", fgColor:{argb:PURPLE} };
        c.alignment = { horizontal:"center", vertical:"middle", wrapText:true };
        c.border = { top:{style:"thin"}, bottom:{style:"thin"}, left:{style:"thin"}, right:{style:"thin"} };
      });
      row.height = 26;
    };

    const setData = (row, i) => {
      row.eachCell(c => {
        c.font = { name:"Arial" };
        c.fill = { type:"pattern", pattern:"solid", fgColor:{argb: i%2===0 ? ROW_EVEN : ROW_ODD} };
        c.border = {
          top:{style:"thin",color:{argb:"FFE1BEE7"}}, bottom:{style:"thin",color:{argb:"FFE1BEE7"}},
          left:{style:"thin",color:{argb:"FFE1BEE7"}}, right:{style:"thin",color:{argb:"FFE1BEE7"}}
        };
      });
    };

    const setSumRow = row => {
      row.eachCell(c => {
        c.font = { name:"Arial", bold:true, color:{argb:DPURPLE} };
        c.fill = { type:"pattern", pattern:"solid", fgColor:{argb:SUMMARY} };
        c.border = { top:{style:"medium",color:{argb:PURPLE}}, bottom:{style:"medium",color:{argb:PURPLE}}, left:{style:"thin"}, right:{style:"thin"} };
      });
    };

    const wb = new ExcelJS.Workbook();
    wb.creator = "LavandaBloom";
    wb.created = new Date();

    // ─── Лист 1: Сводка по всем пользователям ────────────────────────────
    const ws1 = wb.addWorksheet("Все пользователи");
    ws1.mergeCells("A1:J1");
    const t1 = ws1.getCell("A1");
    t1.value = "LavandaBloom — Детальный отчёт по пользователям";
    t1.font = { name:"Arial", bold:true, size:14, color:{argb:DPURPLE} };
    t1.alignment = { horizontal:"center", vertical:"middle" };
    t1.fill = { type:"pattern", pattern:"solid", fgColor:{argb:LIGHT} };
    ws1.getRow(1).height = 32;
    ws1.addRow([]);

    setHeader(ws1.addRow([
      "Имя", "Email", "Заказов",
      "Всего потрачено (₸)", "На товары (₸)", "На доставку (₸)", "Средний чек (₸)",
      "Методы доставки", "Первый заказ", "Последний заказ"
    ]));

    let totOrd = 0, totSpent = 0, totGoods = 0, totDeliv = 0;
    users.forEach((u, i) => {
      const orders = Number(u.orders_cnt)    || 0;
      const spent  = Number(u.total_spent)   || 0;
      const goods  = Number(u.goods_spent)   || 0;
      const deliv  = Number(u.delivery_spent)|| 0;
      const avg    = orders > 0 ? Math.round(spent / orders) : 0;
      totOrd += orders; totSpent += spent; totGoods += goods; totDeliv += deliv;

      const row = ws1.addRow([
        u.name, u.email, orders, spent, goods, deliv, avg,
        u.delivery_methods || "Нет заказов",
        u.first_order ? new Date(u.first_order).toLocaleDateString("ru-RU") : "—",
        u.last_order  ? new Date(u.last_order).toLocaleDateString("ru-RU")  : "—"
      ]);
      setData(row, i);
      row.getCell(3).alignment = {horizontal:"center"};
      [4,5,6,7].forEach(n => row.getCell(n).numFmt = "#,##0");
    });

    const totAvg = totOrd > 0 ? Math.round(totSpent / totOrd) : 0;
    const sRow = ws1.addRow(["ИТОГО", "", totOrd, totSpent, totGoods, totDeliv, totAvg, "", "", ""]);
    setSumRow(sRow);
    sRow.getCell(3).alignment = {horizontal:"center"};
    [4,5,6,7].forEach(n => sRow.getCell(n).numFmt = "#,##0");
    ws1.columns = [{width:22},{width:30},{width:10},{width:22},{width:18},{width:18},{width:16},{width:24},{width:16},{width:16}];

    // ─── Отдельный лист для каждого пользователя у кого есть заказы ──────
    users.forEach(u => {
      const userOrders = ordByUser[u.id] || [];
      if (!userOrders.length) return;

      const sheetName = (u.name || u.email).replace(/[:\\/\?\*\[\]]/g, "").slice(0, 28);
      const ws = wb.addWorksheet(sheetName);

      // Шапка
      ws.mergeCells("A1:F1");
      const ut = ws.getCell("A1");
      ut.value = u.name + "  |  " + u.email;
      ut.font = { name:"Arial", bold:true, size:12, color:{argb:DPURPLE} };
      ut.alignment = { horizontal:"center", vertical:"middle" };
      ut.fill = { type:"pattern", pattern:"solid", fgColor:{argb:LIGHT} };
      ws.getRow(1).height = 28;

      // KPI
      const uOrders = Number(u.orders_cnt)    || 0;
      const uSpent  = Number(u.total_spent)   || 0;
      const uGoods  = Number(u.goods_spent)   || 0;
      const uDeliv  = Number(u.delivery_spent)|| 0;

      ws.addRow([]);
      const kHdr = ws.addRow(["Показатель", "Значение", ""]);
      kHdr.eachCell(c => {
        c.font = { name:"Arial", bold:true, color:{argb:WHITE} };
        c.fill = { type:"pattern", pattern:"solid", fgColor:{argb:PURPLE} };
        c.alignment = { horizontal:"center" };
        c.border = { top:{style:"thin"}, bottom:{style:"thin"}, left:{style:"thin"}, right:{style:"thin"} };
      });

      [
        ["Всего заказов",   uOrders],
        ["Потрачено всего", uSpent],
        ["На товары",       uGoods],
        ["На доставку",     uDeliv],
        ["Средний чек",     uOrders > 0 ? Math.round(uSpent / uOrders) : 0],
      ].forEach((k, i) => {
        const r = ws.addRow([k[0], k[1], i < 1 ? "шт." : "₸"]);
        r.eachCell(c => {
          c.fill = { type:"pattern", pattern:"solid", fgColor:{argb: i%2===0 ? ROW_EVEN : ROW_ODD} };
          c.border = { top:{style:"thin",color:{argb:"FFE1BEE7"}}, bottom:{style:"thin",color:{argb:"FFE1BEE7"}}, left:{style:"thin",color:{argb:"FFE1BEE7"}}, right:{style:"thin",color:{argb:"FFE1BEE7"}} };
        });
        r.getCell(1).font = { name:"Arial", color:{argb:DPURPLE} };
        r.getCell(2).font = { name:"Arial", bold:true, color:{argb:PURPLE} };
        if (typeof k[1] === "number" && i > 0) r.getCell(2).numFmt = "#,##0";
        else r.getCell(2).alignment = {horizontal:"center"};
      });

      ws.addRow([]);

      // Таблица заказов
      setHeader(ws.addRow(["№ зак.", "Дата", "Товары (₸)", "Доставка", "Стоим. доставки (₸)", "Адрес"]));

      userOrders.forEach((o, i) => {
        const dp = Number(o.delivery_price) || 0;
        const tp = Number(o.total_price) || 0;

        const r = ws.addRow([
          o.order_id,
          o.created_at ? new Date(o.created_at).toLocaleString("ru-RU") : "—",
          tp - dp,
          dLabel[o.delivery_method] || "—",
          dp,
          o.delivery_address || "Самовывоз"
        ]);
        setData(r, i * 2);
        r.getCell(3).numFmt = "#,##0";
        r.getCell(5).numFmt = "#,##0";

        // Строки товаров внутри заказа
        const items = (itemsByOrder[o.order_id] || {}).items || [];
        items.forEach(item => {
          const iRow = ws.addRow([
            "", "     └ " + item.product_title,
            Number(item.product_price) * Number(item.quantity),
            "", "",
            Number(item.quantity) + " шт. × " + Number(item.product_price) + " ₸"
          ]);
          iRow.eachCell(c => {
            c.font = { name:"Arial", size:10, italic:true, color:{argb:"FF9C27B0"} };
            c.fill = { type:"pattern", pattern:"solid", fgColor:{argb:"FFFAF5FF"} };
          });
          iRow.getCell(3).numFmt = "#,##0";
        });
      });

      // Итог
      const uSumRow = ws.addRow(["", "ИТОГО", uGoods, "", uDeliv, ""]);
      setSumRow(uSumRow);
      uSumRow.getCell(3).numFmt = "#,##0";
      uSumRow.getCell(5).numFmt = "#,##0";

      ws.columns = [{width:8},{width:24},{width:16},{width:14},{width:22},{width:42}];
    });

    const dateStr = new Date().toISOString().slice(0,10);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="LavandaBloom_users_${dateStr}.xlsx"`);
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error("Ошибка детального отчёта:", err);
    res.status(500).json({ error: err.message });
  }
});


// ── СКЛАДСКОЙ ОТЧЁТ ───────────────────────────────────────────────────────
app.get("/warehouse-report", async (req, res) => {
  const q = (sql, params = []) =>
    new Promise((resolve, reject) =>
      db.query(sql, params, (err, rows) => err ? reject(err) : resolve(rows))
    );

  try {
    // Сколько каждого товара продано (ушло со склада)
    const sold = await q(
      `SELECT p.id, p.title, p.price,
         COALESCE(SUM(ci.quantity), 0) AS sold_qty,
         COALESCE(SUM(ci.quantity * p.price), 0) AS sold_revenue
       FROM products p
       LEFT JOIN cart_items ci ON ci.product_id = p.id
       GROUP BY p.id, p.title, p.price
       ORDER BY sold_qty DESC`
    );

    // Динамика продаж по дням для каждого товара
    const byDay = await q(
      `SELECT DATE(o.created_at) AS day, p.id AS product_id, p.title,
         SUM(ci.quantity) AS qty
       FROM orders o
       JOIN cart_items ci ON ci.user_id = o.user_id
       JOIN products p ON p.id = ci.product_id
       GROUP BY DATE(o.created_at), p.id, p.title
       ORDER BY day DESC, qty DESC`
    );

    // Какие товары заказывали вместе (попарные комбинации)
    const combos = await q(
      `SELECT a.product_id AS id_a, pa.title AS title_a,
         b.product_id AS id_b, pb.title AS title_b,
         COUNT(*) AS times
       FROM cart_items a
       JOIN cart_items b ON a.user_id = b.user_id AND a.product_id < b.product_id
       JOIN products pa ON pa.id = a.product_id
       JOIN products pb ON pb.id = b.product_id
       GROUP BY a.product_id, b.product_id
       ORDER BY times DESC
       LIMIT 20`
    ).catch(() => []);

    // Топ покупатели по каждому товару
    const topBuyers = await q(
      `SELECT p.id AS product_id, p.title, u.name, u.email,
         SUM(ci.quantity) AS qty
       FROM cart_items ci
       JOIN products p ON p.id = ci.product_id
       JOIN users u ON u.id = ci.user_id
       GROUP BY p.id, p.title, u.id, u.name, u.email
       ORDER BY p.id, qty DESC`
    );

    // Группируем топ-покупателей по товару
    const buyersByProduct = {};
    topBuyers.forEach(r => {
      if (!buyersByProduct[r.product_id]) buyersByProduct[r.product_id] = [];
      buyersByProduct[r.product_id].push(r);
    });

    // ─── Excel стили ──────────────────────────────────────────────────────
    const PURPLE   = "FF8E24AA";
    const DPURPLE  = "FF6A1B9A";
    const LIGHT    = "FFF3E5F5";
    const WHITE    = "FFFFFFFF";
    const SUMMARY  = "FFCE93D8";
    const ROW_EVEN = "FFFDF8FF";
    const ROW_ODD  = "FFFCE4EC";
    const GREEN    = "FFE8F5E9";
    const DGREEN   = "FF2E7D32";

    const setHeader = row => {
      row.eachCell(c => {
        c.font = { name:"Arial", bold:true, color:{argb:WHITE} };
        c.fill = { type:"pattern", pattern:"solid", fgColor:{argb:PURPLE} };
        c.alignment = { horizontal:"center", vertical:"middle", wrapText:true };
        c.border = { top:{style:"thin"}, bottom:{style:"thin"}, left:{style:"thin"}, right:{style:"thin"} };
      });
      row.height = 26;
    };

    const setData = (row, i) => {
      row.eachCell(c => {
        c.font = { name:"Arial" };
        c.fill = { type:"pattern", pattern:"solid", fgColor:{argb: i%2===0 ? ROW_EVEN : ROW_ODD} };
        c.border = {
          top:{style:"thin",color:{argb:"FFE1BEE7"}},
          bottom:{style:"thin",color:{argb:"FFE1BEE7"}},
          left:{style:"thin",color:{argb:"FFE1BEE7"}},
          right:{style:"thin",color:{argb:"FFE1BEE7"}}
        };
      });
    };

    const setSumRow = row => {
      row.eachCell(c => {
        c.font = { name:"Arial", bold:true, color:{argb:DPURPLE} };
        c.fill = { type:"pattern", pattern:"solid", fgColor:{argb:SUMMARY} };
        c.border = {
          top:{style:"medium",color:{argb:PURPLE}},
          bottom:{style:"medium",color:{argb:PURPLE}},
          left:{style:"thin"}, right:{style:"thin"}
        };
      });
    };

    const setTitle = (ws, merge, text) => {
      ws.mergeCells(merge);
      const c = ws.getCell(merge.split(":")[0]);
      c.value = text;
      c.font = { name:"Arial", bold:true, size:13, color:{argb:DPURPLE} };
      c.alignment = { horizontal:"center", vertical:"middle" };
      c.fill = { type:"pattern", pattern:"solid", fgColor:{argb:LIGHT} };
      ws.getRow(1).height = 30;
    };

    const wb = new ExcelJS.Workbook();
    wb.creator = "LavandaBloom";
    wb.created = new Date();

    // ─── Лист 1: Движение товаров (сколько ушло) ─────────────────────────
    const ws1 = wb.addWorksheet("Движение товаров");
    setTitle(ws1, "A1:G1", "LavandaBloom — Склад: движение товаров");

    ws1.mergeCells("A2:G2");
    const sub1 = ws1.getCell("A2");
    sub1.value = "Склад: Астана, Жиембет жырау 2  |  " + new Date().toLocaleDateString("ru-RU");
    sub1.font = { name:"Arial", size:10, italic:true, color:{argb:"FF9C27B0"} };
    sub1.alignment = { horizontal:"center" };

    ws1.addRow([]);
    setHeader(ws1.addRow([
      "№", "Название товара", "Цена за шт. (₸)",
      "Продано шт.", "Выручка (₸)",
      "% от общих продаж", "Статус"
    ]));

    const totalSoldQty = sold.reduce((s, p) => s + Number(p.sold_qty), 0);
    const totalSoldRev = sold.reduce((s, p) => s + Number(p.sold_revenue), 0);

    sold.forEach((p, i) => {
      const qty = Number(p.sold_qty);
      const rev = Number(p.sold_revenue);
      const pct = totalSoldQty > 0 ? (qty / totalSoldQty * 100).toFixed(1) : "0.0";
      const status = qty === 0 ? "Нет продаж" : qty >= 10 ? "Хит продаж 🔥" : qty >= 3 ? "Продаётся" : "Мало продаж";

      const row = ws1.addRow([i+1, p.title, Number(p.price), qty, rev, Number(pct), status]);
      setData(row, i);
      row.getCell(3).numFmt = "#,##0";
      row.getCell(4).alignment = {horizontal:"center"};
      row.getCell(5).numFmt = "#,##0";
      row.getCell(6).numFmt = "0.0\"%\"";
      row.getCell(6).alignment = {horizontal:"center"};
      row.getCell(7).alignment = {horizontal:"center"};

      // Хиты — зелёный акцент
      if (qty >= 10) {
        row.getCell(2).font = { name:"Arial", bold:true, color:{argb:DGREEN} };
        row.getCell(7).font = { name:"Arial", bold:true, color:{argb:DGREEN} };
      } else if (qty === 0) {
        row.getCell(7).font = { name:"Arial", italic:true, color:{argb:"FFBDBDBD"} };
      }
    });

    const sumRow1 = ws1.addRow(["", "ИТОГО", "", totalSoldQty, totalSoldRev, "100%", ""]);
    setSumRow(sumRow1);
    sumRow1.getCell(4).alignment = {horizontal:"center"};
    sumRow1.getCell(4).numFmt = "#,##0";
    sumRow1.getCell(5).numFmt = "#,##0";
    sumRow1.getCell(6).alignment = {horizontal:"center"};

    ws1.columns = [{width:5},{width:32},{width:18},{width:14},{width:18},{width:20},{width:16}];

    // ─── Лист 2: Динамика по дням ─────────────────────────────────────────
    const ws2 = wb.addWorksheet("По дням");
    setTitle(ws2, "A1:D1", "LavandaBloom — Склад: продажи по дням");
    ws2.addRow([]);
    setHeader(ws2.addRow(["Дата", "Товар", "Кол-во (шт.)", "Выручка (₸)"]));

    // Группируем по дням
    const dayGroups = {};
    byDay.forEach(r => {
      const d = new Date(r.day).toLocaleDateString("ru-RU");
      if (!dayGroups[d]) dayGroups[d] = [];
      dayGroups[d].push(r);
    });

    let rowIdx = 0;
    Object.entries(dayGroups).forEach(([day, items]) => {
      const dayQty = items.reduce((s, r) => s + Number(r.qty), 0);
      const dayRow = ws2.addRow([day, `Всего за день: ${dayQty} шт.`, dayQty, ""]);
      dayRow.eachCell(c => {
        c.font = { name:"Arial", bold:true, color:{argb:DPURPLE} };
        c.fill = { type:"pattern", pattern:"solid", fgColor:{argb:LIGHT} };
        c.border = { top:{style:"thin",color:{argb:"FFE1BEE7"}}, bottom:{style:"thin",color:{argb:"FFE1BEE7"}}, left:{style:"thin",color:{argb:"FFE1BEE7"}}, right:{style:"thin",color:{argb:"FFE1BEE7"}} };
      });
      dayRow.getCell(3).alignment = {horizontal:"center"};

      items.forEach((r, i) => {
        const itemRow = ws2.addRow(["", "  └ " + r.title, Number(r.qty), Number(r.qty) * (sold.find(p=>p.id===r.product_id)||{price:0}).price]);
        setData(itemRow, rowIdx++);
        itemRow.getCell(3).alignment = {horizontal:"center"};
        itemRow.getCell(4).numFmt = "#,##0";
        itemRow.getCell(2).font = { name:"Arial", size:10, color:{argb:"FF6A1B9A"} };
      });
    });

    ws2.columns = [{width:14},{width:36},{width:14},{width:16}];

    // ─── Лист 3: Топ покупатели по каждому товару ─────────────────────────
    const ws3 = wb.addWorksheet("Кто что брал");
    setTitle(ws3, "A1:E1", "LavandaBloom — Склад: кто что покупал");
    ws3.addRow([]);
    setHeader(ws3.addRow(["Товар", "Покупатель", "Email", "Кол-во (шт.)", "Сумма (₸)"]));

    let ri = 0;
    sold.forEach(p => {
      const buyers = buyersByProduct[p.id] || [];
      if (!buyers.length) return;

      const prodRow = ws3.addRow([p.title, `${buyers.length} покупателей`, "", buyers.reduce((s,b)=>s+Number(b.qty),0), ""]);
      prodRow.eachCell(c => {
        c.font = { name:"Arial", bold:true, color:{argb:DPURPLE} };
        c.fill = { type:"pattern", pattern:"solid", fgColor:{argb:LIGHT} };
        c.border = { top:{style:"thin",color:{argb:"FFE1BEE7"}}, bottom:{style:"thin",color:{argb:"FFE1BEE7"}}, left:{style:"thin",color:{argb:"FFE1BEE7"}}, right:{style:"thin",color:{argb:"FFE1BEE7"}} };
      });
      prodRow.getCell(4).alignment = {horizontal:"center"};

      buyers.forEach((b, i) => {
        const bRow = ws3.addRow(["  └ " + b.title, b.name, b.email, Number(b.qty), Number(b.qty) * Number(p.price)]);
        setData(bRow, ri++);
        bRow.getCell(4).alignment = {horizontal:"center"};
        bRow.getCell(5).numFmt = "#,##0";
        bRow.getCell(1).font = { name:"Arial", size:10, color:{argb:"FF9C27B0"} };
      });
    });

    ws3.columns = [{width:32},{width:22},{width:30},{width:14},{width:16}];

    // ─── Лист 4: Товары заказывали вместе ────────────────────────────────
    if (combos.length > 0) {
      const ws4 = wb.addWorksheet("Заказывали вместе");
      setTitle(ws4, "A1:C1", "LavandaBloom — Склад: часто берут вместе");
      ws4.addRow([]);
      setHeader(ws4.addRow(["Товар А", "Товар Б", "Раз вместе"]));

      combos.forEach((c, i) => {
        const row = ws4.addRow([c.title_a, c.title_b, Number(c.times)]);
        setData(row, i);
        row.getCell(3).alignment = {horizontal:"center"};
        row.getCell(3).numFmt = "#,##0";
      });

      ws4.columns = [{width:36},{width:36},{width:16}];
    }

    const dateStr = new Date().toISOString().slice(0,10);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="LavandaBloom_warehouse_${dateStr}.xlsx"`);
    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error("Ошибка складского отчёта:", err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));