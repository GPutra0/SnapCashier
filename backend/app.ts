import express from "express";
import path from "path";

const app = express();

app.use(express.json());

// API Routes
app.get("/api/health", (req, res) => {
  res.json({ status: "ok", message: "Bakend POS System running" });
});

// Example API for products (read-only for now or could be expanded)
app.get("/api/products", (req, res) => {
    // In a real app, this would fetch from a database
    res.json({ message: "Product API endpoint" });
});

export default app;
