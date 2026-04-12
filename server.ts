import express from "express";
import axios from "axios";
import path from "path";
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// API Route: Generic Proxy for Google Apps Script POST requests
app.post("/api/proxy-apps-script", async (req, res) => {
  const { url, payload } = req.body;
  if (!url) return res.status(400).json({ error: "URL is required" });

  try {
    const jsonString = typeof payload === 'string' ? payload : JSON.stringify(payload);
    const response = await axios.post(url, jsonString, {
      headers: { 'Content-Type': 'text/plain' },
      timeout: 30000,
      validateStatus: () => true // Allow handling all status codes manually
    });

    console.log("GAS Response Status:", response.status);

    if (response.status === 401) {
      return res.status(401).json({ 
        error: "Unauthorized", 
        details: "Lỗi 401: Google Apps Script yêu cầu xác thực hoặc chưa được cấu hình 'Anyone' (Bất kỳ ai) có quyền truy cập. Vui lòng kiểm tra lại phần 'Deploy' trong Google Script." 
      });
    }

    if (typeof response.data === 'string' && response.data.includes('<!DOCTYPE html>')) {
      const errorMatch = response.data.match(/errorMessage">([^<]+)/) || response.data.match(/SyntaxError: ([^<]+)/);
      const errorDetail = errorMatch ? errorMatch[1] : "Lỗi thực thi Script (kiểm tra lại mã GAS)";
      
      // Log the full HTML for debugging in the server console
      console.error("Full GAS Error HTML:", response.data);
      
      return res.status(500).json({ 
        error: "Google Apps Script Error", 
        details: errorDetail,
        debugHtml: response.data // Gửi kèm HTML để người dùng có thể xem chi tiết lỗi trong Console
      });
    }
    res.send(response.data);
  } catch (error: any) {
    res.status(500).json({ error: "Proxy Connection Error", details: error.message });
  }
});

// API Route: Image Proxy to bypass CORS
app.get("/api/proxy-image", async (req, res) => {
  const imageUrl = req.query.url as string;
  if (!imageUrl) {
    return res.status(400).send("URL is required");
  }

  try {
    const response = await axios.get(imageUrl, {
      responseType: "arraybuffer",
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
      },
      timeout: 10000, // 10s timeout
    });

    const contentType = response.headers["content-type"];
    res.setHeader("Content-Type", contentType || "image/jpeg");
    res.setHeader("Cache-Control", "public, max-age=86400"); // Cache for 1 day
    res.send(response.data);
  } catch (error) {
    console.error("Proxy error:", error);
    res.status(500).send("Failed to fetch image");
  }
});

// Vite middleware for development
async function setupServer() {
  if (process.env.NODE_ENV !== "production") {
    const { createServer: createViteServer } = await import("vite");
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    // Serve static files in production
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*all", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }
}

// Export the app for Vercel
export default app;

// Start the server
const PORT = 3000;
setupServer().then(() => {
  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
});
