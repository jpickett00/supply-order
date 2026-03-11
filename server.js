import express from "express";
import cors from "cors";
import axios from "axios";
import dotenv from "dotenv";
import path from "path";
import { fileURLToPath } from 'url';

dotenv.config();

const __filename =
fileURLToPath(import.meta.url);
const __dirname =
path.dirname(__filename);

const app = express();

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

const tenantId = process.env.TENANT_ID;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const userId = process.env.USER_ID;
const tableName = process.env.TABLE_NAME;
const fileName = process.env.FILE_NAME;

let accessToken = "";

async function getAccessToken() {
  const response = await axios.post(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    new URLSearchParams({
      client_id: clientId,
      scope: "https://graph.microsoft.com/.default",
      client_secret: clientSecret,
      grant_type: "client_credentials",
    })
  );

  accessToken = response.data.access_token;
}

async function addToExcel(text) {
  const timestamp = new Date().toLocaleString();

  console.log("QR text received:", text);

  await getAccessToken();

  const url =
    `https://graph.microsoft.com/v1.0/users/${userId}` +
    `/drive/root:/${fileName}:/workbook/tables/${tableName}/rows/add`;

  const res = await axios.post(
    url,
    {
      values: [[text, timestamp]],
    },
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
    }
  );

  console.log("Upload success:", res.status);

  return res.status;
}

app.post("/upload", async (req, res) => {
  const { text } = req.body;

  try {
    const status = await addToExcel(text);
    res.json({ success: true, status });
  } catch (err) {
    console.error("Upload failed:", err.response?.data || err.message);
    res.status(500).json({ error: "Upload failed" });
  }
});

const PORT = process.env.PORT || 10000;

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});