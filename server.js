// Import packages
import express from 'express';
import bodyParser from 'body-parser';
import cors from 'cors';
import axios from 'axios';
import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';

// Enable environment variables
dotenv.config();

// Setup __dirname for ES modules
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Initialize express app
const app = express();
app.use(cors());
app.use(bodyParser.json());

// Serve static files from 'public'
app.use(express.static('public'));

// Serve index.html from /public
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Load environment variables
const tenantId = process.env.TENANT_ID;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const worksheetName = process.env.WORKSHEET_NAME;
const useMeEndpoint = process.env.USE_ME_ENDPOINT === 'true';
const driveRoot = useMeEndpoint ? 'me' : 'users/${process.env.USER_ID}';

let accessToken = '';

// Fetch Microsoft access token
async function getAccessToken() {
  const response = await axios.post(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    new URLSearchParams({
      client_id: clientId,
      scope: 'https://graph.microsoft.com/.default',
      client_secret: clientSecret,
      grant_type: 'client_credentials',
    })
  );
  accessToken = response.data.access_token;
}

// Upload QR data to Excel
async function addToExcel(text) {
  const timestamp = new Date().toLocaleString();
  await getAccessToken();
  console.log(`Uploading to: QRData.xlsx → worksheet '${worksheetName}'`);

  // Create table (ignore if it already exists)
  await axios.post(
    `https://graph.microsoft.com/v1.0/users/${driveRoot}/drive/root:/QRData.xlsx:/workbook/worksheets('${worksheetName}')/tables/add`,
    { address: 'A1:B1', hasHeaders: true },
    { headers: { Authorization: `Bearer ${accessToken}` } }
  ).catch(err => {
    console.error('Table creation error:', err?.response?.data || err.message);
  });

  // Add row
  const res = await axios.post(
    `https://graph.microsoft.com/v1.0/users/${driveRoot}/drive/root:/QRData.xlsx:/workbook/tables/1/rows/add`,
    { values: [[text, timestamp]] },
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
    }
  );

  return res.status;
}

// API endpoint
app.post('/upload', async (req, res) => {
  const { text } = req.body;
  try {
    const status = await addToExcel(text);
    res.status(200).json({ success: true, status });
  } catch (err) {
    console.error(err.response?.data || err.message);
    res.status(500).json({ message: 'Failed to upload to Excel.' });
  }
});

// Start server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`✅ Server running on port ${PORT}`);
});