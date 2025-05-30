// Import packages at the top
import express, { static as expressStatic } from 'express';
import { json } from 'body-parser';
import { post } from 'axios';
import dotenv from 'dotenv';
dotenv.config();

// Create Express app
const app = express();
app.use(json());

// Serve static files (like index.html) from "public" folder
app.use(expressStatic('public'));

// Route to serve the HTML file
app.get('/', (req, res) => {
  res.sendFile(process.cwd() + '/public/index.html'); // Assumes HTML is in a "public" folder
});

// Load environment variables
const tenantId = process.env.TENANT_ID;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const worksheetName = process.env.WORKSHEET_NAME;

let accessToken = '';

// Get access token from Microsoft
async function getAccessToken() {
  const response = await post(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    new URLSearchParams({
      client_id: clientId,
      scope: 'https://graph.microsoft.com/.default',
      client_secret: clientSecret,
      grant_type: 'client_credentials'
    })
  );
  accessToken = response.data.access_token;
}

// Add QR text and timestamp to Excel
async function addToExcel(text) {
  const timestamp = new Date().toLocaleString();
  await getAccessToken();

  // Try to add table (if it exists, skip)
  await post(
    `https://graph.microsoft.com/v1.0/me/drive/root:/QRData.xlsx:/workbook/worksheets('${worksheetName}')/tables/add`,
    { address: 'A1:B1', hasHeaders: true },
    { headers: { Authorization: `Bearer ${accessToken}` } }
  ).catch(() => {}); // Ignore if table already exists

  // Add row
  const res = await post(
    `https://graph.microsoft.com/v1.0/me/drive/root:/QRData.xlsx:/workbook/tables/1/rows/add`,
    { values: [[text, timestamp]] },
    { headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' } }
  );

  return res.status;
}

// Receive QR text from frontend
app.post('/upload', async (req, res) => {
  const { text } = req.body;
  try {
    await addToExcel(text);
    res.json({ message: 'QR code text uploaded to Excel successfully!' });
  } catch (err) {
    console.error(err.response?.data || err.message);
    res.status(500).json({ message: 'Failed to upload to Excel.' });
  }
});

// Start server (Render uses process.env.PORT)
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`✅ Server running on port ${PORT}`);
});