const express = require('express');
const app = express();
app.get('/', (req, res) => {
  res.send('Hello World');
});
const bodyParser = require('body-parser');
const axios = require('axios');
require('dotenv').config();

app.use(bodyParser.json());

const tenantId = process.env.TENANT_ID;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const worksheetName = process.env.WORKSHEET_NAME;
const express = require('express');


// Serve static files (e.g., index.html)
app.use(express.static('index.html')); // or your folder name

// Default route
app.get('/', (req, res) => {
  res.sendFile(__dirname + '/public/index.html');
});

// Render expects you to use process.env.PORT
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});

let accessToken = '';

// Function to get an access token
async function getAccessToken() {
    const response = await axios.post(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
        new URLSearchParams({
            client_id: clientId,
            scope: 'https://graph.microsoft.com/.default',
            client_secret: clientSecret,
            grant_type: 'client_credentials'
        })
    );
    accessToken = response.data.access_token;
}

// Function to add scanned data to Excel
async function addToExcel(text) {
    const timestamp = new Date().toLocaleString();
    await getAccessToken();

    // Try to add a table (if it already exists, ignore the error)
    await axios.post(
        `https://graph.microsoft.com/v1.0/me/drive/root:/QRData.xlsx:/workbook/worksheets('${worksheetName}')/tables/add`,
        { address: 'A1:B1', hasHeaders: true },
        { headers: { Authorization: `Bearer ${accessToken}` } }
    ).catch(() => {}); // Ignore if table already exists

    // Add a new row with scanned text and timestamp
    const res = await axios.post(
        `https://graph.microsoft.com/v1.0/me/drive/root:/QRData.xlsx:/workbook/tables/1/rows/add`,
        { values: [[text, timestamp]] },
        { headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' } }
    );

    return res.status;
}

// Endpoint to receive QR text from frontend
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

