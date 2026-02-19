const express = require('express');
const fetch = require('node-fetch');
const app = express();
const PORT = 3001; // Cambia el puerto si lo necesitas

// Permitir CORS desde tu frontend local
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
  next();
});

app.use(express.json());

const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzvGYWMsmoPmhSrDTsUgZLxp5Zr1zf-ERbg5jOyD58zrYO5vnWh8STguGqEC9fGiSlIbQ/exec';

app.post('/api/caja', async (req, res) => {
  try {
    const response = await fetch(SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(req.body)
    });
    const data = await response.json();
    res.json(data);
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`Proxy backend corriendo en http://localhost:${PORT}`);
});
