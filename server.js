import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import webhookHandler from './api/webhook.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = 3000;
const HOST = '0.0.0.0';

app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true }));

app.all('/api/webhook', async (req, res) => {
  try {
    await webhookHandler(req, res);
  } catch (err) {
    console.error('Webhook error:', err);
    if (!res.headersSent) {
      res.status(500).json({ ok: false, error: err.message || 'Internal Server Error' });
    }
  }
});

app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, HOST, () => {
  console.log(`Server running on http://${HOST}:${PORT}`);
});
