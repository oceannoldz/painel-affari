const express = require('express');
const fetch = require('node-fetch');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;

// Permite CORS do seu domínio!
app.use(cors({
  origin: 'https://oceannoldz.github.io/painel-affari/', // Troque pelo domínio do seu site
}));

app.use(express.json());

// Rota proxy para entregar pedidos
app.post('/api/entregar', async (req, res) => {
  try {
    // Troque pela SUA URL do Apps Script!
    const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzpEw9xdXFOuTxaIiX8CHVxCmnnjtUyzbOMUxYcAQ2fGOr71YfimotW_hYxlDeEQvE-/exec';

    const resposta = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(req.body)
    });

    // O Apps Script pode retornar texto puro
    const text = await resposta.text();
    res.status(200).send(text);

  } catch (err) {
    console.error("Erro ao enviar para Apps Script:", err);
    res.status(500).json({ success: false, error: String(err) });
  }
});

app.listen(PORT, () => {
  console.log(`Backend rodando em http://localhost:${PORT}`);
});
