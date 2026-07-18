export default async function handler(req, res) {
  // Allow CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { type, to, messages } = req.body;
  const token = 'aLu3cQ8JuLSCfI8BP31Pmzr7V0ni2vFrvWGi2C3Kt2IPOSJ6nDzQM6skwyTd7a9i2iUEQVs2bBBxSqF1UQVaAY1GndiNieyDfuZoKv6eWI934g7ynRBmEPW7ykVVaRcic2S0y+gGbciqloKUeZ9TfQdB04t89/1O/w1cDnyilFU=';

  const endpoint = type === 'push'
    ? 'https://api.line.me/v2/bot/message/push'
    : 'https://api.line.me/v2/bot/message/broadcast';

  const body = type === 'push'
    ? { to, messages }
    : { messages };

  try {
    const response = await fetch(endpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token}`
      },
      body: JSON.stringify(body)
    });
    
    let responseText = "";
    try {
      const data = await response.json();
      return res.status(200).json(data);
    } catch(e) {
      responseText = await response.text();
      return res.status(response.status).json({ error: responseText });
    }
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
