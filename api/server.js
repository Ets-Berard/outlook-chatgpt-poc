import express from 'express';
import axios from 'axios';
import cors from 'cors';

const app = express();
app.use(cors());
app.use(express.json());

app.post('/chatgpt', async (req, res) => {
  try {
    const { text } = req.body;
    const response = await axios.post(
      'https://api.openai.com/v1/chat/completions',
      {
        model: 'gpt-4o-mini',
        messages: [{ role: 'user', content: `Résume cet email en français : ${text}` }],
        temperature: 0.5
      },
      { headers: { Authorization: `Bearer ${process.env.OPENAI_API_KEY}` } }
    );
    res.json({ summary: response.data.choices[0].message.content });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.listen(3000, () => console.log('API proxy démarré sur port 3000'));
