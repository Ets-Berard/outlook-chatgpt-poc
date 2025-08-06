import React, { useState } from 'react';

export default function App() {
  const [summary, setSummary] = useState('');
  const [loading, setLoading] = useState(false);

  const summarizeEmail = async () => {
    setLoading(true);
    Office.context.mailbox.item.body.getAsync("text", async (result) => {
      const emailContent = result.value;
      const response = await fetch('https://<TON_URL_PROXY>/chatgpt', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ text: emailContent })
      });
      const data = await response.json();
      setSummary(data.summary);
      setLoading(false);
    });
  };

  return (
    <div style={{ padding: "20px" }}>
      <h2>Résumé de l'email</h2>
      <button onClick={summarizeEmail} disabled={loading}>
        {loading ? 'En cours...' : 'Résumer l’email'}
      </button>
      <div style={{ marginTop: '20px', whiteSpace: 'pre-line' }}>
        {summary}
      </div>
    </div>
  );
}
