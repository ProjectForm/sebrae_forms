export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).end();

  const { email, code } = req.body;

  const response = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${process.env.RESEND_API_KEY}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      from: 'SEBRAE <onboarding@resend.dev>',
      to: email,
      subject: 'Seu código de verificação SEBRAE',
      html: `
        <div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;padding:32px;background:#f0f4fa;border-radius:16px">
          <div style="background:#003F88;padding:20px;border-radius:12px;text-align:center;margin-bottom:24px">
            <span style="color:#fff;font-size:22px;font-weight:700">SEBRAE</span>
          </div>
          <h2 style="color:#003F88;margin-bottom:8px">Código de Verificação</h2>
          <p style="color:#6B7280;margin-bottom:24px">Use o código abaixo para acessar o formulário de adesão:</p>
          <div style="background:#fff;border:2px solid #D1DBF0;border-radius:12px;padding:24px;text-align:center;margin-bottom:24px">
            <span style="font-size:36px;font-weight:700;color:#003F88;letter-spacing:8px">${code}</span>
          </div>
          <p style="color:#6B7280;font-size:12px;text-align:center">Este código expira em 10 minutos. Se não solicitou, ignore este e-mail.</p>
        </div>
      `
    })
  });

  if (!response.ok) return res.status(500).json({ error: 'Erro ao enviar e-mail' });
  return res.status(200).json({ success: true });
}
