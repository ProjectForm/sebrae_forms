export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).end();

  const data = req.body;

  try {
    // 1. Obter token de acesso
    const tokenRes = await fetch(
      `https://login.microsoftonline.com/6d60b55c-2576-4d60-ae33-11df0ea07983/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          grant_type: 'client_credentials',
          client_id: '1385935a-db4a-4514-868e-c76764856c36',
          client_secret: process.env.AZURE_CLIENT_SECRET,
          scope: 'https://graph.microsoft.com/.default'
        })
      }
    );

    const tokenData = await tokenRes.json();
    const token = tokenData.access_token;

    if (!token) {
      console.error('Token error:', tokenData);
      return res.status(500).json({ error: 'Erro ao obter token' });
    }

    // 2. Buscar coluna A para achar ultima linha preenchida (a partir da linha 5)
    const colARes = await fetch(
      `https://graph.microsoft.com/v1.0/users/juanaga@sebraesp.com.br/drive/root:/CONTROLE DE PJs - Faturamento - 2026.xlsx:/workbook/worksheets('DADOS')/range(address='A5:A500')`,
      {
        headers: { Authorization: `Bearer ${token}` }
      }
    );

    const colAData = await colARes.json();
    const rows = colAData.values || [];

    // Encontrar primeira linha vazia na coluna A (a partir da linha 5)
    let nextRow = 5;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] !== null && rows[i][0] !== '') {
        nextRow = 5 + i + 1;
      }
    }

    // 3. Escrever na proxima linha vazia
    const hoje = new Date().toLocaleDateString('pt-BR');
    const values = [[
      data.gestor || '',   // A - Gestor
      hoje,                // B - Data de Atualizacao
      data.razao || '',    // C - Razao Social
      data.cpf || '',      // D - CPF
      data.cel || '',      // E - Celular
      data.cnpj || '',     // F - CNPJ
      data.porte || ''     // G - Porte
    ]];

    const writeRes = await fetch(
      `https://graph.microsoft.com/v1.0/users/juanaga@sebraesp.com.br/drive/root:/CONTROLE DE PJs - Faturamento - 2026.xlsx:/workbook/worksheets('DADOS')/range(address='A${nextRow}:G${nextRow}')`,
      {
        method: 'PATCH',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ values })
      }
    );

    if (!writeRes.ok) {
      const writeErr = await writeRes.json();
      console.error('Write error:', writeErr);
      return res.status(500).json({ error: 'Erro ao escrever na planilha' });
    }

    return res.status(200).json({ success: true });

  } catch (err) {
    console.error('Handler error:', err);
    return res.status(500).json({ error: err.message });
  }
}
