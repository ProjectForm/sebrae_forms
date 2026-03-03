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

    // 2. Buscar o arquivo no OneDrive do usuario
    const driveRes = await fetch(
      `https://graph.microsoft.com/v1.0/users/juanaga@sebraesp.com.br/drive/root:/CONTROLE DE PJs - Faturamento - 2026.xlsx:/workbook/worksheets`,
      {
        headers: { Authorization: `Bearer ${token}` }
      }
    );

    const driveData = await driveRes.json();
    
    if (!driveData.value || driveData.value.length === 0) {
      console.error('Worksheet error:', driveData);
      return res.status(500).json({ error: 'Planilha nao encontrada' });
    }

    const sheetName = driveData.value[0].name;

    // 3. Buscar a ultima linha preenchida (coluna A)
    const rangeRes = await fetch(
      `https://graph.microsoft.com/v1.0/users/juanaga@sebraesp.com.br/drive/root:/CONTROLE DE PJs - Faturamento - 2026.xlsx:/workbook/worksheets('${sheetName}')/usedRange`,
      {
        headers: { Authorization: `Bearer ${token}` }
      }
    );

    const rangeData = await rangeRes.json();
    const lastRow = rangeData.rowCount || 1;
    const nextRow = lastRow + 1;

    // 4. Escrever na proxima linha vazia
    const values = [[
      data.ts || new Date().toLocaleString('pt-BR'),
      data.email || '',
      data.razao || '',
      data.cpf || '',
      data.cel || '',
      data.cnpj || '',
      data.porte || '',
      data.gestor || '',
      data.resp || '',
      data.wa || '',
      data.comp || '',
      data.lgpd || '',
      data.brand || ''
    ]];

    const writeRes = await fetch(
      `https://graph.microsoft.com/v1.0/users/juanaga@sebraesp.com.br/drive/root:/CONTROLE DE PJs - Faturamento - 2026.xlsx:/workbook/worksheets('${sheetName}')/range(address='A${nextRow}:M${nextRow}')`,
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
