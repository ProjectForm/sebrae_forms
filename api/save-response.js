export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).end();

  const data = req.body;

  try {
    // 1. Token
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
      console.error('Token error:', JSON.stringify(tokenData));
      return res.status(500).json({ error: 'Erro ao obter token', detail: tokenData });
    }

    const user = 'juanaga@sebraesp.com.br';
    const fileName = 'CONTROLE DE PJs - Faturamento - 2026.xlsx';
    const sheet = 'DADOS';

    // 2. Buscar arquivo pelo nome na raiz do OneDrive
    const searchRes = await fetch(
      `https://graph.microsoft.com/v1.0/users/${user}/drive/root/children`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const searchData = await searchRes.json();
    console.log('Drive root files:', JSON.stringify(searchData?.value?.map(f => f.name)));

    // Buscar o arquivo na raiz ou em subpastas
    let fileId = null;
    if (searchData.value) {
      const found = searchData.value.find(f => f.name === fileName);
      if (found) fileId = found.id;
    }

    // Se nao achou na raiz, tenta buscar por search
    if (!fileId) {
      const s2 = await fetch(
        `https://graph.microsoft.com/v1.0/users/${user}/drive/root/search(q='CONTROLE DE PJs')`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      const s2data = await s2.json();
      console.log('Search results:', JSON.stringify(s2data?.value?.map(f => f.name)));
      if (s2data.value && s2data.value.length > 0) {
        fileId = s2data.value[0].id;
      }
    }

    if (!fileId) {
      return res.status(500).json({ error: 'Arquivo nao encontrado no OneDrive' });
    }

    console.log('File ID encontrado:', fileId);

    // 3. Buscar coluna A a partir da linha 5
    const colRes = await fetch(
      `https://graph.microsoft.com/v1.0/users/${user}/drive/items/${fileId}/workbook/worksheets('${sheet}')/range(address='A5:A500')`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const colData = await colRes.json();

    if (colData.error) {
      console.error('Range error:', JSON.stringify(colData));
      return res.status(500).json({ error: 'Erro ao ler planilha', detail: colData });
    }

    const rows = colData.values || [];
    let nextRow = 5;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] !== null && rows[i][0] !== '') {
        nextRow = 5 + i + 1;
      }
    }

    console.log('Escrevendo na linha:', nextRow);

    // 4. Escrever
    const hoje = new Date().toLocaleDateString('pt-BR');
    const values = [[
      data.gestor || '',
      hoje,
      data.razao  || '',
      data.cpf    || '',
      data.cel    || '',
      data.cnpj   || '',
      data.porte  || ''
    ]];

    const writeRes = await fetch(
      `https://graph.microsoft.com/v1.0/users/${user}/drive/items/${fileId}/workbook/worksheets('${sheet}')/range(address='A${nextRow}:G${nextRow}')`,
      {
        method: 'PATCH',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ values })
      }
    );

    if (!writeRes.ok) {
      const writeErr = await writeRes.json();
      console.error('Write error:', JSON.stringify(writeErr));
      return res.status(500).json({ error: 'Erro ao escrever', detail: writeErr });
    }

    return res.status(200).json({ success: true, row: nextRow });

  } catch (err) {
    console.error('Handler error:', err.message);
    return res.status(500).json({ error: err.message });
  }
}
