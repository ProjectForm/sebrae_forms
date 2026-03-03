export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).end();

  const data = req.body;

  try {
    // 1. Obter access token usando refresh token
    const tokenRes = await fetch(
      `https://login.microsoftonline.com/common/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          grant_type: 'refresh_token',
          client_id: '1385935a-db4a-4514-868e-c76764856c36',
          refresh_token: process.env.MS_REFRESH_TOKEN,
          scope: 'https://graph.microsoft.com/Files.ReadWrite offline_access'
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
    const sheet = 'DADOS';

    // 2. Buscar arquivo por search
    const searchRes = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/root/search(q='CONTROLE DE PJs - Faturamento - 2026')`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const searchData = await searchRes.json();
    console.log('Search:', JSON.stringify(searchData?.value?.map(f => f.name)));

    if (!searchData.value || searchData.value.length === 0) {
      return res.status(500).json({ error: 'Arquivo nao encontrado' });
    }

    const fileId = searchData.value[0].id;
    console.log('File ID:', fileId);

    // 3. Buscar coluna A a partir da linha 5
    const colRes = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets('${sheet}')/range(address='A5:A500')`,
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

    // 4. Escrever na proxima linha vazia
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
      `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/worksheets('${sheet}')/range(address='A${nextRow}:G${nextRow}')`,
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
      console.error('Write error:', JSON.stringify(writeErr));
      return res.status(500).json({ error: 'Erro ao escrever', detail: writeErr });
    }

    return res.status(200).json({ success: true, row: nextRow });

  } catch (err) {
    console.error('Handler error:', err.message);
    return res.status(500).json({ error: err.message });
  }
}
