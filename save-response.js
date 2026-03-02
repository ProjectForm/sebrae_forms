export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  if (req.method !== 'POST') return res.status(405).end();

  const data = req.body;
  const GITHUB_TOKEN = process.env.GITHUB_TOKEN;
  const REPO = 'Juan-gabrieldev/sebrae_forms';
  const FILE_PATH = 'respostas.json';

  const getRes = await fetch(`https://api.github.com/repos/${REPO}/contents/${FILE_PATH}`, {
    headers: {
      'Authorization': `Bearer ${GITHUB_TOKEN}`,
      'Accept': 'application/vnd.github+json'
    }
  });

  let respostas = [];
  let sha = null;

  if (getRes.ok) {
    const fileData = await getRes.json();
    sha = fileData.sha;
    respostas = JSON.parse(atob(fileData.content.replace(/\n/g, '')));
  }

  respostas.push({ ...data, timestamp: new Date().toISOString() });

  await fetch(`https://api.github.com/repos/${REPO}/contents/${FILE_PATH}`, {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${GITHUB_TOKEN}`,
      'Content-Type': 'application/json',
      'Accept': 'application/vnd.github+json'
    },
    body: JSON.stringify({
      message: 'Nova resposta formulário SEBRAE',
      content: btoa(unescape(encodeURIComponent(JSON.stringify(respostas, null, 2)))),
      ...(sha && { sha })
    })
  });

  return res.status(200).json({ success: true });
}
