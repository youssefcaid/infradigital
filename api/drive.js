const crypto = require("crypto");

const DRIVE_API = "https://www.googleapis.com/drive/v3";
const TOKEN_URL = "https://oauth2.googleapis.com/token";

const ROOTS = {
  "1er année": "1ipcOlzgzo9pqQ8gl8__tYch3oKEKUnbX",
  "2eme année": "1Piu6Tbhhjwd5vzYbze7b6BlJ9yeZSxG3"
};

function createJWT() {
  const now = Math.floor(Date.now() / 1000);
  const header = { alg: "RS256", typ: "JWT" };
  const payload = {
    iss: process.env.GOOGLE_CLIENT_EMAIL,
    scope: "https://www.googleapis.com/auth/drive.readonly",
    aud: TOKEN_URL,
    iat: now,
    exp: now + 3600
  };
  const encode = obj => Buffer.from(JSON.stringify(obj)).toString("base64url");
  const unsignedToken = `${encode(header)}.${encode(payload)}`;
  const privateKey = process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n");
  const signer = crypto.createSign("RSA-SHA256");
  signer.update(unsignedToken);
  const signature = signer.sign(privateKey, "base64url");
  return `${unsignedToken}.${signature}`;
}

let cachedToken = null;
let cachedTokenExpiry = 0;

async function getAccessToken() {
  const now = Math.floor(Date.now() / 1000);
  if (cachedToken && now < cachedTokenExpiry - 60) {
    return cachedToken;
  }
  const jwt = createJWT();
  const response = await fetch(TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer", assertion: jwt })
  });
  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Google Auth Error: ${error}`);
  }
  const data = await response.json();
  cachedToken = data.access_token;
  cachedTokenExpiry = now + (data.expires_in || 3600);
  return cachedToken;
}

async function driveRequest(url, token) {
  const response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Google Drive Error: ${error}`);
  }
  return response.json();
}

async function listFolder(folderId, token) {
  const files = [];
  let pageToken = null;
  do {
    const params = new URLSearchParams({
      q: `'${folderId}' in parents and trashed = false`,
      fields: "nextPageToken,files(id,name,mimeType,size,modifiedTime,webViewLink,webContentLink)",
      pageSize: "1000",
      orderBy: "name"
    });
    if (pageToken) params.set("pageToken", pageToken);
    const data = await driveRequest(`${DRIVE_API}/files?${params}`, token);
    files.push(...(data.files || []));
    pageToken = data.nextPageToken;
  } while (pageToken);
  return files;
}

async function readFolder(folderId, token) {
  const items = await listFolder(folderId, token);
  const results = await Promise.all(items.map(async item => {
    if (item.mimeType === "application/vnd.google-apps.folder") {
      const children = await readFolder(item.id, token);
      return { id: item.id, name: item.name, type: "folder", children };
    }
    return {
      id: item.id, name: item.name, type: "file", mimeType: item.mimeType,
      size: item.size || null, modifiedTime: item.modifiedTime || null,
      viewUrl: item.webViewLink || `https://drive.google.com/file/d/${item.id}/view`,
      downloadUrl: item.webContentLink || null
    };
  }));
  return results;
}

function countPDFs(items) {
  let count = 0;
  for (const item of items) {
    if (item.type === "file") {
      const isPDF = item.mimeType === "application/pdf" || item.name.toLowerCase().endsWith(".pdf");
      if (isPDF) count++;
    } else if (item.type === "folder") {
      count += countPDFs(item.children);
    }
  }
  return count;
}

function findCategory(module, names) {
  const target = names.map(x => x.toLowerCase());
  return module.children.find(item => item.type === "folder" && target.includes(item.name.trim().toLowerCase()));
}

function isEffFolder(name) {
  const n = name.trim().toLowerCase();
  return n === "eff" || n.startsWith("eff -") || n.startsWith("eff-") || n.startsWith("eff ");
}

function collectPdfFiles(items) {
  const files = [];
  for (const item of items) {
    if (item.type === "file") {
      const isPDF = item.mimeType === "application/pdf" || item.name.toLowerCase().endsWith(".pdf");
      if (isPDF) files.push(item);
    } else if (item.type === "folder") {
      files.push(...collectPdfFiles(item.children));
    }
  }
  return files;
}

module.exports = async (req, res) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET");
  res.setHeader("Cache-Control", "s-maxage=600, stale-while-revalidate=86400");
  try {
    const token = await getAccessToken();
    const yearEntries = Object.entries(ROOTS);
    const yearResults = await Promise.all(yearEntries.map(async ([yearName, rootId]) => {
      const rootItems = await readFolder(rootId, token);

      // فولدر EFF غير موجود جوه المودولات، هو فولدر منفصل بجانب المودولات (M201, M202...)
      const effFolder = rootItems.find(item => item.type === "folder" && isEffFolder(item.name));
      const effFiles = effFolder ? collectPdfFiles(effFolder.children) : [];

      const modules = rootItems
        .filter(item => item.type === "folder" && !isEffFolder(item.name))
        .map(module => {
          const cours = findCategory(module, ["cours", "course"]);
          const exercices = findCategory(module, ["exercices", "exercice"]);
          const controle = findCategory(module, ["contrôle", "controle", "contrôles"]);
          const efm = findCategory(module, ["efm"]);
          return {
            id: module.id, name: module.name,
            cours: cours ? countPDFs(cours.children) : 0,
            exercices: exercices ? countPDFs(exercices.children) : 0,
            controle: controle ? countPDFs(controle.children) : 0,
            efm: efm ? countPDFs(efm.children) : 0,
            content: module.children
          };
        });
      return [yearName, { modules, eff: effFiles }];
    }));
    const result = Object.fromEntries(yearResults);
    res.status(200).json({ success: true, updated: new Date().toISOString(), years: result });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, error: error.message });
  }
};