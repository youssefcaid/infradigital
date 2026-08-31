const crypto = require("crypto");

const DRIVE_API = "https://www.googleapis.com/drive/v3";
const TOKEN_URL = "https://oauth2.googleapis.com/token";

const ROOTS = {
  "1er année": "1ipcOlzgzo9pqQ8gl8__tYch3oKEKUnbX",
  "2eme année": "1Piu6Tbhhjwd5vzYbze7b6BlJ9yeZSxG3"
};

// ------------------------------------------------------------------
// 1) Cache ديال access token فالميموري (كان كيتصنع token جديد فكل
//    request، هادشي كيضيف round-trip زايدة ديال OAuth فكل مرة)
// ------------------------------------------------------------------
let cachedToken = null;
let tokenExpiresAt = 0;

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

async function getAccessToken() {
  // إلا كاين token صالح فالكاش، نستعملوه بلا ما نديرو request جديدة
  if (cachedToken && Date.now() < tokenExpiresAt) {
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
  // نخبيو الـ token شوية قبل ما يفوت الوقت ديالو (مارجن ديال دقيقة)
  tokenExpiresAt = Date.now() + (data.expires_in - 60) * 1000;
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

// ------------------------------------------------------------------
// 2) Recursion موازية (parallel) عوض ما تكون سلسلة (sequential)
//    قبل: كل فولدر فرعي كيتسنى لي قبلو يسالي (for...await) —
//    مع M101..M108 وكل وحدة فيها Cours/Exercices/Contrôle/EFM،
//    هادشي كيدير عشرات ديال الـ round-trips وحدة ورا وحدة.
//    دابا: كاع الفولدرات الفرعية كيتقراو فنفس الوقت ب Promise.all.
// ------------------------------------------------------------------
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

async function buildYearResult(yearName, rootId, token) {
  const rootItems = await readFolder(rootId, token);

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

  return { modules, eff: effFiles };
}

// ------------------------------------------------------------------
// 3) Cache ديال النتيجة كاملة فالميموري لمدة دقيقة (60s)
//    باش إلا جا request جديد قريب من لي قبل (نفس الـ warm instance)
//    ما نرجعوش نديرو كاع الـ calls ليGoogle Drive من جديد — كيرجع
//    الجواب فوري بلا تأخير.
// ------------------------------------------------------------------
let cachedResult = null;
let cachedAt = 0;
const CACHE_TTL_MS = 60 * 1000;

module.exports = async (req, res) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET");
  res.setHeader("Cache-Control", "s-maxage=60, stale-while-revalidate=300");
  try {
    if (cachedResult && Date.now() - cachedAt < CACHE_TTL_MS) {
      return res.status(200).json(cachedResult);
    }

    const token = await getAccessToken();

    // الجلب ديال السنتين كيتصاوب فنفس الوقت (parallel) عوض واحدة
    // ورا الأخرى
    const entries = await Promise.all(
      Object.entries(ROOTS).map(async ([yearName, rootId]) => [
        yearName,
        await buildYearResult(yearName, rootId, token)
      ])
    );
    const result = Object.fromEntries(entries);

    const payload = { success: true, updated: new Date().toISOString(), years: result };
    cachedResult = payload;
    cachedAt = Date.now();

    res.status(200).json(payload);
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, error: error.message });
  }
};
