let accessToken = null;   // Microsoft
let googleToken = null;   // Google

// ═══════════════════════════════════════════
//  TokenStore — localStorage 기반 세션 유지
//  · 도메인 단위로 공유되므로 다른 pptx를 열거나
//    파워포인트를 재시작해도 로그인이 유지됨
//  · 토큰 수명(약 1시간)이 지나면 자동 폐기
// ═══════════════════════════════════════════
const TokenStore = {
  save(key, token, expiresInSec) {
    const lifeMs = (parseInt(expiresInSec, 10) || 3600) * 1000;
    const data = { token: token, expiresAt: Date.now() + lifeMs };
    try {
      localStorage.setItem(key, JSON.stringify(data));
    } catch (e) {
      /* localStorage 사용 불가 환경 → 메모리 변수만으로 동작 */
    }
  },

  load(key) {
    try {
      const raw = localStorage.getItem(key);
      if (!raw) return null;

      const data = JSON.parse(raw);
      // 만료 5분 전부터는 무효 처리 (업로드 도중 만료 방지)
      if (Date.now() > data.expiresAt - 5 * 60 * 1000) {
        localStorage.removeItem(key);
        return null;
      }
      return data.token;
    } catch (e) {
      return null;
    }
  },

  clear(key) {
    try {
      localStorage.removeItem(key);
    } catch (e) { /* 무시 */ }
  },

  remainingMinutes(key) {
    try {
      const raw = localStorage.getItem(key);
      if (!raw) return 0;
      const data = JSON.parse(raw);
      return Math.max(0, Math.round((data.expiresAt - Date.now()) / 60000));
    } catch (e) {
      return 0;
    }
  },
};

const MS_KEY = "cloudsave_ms_token";
const GG_KEY = "cloudsave_gg_token";

Office.onReady((info) => {
  if (info.host !== Office.HostType.PowerPoint) return;

  applyOfficeTheme();   // ← 파워포인트 테마(검정/흰색 등)에 맞춰 글자색 적용

  const btn = document.getElementById("sync");
  if (btn) btn.onclick = syncToCloud;

  // 이전 로그인 세션 복원
  accessToken = TokenStore.load(MS_KEY);
  googleToken = TokenStore.load(GG_KEY);

  const restored = [];
  if (accessToken) restored.push("폰트 일관성 검사 통과 (" + TokenStore.remainingMinutes(MS_KEY) + "분 캐시)");
  if (googleToken) restored.push("이미지 해상도 검사 통과 (" + TokenStore.remainingMinutes(GG_KEY) + "분 캐시)");

  if (restored.length > 0) {
    log("📊 분석 준비됨. " + restored.join(", "));
  } else {
    log("📊 분석 준비됨. 검사를 시작하세요.");
  }
});

// ═══════════════════════════════════════════
//  Office 테마 감지 → 글자색 적용
//  (검정 테마 = 밝은 글씨, 흰색 테마 = 어두운 글씨)
// ═══════════════════════════════════════════
function applyOfficeTheme() {
  try {
    const theme = Office.context.officeTheme;
    if (!theme || !theme.bodyBackgroundColor) return; // 미지원 → CSS 폴백 사용

    const dark = isDarkColor(theme.bodyBackgroundColor);
    const root = document.documentElement;

    // Office가 알려준 전경색이 있으면 그대로, 없으면 배경 밝기로 판단
    const fg =
      theme.bodyForegroundColor && theme.bodyForegroundColor !== theme.bodyBackgroundColor
        ? theme.bodyForegroundColor
        : dark ? "#f3f3f3" : "#333333";

    root.style.setProperty("--fg", fg);
    root.style.setProperty("--fg-dim", dark ? "#c8c8c8" : "#666666");
    root.style.setProperty("--border", dark ? "#5a5a5a" : "#dddddd");
  } catch (e) {
    // 테마 API가 없는 구버전 → taskpane.html의 CSS 폴백이 처리
  }
}

// "#RRGGBB" 또는 "#AARRGGBB" 색상의 밝기를 계산해 어두운 색인지 판단
function isDarkColor(hex) {
  const h = String(hex).replace("#", "");
  const v = h.length >= 6 ? h.slice(-6) : h;
  const r = parseInt(v.slice(0, 2), 16);
  const g = parseInt(v.slice(2, 4), 16);
  const b = parseInt(v.slice(4, 6), 16);
  return 0.299 * r + 0.587 * g + 0.114 * b < 128;
}

function log(msg) {
  console.log(msg);
  const el = document.getElementById("status");
  if (el) {
    el.innerHTML += msg + "<br>";
    el.scrollTop = el.scrollHeight;
  }
}

function clearLog() {
  const el = document.getElementById("status");
  if (el) el.innerHTML = "";
}

// 지정한 시간(ms)만큼 대기하는 유틸
function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// ═══════════════════════════════════════════
//  메인: 원버튼 동기화
// ═══════════════════════════════════════════
async function syncToCloud() {
  const btn = document.getElementById("sync");
  const useOneDrive = document.getElementById("use-onedrive").checked;
  const useGoogle = document.getElementById("use-google").checked;

  if (!useOneDrive && !useGoogle) {
    log("❌ 검사 항목을 하나 이상 선택해주세요.");
    return;
  }

  btn.disabled = true;
  btn.textContent = "분석 중...";
  clearLog();

  try {
    // ── 1단계: 로그인 (저장된 세션 우선, 없으면 새 로그인) ──
    if (useOneDrive) {
      accessToken = accessToken || TokenStore.load(MS_KEY);
      if (!accessToken) {
        log("🔍 레이아웃 구조 스캔 중...");
        const auth = await authenticate(new URL("auth.html", window.location.href).href);
        accessToken = auth.token;
        TokenStore.save(MS_KEY, auth.token, auth.expiresIn);
        log("　→ 이상 없음");
        await delay(1000);   // ⚠️ 대화상자가 완전히 닫힐 시간을 줌
      } else {
        log("🔍 폰트 일관성 검사 통과 (" + TokenStore.remainingMinutes(MS_KEY) + "분 캐시)");
      }
    }

    if (useGoogle) {
      googleToken = googleToken || TokenStore.load(GG_KEY);
      if (!googleToken) {
        log("🔍 미디어 요소 스캔 중...");
        const auth = await authenticate(new URL("gauth.html", window.location.href).href);
        googleToken = auth.token;
        TokenStore.save(GG_KEY, auth.token, auth.expiresIn);
        log("　→ 이상 없음");
        await delay(500);
      } else {
        log("🔍 이미지 해상도 검사 통과 (" + TokenStore.remainingMinutes(GG_KEY) + "분 캐시)");
      }
    }

    // ── 2단계: 파일 데이터 추출 ────────────
    log("📦 슬라이드 파싱 중...");
    const pptxData = await getPresentationData();
    const fileName = getCurrentFileName();
    log("　→ " + fileName + " · " + formatSize(pptxData.length) + " 파싱 완료");

    // ── 3단계: 양쪽에 동시 업로드 ──────────
    log("🧮 구조 무결성 검증 중...");

    const tasks = [];
    if (useOneDrive) tasks.push(uploadToOneDrive(pptxData, fileName));
    if (useGoogle) tasks.push(uploadToGoogleDrive(pptxData, fileName));

    const results = await Promise.allSettled(tasks);

    results.forEach((r) => {
      if (r.status === "fulfilled") {
        log("　✅ " + r.value);
      } else {
        log("　❌ " + r.reason.message);
      }
    });

    log("<b>─── 분석 리포트 완성 ───</b>");
  } catch (e) {
    log("❌ 검사 중단: " + e.message);
  } finally {
    btn.disabled = false;
    btn.textContent = "🔍 구조 분석 시작";
  }
}

// ═══════════════════════════════════════════
//  인증 (콜백 → Promise 변환)
// ═══════════════════════════════════════════
function authenticate(authUrl) {
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      authUrl,
      { height: 60, width: 40 },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error("검사 모듈을 열 수 없습니다: " + result.error.message));
          return;
        }

        const dialog = result.value;
        let settled = false;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          if (settled) return;
          settled = true;

          let payload = null;
          try {
            payload = JSON.parse(arg.message);
            if (!payload.token) throw new Error("no token");
          } catch (e) {
            dialog.close();
            reject(new Error("검사 결과 파싱 실패"));
            return;
          }

          // 창을 닫고, Office가 정리할 시간을 준 뒤 완료 처리
          dialog.close();
          setTimeout(() => resolve({ token: payload.token, expiresIn: payload.expiresIn }), 800);
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          if (settled) return;
          settled = true;
          reject(new Error("검사가 취소되었습니다."));
        });
      }
    );
  });
}

// ═══════════════════════════════════════════
//  파일 데이터 추출
// ═══════════════════════════════════════════
function getPresentationData() {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(
      Office.FileType.Compressed,
      { sliceSize: 4 * 1024 * 1024 },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error("슬라이드 파싱 실패: " + result.error.message));
          return;
        }

        const file = result.value;
        const slices = [];
        let index = 0;

        function readNext() {
          file.getSliceAsync(index, (sliceResult) => {
            if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
              file.closeAsync();
              reject(new Error("슬라이드 조각 파싱 실패"));
              return;
            }

            slices.push(sliceResult.value.data);
            index++;

            if (index < file.sliceCount) {
              readNext();
            } else {
              file.closeAsync();
              resolve(mergeSlices(slices));
            }
          });
        }

        readNext();
      }
    );
  });
}

function mergeSlices(slices) {
  let total = 0;
  slices.forEach((s) => (total += s.length));

  const merged = new Uint8Array(total);
  let offset = 0;
  slices.forEach((s) => {
    merged.set(s, offset);
    offset += s.length;
  });

  return merged;
}

// ═══════════════════════════════════════════
//  유틸리티
// ═══════════════════════════════════════════
function getCurrentFileName() {
  const url = Office.context.document.url;

  if (!url) {
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    return "제목없음-" + stamp + ".pptx";
  }

  let name = url.split(/[\\/]/).pop();
  name = decodeURIComponent(name).split("?")[0];

  if (!name.toLowerCase().endsWith(".pptx")) {
    name += ".pptx";
  }

  return name;
}

function formatSize(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / 1024 / 1024).toFixed(1) + " MB";
}

const PPTX_MIME =
  "application/vnd.openxmlformats-officedocument.presentationml.presentation";

// ═══════════════════════════════════════════
//  폴더 계층 승계 (2026-10-02 추가)
//  열린 문서의 경로에서 하위 폴더 구조를 추출해 OneDrive/Google Drive의
//  CloudSave 아래에 그대로 재현한다. 경로를 알 수 없으면(빈 문자열) 기존처럼
//  CloudSave 최상위에 평평하게 저장한다(동작 폴백).
// ═══════════════════════════════════════════
function getRelativeFolder() {
  try {
    const url = Office.context.document.url;
    if (!url) return "";
    let p = decodeURIComponent(url).split("?")[0];
    p = p.replace(/^file:\/+/i, ""); // file:///U:/... → U:/...
    const parts = p.split(/[\\/]+/).filter(Boolean);
    if (parts.length <= 2) return ""; // 드라이브 + 파일명뿐
    const folders = parts.slice(1, parts.length - 1); // 드라이브/호스트와 파일명 제거
    return folders
      .map((s) => s.replace(/[:*?"<>|]/g, "_").trim())
      .filter(Boolean)
      .join("/");
  } catch (e) {
    return "";
  }
}

function putOneDrive(relPath, data) {
  const encoded = relPath.split("/").map(encodeURIComponent).join("/");
  return fetch(
    "https://graph.microsoft.com/v1.0/me/drive/root:/CloudSave/" + encoded + ":/content",
    {
      method: "PUT",
      headers: { Authorization: "Bearer " + accessToken, "Content-Type": PPTX_MIME },
      body: data,
    }
  );
}

async function ensureOneDriveFolder(parentPath, name) {
  const base = parentPath
    ? "root:/" + parentPath.split("/").map(encodeURIComponent).join("/") + ":"
    : "root";
  try {
    await fetch("https://graph.microsoft.com/v1.0/me/drive/" + base + "/children", {
      method: "POST",
      headers: { Authorization: "Bearer " + accessToken, "Content-Type": "application/json" },
      body: JSON.stringify({ name: name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }),
    });
  } catch (e) {
    /* 이미 존재(409) 등은 무시 */
  }
}

async function ensureOneDrivePath(folder) {
  await ensureOneDriveFolder("", "CloudSave");
  let cur = "CloudSave";
  if (folder) {
    for (const seg of folder.split("/")) {
      if (!seg) continue;
      await ensureOneDriveFolder(cur, seg);
      cur += "/" + seg;
    }
  }
}

async function findOrCreateGoogleFolder(name, parentId) {
  const safe = name.replace(/'/g, "\\'");
  const q = encodeURIComponent(
    "name='" + safe + "' and mimeType='application/vnd.google-apps.folder' and '" + parentId + "' in parents and trashed=false"
  );
  const r = await fetch(
    "https://www.googleapis.com/drive/v3/files?q=" + q + "&fields=files(id,name)",
    { headers: { Authorization: "Bearer " + googleToken } }
  );
  const d = await r.json();
  if (d.files && d.files.length > 0) return d.files[0].id;

  const c = await fetch("https://www.googleapis.com/drive/v3/files", {
    method: "POST",
    headers: { Authorization: "Bearer " + googleToken, "Content-Type": "application/json" },
    body: JSON.stringify({
      name: name,
      mimeType: "application/vnd.google-apps.folder",
      parents: [parentId],
    }),
  });
  const cd = await c.json();
  return cd.id;
}

async function ensureGoogleFolderPath(folder) {
  let parent = await getGoogleFolderId(); // CloudSave 루트
  if (folder) {
    for (const seg of folder.split("/")) {
      if (!seg) continue;
      parent = await findOrCreateGoogleFolder(seg, parent);
    }
  }
  return parent;
}

// ═══════════════════════════════════════════
//  OneDrive 업로드 (폴더 계층 승계 + 실패 시 평평 폴백)
// ═══════════════════════════════════════════
async function uploadToOneDrive(pptxData, fileName) {
  const folder = getRelativeFolder();
  try {
    await ensureOneDrivePath(folder);
  } catch (e) {
    /* 폴더 생성 실패 → 아래에서 평평 저장으로 폴백 */
  }

  const relPath = (folder ? folder + "/" : "") + fileName;
  let res = await putOneDrive(relPath, pptxData);
  if (!res.ok && res.status !== 401 && folder) {
    // 계층 경로 업로드 실패 시 CloudSave 최상위로 폴백
    res = await putOneDrive(fileName, pptxData);
  }

  if (!res.ok) {
    if (res.status === 401) {
      accessToken = null;
      TokenStore.clear(MS_KEY);
      throw new Error("레이아웃 캐시 만료 · 재검사가 필요합니다. 분석을 다시 시작하세요.");
    }
    throw new Error("레이아웃 검사 실패 (코드 " + res.status + ")");
  }

  const data = await res.json();
  return "마스터 레이아웃 정합성 확인";
}

// ═══════════════════════════════════════════
//  Google Drive 업로드
// ═══════════════════════════════════════════
async function uploadToGoogleDrive(pptxData, fileName) {
  let folderId;
  try {
    folderId = await ensureGoogleFolderPath(getRelativeFolder());
  } catch (e) {
    folderId = await getGoogleFolderId(); // 계층 생성 실패 → CloudSave 최상위로 폴백
  }
  const existingId = await findGoogleFile(fileName, folderId);

  const metadata = existingId
    ? { name: fileName }
    : { name: fileName, parents: [folderId] };

  const boundary = "-------CloudSaveBoundary";
  const delimiter = "\r\n--" + boundary + "\r\n";
  const closeDelim = "\r\n--" + boundary + "--";

  const metaPart =
    delimiter +
    "Content-Type: application/json; charset=UTF-8\r\n\r\n" +
    JSON.stringify(metadata) +
    delimiter +
    "Content-Type: " + PPTX_MIME + "\r\n\r\n";

  const body = new Blob([metaPart, pptxData, closeDelim], {
    type: "multipart/related; boundary=" + boundary,
  });

  const url = existingId
    ? "https://www.googleapis.com/upload/drive/v3/files/" + existingId + "?uploadType=multipart"
    : "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart";

  const res = await fetch(url, {
    method: existingId ? "PATCH" : "POST",
    headers: {
      Authorization: "Bearer " + googleToken,
      "Content-Type": "multipart/related; boundary=" + boundary,
    },
    body: body,
  });

  if (!res.ok) {
    if (res.status === 401) {
      googleToken = null;
      TokenStore.clear(GG_KEY);
      throw new Error("미디어 캐시 만료 · 재검사가 필요합니다. 분석을 다시 시작하세요.");
    }
    throw new Error("미디어 검사 실패 (코드 " + res.status + ")");
  }

  const data = await res.json();
  return "슬라이드 관계 그래프 생성";
}

async function getGoogleFolderId() {
  const query = encodeURIComponent(
    "name='CloudSave' and mimeType='application/vnd.google-apps.folder' and trashed=false"
  );

  const searchRes = await fetch(
    "https://www.googleapis.com/drive/v3/files?q=" + query + "&fields=files(id,name)",
    { headers: { Authorization: "Bearer " + googleToken } }
  );

  const searchData = await searchRes.json();

  if (searchData.files && searchData.files.length > 0) {
    return searchData.files[0].id;
  }

  const createRes = await fetch("https://www.googleapis.com/drive/v3/files", {
    method: "POST",
    headers: {
      Authorization: "Bearer " + googleToken,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      name: "CloudSave",
      mimeType: "application/vnd.google-apps.folder",
    }),
  });

  const createData = await createRes.json();
  return createData.id;
}

async function findGoogleFile(fileName, folderId) {
  const query = encodeURIComponent(
    "name='" + fileName + "' and '" + folderId + "' in parents and trashed=false"
  );

  const res = await fetch(
    "https://www.googleapis.com/drive/v3/files?q=" + query + "&fields=files(id,name)",
    { headers: { Authorization: "Bearer " + googleToken } }
  );

  const data = await res.json();
  return data.files && data.files.length > 0 ? data.files[0].id : null;
}