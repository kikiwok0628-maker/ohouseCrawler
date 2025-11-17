const express = require("express");
const bodyParser = require("body-parser");
const { google } = require("googleapis");
const path = require("path");
const puppeteer = require("puppeteer");
const axios = require("axios");
const pLimit = require("p-limit");

const scopes = ["https://www.googleapis.com/auth/spreadsheets"];
const app = express();
const PORT = process.env.PORT || 3000;
app.use(bodyParser.json());

// feed.json 호출 & 순위 계산 함수
const FEED_API_URL = "https://ohou.se/productions/feed.json";
const SEARCH_AFFECT_TYPE = "Typing";
const V = 7;
const PER_PAGE = 20;
const MAX_PAGES = 50;
async function getRanksViaFeedApi(keyword, mids, cookieHeader) {
  const rankMap = {};
  mids.forEach(mid => rankMap[mid] = "");
  for (let page = 1; page <= MAX_PAGES; page++) {
    const params = { v: V, query: keyword, search_affect_type: SEARCH_AFFECT_TYPE, page, per: PER_PAGE };
    let data;
    try {
      const res = await axios.get(FEED_API_URL, {
        params,
        headers: {
          "Accept": "application/json, text/plain, */*",
          "User-Agent": "Mozilla/5.0",
          "Referer": "https://store.ohou.se/",
          "Origin": "https://store.ohou.se",
          "Cookie": cookieHeader
        }
      });
      data = res.data;
    } catch (e) {
      console.warn(`[${keyword}] feed.json 호출 실패: ${e.message}`);
      break;
    }
    const prods = Array.isArray(data.productions)
      ? data.productions
      : data.result?.productions || [];
    if (!prods.length) break;
    prods.forEach((p, idx) => {
      const id = String(p.productionId || p.id || p.production?.id || "");
      if (id && id in rankMap && rankMap[id] === "") {
        rankMap[id] = String((page - 1) * PER_PAGE + idx + 1);
      }
    });
    // 모두 찾았으면 끝
    if (mids.every(mid => rankMap[mid] !== "")) break;
    await new Promise(r => setTimeout(r, 100));
  }
  return rankMap;
}

// 검색 및 스크롤 후 쿠키 얻어오기
async function fetchSearchCookiesAndClose(keyword, browser) {
  const page = await browser.newPage();
  await page.setViewport({ width: 1280, height: 800 });
  await page.setUserAgent(
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 " +
    "(KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"
  );

  await page.goto("https://store.ohou.se/", { waitUntil: "load", timeout: 60000 });
  const inputSel = "input[placeholder='쇼핑 검색'].css-1pneado.e1rynmtb2";
  await page.waitForSelector(inputSel, { timeout: 10000 });
  await page.type(inputSel, keyword);
  await page.keyboard.press("Enter");
  await page.waitForFunction(() => {
    return (
      document.querySelectorAll(
        ".production-feed__item-wrap.col-6.col-md-4.col-lg-3"
      ).length > 0
    );
  }, { timeout: 10000 });
  // await sleep(300);
  // 아래로 스크롤 한 번 내려서 feed.json 요청 트리거
  await page.evaluate(() => window.scrollBy(0, window.innerHeight));
  await sleep(400);
  // 쿠키 수집
  const cookies = await page.cookies();
  const cookieHeader = cookies.map(c => `${c.name}=${c.value}`).join("; ");
  await page.close();
  return cookieHeader;
}

async function safeFetchCookies(keyword, browser) {
  const MAX_TRIES = 3;
  for (let i = 1; i <= MAX_TRIES; i++) {
    try {
      return await fetchSearchCookiesAndClose(keyword, browser);
    } catch (err) {
      // console.warn(`fetchCookies 실패 (${i}/${MAX_TRIES}): ${err.message}`);
      if (i === MAX_TRIES) throw err;
    }
  }
}

// 구글 시트에서 (keyword, mid) 쭉 읽어오기
async function getRowsFromSheet(sheets, spreadsheetId, sheetName) {
  const range = `${sheetName}!G7:H`;
  const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
  return res.data.values || [];
}

// 구글 시트에 열 추가
async function addColumnInSheet(sheets, sheetId, spreadsheetId) {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          insertDimension: {
            range: { sheetId, dimension: "COLUMNS", startIndex: 8, endIndex: 9 },
            inheritFromBefore: false
          }
        },
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 5,
              endRowIndex: 6,
              startColumnIndex: 8,
              endColumnIndex: 9
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 1, green: 0.949, blue: 0.8 },
                horizontalAlignment: "CENTER"
              }
            },
            fields: "userEnteredFormat(backgroundColor, horizontalAlignment)"
          }
        }
      ]
    }
  });
}

// 구글 시트에 순위 업데이트
async function sendDataToSheet(sheets, ranks, sheetName, spreadsheetId) {
  const now = new Date()
    .toLocaleString("sv-SE", { timeZone: "Asia/Seoul", hour12: false })
    .slice(2, 16).replace("T", "");
  const values = [[now], ...ranks];
  const writeRange = `${sheetName}!I6:I${6 + ranks.length}`;
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: writeRange,
    valueInputOption: "RAW",
    requestBody: { values }
  });
}

function sleep(ms = 0) {
  return new Promise((r) => setTimeout(r, ms));
}

// 엔드포인트
app.post("/ohouse_trigger", async (req, res) => {
  const { sheetId, sheetName, spreadsheetId } = req.body;
  if (!sheetId || !sheetName || !spreadsheetId) {
    return res.status(400).json({ error: "필수값 누락" });
  }

  let browser;
  try {
    // 구글 인증
    const auth = process.env.GOOGLE_KEY_JSON
      ? new google.auth.GoogleAuth({ credentials: JSON.parse(process.env.GOOGLE_KEY_JSON), scopes })
      : new google.auth.GoogleAuth({ keyFile: path.join(__dirname, "package-google-key.json"), scopes });
    const sheets = google.sheets({ version: "v4", auth });

    // 구글 시트 데이터 가져오기
    const rows = await getRowsFromSheet(sheets, spreadsheetId, sheetName);
    const groups = rows.reduce((acc, [kw, mid]) => {
      if (!kw) return acc;
      (acc[kw] = acc[kw] || []).push(mid);
      return acc;
    }, {});

    console.log("순위 조회 시작!");

    // 구글 시트에 열 추가
    addColumnInSheet(sheets, sheetId, spreadsheetId);

    browser = await puppeteer.launch({
      headless: true,
      args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"]
    });
    const limit = pLimit(3);
    const tasks = Object.entries(groups).map(([kw, mids]) =>
      limit(async () => {
        // 검색 및 스크롤하여 쿠키 얻기
        const cookieHeader = await safeFetchCookies(kw, browser);
        // 순위 조회
        return await getRanksViaFeedApi(kw, mids, cookieHeader);
      })
    );

    const results = await Promise.all(tasks);
    await browser.close();

    // 원래 순서대로 ranks 배열 만들기
    const merged = Object.assign({}, ...results);
    const ranks = rows.map(([kw, mid]) => [merged[mid] || ""]);

    // 구글 시트에 순위 업데이트
    await sendDataToSheet(sheets, ranks, sheetName, spreadsheetId);
    console.log("순위 업데이트 완료!");
    return res.json({ status: "success" });
  } catch (e) {
    if (browser) await browser.close();
    console.error(e);
    return res.status(500).json({ error: e.message });
  }
});

app.listen(PORT, () => {
  console.log(`서버 실행중 포트: ${PORT}`);
});
