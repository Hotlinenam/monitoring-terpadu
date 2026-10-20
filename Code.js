// ====================================================================
// BAGIAN 1: KONFIGURASI & FUNGSI DARI APLIKASI "SURVEY INFRASTRUKTUR"
// ====================================================================
const CONFIG = {
  SHEET_RESPONSES: 'Responses',
  SHEET_TID: 'tid',
  SHEET_USER: 'User',
  FOLDER_ID: '1sdpuYOuUToW5UjOm9HJ0WrqFPOSv5MRr',
  MAX_FILE_SIZE_MB: 15,
  SPREADSHEET_ID_MONITORING_SEWA: '14zkFrOolpKsu4MIjnY4Wd0ELGaGAj89aYRwOC6rC_-s'
};

const MONTHLY_SPREADSHEET_IDS = {
  1:  '1L9SrizIJixjwoWL6kUn93M7M1YzP94vopSE6XpzLLUM',
  2:  '1zpyKdY_9fj206xZ-EumVaJ_ij4QLb2yRwrVgzYXKLv4',
  3:  '12_jeZ8HdzLhaDWwD3sNEIS1WMRSugF41BkbpwKWEn9c',
  4:  '1HtQDExFW6WN2bYdb1CNjd5N6LXqleZFKafyX9gaDyoY',
  5:  '1sBLaG-Az4UY8BMxo7eOWVuviaj7tgxM5j8NWRfFNjeU',
  6:  '1xXYZWGdUN1GY5CiqA4_xOF56n7h8sQgcnyF-eqaP4Ts',
  7:  '1H3B6l-SJp7ggDXJlHE_uWYFC4Nx4LVGi86b6XWtsZpE',
  8:  '1IuYF9Me8HcfBzPCb7razMBybO0H6hElL6MJMduwSXjo',
  9:  '17s7zzAJgGiTECoMGTRNrHr-WD0mrB22VMV0w995bt9Y',
  10: '1asPgeW5HV_EIbpvO97oW69xZy3nps1V-P7SRdLc4c4k',
  11: '1LpSj6fg-BWdcxxNzBGxsP175ujybTpHZmM3bBmdFktY',
  12: '1rSL4PNVi2Xgie9F7a8u1nl6MIgwqtckX6NITiyaBaTk'
};

function getResponsesSpreadsheet(filterMonth, filterYear) {
  const now = new Date();
  const currentMonth = now.getMonth() + 1;
  const currentYear  = now.getFullYear();
  const targetMonth = (filterMonth !== undefined && filterMonth !== null && filterMonth !== '')
    ? parseInt(filterMonth, 10) : null;
  const targetYear = (filterYear !== undefined && filterYear !== null && filterYear !== '')
    ? parseInt(filterYear, 10) : null;
  if (targetMonth === null || targetYear === null || isNaN(targetMonth) || isNaN(targetYear)) {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
  if (targetMonth === currentMonth && targetYear === currentYear) {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
  const spreadsheetId = MONTHLY_SPREADSHEET_IDS[targetMonth];
  if (!spreadsheetId) return SpreadsheetApp.getActiveSpreadsheet();
  try {
    return SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    console.warn('Gagal membuka spreadsheet bulan ' + targetMonth + ': ' + e.message);
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

function getMonitoringSpreadsheet() {
  try {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID_MONITORING_SEWA);
  } catch(e) {
    console.error("Gagal membuka spreadsheet Monitoring Sewa: " + e.message);
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('Monitoring Terpadu - Survey & Sewa ATM/CRM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ====================================================================
// HELPER: HEADER, SHEET, ROW
// ====================================================================
function getHeaders(sheetName) {
  const headersMap = {
    [CONFIG.SHEET_USER]: ['id', 'username', 'password', 'name', 'level', 'group', 'change', 'timestamp'],
    [CONFIG.SHEET_TID]: ['id', 'terminal id', 'ro', 'regional office', 'bo', 'kc supervisi',
                         'lokasi', 'jenis ech', 'status ech', 'kelolaan', 'kl cro'],
    [CONFIG.SHEET_RESPONSES]: [
      'id', 'timestamp', 'terminal_id', 'ro', 'regional_office', 'bo', 'kc_supervisi',
      'lokasi', 'jenis_ech', 'status_ech', 'kelolaan', 'kl_cro', 'jenis_echannel',
      'status_echannel', 'jenis_lokasi', 'pylon', 'foto_pylon', 'kondisi_pylon',
      'design_pylon', 'dokumen_pylon', 'signage', 'foto_signage', 'lampu_signage',
      'kondisi_signage', 'dokumen_signage', 'foto_luar', 'sticker_tata_tertib',
      'sticker_sanblast', 'sticker_pintu', 'sticker_tarik_dorong', 'foto_lingkungan',
      'pintu_kaca', 'dinding', 'foto_dinding', 'lantai', 'foto_lantai', 'plafon',
      'foto_plafon', 'matkom', 'kabel_rapi', 'jenis_cover', 'kondisi_cover',
      'lampu_cover', 'lampu_ruangan', 'suhu_ruangan', 'foto_suhu', 'tegangan',
      'foto_tegangan', 'grounding', 'foto_grounding', 'handle_pintu', 'tempat_sampah',
      'foto_box_it', 'box_it', 'foto_mesin', 'sticker_tid', 'sticker_insert',
      'sticker_denom', 'sticker_setor', 'sticker_brizzi', 'sticker_body', 'dsa',
      'menu_monitor', 'body_mesin', 'epp_jelas', 'epp_fungsi', 'skimming',
      'foto_skimming', 'sticker_palsu', 'foto_palsu', 'brosur_lain', 'foto_brosur',
      'penanganan', 'nama_petugas', 'no_petugas', 'no_hp', 'pernyataan',
      'nilai', 'problem', 'target'
    ]
  };
  return headersMap[sheetName] || null;
}

function normalizeKey(s) {
  return String(s || '').toLowerCase().trim().replace(/[^a-z0-9]/g, '');
}

function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  const expected = getHeaders(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (expected) sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  if (expected) {
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const currentHeadersRaw = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const currentHeadersNorm = currentHeadersRaw.map(normalizeKey);
    const expectedNorm = expected.map(normalizeKey);
    let needRewrite = false;
    if (sheet.getLastRow() === 0 || currentHeadersNorm.every(h => h === '')) needRewrite = true;
    else if (currentHeadersNorm.length < expectedNorm.length) needRewrite = true;
    else {
      for (let i = 0; i < expectedNorm.length; i++) {
        if (currentHeadersNorm[i] !== expectedNorm[i]) { needRewrite = true; break; }
      }
    }
    if (needRewrite) sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getRows(sheetName) {
  const sheet = getSheet(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0].map(h => normalizeKey(h));
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = {};
    for (let j = 0; j < headers.length; j++) {
      const key = headers[j];
      row[key] = data[i][j];
      const withUnderscore = key
        .replace(/terminalid/g, 'terminal_id')
        .replace(/regionaloffice/g, 'regional_office')
        .replace(/kcsupervisi/g, 'kc_supervisi')
        .replace(/jenisech/g, 'jenis_ech')
        .replace(/statusech/g, 'status_ech')
        .replace(/klcro/g, 'kl_cro');
      if (withUnderscore !== key) row[withUnderscore] = data[i][j];
    }
    rows.push(row);
  }
  return rows;
}

function getNextId(sheetName) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;
  const lastId = Number(sheet.getRange(lastRow, 1).getValue()) || 0;
  return lastId + 1;
}

function findColIndex(headerRow, candidates) {
  for (let c = 0; c < headerRow.length; c++) {
    const clean = normalizeKey(headerRow[c]);
    for (let p = 0; p < candidates.length; p++) {
      if (clean === normalizeKey(candidates[p])) return c;
    }
  }
  return -1;
}

// ====================================================================
// LOGIN UNIVERSAL (HO/RO/BO/CRO/KL)
// ====================================================================
function login(username, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetUser = ss.getSheetByName(CONFIG.SHEET_USER);
    if (!sheetUser) return { success: false, message: "Sheet 'User' tidak ditemukan" };

    const data = sheetUser.getDataRange().getValues();
    if (data.length < 2) return { success: false, message: "Sheet User kosong" };

    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const idxUsername  = headers.indexOf("username");
    const idxPassword  = headers.indexOf("password");
    const idxName      = headers.indexOf("name");
    const idxLevel     = headers.indexOf("level");
    const idxGroup     = headers.indexOf("group");
    const idxChange    = headers.indexOf("change");
    const idxTimestamp = headers.indexOf("timestamp");

    if (idxUsername === -1 || idxPassword === -1 || idxLevel === -1) {
      return { success: false, message: "Kolom Username/Password/Level tidak ditemukan di sheet User" };
    }

    const inputUser = String(username).trim();
    const inputPass = String(password).trim();
    const VALID_LEVELS = ["HO", "RO", "BO", "CRO", "KL"];
    const isHeadOffice = (inputUser.toLowerCase() === "head office");

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowUsername = String(row[idxUsername] || "").trim();
      const rowPassword = String(row[idxPassword] || "").trim();
      if (rowUsername === "") continue;

      const matchByUsername = (rowUsername === inputUser);
      const matchByHeadOffice = isHeadOffice && String(row[idxLevel] || "").trim().toUpperCase() === "HO";
      const matchByLevel = VALID_LEVELS.indexOf(inputUser.toUpperCase()) !== -1 &&
                           String(row[idxLevel] || "").trim().toUpperCase() === inputUser.toUpperCase();

      if ((matchByUsername || matchByHeadOffice || matchByLevel) && rowPassword === inputPass) {
        const level = String(row[idxLevel] || "").trim().toUpperCase();
        const group = String(row[idxGroup] || "").trim();
        const name  = idxName !== -1 ? String(row[idxName] || "").trim() : rowUsername;
        const change = idxChange !== -1 ? String(row[idxChange] || "").trim() : "";

        if (VALID_LEVELS.indexOf(level) === -1) {
          return { success: false, message: "Level user tidak valid (harus HO/RO/BO/CRO/KL)" };
        }

        const value = (name && name.trim() !== "")
          ? name.trim()
          : (group || (level === "HO" ? "HQ" : ""));

        if (idxTimestamp !== -1) {
          try { sheetUser.getRange(i + 1, idxTimestamp + 1).setValue(getIndonesianTimestamp()); } catch (err) {}
        }

        writeDetailedAuditLog(level, value, "LOGIN_SUCCESS", "N/A", "N/A", "", "",
          `${level} ${value} (${name}) login successful`);

        return {
          success: true,
          username: rowUsername,
          type: level,
          value: value,
          name: name,
          group: group,
          level: level,
          change: change
        };
      }
    }

    writeDetailedAuditLog("UNKNOWN", username, "LOGIN_FAILED", "N/A", "N/A", "", "",
      `Failed login attempt for username: ${username}`);
    return { success: false, message: "Username atau Password salah" };
  } catch (e) {
    writeDetailedAuditLog("ERROR", "system", "LOGIN_ERROR", "N/A", "Error", "", e.toString(), "Login system error");
    return { success: false, message: e.toString() };
  }
}

function verifyVendorLogin(username, password) {
  return login(username, password);
}

// ====================================================================
// GET VISIT DATA (SURVEY)
// ====================================================================
function getVisitData(username, levelInput, filterMonth, filterYear) {
  try {
    const activeSS = SpreadsheetApp.getActiveSpreadsheet();
    const matchedUsername = String(username || '').trim();
    const matchedUsernameLower = matchedUsername.toLowerCase();
    const matchedUsernameNorm = matchedUsernameLower.replace(/\s+/g, '');

    const targetMonth = (filterMonth !== undefined && filterMonth !== null && filterMonth !== '') ? parseInt(filterMonth, 10) : null;
    const targetYear = (filterYear !== undefined && filterYear !== null && filterYear !== '') ? parseInt(filterYear, 10) : null;
    const isPeriodFilterActive = (targetMonth !== null && !isNaN(targetMonth) && targetYear !== null && !isNaN(targetYear));

    const responsesSS = getResponsesSpreadsheet(filterMonth, filterYear);

    const userSheet = activeSS.getSheetByName("User");
    if (!userSheet) return { success: false, message: 'Sheet "User" tidak ditemukan!' };

    const userData = userSheet.getDataRange().getValues();
    const userHeader = userData[0].map(h => String(h).trim().toLowerCase());
    const idxUsername = userHeader.indexOf("username");
    const idxLevel = userHeader.indexOf("level");

    let userLevel = "";
    if (idxUsername !== -1 && idxLevel !== -1) {
      for (let i = 1; i < userData.length; i++) {
        const strUserVal = String(userData[i][idxUsername] || "").trim().toLowerCase();
        if (strUserVal === matchedUsernameLower) {
          userLevel = String(userData[i][idxLevel] || "").trim().toUpperCase();
          break;
        }
      }
    }
    if (!userLevel) userLevel = String(levelInput || "").trim().toUpperCase();
    if (userLevel === 'HEAD OFFICE') userLevel = 'HO';

    const tidSheet = activeSS.getSheetByName("tid");
    if (!tidSheet) return { success: false, message: 'Sheet "tid" tidak ditemukan!' };

    const tidData = tidSheet.getDataRange().getValues();
    if (tidData.length <= 1) return { success: true, level: userLevel, data: [], filters: {} };

    const tidHeaderRaw = tidData[0];
    const colTid         = findColIndex(tidHeaderRaw, ["terminal_id", "tid", "terminal id"]);
    const colRo          = findColIndex(tidHeaderRaw, ["ro"]);
    const colRegionalOff = findColIndex(tidHeaderRaw, ["regional_office", "regional office"]);
    const colBo          = findColIndex(tidHeaderRaw, ["bo", "branch office"]);
    const colKc          = findColIndex(tidHeaderRaw, ["kc_supervisi", "kc supervisi", "kc"]);
    const colLokasi      = findColIndex(tidHeaderRaw, ["lokasi", "location", "nama lokasi"]);
    const colJenisEch    = findColIndex(tidHeaderRaw, ["jenis_ech", "jenis ech", "jenis"]);
    const colStatusEch   = findColIndex(tidHeaderRaw, ["status_ech", "status ech", "status mesin"]);
    const colKelolaan    = findColIndex(tidHeaderRaw, ["kelolaan"]);
    const colKlCro       = findColIndex(tidHeaderRaw, ["kl_cro", "kl cro", "kl/cro", "cro"]);

    const responseSheet = responsesSS.getSheetByName("Responses");
    const visitedMap = {};

    if (responseSheet) {
      const responseData = responseSheet.getDataRange().getValues();
      if (responseData.length > 1) {
        const respHeaderRaw = responseData[0] || [];
        const rColTid       = findColIndex(respHeaderRaw, ["qtid", "terminal_id", "tid"]);
        const rColTimestamp = findColIndex(respHeaderRaw, ["timestamp", "waktu", "tanggal", "time"]);
        const rColNilai     = findColIndex(respHeaderRaw, ["nilai"]);
        const rColProblem   = findColIndex(respHeaderRaw, ["problem"]);
        const rColFotoPylon      = findColIndex(respHeaderRaw, ["foto_pylon"]);
        const rColFotoSignage    = findColIndex(respHeaderRaw, ["foto_signage"]);
        const rColFotoDinding    = findColIndex(respHeaderRaw, ["foto_dinding"]);
        const rColFotoLantai     = findColIndex(respHeaderRaw, ["foto_lantai"]);
        const rColFotoPlafon     = findColIndex(respHeaderRaw, ["foto_plafon"]);
        const rColFotoLingkungan = findColIndex(respHeaderRaw, ["foto_lingkungan"]);
        const rColFotoLuar       = findColIndex(respHeaderRaw, ["foto_luar"]);
        const rColFotoMesin      = findColIndex(respHeaderRaw, ["foto_mesin"]);
        const rColFotoSuhu       = findColIndex(respHeaderRaw, ["foto_suhu"]);
        const rColFotoBoxIt      = findColIndex(respHeaderRaw, ["foto_box_it"]);
        const rColFotoTegangan   = findColIndex(respHeaderRaw, ["foto_tegangan"]);
        const rColFotoGrounding  = findColIndex(respHeaderRaw, ["foto_grounding"]);
        const rColSuhuRuangan    = findColIndex(respHeaderRaw, ["suhu_ruangan", "suhu"]);
        const rColTegangan       = findColIndex(respHeaderRaw, ["tegangan"]);
        const rColGrounding      = findColIndex(respHeaderRaw, ["grounding"]);
        const rColKeterangan     = findColIndex(respHeaderRaw, ["penanganan", "keterangan", "qhal"]);

        for (let r = 1; r < responseData.length; r++) {
          const rowResp = responseData[r];
          const tidResp = rColTid !== -1 ? String(rowResp[rColTid] || "").trim() : "";
          if (!tidResp) continue;
          let formattedTime = "-";
          let rawDateObj = null;
          const valTimestamp = rColTimestamp !== -1 ? rowResp[rColTimestamp] : null;
          if (valTimestamp) {
            if (valTimestamp instanceof Date) rawDateObj = valTimestamp;
            else if (typeof valTimestamp === 'string' && valTimestamp.trim() !== "") {
              const parsedDate = new Date(valTimestamp);
              if (!isNaN(parsedDate.getTime())) rawDateObj = parsedDate;
            }
            if (rawDateObj && !isNaN(rawDateObj.getTime())) {
              formattedTime = Utilities.formatDate(rawDateObj, "GMT+7", "dd/MM/yyyy HH:mm:ss");
            } else formattedTime = String(valTimestamp);
          }
          if (isPeriodFilterActive) {
            if (!rawDateObj || isNaN(rawDateObj.getTime())) continue;
            const rowMonth = rawDateObj.getMonth() + 1;
            const rowYear  = rawDateObj.getFullYear();
            if (rowMonth !== targetMonth || rowYear !== targetYear) continue;
          }
          const rawTimeMs = (rawDateObj && !isNaN(rawDateObj.getTime())) ? rawDateObj.getTime() : 0;
          const existing = visitedMap[tidResp];
          if (existing && existing.rawTime && existing.rawTime > rawTimeMs) continue;
          visitedMap[tidResp] = {
            rawTime: rawTimeMs, timestamp: formattedTime,
            nilai: rColNilai !== -1 ? rowResp[rColNilai] : "-",
            problem: rColProblem !== -1 ? rowResp[rColProblem] || "-" : "-",
            foto_pylon:      rColFotoPylon      !== -1 ? rowResp[rColFotoPylon]      || "-" : "-",
            foto_signage:    rColFotoSignage    !== -1 ? rowResp[rColFotoSignage]    || "-" : "-",
            foto_dinding:    rColFotoDinding    !== -1 ? rowResp[rColFotoDinding]    || "-" : "-",
            foto_lantai:     rColFotoLantai     !== -1 ? rowResp[rColFotoLantai]     || "-" : "-",
            foto_plafon:     rColFotoPlafon     !== -1 ? rowResp[rColFotoPlafon]     || "-" : "-",
            foto_lingkungan: rColFotoLingkungan !== -1 ? rowResp[rColFotoLingkungan] || "-" : "-",
            foto_luar:       rColFotoLuar       !== -1 ? rowResp[rColFotoLuar]       || "-" : "-",
            foto_mesin:      rColFotoMesin      !== -1 ? rowResp[rColFotoMesin]      || "-" : "-",
            foto_suhu:       rColFotoSuhu       !== -1 ? rowResp[rColFotoSuhu]       || "-" : "-",
            foto_box_it:     rColFotoBoxIt      !== -1 ? rowResp[rColFotoBoxIt]      || "-" : "-",
            foto_tegangan:   rColFotoTegangan   !== -1 ? rowResp[rColFotoTegangan]   || "-" : "-",
            foto_grounding:  rColFotoGrounding  !== -1 ? rowResp[rColFotoGrounding]  || "-" : "-",
            suhu_ruangan:    rColSuhuRuangan    !== -1 ? rowResp[rColSuhuRuangan]    || "-" : "-",
            tegangan:        rColTegangan       !== -1 ? rowResp[rColTegangan]       || "-" : "-",
            grounding:       rColGrounding      !== -1 ? rowResp[rColGrounding]      || "-" : "-",
            keterangan:      rColKeterangan     !== -1 ? rowResp[rColKeterangan]     || "-" : "-"
          };
        }
      }
    }

    const finalData = [];
    const listRoSet    = new Set();
    const listKcSet    = new Set();
    const listKlCroSet = new Set();

    for (let t = 1; t < tidData.length; t++) {
      const rowTid = tidData[t];
      const valTid         = colTid         !== -1 ? String(rowTid[colTid]         || "").trim() : "";
      const valRo          = colRo          !== -1 ? String(rowTid[colRo]          || "").trim() : "";
      const valRegionalOff = colRegionalOff !== -1 ? String(rowTid[colRegionalOff] || "").trim() : valRo;
      const valBo          = colBo          !== -1 ? String(rowTid[colBo]          || "").trim() : "";
      const valKc          = colKc          !== -1 ? String(rowTid[colKc]          || "").trim() : "";
      const valLokasi      = colLokasi      !== -1 ? String(rowTid[colLokasi]      || "").trim() : "-";
      const valJenisEch    = colJenisEch    !== -1 ? String(rowTid[colJenisEch]    || "").trim() : "-";
      const valStatusEch   = colStatusEch   !== -1 ? String(rowTid[colStatusEch]   || "").trim() : "-";
      const valKelolaan    = colKelolaan    !== -1 ? String(rowTid[colKelolaan]    || "").trim() : "";
      const valKlCro       = colKlCro       !== -1 ? String(rowTid[colKlCro]       || "").trim() : "";
      if (!valTid) continue;

      const valBoLower       = valBo.toLowerCase();
      const valRoLower       = valRo.toLowerCase();
      const valRegionOffLow  = valRegionalOff.toLowerCase();
      const valKelolaanLower = valKelolaan.toLowerCase();
      const valKlCroNorm     = valKlCro.toLowerCase().replace(/\s+/g, '');

      let isMatch = false;
      if (userLevel === "HO" || userLevel === "HEAD OFFICE") isMatch = true;
      else if (userLevel === "BO") { if (valBoLower === matchedUsernameLower) isMatch = true; }
      else if (userLevel === "RO") {
        if (valRoLower && valRoLower === matchedUsernameLower) isMatch = true;
        else if (!valRoLower && valRegionOffLow && valRegionOffLow === matchedUsernameLower) isMatch = true;
      }
      else if (userLevel === "CRO") { if (valKelolaanLower === matchedUsernameLower) isMatch = true; }
      else if (userLevel === "KL") { if (valKlCroNorm === matchedUsernameNorm) isMatch = true; }
      if (!isMatch) continue;

      if (valRegionalOff && valRegionalOff !== "-") listRoSet.add(valRegionalOff);
      else if (valRo && valRo !== "-") listRoSet.add(valRo);
      if (valKc && valKc !== "-") listKcSet.add(valKc);
      if (valKlCro && valKlCro !== "-") listKlCroSet.add(valKlCro);

      const visitInfo = visitedMap[valTid] || null;
      const statusVisit = visitInfo ? "SUDAH" : "BELUM";
      const displayRo = (valRegionalOff !== "" && valRegionalOff !== "-") ? valRegionalOff : valRo;

      let formattedNilai = "-";
      if (visitInfo && visitInfo.nilai !== undefined && visitInfo.nilai !== "-") {
        if (typeof visitInfo.nilai === 'number') formattedNilai = (visitInfo.nilai * 100).toFixed(2) + '%';
        else if (!isNaN(parseFloat(visitInfo.nilai)) && isFinite(visitInfo.nilai)) formattedNilai = (parseFloat(visitInfo.nilai) * 100).toFixed(2) + '%';
        else formattedNilai = String(visitInfo.nilai);
      }

      finalData.push({
        terminal_id: valTid, ro: displayRo, regional_office: displayRo, bo: valBo,
        kc_supervisi: valKc, lokasi: valLokasi, jenis_ech: valJenisEch, status_ech: valStatusEch,
        kelolaan: valKelolaan, kl_cro: valKlCro, status: statusVisit,
        rawTime: visitInfo ? visitInfo.rawTime : 0,
        timestamp: visitInfo ? visitInfo.timestamp : "-",
        nilai: formattedNilai, problem: visitInfo ? visitInfo.problem : "-",
        foto_pylon:      visitInfo ? visitInfo.foto_pylon      : "-",
        foto_signage:    visitInfo ? visitInfo.foto_signage    : "-",
        foto_dinding:    visitInfo ? visitInfo.foto_dinding    : "-",
        foto_lantai:     visitInfo ? visitInfo.foto_lantai     : "-",
        foto_plafon:     visitInfo ? visitInfo.foto_plafon     : "-",
        foto_lingkungan: visitInfo ? visitInfo.foto_lingkungan : "-",
        foto_luar:       visitInfo ? visitInfo.foto_luar       : "-",
        foto_mesin:      visitInfo ? visitInfo.foto_mesin      : "-",
        foto_suhu:       visitInfo ? visitInfo.foto_suhu       : "-",
        foto_box_it:     visitInfo ? visitInfo.foto_box_it     : "-",
        suhu_ruangan:    visitInfo ? visitInfo.suhu_ruangan    : "-",
        tegangan:        visitInfo ? visitInfo.tegangan        : "-",
        grounding:       visitInfo ? visitInfo.grounding       : "-",
        keterangan:      visitInfo ? visitInfo.keterangan      : "-",
        foto_tegangan:   visitInfo ? visitInfo.foto_tegangan   : "-",
        foto_grounding:  visitInfo ? visitInfo.foto_grounding  : "-"
      });
    }

    finalData.sort(function(a, b) {
      if (b.rawTime !== a.rawTime) return b.rawTime - a.rawTime;
      const roCompare = String(a.regional_office).localeCompare(String(b.regional_office), undefined, { numeric: true, sensitivity: 'base' });
      if (roCompare !== 0) return roCompare;
      const kcCompare = String(a.kc_supervisi).localeCompare(String(b.kc_supervisi), undefined, { numeric: true, sensitivity: 'base' });
      if (kcCompare !== 0) return kcCompare;
      return String(a.terminal_id).localeCompare(String(b.terminal_id), undefined, { numeric: true, sensitivity: 'base' });
    });

    return {
      success: true, level: userLevel, data: finalData,
      filters: {
        regional_office: Array.from(listRoSet).sort(),
        kc_supervisi:    Array.from(listKcSet).sort(),
        kl_cro:          Array.from(listKlCroSet).sort()
      }
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ====================================================================
// GET DASHBOARD DATA (SURVEY)
// ====================================================================
function getDashboardData(username, levelInput, filterMonth, filterYear, filterRo, filterKc) {
  try {
    const activeSS = SpreadsheetApp.getActiveSpreadsheet();
    const matchedUsername = String(username || "").trim();
    const matchedUsernameLower = matchedUsername.toLowerCase();

    function formatPersen(val) {
      const num = parseFloat(val);
      if (isNaN(num)) return "0,00%";
      return num.toFixed(2).replace('.', ',') + '%';
    }

    const responsesSS = getResponsesSpreadsheet(filterMonth, filterYear);
    const userSheet = activeSS.getSheetByName("User");
    let userLevel = "";
    if (userSheet) {
      const userData = userSheet.getDataRange().getValues();
      const userHeader = userData[0].map(h => String(h).trim().toLowerCase());
      const idxUsername = userHeader.indexOf("username");
      const idxLevel = userHeader.indexOf("level");
      if (idxUsername !== -1 && idxLevel !== -1) {
        for (let i = 1; i < userData.length; i++) {
          if (String(userData[i][idxUsername] || "").trim().toLowerCase() === matchedUsernameLower) {
            userLevel = String(userData[i][idxLevel] || "").trim().toUpperCase();
            break;
          }
        }
      }
    }
    if (!userLevel) userLevel = String(levelInput || "").trim().toUpperCase();

    const tidSheet = activeSS.getSheetByName("tid");
    if (!tidSheet) return { success: false, message: 'Sheet "tid" tidak ditemukan!' };

    const tidData = tidSheet.getDataRange().getValues();
    if (tidData.length <= 1) {
      return {
        success: true, userLevel: userLevel,
        filters: { regional_office: [], kc_supervisi: [] },
        summary: { totalTerminals: 0, totalAtm: 0, totalCrm: 0,
          visitedAtm: 0, unvisitedAtm: 0, pctAtm: "0,00%",
          visitedCrm: 0, unvisitedCrm: 0, pctCrm: "0,00%",
          problemCount: 0, nilaiAtm: "0,00%", nilaiCrm: "0,00%", totalNilai: "0,00%" }
      };
    }

    const tidHeaderRaw = tidData[0];
    const colTid       = findColIndex(tidHeaderRaw, ["terminal_id", "tid", "terminal id"]);
    const colRo        = findColIndex(tidHeaderRaw, ["ro"]);
    const colRegionOff = findColIndex(tidHeaderRaw, ["regional_office", "regional office"]);
    const colBo        = findColIndex(tidHeaderRaw, ["bo", "branch office"]);
    const colKc        = findColIndex(tidHeaderRaw, ["kc_supervisi", "kc supervisi", "kc"]);
    const colJenisEch  = findColIndex(tidHeaderRaw, ["jenis_ech", "jenis ech", "jenis"]);
    const colKelolaan  = findColIndex(tidHeaderRaw, ["kelolaan"]);
    const colKlCro     = findColIndex(tidHeaderRaw, ["kl_cro", "kl cro", "kl/cro", "cro"]);

    const filteredTidData = [];
    const listRoSet = new Set();
    const listKcSet = new Set();
    const allowedTidMap = {};

    for (let t = 1; t < tidData.length; t++) {
      const rowTid = tidData[t];
      const valTid       = colTid !== -1 ? String(rowTid[colTid] || "").trim() : "";
      const valRo        = colRo !== -1 ? String(rowTid[colRo] || "").trim() : "";
      const valRegionOff = colRegionOff !== -1 ? String(rowTid[colRegionOff] || "").trim() : valRo;
      const valBo        = colBo !== -1 ? String(rowTid[colBo] || "").trim() : "";
      const valKc        = colKc !== -1 ? String(rowTid[colKc] || "").trim() : "";
      const valJenisEch  = colJenisEch !== -1 ? String(rowTid[colJenisEch] || "").trim().toUpperCase() : "";
      const valKelolaan  = colKelolaan !== -1 ? String(rowTid[colKelolaan] || "").trim() : "";
      const valKlCro     = colKlCro !== -1 ? String(rowTid[colKlCro] || "").trim() : "";

      let isMatch = false;
      if (userLevel === "HO" || userLevel === "HEAD OFFICE") isMatch = true;
      else if (userLevel === "BO") { if (valBo.toLowerCase() === matchedUsernameLower) isMatch = true; }
      else if (userLevel === "RO") { if (valRo.toLowerCase() === matchedUsernameLower) isMatch = true; }
      else if (userLevel === "CRO") { if (valKelolaan.toLowerCase() === matchedUsernameLower) isMatch = true; }
      else if (userLevel === "KL") {
        const klCroNorm = valKlCro.toLowerCase().replace(/\s+/g, '');
        const usernameNorm = matchedUsernameLower.replace(/\s+/g, '');
        if (klCroNorm === usernameNorm) isMatch = true;
      }
      if (!isMatch) continue;

      const displayRo = (valRegionOff !== "" && valRegionOff !== "-") ? valRegionOff : valRo;
      if (filterRo && displayRo.toLowerCase() !== String(filterRo).trim().toLowerCase()) continue;
      if (filterKc && valKc.toLowerCase() !== String(filterKc).trim().toLowerCase()) continue;

      if (displayRo && displayRo !== "-") listRoSet.add(displayRo);
      if (valKc && valKc !== "-") listKcSet.add(valKc);
      if (valTid) allowedTidMap[valTid] = valJenisEch;

      filteredTidData.push({ terminal_id: valTid, jenis_ech: valJenisEch, ro: displayRo, kc_supervisi: valKc });
    }

    const responseSheet = responsesSS.getSheetByName("Responses");
    const visitedMap = {};
    const targetMonth = filterMonth !== undefined && filterMonth !== "" ? parseInt(filterMonth, 10) : (new Date().getMonth() + 1);
    const targetYear  = filterYear !== undefined && filterYear !== "" ? parseInt(filterYear, 10) : new Date().getFullYear();

    if (responseSheet) {
      const responseData = responseSheet.getDataRange().getValues();
      if (responseData.length > 1) {
        const respHeader = responseData[0];
        const rColTid       = findColIndex(respHeader, ["terminal_id", "qtid", "tid"]);
        const rColTimestamp = findColIndex(respHeader, ["timestamp"]);
        const rColNilai     = findColIndex(respHeader, ["nilai"]);
        const rColProblem   = findColIndex(respHeader, ["problem"]);
        const rColJenisEch  = findColIndex(respHeader, ["jenis_echannel", "jenis_ech"]);

        if (rColTid !== -1) {
          for (let r = 1; r < responseData.length; r++) {
            const rowResp = responseData[r];
            const tidResp = String(rowResp[rColTid] || "").trim();
            if (!tidResp) continue;
            if (!allowedTidMap.hasOwnProperty(tidResp)) continue;

            let rawDateObj = null;
            const tsVal = rColTimestamp !== -1 ? rowResp[rColTimestamp] : null;
            if (tsVal instanceof Date) rawDateObj = tsVal;
            else if (typeof tsVal === 'string' && tsVal.trim() !== "") {
              const parsed = new Date(tsVal);
              if (!isNaN(parsed.getTime())) rawDateObj = parsed;
            }
            if (rawDateObj) {
              if ((rawDateObj.getMonth() + 1) !== targetMonth || rawDateObj.getFullYear() !== targetYear) continue;
            }

            const jenisEchResponses = rColJenisEch !== -1 ? String(rowResp[rColJenisEch] || "").trim().toUpperCase() : "";
            visitedMap[tidResp] = {
              jenis_echannel_responses: jenisEchResponses,
              nilai: rColNilai !== -1 ? rowResp[rColNilai] : "-",
              problem: rColProblem !== -1 ? rowResp[rColProblem] || "-" : "-"
            };
          }
        }
      }
    }

    let totalAtm = 0, totalCrm = 0;
    let visitedAtm = 0, visitedCrm = 0;
    let sumNilaiAtm = 0, countNilaiAtm = 0;
    let sumNilaiCrm = 0, countNilaiCrm = 0;
    let countProblemTerminals = 0;

    filteredTidData.forEach(function(row) {
      const jEchTid = String(row.jenis_ech || "").trim().toUpperCase();
      const visitInfo = visitedMap[row.terminal_id];
      const isVisited = !!visitInfo;
      if (jEchTid === 'ATM') totalAtm++;
      else if (jEchTid === 'CRM') totalCrm++;

      if (isVisited) {
        const jEchResp = String(visitInfo.jenis_echannel_responses || "").trim().toUpperCase();
        if (jEchResp === 'ATM') {
          visitedAtm++;
          if (visitInfo.nilai !== undefined && visitInfo.nilai !== "-" && visitInfo.nilai !== "") {
            const numVal = parseFloat(String(visitInfo.nilai).replace('%', '').replace(',', '.'));
            if (!isNaN(numVal)) { sumNilaiAtm += (numVal * 100); countNilaiAtm++; }
          }
        } else if (jEchResp === 'CRM') {
          visitedCrm++;
          if (visitInfo.nilai !== undefined && visitInfo.nilai !== "-" && visitInfo.nilai !== "") {
            const numVal = parseFloat(String(visitInfo.nilai).replace('%', '').replace(',', '.'));
            if (!isNaN(numVal)) { sumNilaiCrm += (numVal * 100); countNilaiCrm++; }
          }
        }
        if (visitInfo.problem && visitInfo.problem !== '-' && visitInfo.problem.trim() !== '') {
          countProblemTerminals++;
        }
      }
    });

    const unvisitedAtm = totalAtm - visitedAtm;
    const unvisitedCrm = totalCrm - visitedCrm;
    const avgNilaiAtm = countNilaiAtm > 0 ? (sumNilaiAtm / countNilaiAtm) : 0;
    const avgNilaiCrm = countNilaiCrm > 0 ? (sumNilaiCrm / countNilaiCrm) : 0;
    let totalNilai = 0;
    if (avgNilaiAtm > 0 && avgNilaiCrm > 0) totalNilai = (avgNilaiAtm + avgNilaiCrm) / 2;
    else if (avgNilaiAtm > 0) totalNilai = avgNilaiAtm;
    else if (avgNilaiCrm > 0) totalNilai = avgNilaiCrm;

    const pctAtm = totalAtm > 0 ? (visitedAtm / totalAtm) * 100 : 0;
    const pctCrm = totalCrm > 0 ? (visitedCrm / totalCrm) * 100 : 0;

    return {
      success: true, userLevel: userLevel,
      filters: { regional_office: Array.from(listRoSet).sort(), kc_supervisi: Array.from(listKcSet).sort() },
      summary: {
        totalTerminals: totalAtm + totalCrm, totalAtm: totalAtm, totalCrm: totalCrm,
        visitedAtm: visitedAtm, unvisitedAtm: unvisitedAtm, pctAtm: formatPersen(pctAtm),
        visitedCrm: visitedCrm, unvisitedCrm: unvisitedCrm, pctCrm: formatPersen(pctCrm),
        problemCount: countProblemTerminals,
        nilaiAtm: formatPersen(avgNilaiAtm), nilaiCrm: formatPersen(avgNilaiCrm),
        totalNilai: formatPersen(totalNilai)
      }
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ====================================================================
// KELOLAAN DATA (HO)
// ====================================================================
function getKelolaanData(username, levelInput) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const matchedUsernameLower = String(username || '').trim().toLowerCase();
    let userLevel = String(levelInput || '').trim().toUpperCase();
    if (userLevel === 'HEAD OFFICE') userLevel = 'HO';
    const userSheet = ss.getSheetByName('User');
    if (userSheet) {
      const userData = userSheet.getDataRange().getValues();
      const userHeader = userData[0].map(h => String(h).trim().toLowerCase());
      const idxUsername = userHeader.indexOf('username');
      const idxLevel = userHeader.indexOf('level');
      if (idxUsername !== -1 && idxLevel !== -1) {
        for (let i = 1; i < userData.length; i++) {
          if (String(userData[i][idxUsername] || '').trim().toLowerCase() === matchedUsernameLower) {
            userLevel = String(userData[i][idxLevel] || '').trim().toUpperCase();
            break;
          }
        }
      }
    }
    if (userLevel === 'HEAD OFFICE') userLevel = 'HO';
    if (userLevel !== 'HO') return { success: false, message: 'Akses Data Kelolaan hanya untuk level Head Office.' };

    const tidSheet = ss.getSheetByName('tid');
    if (!tidSheet) return { success: false, message: 'Sheet "tid" tidak ditemukan!' };
    const tidData = tidSheet.getDataRange().getValues();
    if (tidData.length <= 1) return { success: true, level: userLevel, data: [] };

    const headers = tidData[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, '_'));
    function findCol(names) {
      for (let n = 0; n < names.length; n++) {
        const cleanTarget = normalizeKey(names[n]);
        for (let c = 0; c < headers.length; c++) {
          if (normalizeKey(headers[c]) === cleanTarget) return c;
        }
      }
      return -1;
    }
    const colId        = findCol(['id']);
    const colTid       = findCol(['terminal_id', 'terminal id', 'tid']);
    const colRo        = findCol(['ro']);
    const colRegionOff = findCol(['regional_office', 'regional office']);
    const colBo        = findCol(['bo']);
    const colKc        = findCol(['kc_supervisi', 'kc supervisi']);
    const colLokasi    = findCol(['lokasi']);
    const colJenisEch  = findCol(['jenis_ech', 'jenis ech']);
    const colStatusEch = findCol(['status_ech', 'status ech']);
    const colKelolaan  = findCol(['kelolaan']);
    const colKlCro     = findCol(['kl_cro', 'kl cro']);

    const result = [];
    let rowIndex = 0;
    for (let t = 1; t < tidData.length; t++) {
      const row = tidData[t];
      rowIndex++;
      const valRo          = colRo !== -1 ? String(row[colRo] || '').trim() : '';
      const valRegionalOff = colRegionOff !== -1 ? String(row[colRegionOff] || '').trim() : '';
      const rawJenis = colJenisEch !== -1 ? String(row[colJenisEch] || '').trim() : '';
      let jenisNorm = 'ATM';
      if (rawJenis.toUpperCase().indexOf('CRM') !== -1) jenisNorm = 'CRM';
      const rawStatus = colStatusEch !== -1 ? String(row[colStatusEch] || '').trim() : '';
      let statusNorm = 'ONSITE';
      const statusUpper = rawStatus.toUpperCase().replace(/\s+/g, '');
      if (statusUpper.indexOf('OFFSITE') !== -1) statusNorm = 'OFFSITE';

      result.push({
        no: rowIndex, row_index: t + 1,
        id: colId !== -1 ? row[colId] : '',
        terminal_id: colTid !== -1 ? String(row[colTid] || '').trim() : '',
        ro: valRo, regional_office: valRegionalOff,
        bo: colBo !== -1 ? String(row[colBo] || '').trim() : '',
        kc_supervisi: colKc !== -1 ? String(row[colKc] || '').trim() : '',
        lokasi: colLokasi !== -1 ? String(row[colLokasi] || '').trim() : '',
        jenis_ech: jenisNorm, status_ech: statusNorm,
        kelolaan: colKelolaan !== -1 ? String(row[colKelolaan] || '').trim() : '',
        kl_cro: colKlCro !== -1 ? String(row[colKlCro] || '').trim() : '',
        editable: true
      });
    }
    return { success: true, level: userLevel, data: result };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function updateKelolaanData(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('tid');
    if (!sheet) return { success: false, message: 'Sheet "tid" tidak ditemukan!' };
    const rowIndex = parseInt(payload.row_index, 10);
    if (!rowIndex || rowIndex < 2) return { success: false, message: 'Baris tidak valid.' };
    const jenisEch = String(payload.jenis_ech || '').trim().toUpperCase();
    const statusEch = String(payload.status_ech || '').trim().toUpperCase();
    payload.jenis_ech = jenisEch;
    payload.status_ech = statusEch;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      .map(h => String(h).trim().toLowerCase().replace(/\s+/g, '_'));
    function findCol(names) {
      for (let n = 0; n < names.length; n++) {
        const cleanTarget = normalizeKey(names[n]);
        for (let c = 0; c < headers.length; c++) {
          if (normalizeKey(headers[c]) === cleanTarget) return c + 1;
        }
      }
      return -1;
    }
    const fieldMapping = {
      terminal_id:     findCol(['terminal_id', 'terminal id', 'tid']),
      ro:              findCol(['ro']),
      regional_office: findCol(['regional_office', 'regional office']),
      kc_supervisi:    findCol(['kc_supervisi', 'kc supervisi']),
      lokasi:          findCol(['lokasi']),
      jenis_ech:       findCol(['jenis_ech', 'jenis ech']),
      status_ech:      findCol(['status_ech', 'status ech']),
      kelolaan:        findCol(['kelolaan']),
      kl_cro:          findCol(['kl_cro', 'kl cro'])
    };
    Object.keys(fieldMapping).forEach(function(key) {
      const colIdx = fieldMapping[key];
      if (colIdx !== -1 && payload[key] !== undefined) sheet.getRange(rowIndex, colIdx).setValue(payload[key]);
    });
    SpreadsheetApp.flush();
    return { success: true, message: 'Data kelolaan berhasil diperbarui.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteKelolaanData(rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('tid');
    if (!sheet) return { success: false, message: 'Sheet "tid" tidak ditemukan!' };
    const idx = parseInt(rowIndex, 10);
    if (!idx || idx < 2) return { success: false, message: 'Baris tidak valid.' };
    sheet.deleteRow(idx);
    SpreadsheetApp.flush();
    return { success: true, message: 'Data kelolaan berhasil dihapus.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function addKelolaanData(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('tid');
    if (!sheet) return { success: false, message: 'Sheet "tid" tidak ditemukan!' };
    const terminalId = String(payload.terminal_id || '').trim();
    if (!terminalId) return { success: false, message: 'Terminal ID wajib diisi.' };
    const data = sheet.getDataRange().getValues();
    const headerRow = data[0];
    function findCol(names) {
      for (let n = 0; n < names.length; n++) {
        const cleanTarget = normalizeKey(names[n]);
        for (let c = 0; c < headerRow.length; c++) {
          if (normalizeKey(headerRow[c]) === cleanTarget) return c;
        }
      }
      return -1;
    }
    const colTid = findCol(['terminal_id', 'terminal id', 'tid']);
    if (colTid === -1) return { success: false, message: 'Kolom "terminal id" tidak ditemukan di sheet tid.' };
    for (let i = 1; i < data.length; i++) {
      const existingTid = String(data[i][colTid] || '').trim().toLowerCase();
      if (existingTid === terminalId.toLowerCase()) {
        return { success: false, message: 'Terminal ID "' + terminalId + '" sudah ada di database.' };
      }
    }
    const colId = findCol(['id']);
    let nextId = 1;
    if (colId !== -1) {
      for (let i = data.length - 1; i >= 1; i--) {
        const valId = Number(data[i][colId]);
        if (!isNaN(valId) && valId >= nextId) { nextId = valId + 1; break; }
      }
    }
    const fieldMapping = {
      id: colId, terminal_id: colTid,
      ro: findCol(['ro']),
      regional_office: findCol(['regional_office', 'regional office']),
      bo: findCol(['bo']),
      kc_supervisi: findCol(['kc_supervisi', 'kc supervisi']),
      lokasi: findCol(['lokasi']),
      jenis_ech: findCol(['jenis_ech', 'jenis ech']),
      status_ech: findCol(['status_ech', 'status ech']),
      kelolaan: findCol(['kelolaan']),
      kl_cro: findCol(['kl_cro', 'kl cro'])
    };
    const newRow = new Array(sheet.getLastColumn()).fill('');
    if (fieldMapping.id !== -1) newRow[fieldMapping.id] = nextId;
    if (fieldMapping.terminal_id !== -1) newRow[fieldMapping.terminal_id] = terminalId;
    if (fieldMapping.ro !== -1) newRow[fieldMapping.ro] = payload.ro || '';
    if (fieldMapping.regional_office !== -1) newRow[fieldMapping.regional_office] = payload.regional_office || payload.ro || '';
    if (fieldMapping.bo !== -1) newRow[fieldMapping.bo] = payload.bo || '';
    if (fieldMapping.kc_supervisi !== -1) newRow[fieldMapping.kc_supervisi] = payload.kc_supervisi || '';
    if (fieldMapping.lokasi !== -1) newRow[fieldMapping.lokasi] = payload.lokasi || '';
    if (fieldMapping.jenis_ech !== -1) newRow[fieldMapping.jenis_ech] = payload.jenis_ech || 'ATM';
    if (fieldMapping.status_ech !== -1) newRow[fieldMapping.status_ech] = payload.status_ech || 'ONSITE';
    if (fieldMapping.kelolaan !== -1) newRow[fieldMapping.kelolaan] = payload.kelolaan || '';
    if (fieldMapping.kl_cro !== -1) newRow[fieldMapping.kl_cro] = payload.kl_cro || '';
    sheet.appendRow(newRow);
    SpreadsheetApp.flush();
    return { success: true, message: 'Data "' + terminalId + '" berhasil ditambahkan!' };
  } catch (e) {
    return { success: false, message: 'Gagal menambahkan data: ' + e.message };
  }
}

// ====================================================================
// EXPORT FOTO & CSV (SURVEY)
// ====================================================================
function getExportPhotosList(singleDateStr) {
  try {
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const files = folder.getFiles();
    const targetDateStr = String(singleDateStr || '').trim();
    const photoList = [];
    const timeZone = Session.getScriptTimeZone();
    const MAX_FILES = 500;
    while (files.hasNext() && photoList.length < MAX_FILES) {
      const file = files.next();
      if (file.getName().toLowerCase().endsWith('.zip')) continue;
      const createdDate = file.getDateCreated();
      const fileDateStr = Utilities.formatDate(createdDate, timeZone, 'yyyy-MM-dd');
      if (targetDateStr && fileDateStr !== targetDateStr) continue;
      let size = 0;
      try { size = file.getSize(); } catch (e) { size = 0; }
      if (size > 4 * 1024 * 1024) continue;
      photoList.push({ id: file.getId(), name: file.getName(), mimeType: file.getMimeType(), size: size });
    }
    return { success: true, photos: photoList, total: photoList.length, limited: photoList.length >= MAX_FILES };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getPhotoBase64ById(fileId) {
  try {
    if (!fileId) return { success: false, message: 'File ID kosong' };
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const bytes = blob.getBytes();
    if (bytes.length > 4 * 1024 * 1024) return { success: false, message: 'File terlalu besar (>4MB)' };
    return { success: true, id: fileId, name: file.getName(), mimeType: file.getMimeType(), size: bytes.length, base64: Utilities.base64Encode(bytes) };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getSinglePhotoData(photoName) {
  try {
    if (!photoName || photoName === '-' || String(photoName).trim() === '') return { success: false, message: 'Nama file kosong' };
    const cleanName = String(photoName).trim();
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const files = folder.getFilesByName(cleanName);
    if (!files.hasNext()) return { success: false, message: 'File tidak ditemukan: ' + cleanName };
    const file = files.next();
    const blob = file.getBlob();
    const bytes = blob.getBytes();
    if (bytes.length > 4 * 1024 * 1024) return { success: false, message: 'File terlalu besar (>4MB)' };
    return { success: true, photo: { name: file.getName(), mimeType: file.getMimeType(), size: bytes.length, base64: Utilities.base64Encode(bytes) } };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getExportResponsesData(singleDateStr) {
  try {
    const sheet = getSheet(CONFIG.SHEET_RESPONSES);
    if (!sheet) return { success: false, message: 'Sheet Responses tidak ditemukan' };
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, headers: data[0] || [], rows: [] };
    const headers = data[0].map(function(h) { return String(h || ''); });
    const targetDateStr = String(singleDateStr || '').trim();
    const timeZone = Session.getScriptTimeZone();
    let tsIdx = -1;
    for (let i = 0; i < headers.length; i++) {
      if (normalizeKey(headers[i]) === 'timestamp') { tsIdx = i; break; }
    }
    const rows = [];
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      if (targetDateStr && tsIdx !== -1) {
        const tsVal = row[tsIdx];
        let ok = false;
        if (tsVal instanceof Date) ok = (Utilities.formatDate(tsVal, timeZone, 'yyyy-MM-dd') === targetDateStr);
        else if (typeof tsVal === 'string') {
          const parts = tsVal.split(' ')[0].split('/');
          if (parts.length === 3) {
            const normalized = parts[2] + '-' + parts[1].padStart(2,'0') + '-' + parts[0].padStart(2,'0');
            ok = (normalized === targetDateStr);
          }
        }
        if (!ok) continue;
      }
      rows.push(row);
    }
    return { success: true, headers: headers, rows: rows };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ====================================================================
// VERIFIKASI TID & SAVING RESPONSES (SURVEY)
// ====================================================================
function verifyTid(tid) {
  const tidStr = String(tid || '').trim().toLowerCase();
  if (!tidStr) return { valid: false, data: null };
  const terminals = getRows(CONFIG.SHEET_TID);
  const terminal = terminals.find(t => {
    const candidates = [t.terminal_id, t['terminal id'], t.tid, t.terminal_id_, t.terminalid];
    for (let i = 0; i < candidates.length; i++) {
      if (candidates[i] !== undefined && String(candidates[i]).trim().toLowerCase() === tidStr) return true;
    }
    return false;
  }) || null;
  return { valid: !!terminal, data: terminal };
}

function saveResponse(formData, fileMetadata) {
  const lock = LockService.getPublicLock();
  try {
    lock.waitLock(30000);
    const sheet = getSheet(CONFIG.SHEET_RESPONSES);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const terminalId = formData.qtid;
    const verifyResult = verifyTid(terminalId);
    if (!verifyResult.valid) return { success: false, message: 'Terminal ID tidak ditemukan di database' };
    const terminal = verifyResult.data;
    const now = new Date();
    const pad = (num) => String(num).padStart(2, '0');
    const formattedTimestamp = `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
    const currentNextId = getNextId(CONFIG.SHEET_RESPONSES);
    const dataMap = {
      'id': currentNextId, 'timestamp': formattedTimestamp, 'terminal_id': terminalId,
      'ro': terminal.ro || terminal.regional_office || '',
      'regional_office': terminal.regional_office || terminal.ro || '',
      'bo': terminal.bo || '',
      'kc_supervisi': terminal.kc_supervisi || terminal.kc_supervisi_ || '',
      'lokasi': terminal.lokasi || '',
      'jenis_ech': terminal.jenis_ech || terminal.jenis_ech_ || '',
      'status_ech': terminal.status_ech || terminal.status_ech_ || '',
      'kelolaan': terminal.kelolaan || '',
      'kl_cro': terminal.kl_cro || terminal.kl_cro_ || '',
      'jenis_echannel': formData.qech || '',
      'status_echannel': formData.start || '',
      'jenis_lokasi': formData.qJenisLokasi || ''
    };
    const mapping = {
      'qPylon': 'pylon', 'qkondisipylon': 'kondisi_pylon', 'qdesignpylon': 'design_pylon',
      'qSignage': 'signage', 'qkondisisignage': 'lampu_signage', 'qdesignsignage': 'kondisi_signage',
      'qstickertatatertib': 'sticker_tata_tertib', 'qsanblast': 'sticker_sanblast',
      'qstickerpintu': 'sticker_pintu', 'qtarikdorong': 'sticker_tarik_dorong',
      'qpintukaca': 'pintu_kaca', 'qdinding': 'dinding', 'qlantai': 'lantai', 'qplafon': 'plafon',
      'qmatkom': 'matkom', 'qkabel': 'kabel_rapi', 'qcover': 'jenis_cover',
      'qkondisicover': 'kondisi_cover', 'qlampucover': 'lampu_cover', 'qlampu': 'lampu_ruangan',
      'qsuhu': 'suhu_ruangan', 'qtegangan': 'tegangan', 'qgrounding': 'grounding',
      'qhandle': 'handle_pintu', 'qsampah': 'tempat_sampah', 'qboxit': 'box_it',
      'qfotomesin': 'foto_mesin', 'qstickertid': 'sticker_tid', 'qinsert': 'sticker_insert',
      'qdenom': 'sticker_denom', 'qstickersetor': 'sticker_setor', 'qbrizzi': 'sticker_brizzi',
      'qstickerbody': 'sticker_body', 'qdsa': 'dsa', 'qmenu': 'menu_monitor',
      'qbody': 'body_mesin', 'qangka': 'epp_jelas', 'qtombol': 'epp_fungsi',
      'qskimming': 'skimming', 'qpalsu': 'sticker_palsu', 'qfotopalsu': 'foto_palsu',
      'qbrosur': 'brosur_lain', 'qfotobrosur': 'foto_brosur', 'qhal': 'penanganan',
      'qnama': 'nama_petugas', 'qnomor': 'no_petugas', 'qhp': 'no_hp', 'qpernyataan': 'pernyataan'
    };
    for (const [formKey, dbCol] of Object.entries(mapping)) {
      if (formData[formKey] !== undefined && formData[formKey] !== null && formData[formKey] !== '') {
        dataMap[dbCol] = formData[formKey];
      }
    }
    if (formData.qsuhu) dataMap['suhu_ruangan'] = formData.qsuhu;
    if (fileMetadata) {
      for (const [dbCol, filename] of Object.entries(fileMetadata)) {
        if (filename) dataMap[dbCol] = filename;
      }
    }
    const calcResult = calculateScoreAndProblems(dataMap);
    dataMap['nilai'] = Math.round(calcResult.score * 10000) / 10000;
    dataMap['problem'] = calcResult.problems.join('; ');
    const finalRowValues = headers.map(h => {
      const sanitizedHeader = String(h).toLowerCase().trim().replace(/\s+/g, '_');
      return dataMap[sanitizedHeader] !== undefined ? dataMap[sanitizedHeader] : '';
    });
    sheet.appendRow(finalRowValues);
    SpreadsheetApp.flush();
    return { success: true, message: 'Data survey berhasil!' };
  } catch (e) {
    return { success: false, message: 'Sistem sibuk, gagal memproses survey: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function uploadSingleFile(fileDataObj, filename) {
  try {
    if (!fileDataObj || !fileDataObj.data) throw new Error('Data file tidak valid');
    if (!filename || filename.trim() === '') throw new Error('Nama file tidak boleh kosong');
    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    const base64Data = fileDataObj.data.split(',')[1];
    if (!base64Data) throw new Error('Format data base64 tidak valid');
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), fileDataObj.mimeType);
    const sizeInMB = blob.getBytes().length / (1024 * 1024);
    if (sizeInMB > CONFIG.MAX_FILE_SIZE_MB) throw new Error('Ukuran file melebihi batas maksimum ' + CONFIG.MAX_FILE_SIZE_MB + 'MB');
    blob.setName(filename);
    folder.createFile(blob);
    return { success: true };
  } catch (e) {
    console.error('Gagal upload lampiran ' + filename + ': ' + e.message);
    throw new Error(e.message);
  }
}

function calculateScoreAndProblems(data) {
  let totalScore = 0;
  const problems = [];
  function checkQuestion(value, validValue, problemText, scorePoints) {
    const strVal = String(value || '').trim();
    if (!strVal) return;
    if (strVal.toLowerCase() === String(validValue).trim().toLowerCase()) totalScore += (scorePoints !== undefined ? scorePoints : 1);
    else if (problemText && problemText !== '-') problems.push(problemText);
  }
  if (data.pylon === 'Ya') {
    checkQuestion(data.pylon, 'Ya', 'Pylon tidak ada');
    checkQuestion(data.kondisi_pylon, 'Baik', 'Pylon rusak');
    checkQuestion(data.design_pylon, 'Ya', 'Pylon terdapat logo BRI Agro / Syariah');
  } else if (data.pylon === 'Tidak') problems.push('Pylon tidak ada');
  if (data.signage === 'Ya') {
    checkQuestion(data.signage, 'Ya', 'Signage tidak ada');
    checkQuestion(data.lampu_signage, 'Menyala saat malam', 'Lampu Signage mati saat malam');
    checkQuestion(data.kondisi_signage, 'Baik', 'Signage rusak');
  } else if (data.signage === 'Tidak') problems.push('Signage tidak ada');
  checkQuestion(data.sticker_tata_tertib, 'Baik', 'Sticker Tata Tertib rusak');
  checkQuestion(data.sticker_sanblast, 'Baik', 'Sticker Sanblast rusak');
  checkQuestion(data.sticker_pintu, 'Baik', 'Sticker Pintu rusak');
  checkQuestion(data.sticker_tarik_dorong, 'Baik', 'Sticker Tarik Dorong rusak');
  checkQuestion(data.pintu_kaca, 'Bersih', 'Pintu kaca kotor');
  checkQuestion(data.dinding, 'Bersih', 'Dinding kotor');
  checkQuestion(data.lantai, 'Bersih', 'Lantai kotor');
  checkQuestion(data.plafon, 'Bersih', 'Plafon kotor');
  checkQuestion(data.matkom, 'Ya', 'Matkom tidak sesuai');
  checkQuestion(data.kabel_rapi, 'Rapi', 'Kabel berantakan');
  if (data.jenis_cover && data.jenis_cover !== 'Tidak ada') {
    checkQuestion(data.kondisi_cover, 'Baik', 'Cover rusak');
    checkQuestion(data.lampu_cover, 'Menyala', 'Lampu Cover mati');
  } else if (data.jenis_cover === 'Tidak ada') problems.push('Cover tidak ada');
  checkQuestion(data.lampu_ruangan, 'Terang', 'Lampu ruangan redup');
  if (data.suhu_ruangan === '< 23°C' || data.suhu_ruangan === '23°C - 25°C') totalScore += 1;
  else problems.push('Ruangan panas / AC tidak ada');
  checkQuestion(data.tegangan, '180 V - 260 V', 'Tegangan tidak stabil');
  checkQuestion(data.grounding, '< 1 V', 'Grounding > 1V');
  checkQuestion(data.handle_pintu, 'Baik', 'Handle pintu rusak');
  checkQuestion(data.tempat_sampah, 'Bagus', 'Tempat sampah rusak');
  checkQuestion(data.box_it, 'Ya', 'Box IT tidak terkunci');
  checkQuestion(data.sticker_tid, 'Bersih', 'Sticker TID kotor');
  checkQuestion(data.sticker_insert, 'Bersih', 'Sticker Insert Card kotor');
  if (data.jenis_echannel === 'ATM') checkQuestion(data.sticker_denom, 'Bersih', 'Sticker Denom kotor');
  else if (data.jenis_echannel === 'CRM') {
    checkQuestion(data.sticker_setor, 'Bersih', 'Sticker Setor Tarik kotor');
    checkQuestion(data.dsa, 'Baik', 'DSA rusak');
  }
  checkQuestion(data.sticker_brizzi, 'Bersih', 'Sticker BRIZZI kotor');
  checkQuestion(data.sticker_body, 'Bersih', 'Sticker Body kotor');
  checkQuestion(data.menu_monitor, 'Jelas', 'Layar monitor silau');
  checkQuestion(data.body_mesin, 'Baik', 'Body mesin rusak');
  checkQuestion(data.epp_jelas, 'Jelas', 'EPP/Softkey tidak jelas');
  checkQuestion(data.epp_fungsi, 'Ya', 'Tombol EPP/Softkey tidak berfungsi');
  checkQuestion(data.skimming, 'Tidak', 'Terindikasi alat skimming');
  checkQuestion(data.sticker_palsu, 'Tidak', 'Terdapat sticker informasi palsu');
  checkQuestion(data.brosur_lain, 'Tidak', 'Terdapat brosur / sticker non-BRI');
  let divider = 35;
  const jenisEch = String(data.jenis_echannel || '').trim().toUpperCase();
  const statusEch = String(data.status_echannel || '').trim().toUpperCase();
  const jenisLokasi = String(data.jenis_lokasi || '').trim().toLowerCase();
  if (statusEch === 'ONSITE') {
    if (jenisLokasi === 'ruangan khusus') divider = (jenisEch === 'ATM') ? 35 : 36;
    else if (jenisLokasi === 'ruangan sharing') divider = (jenisEch === 'ATM') ? 24 : 25;
  } else if (statusEch === 'OFFSITE') {
    if (jenisLokasi === 'ruangan khusus') divider = (jenisEch === 'ATM') ? 38 : 39;
    else if (jenisLokasi === 'ruangan sharing') divider = (jenisEch === 'ATM') ? 27 : 28;
  }
  const finalScore = divider > 0 ? (totalScore / divider) : 0;
  return { score: finalScore, problems: problems };
}

// ====================================================================
// CHANGE PASSWORD
// ====================================================================
function changeUserPassword(username, oldPassword, newPassword) {
  try {
    const sheet = getSheet(CONFIG.SHEET_USER);
    if (!sheet) return { success: false, message: 'Sheet User tidak ditemukan.' };
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: false, message: 'Data User kosong.' };
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    const usernameIdx = headers.indexOf('username');
    const passwordIdx = headers.indexOf('password');
    let changeIdx = headers.indexOf('change');
    let timestampIdx = headers.indexOf('timestamp');
    if (usernameIdx === -1 || passwordIdx === -1) return { success: false, message: 'Kolom Username atau Password tidak ditemukan.' };
    if (changeIdx === -1) { changeIdx = headers.length; sheet.getRange(1, changeIdx + 1).setValue('Change'); }
    if (timestampIdx === -1) { timestampIdx = headers.length + (changeIdx === headers.length ? 1 : 0); sheet.getRange(1, timestampIdx + 1).setValue('timestamp'); }
    const cleanInputUser = String(username || '').trim().toLowerCase();
    const cleanOldPass = String(oldPassword || '').trim();
    const cleanNewPass = String(newPassword || '').trim();
    if (cleanNewPass.length < 4) return { success: false, message: 'Password baru minimal 4 karakter.' };
    for (let i = 1; i < data.length; i++) {
      const dbUser = String(data[i][usernameIdx] || '').trim().toLowerCase();
      const dbPass = String(data[i][passwordIdx] || '').trim();
      const dbChange = String(data[i][changeIdx] || '').trim();
      const isPasswordValid = (dbPass === cleanOldPass) || (dbChange !== '' && dbChange === cleanOldPass);
      if (dbUser === cleanInputUser && isPasswordValid) {
        const now = new Date();
        const pad = (num) => String(num).padStart(2, '0');
        const formattedTimestamp = `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
        sheet.getRange(i + 1, changeIdx + 1).setValue(cleanNewPass);
        sheet.getRange(i + 1, timestampIdx + 1).setValue(formattedTimestamp);
        SpreadsheetApp.flush();
        return { success: true, message: 'Password berhasil diperbarui!' };
      }
    }
    return { success: false, message: 'Password lama Anda salah.' };
  } catch (e) {
    return { success: false, message: 'Gagal memperbarui password: ' + e.message };
  }
}

// ====================================================================
// BAGIAN 2: FUNGSI MONITORING SEWA
// ====================================================================
function getSessionId() {
  return Session.getActiveUser().getEmail() + "_" + new Date().getTime();
}

function getIndonesianTime() {
  return Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
}

function getIndonesianTimestamp() {
  return Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
}

function getCachedData(cacheKey) {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
    return null;
  } catch (e) { return null; }
}

function setCachedData(cacheKey, data, expirationSeconds = 60) {
  try {
    CacheService.getScriptCache().put(cacheKey, JSON.stringify(data), expirationSeconds);
  } catch (e) {}
}

function clearCache() {
  try {
    CacheService.getScriptCache().removeAll();
    return { success: true, message: "Cache cleared" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function cleanEmptyColumns(headers, data) {
  let validIndices = [];
  for (let i = 0; i < headers.length; i++) {
    let headerStr = String(headers[i] || "").trim();
    if (headerStr !== "" && headerStr !== "null" && headerStr !== "undefined") validIndices.push(i);
  }
  if (validIndices.length === headers.length) return { headers: headers, data: data };
  let newHeaders = validIndices.map(i => headers[i]);
  let newData = data.map(row => {
    if (!row || row.length === 0) return [];
    return validIndices.map(i => (i < row.length) ? row[i] : "");
  });
  return { headers: newHeaders, data: newData };
}

function parseDate(dateStr) {
  if (!dateStr) return null;
  if (dateStr instanceof Date) {
    var d = new Date(dateStr); d.setHours(0, 0, 0, 0); return d;
  }
  let parts = String(dateStr).split('/');
  if (parts.length === 3) {
    let day = parseInt(parts[0]), month = parseInt(parts[1]) - 1, year = parseInt(parts[2]);
    if (!isNaN(day) && !isNaN(month) && !isNaN(year)) { var d = new Date(year, month, day); d.setHours(0, 0, 0, 0); return d; }
  }
  let parts2 = String(dateStr).split('-');
  if (parts2.length === 3) {
    let year = parseInt(parts2[0]), month = parseInt(parts2[1]) - 1, day = parseInt(parts2[2]);
    if (!isNaN(day) && !isNaN(month) && !isNaN(year)) { var d = new Date(year, month, day); d.setHours(0, 0, 0, 0); return d; }
  }
  return null;
}

function validateRowData(rowData, isFree) {
  if (isFree || String(rowData["Kategori Lokasi"] || "").trim().toUpperCase() === "FREE") return [];
  var errors = [];
  if (!rowData["Terminal ID"] || String(rowData["Terminal ID"]).trim() === "") errors.push("Terminal ID wajib diisi");
  if (!rowData["Jenis Mesin"] || String(rowData["Jenis Mesin"]).trim() === "") errors.push("Jenis Mesin wajib diisi");
  if (!rowData["Status Sewa"] || String(rowData["Status Sewa"]).trim() === "") errors.push("Status Sewa wajib diisi");
  if (!rowData["Kategori Lokasi"] || String(rowData["Kategori Lokasi"]).trim() === "") errors.push("Kategori Lokasi wajib diisi");
  var hargaTerbaru = parseFloat(String(rowData["Harga Sewa Terbaru"] || "0").replace(/[^0-9.-]/g, ''));
  if (isNaN(hargaTerbaru) || hargaTerbaru <= 0) errors.push("Harga Sewa Terbaru wajib diisi dan tidak boleh 0");
  if (!rowData["Mekanisme Pembayaran"] || String(rowData["Mekanisme Pembayaran"]).trim() === "") errors.push("Mekanisme Pembayaran wajib diisi");
  var periodeBulan = parseInt(String(rowData["Periode Bulan"] || "0").replace(/[^0-9]/g, ''));
  if (isNaN(periodeBulan) || periodeBulan <= 0) errors.push("Periode Bulan wajib diisi dan harus berupa angka positif");
  if (!rowData["Tanggal Awal"] || String(rowData["Tanggal Awal"]).trim() === "") errors.push("Tanggal Awal Sewa wajib diisi");
  else { var tglAwal = parseDate(rowData["Tanggal Awal"]); if (!tglAwal) errors.push("Format Tanggal Awal tidak valid (gunakan DD/MM/YYYY)"); }
  if (!rowData["Tanggal Berakhir"] || String(rowData["Tanggal Berakhir"]).trim() === "") errors.push("Tanggal Berakhir Sewa wajib diisi");
  else {
    var tglBerakhir = parseDate(rowData["Tanggal Berakhir"]);
    if (!tglBerakhir) errors.push("Format Tanggal Berakhir tidak valid (gunakan DD/MM/YYYY)");
    else {
      var today = new Date(); var todayDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());
      if (tglBerakhir < todayDate) errors.push("Tanggal berakhir sewa sudah expired.");
      if (rowData["Tanggal Awal"]) {
        var tglAwal2 = parseDate(rowData["Tanggal Awal"]);
        if (tglAwal2 && tglBerakhir < tglAwal2) errors.push("Tanggal Berakhir tidak boleh lebih kecil dari Tanggal Awal");
      }
    }
  }
  if (rowData["Kategori Lokasi"] === "Kerjasama") {
    if (!rowData["Keterangan Kerjasama"] || String(rowData["Keterangan Kerjasama"]).trim() === "") errors.push("Keterangan Kerjasama wajib diisi untuk kategori Kerjasama");
  }
  return errors;
}

// ====================================================================
// GET DATA WITH FILTERS (SEWA)
// ====================================================================
function getDataWithFilters(loginData, filters) {
  try {
    const type = loginData.type || loginData.loginType;
    const value = loginData.value || loginData.lookupValue;
    if (!type || !value) return { headers: [], rows: [], error: "Parameter login tidak lengkap" };
    const ss = getMonitoringSpreadsheet();
    const sheet = ss.getSheetByName("Data");
    if (!sheet) return { headers: [], rows: [], error: "Sheet 'Data' tidak ditemukan" };
    let range = sheet.getDataRange();
    if (range.getNumRows() < 2) return { headers: [], rows: [], error: "Sheet Data kosong" };
    let data = range.getValues();
    let headers = data[0];
    while (headers.length > 0 && (headers[0] === "" || headers[0] === null)) { headers.shift(); for (let i = 0; i < data.length; i++) data[i].shift(); }
    while (headers.length > 0 && (headers[headers.length-1] === "" || headers[headers.length-1] === null)) { headers.pop(); for (let i = 0; i < data.length; i++) data[i].pop(); }
    let cleaned = cleanEmptyColumns(headers, data);
    headers = cleaned.headers; data = cleaned.data;
    let regionalIdx = headers.indexOf("Regional Office");
    let kcIdx = headers.indexOf("KC Supervisi");
    let tglBerakhirIdx = headers.indexOf("Tanggal Berakhir");
    let terminalIdIdx = headers.indexOf("Terminal ID");
    let nominalIdx = headers.indexOf("Jumlah Pembayaran");
    let jenisMesinIdx = headers.indexOf("Jenis Mesin");
    let kategoriLokasiIdx = headers.indexOf("Kategori Lokasi");
    let hargaTerbaruIdx = headers.indexOf("Harga Sewa Terbaru");
    if (regionalIdx === -1 || kcIdx === -1 || tglBerakhirIdx === -1) return { headers: headers, rows: [], error: "Kolom penting pada sheet Data tidak ditemukan" };
    const filterRegional = filters.regionalOffice && filters.regionalOffice !== "Semua" ? filters.regionalOffice : null;
    const filterKC = filters.kcSupervisi && filters.kcSupervisi !== "Semua" ? filters.kcSupervisi : null;
    const filterJatuhTempo = filters.jatuhTempo && filters.jatuhTempo !== "Semua" ? filters.jatuhTempo : null;
    const filterTahun = filters.tahun && filters.tahun !== "Semua" ? parseInt(filters.tahun) : null;
    var today = new Date(); var todayDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());

    let accessibleRows = [];
    for (let i = 1; i < data.length; i++) {
      let row = data[i]; let include = true;
      if (type !== "HO") {
        if (type === "RO" && String(row[regionalIdx]).trim().toLowerCase() !== String(value).trim().toLowerCase()) include = false;
        if (type === "BO" && String(row[kcIdx]).trim().toLowerCase() !== String(value).trim().toLowerCase()) include = false;
      }
      if (filterRegional && String(row[regionalIdx]).trim().toLowerCase() !== String(filterRegional).trim().toLowerCase()) include = false;
      if (filterKC && String(row[kcIdx]).trim().toLowerCase() !== String(filterKC).trim().toLowerCase()) include = false;
      if (include) accessibleRows.push(row);
    }

    accessibleRows.sort(function(a, b) {
      var regionalA = String(a[regionalIdx] || "").trim().toLowerCase();
      var regionalB = String(b[regionalIdx] || "").trim().toLowerCase();
      if (regionalA < regionalB) return -1;
      if (regionalA > regionalB) return 1;
      var kcA = String(a[kcIdx] || "").trim().toLowerCase();
      var kcB = String(b[kcIdx] || "").trim().toLowerCase();
      if (kcA < kcB) return -1;
      if (kcA > kcB) return 1;
      var tidA = String(a[terminalIdIdx] || "").trim();
      var tidB = String(b[terminalIdIdx] || "").trim();
      var numA = parseInt(tidA); var numB = parseInt(tidB);
      if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
      return tidA.localeCompare(tidB);
    });

    let count3BulanAll = 0, count2BulanAll = 0, count1BulanAll = 0;
    let countExpiredAll = 0, countTanpaTanggalAll = 0, countBelumAdaNominalAll = 0;
    let countFreeAll = 0;

    accessibleRows.forEach(row => {
      let kategoriLokasi = String(row[kategoriLokasiIdx] || '').trim().toUpperCase();
      let isFree = kategoriLokasi === "FREE";
      if (isFree) { countFreeAll++; return; }
      let tglValue = row[tglBerakhirIdx]; let tglDate = parseDate(tglValue);
      let isTglKosong = (tglValue === "" || tglValue === null || tglValue === undefined || String(tglValue).trim() === "" || String(tglValue).trim() === "0" || !tglDate);
      if (isTglKosong) { countTanpaTanggalAll++; return; }
      let hargaTerbaruValue = row[hargaTerbaruIdx];
      let isHargaKosong = false;
      if (hargaTerbaruIdx !== -1) {
        let harga = parseFloat(String(hargaTerbaruValue).replace(/[^0-9.-]/g, ''));
        isHargaKosong = (harga === 0 || isNaN(harga) || hargaTerbaruValue === "" || hargaTerbaruValue === null || hargaTerbaruValue === undefined);
      } else isHargaKosong = true;
      if (isHargaKosong) { countBelumAdaNominalAll++; return; }
      if (tglDate) {
        var tglCompare = new Date(tglDate); tglCompare.setHours(0, 0, 0, 0);
        var diffDays = Math.ceil((tglCompare - todayDate) / (1000 * 60 * 60 * 24));
        if (diffDays < 0) countExpiredAll++;
        else if (diffDays >= 0 && diffDays <= 30) count1BulanAll++;
        else if (diffDays > 30 && diffDays <= 60) count2BulanAll++;
        else if (diffDays > 60 && diffDays <= 90) count3BulanAll++;
      }
    });

    let filteredRows = [];
    let totalNominal = 0;
    let uniqueTerminals = new Set();
    let countATM = 0, countCRM = 0;

    accessibleRows.forEach(row => {
      let tglValue = row[tglBerakhirIdx]; let tglDate = parseDate(tglValue);
      let isTglKosong = (tglValue === "" || tglValue === null || tglValue === undefined || String(tglValue).trim() === "" || String(tglValue).trim() === "0" || !tglDate);
      let kategoriLokasi = String(row[kategoriLokasiIdx] || '').trim().toUpperCase();
      let isFree = kategoriLokasi === "FREE";
      let terminalId = String(row[terminalIdIdx] || '').trim();
      let nominal = 0;
      if (nominalIdx !== -1 && row[nominalIdx]) nominal = parseFloat(String(row[nominalIdx]).replace(/[^0-9.-]/g, ''));
      let hargaTerbaruValue = row[hargaTerbaruIdx];
      let isHargaKosong = false;
      if (hargaTerbaruIdx !== -1) {
        let harga = parseFloat(String(hargaTerbaruValue).replace(/[^0-9.-]/g, ''));
        isHargaKosong = (harga === 0 || isNaN(harga) || hargaTerbaruValue === "" || hargaTerbaruValue === null || hargaTerbaruValue === undefined);
      } else isHargaKosong = true;
      let kategori = '';
      if (isFree) kategori = 'Free';
      else if (isTglKosong) kategori = 'Tidak Ada Tanggal';
      else if (isHargaKosong) kategori = 'Belum Ada Nominal';
      else if (tglDate) {
        var tglCompare = new Date(tglDate); tglCompare.setHours(0, 0, 0, 0);
        var diffDays = Math.ceil((tglCompare - todayDate) / (1000 * 60 * 60 * 24));
        if (diffDays < 0) kategori = 'Expired';
        else if (diffDays >= 0 && diffDays <= 30) kategori = '1 Bulan';
        else if (diffDays > 30 && diffDays <= 60) kategori = '2 Bulan';
        else if (diffDays > 60 && diffDays <= 90) kategori = '3 Bulan';
        else kategori = 'Lainnya (>90 hari)';
      }
      let includeInFilter = true;
      if (filterJatuhTempo) {
        switch(filterJatuhTempo) {
          case "Belum ada Nominal": if (kategori !== 'Belum Ada Nominal') includeInFilter = false; break;
          case "Belum ada tanggal": if (kategori !== 'Tidak Ada Tanggal') includeInFilter = false; break;
          case "Expired": if (kategori !== 'Expired') includeInFilter = false; break;
          case "1 bulan": if (kategori !== '1 Bulan') includeInFilter = false; break;
          case "2 bulan": if (kategori !== '2 Bulan') includeInFilter = false; break;
          case "3 bulan": if (kategori !== '3 Bulan') includeInFilter = false; break;
          case "Free": if (kategori !== 'Free') includeInFilter = false; break;
        }
      }
      if (filterTahun && includeInFilter) {
        if (isFree || isTglKosong) {}
        else if (tglDate && tglDate.getFullYear() !== filterTahun) includeInFilter = false;
      }
      if (!includeInFilter) return;
      let stringRow = row.map(cell => {
        if (cell === null || cell === undefined) return "";
        if (cell instanceof Date) return `${cell.getDate().toString().padStart(2,'0')}/${(cell.getMonth()+1).toString().padStart(2,'0')}/${cell.getFullYear()}`;
        return String(cell);
      });
      filteredRows.push(stringRow);
      if (terminalIdIdx !== -1 && terminalId) uniqueTerminals.add(terminalId);
      if (!isFree) {
        var nominalValue = 0;
        if (nominalIdx !== -1 && row[nominalIdx]) nominalValue = parseFloat(String(row[nominalIdx]).replace(/[^0-9.-]/g, ''));
        if ((isNaN(nominalValue) || nominalValue === 0) && hargaTerbaruIdx !== -1 && row[hargaTerbaruIdx]) nominalValue = parseFloat(String(row[hargaTerbaruIdx]).replace(/[^0-9.-]/g, ''));
        if (!isNaN(nominalValue) && nominalValue > 0) totalNominal += nominalValue;
      }
      if (jenisMesinIdx !== -1 && row[jenisMesinIdx]) {
        let machineType = String(row[jenisMesinIdx]).trim().toUpperCase();
        if (machineType === "ATM") countATM++;
        else if (machineType === "CRM") countCRM++;
      }
    });

    return {
      headers: headers, rows: filteredRows,
      summary: {
        totalTerminals: filteredRows.length, totalNominal: totalNominal,
        totalATM: countATM, totalCRM: countCRM,
        count3Bulan: count3BulanAll, count2Bulan: count2BulanAll, count1Bulan: count1BulanAll,
        countExpired: countExpiredAll, countTanpaTanggal: countTanpaTanggalAll,
        countBelumAdaNominal: countBelumAdaNominalAll, countFree: countFreeAll
      }
    };
  } catch (e) {
    console.error("❌ Error getDataWithFilters:", e);
    return { headers: [], rows: [], error: e.toString() };
  }
}

// ====================================================================
// GET ALL DATA FOR MANAGEMENT (SEWA)
// ====================================================================
function getAllDataForManagement(loginData) {
  try {
    const type = loginData.type;
    const value = loginData.value;
    if (!type || !value) return { headers: [], rows: [], error: "Invalid login" };
    const ss = getMonitoringSpreadsheet();
    const sheet = ss.getSheetByName("Data");
    if (!sheet) return { headers: [], rows: [], error: "Sheet Data tidak ditemukan" };
    let range = sheet.getDataRange();
    if (range.getNumRows() < 2) return { headers: [], rows: [], error: "Data kosong" };
    let data = range.getValues();
    let headers = data[0];
    while (headers.length > 0 && (headers[0] === "" || headers[0] === null)) { headers.shift(); for (let i=0;i<data.length;i++) data[i].shift(); }
    while (headers.length > 0 && (headers[headers.length-1] === "" || headers[headers.length-1] === null)) { headers.pop(); for (let i=0;i<data.length;i++) data[i].pop(); }
    let cleaned = cleanEmptyColumns(headers, data);
    headers = cleaned.headers; data = cleaned.data;
    let regionalIdx = headers.indexOf("Regional Office");
    let kcIdx = headers.indexOf("KC Supervisi");
    if (regionalIdx === -1 || kcIdx === -1) return { headers, rows: [], error: "Kolom Regional Office/KC Supervisi tidak ada" };
    const filteredRows = [];
    const originalIndices = [];
    if (type === "HO") {
      for (let i = 1; i < data.length; i++) { filteredRows.push(data[i]); originalIndices.push(i+1); }
    } else {
      for (let i = 1; i < data.length; i++) {
        let match = false;
        if (type === "RO") { if (String(data[i][regionalIdx]).trim().toLowerCase() === String(value).trim().toLowerCase()) match = true; }
        else if (type === "BO") { if (String(data[i][kcIdx]).trim().toLowerCase() === String(value).trim().toLowerCase()) match = true; }
        if (match) { filteredRows.push(data[i]); originalIndices.push(i+1); }
      }
    }
    return { headers: headers, rows: filteredRows, originalIndices: originalIndices };
  } catch (e) {
    return { headers: [], rows: [], error: e.toString() };
  }
}

function getOrInsertTimestampColumn(sheet) {
  let lastCol = sheet.getLastColumn();
  if (lastCol === 0) return -1;
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let tsIdx = headers.findIndex(h => String(h).trim().toLowerCase() === 'timestamp');
  if (tsIdx === -1) {
    lastCol++;
    sheet.getRange(1, lastCol).setValue("Timestamp");
    sheet.getRange(1, lastCol).setFontWeight("bold");
    return lastCol;
  }
  return tsIdx + 1;
}

// ====================================================================
// ADD NEW ROW (SEWA)
// ====================================================================
function addNewRow(loginData, rowData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    if (!loginData) return { success: false, message: "Login data tidak ditemukan" };
    if (!rowData) return { success: false, message: "Data yang akan disimpan tidak ditemukan" };
    const ss = getMonitoringSpreadsheet();
    const sheet = ss.getSheetByName("Data");
    if (!sheet) return { success: false, message: "Sheet Data tidak ditemukan" };
    const terminalId = String(rowData["Terminal ID"] || "").trim();
    if (!terminalId) return { success: false, message: "Terminal ID wajib diisi" };
    const isFree = String(rowData["Kategori Lokasi"] || "").trim().toUpperCase() === "FREE";
    const validationErrors = validateRowData(rowData, isFree);
    if (validationErrors.length > 0) return { success: false, message: "Validasi gagal:\n- " + validationErrors.join("\n- ") };
    let lastCol = sheet.getLastColumn();
    if (lastCol === 0) return { success: false, message: "Sheet Data kosong, tidak ada header" };
    let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    while (headers.length > 0 && (headers[0] === "" || headers[0] === null)) headers.shift();
    while (headers.length > 0 && (headers[headers.length - 1] === "" || headers[headers.length - 1] === null)) headers.pop();
    const terminalIdx = headers.indexOf("Terminal ID");
    if (terminalIdx === -1) return { success: false, message: 'Kolom "Terminal ID" tidak ditemukan di Sheet Data' };
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const existingTerminalIds = sheet.getRange(2, terminalIdx + 1, lastRow - 1, 1).getValues();
      const normalizedInput = terminalId.trim().toUpperCase();
      for (let i = 0; i < existingTerminalIds.length; i++) {
        const existingId = String(existingTerminalIds[i][0] || "").trim().toUpperCase();
        if (existingId !== "" && existingId === normalizedInput) {
          return { success: false, duplicate: true, message: "❌ Data tidak disimpan.\n\nTerminal ID " + terminalId + " sudah terdaftar di Sheet Data." };
        }
      }
    }
    rowData["Jenis Lokasi"] = "OFFSITE";
    let newRow = [];
    for (let i = 0; i < headers.length; i++) {
      const headerName = headers[i];
      let value = rowData[headerName];
      if (value === undefined || value === null) value = "";
      newRow.push(value);
    }
    newRow[terminalIdx] = terminalId;
    sheet.appendRow(newRow);
    SpreadsheetApp.flush();
    const newRowIndex = sheet.getLastRow();
    const timestamp = getIndonesianTimestamp();
    const tsColIndex = getOrInsertTimestampColumn(sheet);
    if (tsColIndex !== -1) { sheet.getRange(newRowIndex, tsColIndex).setValue(timestamp); SpreadsheetApp.flush(); }
    try { CacheService.getScriptCache().removeAll(); } catch (cacheError) {}
    return { success: true, verified: true, duplicate: false, rowIndex: newRowIndex, message: "✅ Data berhasil disimpan!\n\nTerminal ID: " + terminalId };
  } catch (e) {
    return { success: false, message: "❌ Gagal menyimpan data: " + e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (lockError) {}
  }
}

// ====================================================================
// UPDATE PERPANJANGAN DATA (SEWA)
// ====================================================================
function updatePerpanjanganData(loginData, rowIndex, updatedData) {
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  try {
    if (!loginData || typeof loginData !== 'object') return { success: false, message: "Parameter loginData tidak valid" };
    const type = loginData.type;
    const value = loginData.value;
    if (!type || !value) return { success: false, message: "Type atau value login tidak lengkap" };
    if (rowIndex === undefined || rowIndex === null || rowIndex < 1) return { success: false, message: "Row index tidak valid" };
    if (!updatedData || typeof updatedData !== 'object') return { success: false, message: "Data update tidak valid" };
    try { lock.waitLock(30000); lockAcquired = true; } catch (lockError) { return { success: false, message: "Sistem sedang sibuk, silakan coba lagi." }; }
    const ss = getMonitoringSpreadsheet();
    const sheet = ss.getSheetByName("Data");
    if (!sheet) return { success: false, message: "Sheet Data tidak ditemukan" };
    const lastColumn = sheet.getLastColumn();
    if (lastColumn < 1) return { success: false, message: "Sheet Data kosong" };
    let headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    headers = headers.map(header => String(header || "").trim());
    const regionalIdx = headers.indexOf("Regional Office");
    const kcIdx = headers.indexOf("KC Supervisi");
    const rowData = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    if (!rowData || rowData.length === 0) return { success: false, message: `Data pada baris ${rowIndex} tidak ditemukan` };
    let hasAccess = false;
    if (type === "HO") hasAccess = true;
    else if (type === "RO" && regionalIdx !== -1) hasAccess = (String(rowData[regionalIdx] || "").trim().toLowerCase() === String(value).trim().toLowerCase());
    else if (type === "BO" && kcIdx !== -1) hasAccess = (String(rowData[kcIdx] || "").trim().toLowerCase() === String(value).trim().toLowerCase());
    if (!hasAccess) return { success: false, message: `Anda tidak memiliki akses untuk mengubah data ini.` };
    const isFree = String(updatedData["Kategori Lokasi"] || rowData[headers.indexOf("Kategori Lokasi")] || "").trim().toUpperCase() === "FREE";
    const validationErrors = validateRowData(updatedData, isFree);
    if (validationErrors && validationErrors.length > 0) return { success: false, message: "Validasi gagal:\n- " + validationErrors.join("\n- ") };
    let changesLogged = []; let fieldsChanged = []; let beforeSnapshot = {};
    for (let i = 0; i < headers.length; i++) beforeSnapshot[headers[i]] = rowData[i];
    for (const key in updatedData) {
      if (!updatedData.hasOwnProperty(key)) continue;
      const colIdx = headers.indexOf(key);
      if (colIdx === -1) continue;
      const oldValue = String(rowData[colIdx] || "").trim();
      const newValue = String(updatedData[key] || "").trim();
      if (oldValue !== newValue) {
        fieldsChanged.push(key);
        changesLogged.push(`${key}: "${oldValue}" → "${newValue}"`);
        sheet.getRange(rowIndex, colIdx + 1).setValue(updatedData[key]);
      }
    }
    if (changesLogged.length > 0) {
      const timestamp = getIndonesianTimestamp();
      const tsColIndex = getOrInsertTimestampColumn(sheet);
      if (tsColIndex !== -1) sheet.getRange(rowIndex, tsColIndex).setValue(timestamp);
      SpreadsheetApp.flush();
      try { CacheService.getScriptCache().removeAll(); } catch (cacheError) {}
      return { success: true, message: `✅ Berhasil mengupdate ${changesLogged.length} field`, verified: true, rowIndex: rowIndex, fieldsChanged: fieldsChanged.length, changes: changesLogged };
    }
    return { success: true, message: "Tidak ada perubahan yang disimpan" };
  } catch (error) {
    return { success: false, message: "❌ Gagal mengupdate data: " + error.toString() };
  } finally {
    if (lockAcquired) { try { lock.releaseLock(); } catch (lockError) {} }
  }
}

// ====================================================================
// UPDATE TERMINAL ID ONLY (SEWA)
// ====================================================================
function updateTerminalIdOnly(loginData, rowIndex, oldTerminalId, newTerminalId) {
  try {
    const type = loginData.type;
    const value = loginData.value;
    if (!type || !value || !rowIndex || !newTerminalId) return { success: false, message: "Parameter tidak lengkap" };
    const ss = getMonitoringSpreadsheet();
    const sheet = ss.getSheetByName("Data");
    if (!sheet) return { success: false, message: "Sheet Data tidak ditemukan" };
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let terminalIdx = headers.indexOf("Terminal ID");
    let regionalIdx = headers.indexOf("Regional Office");
    let kcIdx = headers.indexOf("KC Supervisi");
    if (terminalIdx === -1) return { success: false, message: "Kolom Terminal ID tidak ditemukan" };
    let currentRowData = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    if (type !== "HO" && regionalIdx !== -1 && kcIdx !== -1) {
      let regional = currentRowData[regionalIdx];
      let kc = currentRowData[kcIdx];
      let hasAccess = false;
      if (type === "RO" && String(regional).trim().toLowerCase() === String(value).trim().toLowerCase()) hasAccess = true;
      if (type === "BO" && String(kc).trim().toLowerCase() === String(value).trim().toLowerCase()) hasAccess = true;
      if (!hasAccess) return { success: false, message: "Tidak memiliki akses ke data ini" };
    }
    let fullData = sheet.getDataRange().getValues();
    for (let i = 1; i < fullData.length; i++) {
      if (i === (rowIndex - 1)) continue;
      if (String(fullData[i][terminalIdx]).trim() === String(newTerminalId).trim()) return { success: false, message: "Terminal ID Baru sudah terdaftar di sistem pada perangkat lain" };
    }
    sheet.getRange(rowIndex, terminalIdx + 1).setValue(String(newTerminalId).trim());
    SpreadsheetApp.flush();
    try { CacheService.getScriptCache().removeAll(); } catch(err) {}
    return { success: true, message: "✅ Terminal ID berhasil diubah menjadi: " + newTerminalId, verified: true, rowIndex: rowIndex };
  } catch(e) {
    return { success: false, message: "❌ Gagal mengubah Terminal ID: " + e.toString() };
  }
}

// ====================================================================
// DELETE ROW DATA (SEWA)
// ====================================================================
function deleteRowData(loginData, rowIndex) {
  try {
    const type = loginData.type;
    const value = loginData.value;
    if (!type || !value || !rowIndex) return { success: false, message: "Parameter tidak lengkap" };
    const ss = getMonitoringSpreadsheet();
    const sheet = ss.getSheetByName("Data");
    if (!sheet) return { success: false, message: "Sheet Data tidak ditemukan" };
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    while (headers.length > 0 && headers[headers.length-1] === "") headers.pop();
    while (headers.length > 0 && headers[0] === "") headers.shift();
    let rowData = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    let regionalIdx = headers.indexOf("Regional Office");
    let kcIdx = headers.indexOf("KC Supervisi");
    let hasAccess = false;
    if (type === "HO") hasAccess = true;
    else if (type === "RO" && regionalIdx !== -1) {
      if (String(rowData[regionalIdx]).trim().toLowerCase() === String(value).trim().toLowerCase()) hasAccess = true;
    } else if (type === "BO" && kcIdx !== -1) {
      if (String(rowData[kcIdx]).trim().toLowerCase() === String(value).trim().toLowerCase()) hasAccess = true;
    }
    if (!hasAccess) return { success: false, message: "Anda tidak memiliki akses untuk menghapus data ini." };
    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();
    try { CacheService.getScriptCache().removeAll(); } catch(err) {}
    return { success: true, message: "✅ Data berhasil dihapus", verified: true };
  } catch(e) {
    return { success: false, message: "❌ Gagal menghapus data: " + e.toString() };
  }
}

// ====================================================================
// UPDATE KERJASAMA DATA (SEWA)
// ====================================================================
function updateKerjasamaData(loginData, rowIndex, updatedData) {
  try {
    const type = loginData.type;
    const value = loginData.value;
    if (!type || !value || !rowIndex) return { success: false, message: "Parameter tidak lengkap" };
    const ss = getMonitoringSpreadsheet();
    const sheet = ss.getSheetByName("Data");
    if (!sheet) return { success: false, message: "Sheet Data tidak ditemukan" };
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    while (headers.length > 0 && headers[headers.length-1] === "") headers.pop();
    while (headers.length > 0 && headers[0] === "") headers.shift();
    let cleaned = cleanEmptyColumns(headers, [headers]);
    headers = cleaned.headers;
    let oldRowData = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    let regionalIdx = headers.indexOf("Regional Office");
    let kcIdx = headers.indexOf("KC Supervisi");
    let hasAccess = false;
    if (type === "HO") hasAccess = true;
    else if (type === "RO" && regionalIdx !== -1) {
      if (String(oldRowData[regionalIdx]).trim().toLowerCase() === String(value).trim().toLowerCase()) hasAccess = true;
    } else if (type === "BO" && kcIdx !== -1) {
      if (String(oldRowData[kcIdx]).trim().toLowerCase() === String(value).trim().toLowerCase()) hasAccess = true;
    }
    if (!hasAccess) return { success: false, message: "Anda tidak memiliki akses untuk mengubah data ini." };
    let changesLogged = []; let fieldsChanged = [];
    for (let key in updatedData) {
      let colIdx = headers.indexOf(key);
      if (colIdx !== -1) {
        let oldValue = String(oldRowData[colIdx] || "").trim();
        let newValCompare = String(updatedData[key]).trim();
        if (oldValue !== newValCompare) {
          fieldsChanged.push(key);
          changesLogged.push(`${key}: "${oldValue}" → "${newValCompare}"`);
          sheet.getRange(rowIndex, colIdx+1).setValue(updatedData[key]);
        }
      }
    }
    if (changesLogged.length > 0) {
      const timestamp = getIndonesianTimestamp();
      let tsColIndex = getOrInsertTimestampColumn(sheet);
      if (tsColIndex !== -1) sheet.getRange(rowIndex, tsColIndex).setValue(timestamp);
      SpreadsheetApp.flush();
      try { CacheService.getScriptCache().removeAll(); } catch(err) {}
      return { success: true, message: `✅ Berhasil mengupdate ${fieldsChanged.length} field data kerjasama`, verified: true, fieldsChanged: fieldsChanged.length };
    }
    return { success: true, message: "Tidak ada perubahan yang disimpan" };
  } catch(e) {
    return { success: false, message: "❌ Gagal mengupdate data: " + e.toString() };
  }
}

// ====================================================================
// UPDATE MASSAL INDOMARET (SEWA)
// ====================================================================
function updateMassalIndomaret(loginData) {
  try {
    const type = loginData.type;
    if (type !== 'HO') return { success: false, message: "Hanya Head Office yang dapat melakukan update massal Indomaret." };
    const ss = getMonitoringSpreadsheet();
    const sheet = ss.getSheetByName("Data");
    if (!sheet) return { success: false, message: "Sheet 'Data' tidak ditemukan" };
    let range = sheet.getDataRange();
    if (range.getNumRows() < 2) return { success: false, message: "Sheet Data kosong" };
    let data = range.getValues(); let headers = data[0];
    var colKerjasama = headers.indexOf("Kerjasama");
    var colJenisLokasi = headers.indexOf("Jenis Lokasi");
    var colStatusSewa = headers.indexOf("Status Sewa");
    var colKategoriLokasi = headers.indexOf("Kategori Lokasi");
    var colKeteranganKerjasama = headers.indexOf("Keterangan Kerjasama");
    var colHargaSebelumnya = headers.indexOf("Harga Sewa Sebelumnya");
    var colHargaTerbaru = headers.indexOf("Harga Sewa Terbaru");
    var colPeriodeBulan = headers.indexOf("Periode Bulan");
    var colMekanisme = headers.indexOf("Mekanisme Pembayaran");
    var colTanggalAwal = headers.indexOf("Tanggal Awal");
    var colTanggalBerakhir = headers.indexOf("Tanggal Berakhir");
    var colKeterangan = headers.indexOf("Keterangan");
    var colTimestamp = headers.indexOf("Timestamp");
    if (colKerjasama === -1) return { success: false, message: "Kolom 'Kerjasama' tidak ditemukan." };
    const currentYear = new Date().getFullYear();
    const tanggalAwal = "01/01/" + currentYear;
    const tanggalBerakhir = "31/12/" + currentYear;
    const hargaIndomaret = "18000000";
    let rowsToUpdate = []; let rowNumbers = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowKerjasama = String(row[colKerjasama] || '').trim().toLowerCase();
      if (rowKerjasama === "indomaret") { rowsToUpdate.push(row); rowNumbers.push(i + 1); }
    }
    if (rowsToUpdate.length === 0) return { success: false, message: "Tidak ditemukan data dengan nilai 'Indomaret' di kolom Kerjasama." };
    var totalUpdated = 0; var timestamp = getIndonesianTimestamp();
    for (var j = 0; j < rowsToUpdate.length; j++) {
      var rowNum = rowNumbers[j];
      if (colJenisLokasi !== -1) sheet.getRange(rowNum, colJenisLokasi + 1).setValue("OFFSITE");
      if (colStatusSewa !== -1) sheet.getRange(rowNum, colStatusSewa + 1).setValue("SW 2 - Perpanjangan");
      if (colKategoriLokasi !== -1) sheet.getRange(rowNum, colKategoriLokasi + 1).setValue("Kerjasama");
      if (colKeteranganKerjasama !== -1) sheet.getRange(rowNum, colKeteranganKerjasama + 1).setValue("Indomaret");
      if (colHargaSebelumnya !== -1) sheet.getRange(rowNum, colHargaSebelumnya + 1).setValue(hargaIndomaret);
      if (colHargaTerbaru !== -1) sheet.getRange(rowNum, colHargaTerbaru + 1).setValue(hargaIndomaret);
      if (colPeriodeBulan !== -1) sheet.getRange(rowNum, colPeriodeBulan + 1).setValue("12");
      if (colMekanisme !== -1) sheet.getRange(rowNum, colMekanisme + 1).setValue("FULL");
      if (colTanggalAwal !== -1) sheet.getRange(rowNum, colTanggalAwal + 1).setValue(tanggalAwal);
      if (colTanggalBerakhir !== -1) sheet.getRange(rowNum, colTanggalBerakhir + 1).setValue(tanggalBerakhir);
      if (colKeterangan !== -1) sheet.getRange(rowNum, colKeterangan + 1).setValue("Kerjasama Terpusat");
      if (colTimestamp !== -1) sheet.getRange(rowNum, colTimestamp + 1).setValue(timestamp);
      totalUpdated++;
      if (totalUpdated % 50 === 0) SpreadsheetApp.flush();
    }
    SpreadsheetApp.flush();
    try { CacheService.getScriptCache().removeAll(); } catch(err) {}
    return { success: true, message: "Berhasil mengupdate " + totalUpdated + " data Indomaret secara massal.", updatedCount: totalUpdated, timestamp: timestamp };
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

// ====================================================================
// GET DATA BY TERMINAL ID (SEWA)
// ====================================================================
function getDataByTerminalId(loginData, terminalId) {
  try {
    const type = loginData.type;
    const value = loginData.value;
    if (!type || !value || !terminalId) return { success: false, message: "Parameter tidak lengkap" };
    const ss = getMonitoringSpreadsheet();
    const sheet = ss.getSheetByName("Data");
    if (!sheet) return { success: false, message: "Sheet Data tidak ditemukan" };
    let range = sheet.getDataRange();
    if (range.getNumRows() < 2) return { success: false, message: "Data kosong" };
    let data = range.getValues(); let headers = data[0];
    while (headers.length > 0 && (headers[0] === "" || headers[0] === null)) { headers.shift(); for (let i = 0; i < data.length; i++) data[i].shift(); }
    while (headers.length > 0 && (headers[headers.length-1] === "" || headers[headers.length-1] === null)) { headers.pop(); for (let i = 0; i < data.length; i++) data[i].pop(); }
    let cleaned = cleanEmptyColumns(headers, data);
    headers = cleaned.headers; data = cleaned.data;
    let terminalIdx = headers.indexOf("Terminal ID");
    let regionalIdx = headers.indexOf("Regional Office");
    let kcIdx = headers.indexOf("KC Supervisi");
    let kerjasamaIdx = headers.indexOf("Kerjasama");
    if (terminalIdx === -1) return { success: false, message: "Kolom Terminal ID tidak ditemukan" };
    let foundRow = null; let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][terminalIdx]).trim() === String(terminalId).trim()) { foundRow = data[i]; rowIndex = i + 1; break; }
    }
    if (!foundRow) return { success: false, message: "Terminal ID tidak ditemukan" };
    let hasAccess = false;
    if (type === "HO") hasAccess = true;
    else if (type === "RO") {
      if (String(foundRow[regionalIdx]).trim().toLowerCase() === String(value).trim().toLowerCase()) hasAccess = true;
    } else if (type === "BO") {
      if (String(foundRow[kcIdx]).trim().toLowerCase() === String(value).trim().toLowerCase()) hasAccess = true;
    }
    if (!hasAccess) return { success: false, message: "Anda tidak memiliki akses ke data ini." };
    let resultData = {};
    for (let i = 0; i < headers.length; i++) {
      let val = foundRow[i];
      if (val === null || val === undefined) resultData[headers[i]] = "";
      else if (val instanceof Date) resultData[headers[i]] = `${val.getDate().toString().padStart(2,'0')}/${(val.getMonth()+1).toString().padStart(2,'0')}/${val.getFullYear()}`;
      else resultData[headers[i]] = String(val);
    }
    if (kerjasamaIdx !== -1) resultData["_kerjasama"] = String(foundRow[kerjasamaIdx] || '').trim();
    return { success: true, data: resultData, rowIndex: rowIndex, headers: headers };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ====================================================================
// GET AVAILABLE YEARS
// ====================================================================
function getAvailableYears() {
  try {
    let currentYear = new Date().getFullYear();
    let years = [currentYear - 1, currentYear];
    for (let i = 1; i <= 5; i++) years.push(currentYear + i);
    return years.sort((a, b) => a - b);
  } catch(e) { return []; }
}

// ====================================================================
// GET REGIONAL OFFICES
// ====================================================================
function getRegionalOffices() {
  const cacheKey = "cached_regional_offices";
  const cached = getCachedData(cacheKey);
  if (cached) return cached;
  try {
    const ss = getMonitoringSpreadsheet();
    const sheet = ss.getSheetByName("Data");
    if (!sheet) return [];
    let data = sheet.getDataRange().getValues();
    let headers = data[0];
    let regionalIdx = headers.indexOf("Regional Office");
    if (regionalIdx === -1) return [];
    let regionals = new Set();
    for (let i = 1; i < data.length; i++) {
      let regional = String(data[i][regionalIdx]).trim();
      if (regional) regionals.add(regional);
    }
    let result = Array.from(regionals).sort();
    setCachedData(cacheKey, result, 300);
    return result;
  } catch(e) { return []; }
}

// ==========================================
// GET KC SUPERVISI — dari kolom "KC Supervisi" sheet Data
// Final: case-insensitive, trim, dedupe case-insensitive, sort
// ✅ FIX: pakai getMonitoringSpreadsheet() supaya baca spreadsheet Data yang benar
// ==========================================
function getKCSupervisi(regionalOffice = null) {
  try {
    const ss = getMonitoringSpreadsheet();   // ✅ BENAR
    const sheet = ss.getSheetByName("Data");
    if (!sheet) return [];
    let data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    let headers = data[0];
    let regionalIdx = headers.indexOf("Regional Office");
    let kcIdx = headers.indexOf("KC Supervisi");
    if (regionalIdx === -1 || kcIdx === -1) return [];

    const filterRegionalRaw = (regionalOffice === null || regionalOffice === undefined || regionalOffice === "" || regionalOffice === "Semua")
      ? null
      : String(regionalOffice).trim().toLowerCase();

    let kcMap = {};
    for (let i = 1; i < data.length; i++) {
      const regional = String(data[i][regionalIdx] || "").trim();
      const kc       = String(data[i][kcIdx] || "").trim();
      if (!kc) continue;
      if (filterRegionalRaw && regional.toLowerCase() !== filterRegionalRaw) continue;
      const key = kc.toLowerCase();
      if (!kcMap[key]) kcMap[key] = kc;
    }
    return Object.keys(kcMap).map(k => kcMap[k]).sort();
  } catch(e) {
    console.error("❌ Error getKCSupervisi:", e);
    return [];
  }
}

function getKCSupervisiByRegionalForRO(regionalOffice) { return getKCSupervisi(regionalOffice); }
function getKCSupervisiByRegional(regionalOffice)      { return getKCSupervisi(regionalOffice); }

// ====================================================================
// GET USER DATA FOR FORM
// ====================================================================
function getUserDataForForm(loginData) {
  try {
    const type = loginData.type;
    const value = loginData.value;
    const ss = getMonitoringSpreadsheet();
    const sheet = ss.getSheetByName("Data");
    if (!sheet) return { error: "Sheet 'Data' tidak ditemukan" };
    let data = sheet.getDataRange().getValues();
    if (data.length < 2) return { error: "Sheet Data kosong" };
    let headers = data[0];
    let regionalIdx = headers.indexOf("Regional Office");
    let kcIdx = headers.indexOf("KC Supervisi");
    if (regionalIdx === -1 || kcIdx === -1) return { error: "Kolom 'Regional Office' atau 'KC Supervisi' tidak ditemukan" };
    if (type === "HO") {
      let regionals = new Set();
      for (let i = 1; i < data.length; i++) {
        let regional = String(data[i][regionalIdx]).trim();
        if (regional) regionals.add(regional);
      }
      return { regionalList: Array.from(regionals).sort(), userType: "HO" };
    }
    if (type === "RO") {
      let kcList = new Set();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][regionalIdx]).trim().toLowerCase() === String(value).trim().toLowerCase()) {
          let kc = String(data[i][kcIdx]).trim();
          if (kc) kcList.add(kc);
        }
      }
      return { regionalOffice: value, kcList: Array.from(kcList).sort(), userType: "RO" };
    }
    if (type === "BO") {
      let targetRegional = "";
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][kcIdx]).trim().toLowerCase() === String(value).trim().toLowerCase()) {
          targetRegional = String(data[i][regionalIdx]).trim(); break;
        }
      }
      return { regionalOffice: targetRegional, kcSupervisi: value, userType: "BO" };
    }
    return { error: "Tipe user tidak dikenal" };
  } catch (e) { return { error: e.toString() }; }
}

function getExportData(loginData, filters) {
  const result = getDataWithFilters(loginData, filters);
  if (result.error) throw new Error(result.error);
  result.exportTimestamp = getIndonesianTimestamp();
  return { headers: result.headers, rows: result.rows, exportTimestamp: result.exportTimestamp };
}

// ====================================================================
// AUDIT LOG
// ====================================================================
function writeDetailedAuditLog(userType, userValue, action, terminalId, fieldChanged, oldValue, newValue, details) {
  try {
    const ss = getMonitoringSpreadsheet();
    let auditSheet = ss.getSheetByName("AuditLog");
    if (!auditSheet) {
      auditSheet = ss.insertSheet("AuditLog");
      auditSheet.getRange(1, 1, 1, 9).setValues([["Timestamp", "User Type", "Username", "Action", "Terminal ID", "Field Changed", "Old Value", "New Value", "Details"]]);
      auditSheet.getRange(1, 1, 1, 9).setFontWeight("bold");
      auditSheet.setFrozenRows(1);
      auditSheet.getRange("A:I").setWrap(true);
    }
    const timestamp = getIndonesianTimestamp();
    auditSheet.appendRow([timestamp, userType, userValue, action, terminalId, fieldChanged,
      String(oldValue || "").substring(0, 500), String(newValue || "").substring(0, 500), String(details || "").substring(0, 500)]);
    return true;
  } catch(e) {
    console.error("Audit log error: " + e.toString());
    return false;
  }
}

// ====================================================================
// REKAPITULASI (SEWA) — ✅ FIXED: filterJatuhTempo sudah dipakai
// ====================================================================
function getRekapitulasiData(loginData, filters) {
  try {
    const type = loginData.type;
    const value = loginData.value;
    if (type === 'BO') return { error: "Menu Rekapitulasi hanya untuk Head Office dan Regional Office." };
    const ss = getMonitoringSpreadsheet();
    const sheet = ss.getSheetByName("Data");
    if (!sheet) return { error: "Sheet 'Data' tidak ditemukan" };
    let range = sheet.getDataRange();
    if (range.getNumRows() < 2) return { error: "Sheet Data kosong" };
    let data = range.getValues(); let headers = data[0];
    while (headers.length > 0 && (headers[0] === "" || headers[0] === null)) { headers.shift(); for (let i = 0; i < data.length; i++) data[i].shift(); }
    while (headers.length > 0 && (headers[headers.length-1] === "" || headers[headers.length-1] === null)) { headers.pop(); for (let i = 0; i < data.length; i++) data[i].pop(); }
    let cleaned = cleanEmptyColumns(headers, data);
    headers = cleaned.headers; data = cleaned.data;
    let regionalIdx = headers.indexOf("Regional Office");
    let kcIdx = headers.indexOf("KC Supervisi");
    let tglBerakhirIdx = headers.indexOf("Tanggal Berakhir");
    let nominalIdx = headers.indexOf("Jumlah Pembayaran");
    let terminalIdIdx = headers.indexOf("Terminal ID");
    let hargaTerbaruIdx = headers.indexOf("Harga Sewa Terbaru");
    let kategoriLokasiIdx = headers.indexOf("Kategori Lokasi");
    if (regionalIdx === -1) return { error: "Kolom 'Regional Office' tidak ditemukan" };
    if (kcIdx === -1) return { error: "Kolom 'KC Supervisi' tidak ditemukan" };
    if (tglBerakhirIdx === -1) return { error: "Kolom 'Tanggal Berakhir' tidak ditemukan" };
    if (terminalIdIdx === -1) return { error: "Kolom 'Terminal ID' tidak ditemukan" };
    const filterJatuhTempo = filters.jatuhTempo && filters.jatuhTempo !== "Semua" ? filters.jatuhTempo : null;
    const filterTahun = filters.tahun && filters.tahun !== "Semua" ? parseInt(filters.tahun) : null;
    var today = new Date(); var todayDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    let accessibleData = [];
    for (let i = 1; i < data.length; i++) {
      let row = data[i]; let include = true;
      if (type === "RO") {
        if (String(row[regionalIdx]).trim().toLowerCase() !== String(value).trim().toLowerCase()) include = false;
      }
      if (include) accessibleData.push(row);
    }
    let rekapData = []; let totalAllTID = 0; let title = "";
    if (type === "HO") {
      title = "Rekapitulasi per Regional Office";
      let regionals = new Set();
      for (let i = 0; i < accessibleData.length; i++) {
        let regional = String(accessibleData[i][regionalIdx]).trim();
        if (regional) regionals.add(regional);
      }
      regionals.forEach(regional => {
        let result = hitungRekapitulasiPerKC(accessibleData, headers, regional, null, regionalIdx, kcIdx, tglBerakhirIdx, nominalIdx, terminalIdIdx, hargaTerbaruIdx, kategoriLokasiIdx, filterJatuhTempo, filterTahun, todayDate);
        if (result) { rekapData.push(result); totalAllTID += result.totalTID; }
      });
    } else if (type === "RO") {
      title = `Rekapitulasi per KC Supervisi - ${value}`;
      let kcList = new Set();
      for (let i = 0; i < accessibleData.length; i++) {
        let kc = String(accessibleData[i][kcIdx]).trim();
        if (kc) kcList.add(kc);
      }
      kcList.forEach(kc => {
        let result = hitungRekapitulasiPerKC(accessibleData, headers, value, kc, regionalIdx, kcIdx, tglBerakhirIdx, nominalIdx, terminalIdIdx, hargaTerbaruIdx, kategoriLokasiIdx, filterJatuhTempo, filterTahun, todayDate);
        if (result) { rekapData.push(result); totalAllTID += result.totalTID; }
      });
    }
    rekapData.sort((a, b) => a.prosentase - b.prosentase);
    let total3Bulan = 0, total2Bulan = 0, total1Bulan = 0;
    let totalExpired = 0, totalTanpaTanggal = 0, totalBelumAdaNominal = 0;
    let totalFree = 0; let totalProsentase = 0;
    rekapData.forEach(item => {
      total3Bulan += item.count3Bulan; total2Bulan += item.count2Bulan; total1Bulan += item.count1Bulan;
      totalExpired += item.countExpired; totalTanpaTanggal += item.countTanpaTanggal;
      totalBelumAdaNominal += item.countBelumAdaNominal; totalFree += item.countFree || 0;
      totalProsentase += item.prosentase;
    });
    let avgProsentase = rekapData.length > 0 ? totalProsentase / rekapData.length : 0;
    return {
      rekapData: rekapData, title: title, userType: type, userValue: value,
      summary: {
        totalRegional: rekapData.length, totalTID: totalAllTID,
        total3Bulan: total3Bulan, total2Bulan: total2Bulan, total1Bulan: total1Bulan,
        totalExpired: totalExpired, totalTanpaTanggal: totalTanpaTanggal,
        totalBelumAdaNominal: totalBelumAdaNominal, totalFree: totalFree,
        totalProsentase: avgProsentase
      }
    };
  } catch (e) { return { error: e.toString() }; }
}

function hitungRekapitulasiPerKC(accessibleData, headers, regional, kc, regionalIdx, kcIdx, tglBerakhirIdx, nominalIdx, terminalIdIdx, hargaTerbaruIdx, kategoriLokasiIdx, filterJatuhTempo, filterTahun, todayDate) {
  let count3Bulan = 0, count2Bulan = 0, count1Bulan = 0;
  let countExpired = 0, countTanpaTanggal = 0, countBelumAdaNominal = 0, countFree = 0;
  let uniqueTerminals = new Set();
  for (let i = 0; i < accessibleData.length; i++) {
    let row = accessibleData[i];
    if (String(row[regionalIdx]).trim().toLowerCase() !== String(regional).trim().toLowerCase()) continue;
    if (kc && String(row[kcIdx]).trim().toLowerCase() !== String(kc).trim().toLowerCase()) continue;
    let includeInFilter = true;
    let tglValue = row[tglBerakhirIdx]; let tglDate = parseDate(tglValue);
    let isTglKosong = (tglValue === "" || tglValue === null || tglValue === undefined || String(tglValue).trim() === "" || String(tglValue).trim() === "0" || !tglDate);
    let terminalId = String(row[terminalIdIdx] || "").trim();
    let kategoriLokasi = String(row[kategoriLokasiIdx] || '').trim().toUpperCase();
    let isFree = kategoriLokasi === "FREE";
    if (terminalId) uniqueTerminals.add(terminalId);
    let kategori = '';
    if (isFree) kategori = 'Free';
    else if (isTglKosong) kategori = 'Tidak Ada Tanggal';
    else {
      let hargaTerbaruValue = row[hargaTerbaruIdx];
      let isHargaKosong = false;
      if (hargaTerbaruIdx !== -1) {
        let harga = parseFloat(String(hargaTerbaruValue).replace(/[^0-9.-]/g, ''));
        isHargaKosong = (harga === 0 || isNaN(harga) || hargaTerbaruValue === "" || hargaTerbaruValue === null || hargaTerbaruValue === undefined);
      } else isHargaKosong = true;
      if (isHargaKosong) kategori = 'Belum Ada Nominal';
      else if (tglDate) {
        var tglCompare = new Date(tglDate); tglCompare.setHours(0, 0, 0, 0);
        var diffDays = Math.ceil((tglCompare - todayDate) / (1000 * 60 * 60 * 24));
        if (diffDays < 0) kategori = 'Expired';
        else if (diffDays >= 0 && diffDays <= 30) kategori = '1 Bulan';
        else if (diffDays > 30 && diffDays <= 60) kategori = '2 Bulan';
        else if (diffDays > 60 && diffDays <= 90) kategori = '3 Bulan';
        else kategori = 'Lainnya (>90 hari)';
      }
    }

    // ✅ FIX: Filter Jatuh Tempo sekarang BENAR-BENAR dipakai
    if (filterJatuhTempo) {
      var isMatchJatuhTempo = false;
      switch(filterJatuhTempo) {
        case "Belum ada Nominal": if (kategori === 'Belum Ada Nominal') isMatchJatuhTempo = true; break;
        case "Belum ada tanggal": if (kategori === 'Tidak Ada Tanggal') isMatchJatuhTempo = true; break;
        case "Expired": if (kategori === 'Expired') isMatchJatuhTempo = true; break;
        case "1 bulan": if (kategori === '1 Bulan') isMatchJatuhTempo = true; break;
        case "2 bulan": if (kategori === '2 Bulan') isMatchJatuhTempo = true; break;
        case "3 bulan": if (kategori === '3 Bulan') isMatchJatuhTempo = true; break;
        case "Free": if (kategori === 'Free') isMatchJatuhTempo = true; break;
      }
      if (!isMatchJatuhTempo) continue;
    }

    if (filterTahun) {
      if (isFree || isTglKosong) {}
      else if (tglDate && tglDate.getFullYear() !== filterTahun) includeInFilter = false;
    }
    if (!includeInFilter) continue;
    switch(kategori) {
      case 'Tidak Ada Tanggal': countTanpaTanggal++; break;
      case 'Belum Ada Nominal': countBelumAdaNominal++; break;
      case 'Expired': countExpired++; break;
      case '1 Bulan': count1Bulan++; break;
      case '2 Bulan': count2Bulan++; break;
      case '3 Bulan': count3Bulan++; break;
      case 'Free': countFree++; break;
    }
  }
  let totalTID = uniqueTerminals.size;
  let prosentase = 0;
  let pembagi = countTanpaTanggal + countBelumAdaNominal + countExpired;
  if (totalTID > 0 && pembagi > 0) prosentase = (pembagi / totalTID) * 100;
  let label = kc ? kc : regional;
  return { label, regional, kc, totalTID, count3Bulan, count2Bulan, count1Bulan, countExpired, countTanpaTanggal, countBelumAdaNominal, countFree, prosentase, total: count3Bulan + count2Bulan + count1Bulan + countExpired };
}

// ====================================================================
// GET KERJASAMA DATA (SEWA)
// ====================================================================
function getKerjasamaData(loginData, filters) {
  try {
    const type = loginData.type;
    const value = loginData.value;
    if (type !== 'HO' && type !== 'RO' && type !== 'BO') return { error: "Menu Kerjasama hanya untuk HO/RO/BO." };
    const ss = getMonitoringSpreadsheet();
    const sheet = ss.getSheetByName("Data");
    if (!sheet) return { error: "Sheet 'Data' tidak ditemukan" };
    let range = sheet.getDataRange();
    if (range.getNumRows() < 2) return { error: "Sheet Data kosong" };
    let data = range.getValues(); let headers = data[0];
    let cleaned = cleanEmptyColumns(headers, data);
    headers = cleaned.headers; data = cleaned.data;
    let timestampIdx = -1;
    let timestampVariants = ["Timestamp", "Waktu Update", "Last Updated", "Terakhir Diperbarui", "LastModified"];
    for (let i = 0; i < headers.length; i++) {
      let hLower = String(headers[i]).toLowerCase().trim();
      if (timestampVariants.some(v => hLower === v.toLowerCase())) { timestampIdx = i; break; }
    }
    let kerjasamaIdx = headers.indexOf("Kerjasama");
    let regionalIdx = headers.indexOf("Regional Office");
    let kcIdx = headers.indexOf("KC Supervisi");
    let terminalIdIdx = headers.indexOf("Terminal ID");
    let hargaSebelumnyaIdx = headers.indexOf("Harga Sewa Sebelumnya");
    let hargaTerbaruIdx = headers.indexOf("Harga Sewa Terbaru");
    let kenaikanIdx = headers.indexOf("Kenaikan");
    let jumlahPembayaranIdx = headers.indexOf("Jumlah Pembayaran");
    let kategoriLokasiIdx = headers.indexOf("Kategori Lokasi");
    const filterKerjasama = filters.kerjasama && filters.kerjasama !== "Semua" ? filters.kerjasama.trim().toLowerCase() : null;
    const filterRegional = filters.regionalOffice && filters.regionalOffice !== "Semua" ? filters.regionalOffice.trim().toLowerCase() : null;
    const filterKC = filters.kcSupervisi && filters.kcSupervisi !== "Semua" ? filters.kcSupervisi.trim().toLowerCase() : null;
    let rowsWithData = []; let totalNominal = 0; let uniqueTerminals = new Set();
    for (let i = 1; i < data.length; i++) {
      let row = data[i]; let include = true;
      let rowKerjasama = kerjasamaIdx !== -1 ? String(row[kerjasamaIdx] || '').trim().toLowerCase() : '';
      let rowRegional = regionalIdx !== -1 ? String(row[regionalIdx] || '').trim().toLowerCase() : '';
      let rowKc = kcIdx !== -1 ? String(row[kcIdx] || '').trim().toLowerCase() : '';
      if (filterKerjasama) { if (rowKerjasama !== filterKerjasama) include = false; }
      else { if (!rowKerjasama) include = false; }
      if (type === "RO" && rowRegional !== String(value).trim().toLowerCase()) include = false;
      if (type === "BO" && rowKc !== String(value).trim().toLowerCase()) include = false;
      if (filterRegional && rowRegional !== filterRegional) include = false;
      if (filterKC && rowKc !== filterKC) include = false;
      if (include) {
        let terminalId = terminalIdIdx !== -1 ? String(row[terminalIdIdx] || '').trim() : '';
        let kenaikanValue = 0;
        if (hargaSebelumnyaIdx !== -1 && hargaTerbaruIdx !== -1) {
          let hargaSebelumnya = parseFloat(String(row[hargaSebelumnyaIdx]).replace(/[^0-9.-]/g, ''));
          let hargaTerbaru = parseFloat(String(row[hargaTerbaruIdx]).replace(/[^0-9.-]/g, ''));
          if (!isNaN(hargaSebelumnya) && !isNaN(hargaTerbaru) && hargaSebelumnya > 0) kenaikanValue = (hargaTerbaru - hargaSebelumnya) / hargaSebelumnya;
        }
        let timestampVal = "";
        if (timestampIdx !== -1) {
          let val = row[timestampIdx];
          if (val && String(val).trim() !== "") {
            timestampVal = (val instanceof Date) ? `${val.getDate().toString().padStart(2,'0')}/${(val.getMonth()+1).toString().padStart(2,'0')}/${val.getFullYear()} ${val.getHours().toString().padStart(2,'0')}:${val.getMinutes().toString().padStart(2,'0')}` : String(val);
          }
        }
        let stringRow = row.map(cell => {
          if (cell === null || cell === undefined) return "";
          if (cell instanceof Date) return `${cell.getDate().toString().padStart(2,'0')}/${(cell.getMonth()+1).toString().padStart(2,'0')}/${cell.getFullYear()}`;
          return String(cell);
        });
        if (kenaikanIdx !== -1) stringRow[kenaikanIdx] = kenaikanValue;
        rowsWithData.push({ row: stringRow, terminalId: terminalId, timestamp: timestampVal, rawTimestamp: timestampIdx !== -1 ? new Date(row[timestampIdx]).getTime() : 0, kerjasama: rowKerjasama });
        if (terminalId) uniqueTerminals.add(terminalId);
        if (jumlahPembayaranIdx !== -1 && row[jumlahPembayaranIdx]) {
          let nominal = parseFloat(String(row[jumlahPembayaranIdx]).replace(/[^0-9.-]/g, ''));
          if (!isNaN(nominal)) totalNominal += nominal;
        }
      }
    }
    rowsWithData.sort((a, b) => (isNaN(a.rawTimestamp) ? 0 : a.rawTimestamp) - (isNaN(b.rawTimestamp) ? 0 : b.rawTimestamp));
    let displayHeaders = ["No", "Action", "Timestamp", "Regional Office", "KC Supervisi", "Terminal ID", "Nama Lokasi", "Jenis Mesin", "Jenis Lokasi", "Status Sewa", "Kategori Lokasi", "Keterangan Kerjasama", "Harga Sewa Sebelumnya", "Harga Sewa Terbaru", "Kenaikan", "Periode Bulan", "Periode Tahun", "Jumlah Pembayaran", "Mekanisme Pembayaran", "Tanggal Awal", "Tanggal Berakhir", "Keterangan", "Kerjasama"];
    let validHeaders = displayHeaders.filter(h => h === "Action" || h === "No" || headers.indexOf(h) !== -1);
    let displayRows = rowsWithData.map((item, index) => {
      let newRow = [];
      validHeaders.forEach(function(h) {
        if (h === "Action") newRow.push(`<button class="edit-btn-kerjasama" data-terminal="${escapeHtml(item.terminalId)}">✏️ Edit</button>`);
        else if (h === "No") newRow.push(index + 1);
        else if (h === "Timestamp") newRow.push(item.timestamp || "");
        else {
          let idx = headers.indexOf(h);
          if (idx !== -1) {
            let val = item.row[idx];
            if (h === "Kenaikan") {
              let num = parseFloat(String(val).replace(/[^0-9.-]/g, ''));
              newRow.push(!isNaN(num) ? (num * 100).toFixed(2) + '%' : '0%');
            } else newRow.push(val);
          } else newRow.push("");
        }
      });
      return newRow;
    });
    return { headers: validHeaders, rows: displayRows, summary: { totalTerminals: uniqueTerminals.size, totalNominal: totalNominal }, showAction: true, isIndomaret: false };
  } catch (e) { return { error: e.toString() }; }
}

function exportKerjasamaToExcel(loginData, filters) {
  const result = getKerjasamaData(loginData, filters);
  if (result.error) throw new Error(result.error);
  result.exportTimestamp = getIndonesianTimestamp();
  return { headers: result.headers, rows: result.rows, summary: result.summary, exportTimestamp: result.exportTimestamp };
}

function escapeHtml(str) {
  if (!str) return "";
  return String(str).replace(/[&<>]/g, function(m) { return {'&': '&amp;', '<': '&lt;', '>': '&gt;'}[m] || m; });
}

function checkTerminalIdExists(terminalId) {
  try {
    if (!terminalId) return { exists: false, error: "Terminal ID tidak boleh kosong" };
    const ss = getMonitoringSpreadsheet();
    const sheet = ss.getSheetByName("Data");
    if (!sheet) return { exists: false, error: "Sheet 'Data' tidak ditemukan" };
    let data = sheet.getDataRange().getValues();
    if (data.length < 2) return { exists: false, message: "Sheet Data kosong" };
    let headers = data[0];
    let terminalIdx = headers.indexOf("Terminal ID");
    if (terminalIdx === -1) return { exists: false, error: "Kolom 'Terminal ID' tidak ditemukan" };
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][terminalIdx] || "").trim().toLowerCase() === String(terminalId).trim().toLowerCase()) {
        return { exists: true, message: "Terminal ID sudah terdaftar" };
      }
    }
    return { exists: false, message: "Terminal ID tersedia" };
  } catch (e) { return { exists: false, error: e.toString() }; }
}