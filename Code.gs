/**
 * ============================================================================
 * ระบบเฝ้าระวังคุณภาพอากาศและฝุ่นละออง PM 2.5 โรงพยาบาลโพนสวรรค์
 * Google Apps Script Backend (Code.gs) — ฐานข้อมูลจริง Phonsawan_Hospital_PM25_DB
 * ============================================================================
 * เงื่อนไขสำคัญ: ใช้ข้อมูลจริงจาก Air4Thai และฐานข้อมูลจริงใน Google Sheets เท่านั้น
 * ไม่มีการสร้างหรือเติมข้อมูลจำลอง (No Mock/Dummy Data)
 */

const CONFIG = {
  AIR4THAI_URL: 'http://air4thai.pcd.go.th/services/getNewAQI_JSON.php',
  AIR4THAI_URL_HTTPS: 'https://air4thai.pcd.go.th/services/getNewAQI_JSON.php',
  HOURLY_SHEET: 'Hourly_Logs',
  DAILY_SHEET: 'Daily_Summary',
  TIMEZONE: 'Asia/Bangkok',
  HOSPITAL_NAME: 'โรงพยาบาลโพนสวรรค์',
  HOSPITAL_PROVINCE: 'จังหวัดนครพนม',
  HOSPITAL_LAT: 17.4547,
  HOSPITAL_LON: 104.4581,
  PRIMARY_STATION_IDS: ['88t'],
  PRIMARY_KEYWORDS: ['นครพนม', 'nakhon phanom'],
  BACKUP_STATION_IDS: ['90t', '102t', '106t', '107t'],
  STALE_HOURS: 6,
  CACHE_SECONDS: 60
};

const HOURLY_HEADERS = [
  'Timestamp_Saved',
  'Date',
  'Time',
  'Station_ID',
  'Station_Name',
  'Area',
  'Is_Fallback',
  'PM25',
  'AQI',
  'AQI_Param',
  'PM10',
  'NO2',
  'Level_Num',
  'Level_Label'
];

const DAILY_HEADERS = [
  'Date',
  'Station_ID',
  'Station_Name',
  'PM25_Avg',
  'PM25_Min',
  'PM25_Max',
  'Peak_Time',
  'AQI_Max',
  'Hours_Count',
  'Level_Num',
  'Level_Label',
  'Updated_At'
];

/**
 * ดึงค่า PM 2.5 และ AQI จริงจาก Air4Thai ของสถานีจังหวัดนครพนม
 * บันทึกลงชีต Hourly_Logs (ป้องกันบันทึกซ้ำหากวัน-เวลาเดิมมีอยู่แล้ว)
 * และอัปเดตสรุปค่าเฉลี่ยรายวันลงชีต Daily_Summary
 */
function fetchAndSaveAir4Thai() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('ไม่พบ Google Spreadsheet กรุณาเปิดสคริปต์จาก Extensions > Apps Script');
  }

  try {
    if (ss.getSpreadsheetTimeZone() !== CONFIG.TIMEZONE) {
      ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);
    }
  } catch (e) {}

  const hourlySheet = ensureSheet_(ss, CONFIG.HOURLY_SHEET, HOURLY_HEADERS);
  const dailySheet = ensureSheet_(ss, CONFIG.DAILY_SHEET, DAILY_HEADERS);

  const rawStations = fetchAir4ThaiStations_();
  const selection = selectNakhonPhanomStation_(rawStations);
  const st = selection.active;

  const dataDate = st.date || Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const dataTime = st.time || Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'HH:00');
  const savedAt = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

  const isDup = isDuplicateHourly_(hourlySheet, dataDate, dataTime, st.stationID);
  if (!isDup) {
    const nextRow = hourlySheet.getLastRow() + 1;
    hourlySheet.getRange(nextRow, 1, 1, 3).setNumberFormat('@');
    hourlySheet.getRange(nextRow, 1, 1, HOURLY_HEADERS.length).setValues([[
      savedAt,
      dataDate,
      dataTime,
      sanitizeSheetCell_(st.stationID),
      sanitizeSheetCell_(st.nameTH),
      sanitizeSheetCell_(st.areaTH),
      st.isFallback ? 'YES' : 'NO',
      st.pm25,
      st.aqi !== null ? st.aqi : '',
      sanitizeSheetCell_(st.aqiParam || ''),
      st.pm10 !== null ? st.pm10 : '',
      st.no2 !== null ? st.no2 : '',
      st.level.level,
      sanitizeSheetCell_(st.level.label)
    ]]);
  }

  const dailySummary = upsertDailySummary_(hourlySheet, dailySheet, dataDate);

  try {
    CacheService.getScriptCache().remove('phonsawan_pm25_api_v2');
  } catch (e) {}

  const result = {
    ok: true,
    saved: !isDup,
    message: isDup
      ? 'วัน-เวลาเดิมถูกบันทึกไว้แล้วใน Hourly_Logs (อัปเดต Daily_Summary เรียบร้อย)'
      : 'บันทึกข้อมูลจริงลง Hourly_Logs และอัปเดต Daily_Summary เรียบร้อย',
    stationID: st.stationID,
    stationName: st.nameTH,
    date: dataDate,
    time: dataTime,
    pm25: st.pm25,
    aqi: st.aqi,
    dailySummary: dailySummary
  };
  Logger.log(JSON.stringify(result));
  return result;
}

/**
 * Web App Endpoint (HTTPS JSON) แก้ปัญหา CORS และ Mixed Content
 * ส่งออกค่าจริงแบบเรียลไทม์จาก Air4Thai + ข้อมูลประวัติจริงทั้งหมดที่มีใน Google Sheets
 */
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const requestedNoCache = params.nocache === '1' || params.nocache === 'true';

  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'phonsawan_pm25_api_v2';
    const rateKey = 'phonsawan_pm25_nocache_rl_v1';

    // จำกัดความถี่ nocache เพื่อป้องกันการยิงซ้ำถี่เกินไป (อย่างน้อยห่างกัน 15 วินาที)
    const allowBypass = requestedNoCache && !cache.get(rateKey);
    const cached = cache.get(cacheKey);
    if (cached && !allowBypass) {
      return ContentService
        .createTextOutput(cached)
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (allowBypass) {
      cache.put(rateKey, '1', 15);
    }

    const rawStations = fetchAir4ThaiStations_();
    const selection = selectNakhonPhanomStation_(rawStations);
    const history = readRealHistoryFromSheets_();

    const payload = {
      ok: true,
      source: 'air4thai_realtime_and_google_sheets',
      fetchedAt: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss+07:00"),
      hospital: {
        name: CONFIG.HOSPITAL_NAME,
        province: CONFIG.HOSPITAL_PROVINCE,
        lat: CONFIG.HOSPITAL_LAT,
        lon: CONFIG.HOSPITAL_LON
      },
      station: selection.active,
      nearbyStations: selection.nearbyStations,
      history: history
    };

    const jsonStr = JSON.stringify(payload);
    if (jsonStr.length < 95000) {
      cache.put(cacheKey, jsonStr, CONFIG.CACHE_SECONDS);
    }

    return ContentService
      .createTextOutput(jsonStr)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('doGet error: ' + (err && err.stack ? err.stack : String(err)));
    return ContentService
      .createTextOutput(JSON.stringify({
        ok: false,
        error: 'ไม่สามารถดึงข้อมูลคุณภาพอากาศได้ในขณะนี้',
        fetchedAt: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss+07:00")
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ตั้งค่า Time-driven Trigger ให้รัน fetchAndSaveAir4Thai() อัตโนมัติทุก 1 ชั่วโมง
 */
function setupHourlyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === 'fetchAndSaveAir4Thai') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('fetchAndSaveAir4Thai')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('ตั้งค่า Trigger รายชั่วโมงสำหรับ fetchAndSaveAir4Thai() เรียบร้อยแล้ว');
}

/* ==================== Air4Thai Fetch & Station Filter ==================== */

function fetchAir4ThaiStations_() {
  const opts = {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      'Accept': 'application/json',
      'User-Agent': 'PhonsawanHospital-PM25-GAS/2.0'
    }
  };

  let res = UrlFetchApp.fetch(CONFIG.AIR4THAI_URL, opts);
  if (res.getResponseCode() !== 200) {
    res = UrlFetchApp.fetch(CONFIG.AIR4THAI_URL_HTTPS, opts);
  }
  if (res.getResponseCode() !== 200) {
    throw new Error('ไม่สามารถเชื่อมต่อ Air4Thai ได้ (HTTP ' + res.getResponseCode() + ')');
  }

  const parsed = JSON.parse(res.getContentText('UTF-8'));
  if (!parsed || !Array.isArray(parsed.stations) || parsed.stations.length === 0) {
    throw new Error('โครงสร้าง JSON จาก Air4Thai ไม่ถูกต้อง');
  }
  return parsed.stations;
}

function selectNakhonPhanomStation_(rawStations) {
  const now = new Date();
  const list = rawStations.map(function (s) {
    return normalizeStation_(s, now);
  }).sort(function (a, b) {
    return a.distanceKm - b.distanceKm;
  });

  const nkpList = list.filter(function (st) {
    if (CONFIG.PRIMARY_STATION_IDS.indexOf(st.stationID) !== -1) return true;
    const text = (st.areaTH + ' ' + st.areaEN + ' ' + st.nameTH).toLowerCase();
    return CONFIG.PRIMARY_KEYWORDS.some(function (kw) {
      return text.indexOf(kw.toLowerCase()) !== -1;
    });
  });

  const primary = nkpList.length > 0 ? nkpList[0] : null;
  const freshPrimary = nkpList.find(function (st) {
    return st.hasValidPM25 && !st.isStale;
  });

  let active = null;
  if (freshPrimary) {
    active = JSON.parse(JSON.stringify(freshPrimary));
    active.isFallback = false;
    active.fallbackReason = '';
  } else {
    const reason = primary
      ? (!primary.hasValidPM25
        ? 'สถานีหลัก (' + primary.nameTH + ') ไม่มีค่า PM 2.5 ในรอบล่าสุด'
        : 'สถานีหลัก (' + primary.nameTH + ') ไม่อัปเดตเกิน ' + CONFIG.STALE_HOURS + ' ชม.')
      : 'ไม่พบสถานีจังหวัดนครพนมในข้อมูลล่าสุด';

    const freshBackups = list.filter(function (st) {
      return st.hasValidPM25 && !st.isStale && st.distanceKm <= 250;
    });
    const preferred = freshBackups.find(function (st) {
      return CONFIG.BACKUP_STATION_IDS.indexOf(st.stationID) !== -1;
    });
    const chosen = preferred || freshBackups[0] || null;

    if (chosen) {
      active = JSON.parse(JSON.stringify(chosen));
      active.isFallback = true;
      active.fallbackReason = reason + ' ระบบจึงใช้ข้อมูลสำรองจากสถานีใกล้เคียง (' + active.nameTH + ' ห่าง ' + active.distanceKm + ' กม.)';
    } else if (primary && primary.hasValidPM25) {
      active = JSON.parse(JSON.stringify(primary));
      active.isFallback = false;
      active.fallbackReason = 'หมายเหตุ: ข้อมูลจากสถานีหลักอาจล่าช้ากว่าปกติ';
    } else {
      throw new Error('ไม่พบสถานีตรวจวัดที่มีค่า PM 2.5 ที่ใช้งานได้');
    }
  }

  const nearbyStations = list.slice(0, 4).map(function (st) {
    return {
      stationID: st.stationID,
      nameTH: st.nameTH,
      areaTH: st.areaTH,
      distanceKm: st.distanceKm,
      date: st.date,
      time: st.time,
      pm25: st.pm25,
      aqi: st.aqi,
      isActive: st.stationID === active.stationID,
      level: st.level
    };
  });

  return { active: active, nearbyStations: nearbyStations };
}

function normalizeStation_(s, now) {
  const last = s.AQILast || {};
  const lat = parseFloat(s.lat);
  const lon = parseFloat(s.long);
  const distanceKm = (!isNaN(lat) && !isNaN(lon))
    ? round1_(haversineKm_(CONFIG.HOSPITAL_LAT, CONFIG.HOSPITAL_LON, lat, lon))
    : 9999;

  const pm25 = parseNum_(last.PM25 && last.PM25.value);
  const pm25Aqi = parseNum_(last.PM25 && last.PM25.aqi);
  const aqi = parseNum_(last.AQI && last.AQI.aqi);
  const dateStr = String(last.date || '').trim();
  const timeStr = String(last.time || '').trim();
  const ageHours = computeAgeHours_(dateStr, timeStr, now);

  return {
    stationID: String(s.stationID || ''),
    nameTH: String(s.nameTH || '').trim(),
    nameEN: String(s.nameEN || '').trim(),
    areaTH: String(s.areaTH || '').trim(),
    areaEN: String(s.areaEN || '').trim(),
    lat: isNaN(lat) ? null : lat,
    long: isNaN(lon) ? null : lon,
    distanceKm: distanceKm,
    date: dateStr,
    time: timeStr,
    ageHours: ageHours !== null ? round1_(ageHours) : null,
    isStale: ageHours === null ? true : ageHours > CONFIG.STALE_HOURS,
    hasValidPM25: pm25 !== null,
    pm25: pm25,
    pm25Aqi: pm25Aqi,
    pm10: parseNum_(last.PM10 && last.PM10.value),
    no2: parseNum_(last.NO2 && last.NO2.value),
    o3: parseNum_(last.O3 && last.O3.value),
    co: parseNum_(last.CO && last.CO.value),
    so2: parseNum_(last.SO2 && last.SO2.value),
    aqi: aqi !== null ? aqi : pm25Aqi,
    aqiParam: String((last.AQI && last.AQI.param) || 'PM25'),
    level: classifyPM25_(pm25)
  };
}

/* ==================== Google Sheets Real Database ==================== */

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#0f766e')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

function isDuplicateHourly_(sheet, dateStr, timeStr, stationID) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;

  const count = Math.min(72, lastRow - 1);
  const rows = sheet.getRange(lastRow - count + 1, 2, count, 3).getValues();
  for (let i = rows.length - 1; i >= 0; i--) {
    if (
      formatCellDate_(rows[i][0]) === dateStr &&
      formatCellTime_(rows[i][1]) === timeStr &&
      String(rows[i][2] || '').trim() === stationID
    ) {
      return true;
    }
  }
  return false;
}

function upsertDailySummary_(hourlySheet, dailySheet, targetDate) {
  const lastRow = hourlySheet.getLastRow();
  if (lastRow <= 1) return null;

  const rows = hourlySheet.getRange(2, 2, lastRow - 1, 8).getValues();
  let sum = 0;
  let count = 0;
  let minVal = Infinity;
  let maxVal = -Infinity;
  let peakTime = '';
  let maxAqi = null;
  let stationID = '';
  let stationName = '';

  for (let i = 0; i < rows.length; i++) {
    if (formatCellDate_(rows[i][0]) !== targetDate) continue;
    const pm25 = parseNum_(rows[i][6]);
    if (pm25 === null) continue;

    const t = formatCellTime_(rows[i][1]);
    const aqi = parseNum_(rows[i][7]);
    stationID = String(rows[i][2] || stationID);
    stationName = String(rows[i][3] || stationName);

    sum += pm25;
    count++;
    if (pm25 < minVal) minVal = pm25;
    if (pm25 > maxVal) {
      maxVal = pm25;
      peakTime = t;
    }
    if (aqi !== null && (maxAqi === null || aqi > maxAqi)) {
      maxAqi = aqi;
    }
  }

  if (count === 0) return null;

  const avg = round1_(sum / count);
  const lv = classifyPM25_(avg);
  const updatedAt = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

  const rowValues = [
    targetDate,
    stationID,
    stationName,
    avg,
    round1_(minVal),
    round1_(maxVal),
    peakTime,
    maxAqi !== null ? maxAqi : '',
    count,
    lv.level,
    lv.label,
    updatedAt
  ];

  const dLast = dailySheet.getLastRow();
  let foundRow = -1;
  if (dLast > 1) {
    const dDates = dailySheet.getRange(2, 1, dLast - 1, 1).getValues();
    for (let j = 0; j < dDates.length; j++) {
      if (formatCellDate_(dDates[j][0]) === targetDate) {
        foundRow = j + 2;
        break;
      }
    }
  }

  const targetRow = foundRow !== -1 ? foundRow : dLast + 1;
  dailySheet.getRange(targetRow, 1).setNumberFormat('@');
  dailySheet.getRange(targetRow, 7).setNumberFormat('@');
  dailySheet.getRange(targetRow, 12).setNumberFormat('@');
  dailySheet.getRange(targetRow, 1, 1, DAILY_HEADERS.length).setValues([rowValues]);

  return {
    date: targetDate,
    stationID: stationID,
    stationName: stationName,
    avg: avg,
    min: round1_(minVal),
    max: round1_(maxVal),
    peakTime: peakTime,
    aqiMax: maxAqi,
    hoursCount: count,
    level: lv
  };
}

/**
 * อ่านประวัติจริงจากชีต Hourly_Logs และ Daily_Summary เท่านั้น
 * จัดกลุ่มสำหรับกราฟ 4 มุมมอง: รายวัน (daily), รายสัปดาห์ (weekly), รายเดือน (monthly), รายปี (yearly)
 */
function readRealHistoryFromSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hourlyAll = [];
  const dailyAll = [];

  if (ss) {
    const hSheet = ss.getSheetByName(CONFIG.HOURLY_SHEET);
    if (hSheet && hSheet.getLastRow() > 1) {
      const hRows = hSheet.getRange(2, 1, hSheet.getLastRow() - 1, HOURLY_HEADERS.length).getValues();
      hRows.forEach(function (r) {
        const d = formatCellDate_(r[1]);
        const t = formatCellTime_(r[2]);
        const pm25 = parseNum_(r[7]);
        if (!d || pm25 === null) return;
        hourlyAll.push({
          savedAt: String(r[0] || ''),
          date: d,
          time: t,
          datetime: d + ' ' + t,
          stationID: String(r[3] || ''),
          stationName: String(r[4] || ''),
          area: String(r[5] || ''),
          isFallback: String(r[6] || '') === 'YES',
          pm25: pm25,
          aqi: parseNum_(r[8]),
          aqiParam: String(r[9] || ''),
          pm10: parseNum_(r[10]),
          no2: parseNum_(r[11]),
          level: classifyPM25_(pm25)
        });
      });
    }

    const dSheet = ss.getSheetByName(CONFIG.DAILY_SHEET);
    if (dSheet && dSheet.getLastRow() > 1) {
      const dRows = dSheet.getRange(2, 1, dSheet.getLastRow() - 1, DAILY_HEADERS.length).getValues();
      dRows.forEach(function (r) {
        const d = formatCellDate_(r[0]);
        const avg = parseNum_(r[3]);
        if (!d || avg === null) return;
        dailyAll.push({
          date: d,
          stationID: String(r[1] || ''),
          stationName: String(r[2] || ''),
          avg: avg,
          min: parseNum_(r[4]),
          max: parseNum_(r[5]),
          peakTime: formatCellTime_(r[6]),
          aqiMax: parseNum_(r[7]),
          hoursCount: parseInt(r[8], 10) || 0,
          level: classifyPM25_(avg)
        });
      });
    }
  }

  hourlyAll.sort(function (a, b) {
    return a.datetime.localeCompare(b.datetime);
  });
  dailyAll.sort(function (a, b) {
    return a.date.localeCompare(b.date);
  });

  const now = new Date();
  const cutoff7 = Utilities.formatDate(new Date(now.getTime() - 6 * 86400000), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const cutoff30 = Utilities.formatDate(new Date(now.getTime() - 29 * 86400000), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const cutoffYear = Utilities.formatDate(new Date(now.getTime() - 365 * 86400000), CONFIG.TIMEZONE, 'yyyy-MM-dd');

  // 1. รายวัน (daily): จุดข้อมูลรายชั่วโมงจริงของวันล่าสุดที่มีใน Google Sheets (สูงสุด 24 ชม. ล่าสุด)
  const latestDate = hourlyAll.length > 0 ? hourlyAll[hourlyAll.length - 1].date : null;
  const dailyHourly = latestDate
    ? hourlyAll.filter(function (h) { return h.date === latestDate; })
    : [];

  // 2. รายสัปดาห์ (weekly): ค่าเฉลี่ยรายวันจริงย้อนหลังไม่เกิน 7 วันที่มีใน Daily_Summary
  const weekly = dailyAll.filter(function (d) { return d.date >= cutoff7; });

  // 3. รายเดือน (monthly): ค่าเฉลี่ยรายวันจริงย้อนหลังไม่เกิน 30 วันที่มีใน Daily_Summary
  const monthly = dailyAll.filter(function (d) { return d.date >= cutoff30; });

  // 4. รายปี (yearly): ค่าเฉลี่ยรายเดือนจริง (YYYY-MM) คำนวณจาก Daily_Summary ที่มีจริงในรอบ 12 เดือน
  const monthMap = {};
  dailyAll.forEach(function (d) {
    if (d.date < cutoffYear) return;
    const ym = d.date.substring(0, 7);
    if (!monthMap[ym]) {
      monthMap[ym] = { month: ym, sum: 0, count: 0, min: Infinity, max: -Infinity, daysCount: 0 };
    }
    const m = monthMap[ym];
    m.sum += d.avg;
    m.count += 1;
    m.daysCount += 1;
    if (d.min !== null && d.min < m.min) m.min = d.min;
    if (d.max !== null && d.max > m.max) m.max = d.max;
  });

  const yearly = Object.keys(monthMap).sort().map(function (ym) {
    const m = monthMap[ym];
    const avg = round1_(m.sum / m.count);
    return {
      month: ym,
      avg: avg,
      min: m.min === Infinity ? avg : round1_(m.min),
      max: m.max === -Infinity ? avg : round1_(m.max),
      daysCount: m.daysCount,
      level: classifyPM25_(avg)
    };
  });

  return {
    stats: {
      totalHourlyRecords: hourlyAll.length,
      totalDaysRecorded: dailyAll.length,
      firstDate: dailyAll.length > 0 ? dailyAll[0].date : (hourlyAll.length > 0 ? hourlyAll[0].date : null),
      lastDate: dailyAll.length > 0 ? dailyAll[dailyAll.length - 1].date : latestDate
    },
    daily: dailyHourly,
    weekly: weekly,
    monthly: monthly,
    yearly: yearly,
    hourlyLogs: hourlyAll.slice(-168).reverse(),
    dailySummary: dailyAll.slice(-365).reverse()
  };
}

/* ==================== Helpers & PCD 5-Level Standard ==================== */

function classifyPM25_(pm25) {
  if (pm25 === null || pm25 === undefined || isNaN(pm25) || pm25 < 0) {
    return { level: 0, key: 'unknown', colorName: 'เทา', hex: '#64748b', label: 'ไม่มีข้อมูล' };
  }
  if (pm25 <= 15.0) return { level: 1, key: 'blue',   colorName: 'สีฟ้า',    hex: '#0284c7', label: 'คุณภาพอากาศดีมาก' };
  if (pm25 <= 25.0) return { level: 2, key: 'green',  colorName: 'สีเขียว',  hex: '#16a34a', label: 'คุณภาพอากาศดี' };
  if (pm25 <= 37.5) return { level: 3, key: 'yellow', colorName: 'สีเหลือง', hex: '#ca8a04', label: 'ปานกลาง' };
  if (pm25 <= 75.0) return { level: 4, key: 'orange', colorName: 'สีส้ม',    hex: '#ea580c', label: 'เริ่มมีผลกระทบต่อสุขภาพ' };
  return { level: 5, key: 'red', colorName: 'สีแดง', hex: '#dc2626', label: 'มีผลกระทบต่อสุขภาพ' };
}

function parseNum_(val) {
  if (val === null || val === undefined) return null;
  const s = String(val).trim();
  if (s === '' || s === '-' || s.toUpperCase() === 'N/A') return null;
  const n = Number(s);
  return (isNaN(n) || n < 0) ? null : n;
}

function computeAgeHours_(dateStr, timeStr, now) {
  if (!dateStr || !timeStr) return null;
  const t = timeStr.length === 5 ? timeStr + ':00' : timeStr;
  const dt = new Date(dateStr + 'T' + t + '+07:00');
  if (isNaN(dt.getTime())) return null;
  return Math.max(0, (now.getTime() - dt.getTime()) / 3600000);
}

function haversineKm_(lat1, lon1, lat2, lon2) {
  const R = 6371;
  const r = function (d) { return d * Math.PI / 180; };
  const dLat = r(lat2 - lat1);
  const dLon = r(lon2 - lon1);
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2)
    + Math.cos(r(lat1)) * Math.cos(r(lat2)) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
  return 2 * R * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function round1_(n) {
  return Math.round(n * 10) / 10;
}

function formatCellDate_(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  }
  return String(val || '').trim().substring(0, 10);
}

function formatCellTime_(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, CONFIG.TIMEZONE, 'HH:mm');
  }
  return String(val || '').trim().substring(0, 5);
}

/**
 * ป้องกัน Spreadsheet Formula Injection (=, +, -, @) ก่อนบันทึกลง Google Sheets
 */
function sanitizeSheetCell_(val) {
  const s = String(val === null || val === undefined ? '' : val)
    .replace(/[\r\n]+/g, ' ')
    .trim()
    .substring(0, 200);
  if (/^[=+\-@\t\r]/.test(s)) {
    return "'" + s;
  }
  return s;
}
