const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;
const HOST = '0.0.0.0';
const TIMEZONE = 'Asia/Hong_Kong';
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'bus2026';

// GitHub 自動備份設定（免費方案持久化儲存）
const GITHUB_TOKEN = (process.env.GITHUB_TOKEN || '').trim();
const GITHUB_REPO = (process.env.GITHUB_REPO || '').trim();
const GITHUB_BRANCH = (process.env.GITHUB_BRANCH || 'main').trim();
const GITHUB_BACKUP_INTERVAL_MS = 3 * 60 * 1000;

const KMB_API = 'https://data.etabus.gov.hk/v1/transport/kmb';
const CTB_API = 'https://rt.data.gov.hk/v1/transport/citybus-nwfb';

async function safeFetchJson(url, retries = 2) {
    for (let i = 0; i <= retries; i++) {
        try {
            const res = await fetch(url, {
                headers: {
                    'Accept': 'application/json',
                    'User-Agent': 'HK-Bus-Monitor/1.0'
                }
            });
            const text = await res.text();
            if (!res.ok || text.trim().startsWith('<')) {
                throw new Error(`HTTP ${res.status} (收到 HTML 非 JSON 回應)`);
            }
            return JSON.parse(text);
        } catch (err) {
            if (i === retries) return null;
            await new Promise(r => setTimeout(r, 800 * (i + 1)));
        }
    }
    return null;
}

const DATA_DIR = process.env.DATA_DIR || __dirname;
if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
}

const BOOKMARKS_FILE = path.join(DATA_DIR, 'bookmarks.json');
const SHAPES_FILE = path.join(__dirname, 'bus_shapes.json');

const ROLLING_48H_EXCEL_FILENAME = 'bus_arrival_records_48h_rolling.xlsx';
const ROLLING_48H_EXCEL_FILE = path.join(DATA_DIR, ROLLING_48H_EXCEL_FILENAME);
const ROLLING_48H_CSV_FILENAME = 'bus_arrival_records_48h_rolling.csv';
const ROLLING_48H_CSV_FILE = path.join(DATA_DIR, ROLLING_48H_CSV_FILENAME);
const ROLLING_48H_TRIPS_FILE = path.join(DATA_DIR, 'trip_records_48h_rolling.json');

const RETENTION_MS = 48 * 3600 * 1000;

function getHkTimeParts(ts = Date.now()) {
    const formatter = new Intl.DateTimeFormat('en-GB', {
        timeZone: TIMEZONE,
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false
    });
    const parts = {};
    formatter.formatToParts(new Date(ts)).forEach(p => {
        if (p.type !== 'literal') parts[p.type] = p.value;
    });
    return parts;
}

function formatTime(ts) {
    if (!ts) return '';
    const p = getHkTimeParts(ts);
    return `${p.hour}:${p.minute}:${p.second}`;
}

function formatDate(ts) {
    const p = getHkTimeParts(ts);
    return `${p.year}-${p.month}-${p.day}`;
}

function getStartupSessionTag() {
    const p = getHkTimeParts(Date.now());
    const yy = p.year.slice(-2);
    return `${p.day}${p.month}${yy}_${p.hour}${p.minute}${p.second}`;
}

const SESSION_TAG = getStartupSessionTag();
const SESSION_EXCEL_FILENAME = `bus_arrival_records_${SESSION_TAG}.xlsx`;
const SESSION_EXCEL_FILE = path.join(DATA_DIR, SESSION_EXCEL_FILENAME);

app.use(cors());
app.use(express.json());
app.use(express.static(__dirname));

let busShapesFeatures = [];
if (fs.existsSync(SHAPES_FILE)) {
    try {
        const parsed = JSON.parse(fs.readFileSync(SHAPES_FILE, 'utf8'));
        busShapesFeatures = parsed.features || [];
        console.log(`🗺️ 已載入 ${busShapesFeatures.length} 條官方道路軌跡`);
    } catch (e) {
        console.error('⚠️ 載入 bus_shapes.json 失敗:', e.message);
    }
}

function loadInitialBookmarks() {
    if (fs.existsSync(BOOKMARKS_FILE)) {
        try {
            return JSON.parse(fs.readFileSync(BOOKMARKS_FILE, 'utf8'));
        } catch (e) { /* ignore */ }
    }
    const envRaw = (process.env.DEFAULT_BOOKMARKS || '').trim();
    if (!envRaw) return [];

    const nowStr = `${formatDate(Date.now())} ${formatTime(Date.now())}`;
    return envRaw.split(',').map(entry => {
        const [op, route, dirParam = 'outbound', serviceType = '1'] = entry.trim().split(':');
        if (!op || !route) return null;
        const operator = op.toUpperCase();
        return {
            id: `${operator}_${route.toUpperCase()}_${dirParam}_${serviceType}`,
            operator,
            route: route.toUpperCase(),
            dirParam,
            serviceType,
            origTc: '(雲端預設)',
            destTc: '(雲端預設)',
            origEn: '(Cloud Default)',
            destEn: '(Cloud Default)',
            addedAt: nowStr
        };
    }).filter(Boolean);
}

function loadInitialTrips() {
    const cutoff = Date.now() - RETENTION_MS;
    const loaded = {};

    if (fs.existsSync(ROLLING_48H_TRIPS_FILE)) {
        try {
            const parsed = JSON.parse(fs.readFileSync(ROLLING_48H_TRIPS_FILE, 'utf8'));
            Object.entries(parsed).forEach(([id, t]) => {
                if (t && t.departureTs >= cutoff) {
                    loaded[id] = t;
                }
            });
        } catch (e) {
            console.error('⚠️ 讀取 48 小時歷史班次失敗:', e.message);
        }
    }
    return loaded;
}

let bookmarks = loadInitialBookmarks();
let trips = loadInitialTrips();
let activeStopTrackers = {};

let routeStopsCache = {};
let kmbStopsMap = {};
let ctbStopsMap = {};
let isPolling = false;
let drainMode = false;
let lastPollTimeStr = '尚未執行';
let lastGithubBackupTs = 0;
let lastGithubBackupError = '';
let isBackingUpGithub = false;

function saveBookmarks() {
    fs.writeFileSync(BOOKMARKS_FILE, JSON.stringify(bookmarks, null, 2));
}

// ==========================================
// GitHub 自動備份與啟動還原模組
// ==========================================
async function fetchFileFromGithub(repoFilePath) {
    if (!GITHUB_TOKEN || !GITHUB_REPO) return null;
    const apiUrl = `https://api.github.com/repos/${GITHUB_REPO}/contents/${repoFilePath}?ref=${encodeURIComponent(GITHUB_BRANCH)}`;
    try {
        const res = await fetch(apiUrl, {
            headers: {
                'Authorization': `Bearer ${GITHUB_TOKEN}`,
                'Accept': 'application/vnd.github+json',
                'User-Agent': 'bus-monitor-app'
            }
        });
        if (!res.ok) {
            const errText = await res.text();
            console.warn(`⚠️ GitHub 讀取 [${repoFilePath}] HTTP ${res.status}: ${errText}`);
            return null;
        }
        const data = await res.json();
        if (!data.content) return null;
        return Buffer.from(data.content, 'base64').toString('utf8');
    } catch (err) {
        console.warn(`⚠️ GitHub 讀取連線異常 [${repoFilePath}]:`, err.message);
        return null;
    }
}

async function uploadFileToGithub(localFilePath, repoFilePath) {
    if (!GITHUB_TOKEN || !GITHUB_REPO) {
        lastGithubBackupError = '未設定 GITHUB_TOKEN 或 GITHUB_REPO';
        return false;
    }
    if (!fs.existsSync(localFilePath)) {
        lastGithubBackupError = `找不到本地檔案：${localFilePath}`;
        console.warn(`⚠️ ${lastGithubBackupError}`);
        return false;
    }

    const contentBase64 = fs.readFileSync(localFilePath).toString('base64');
    const apiUrl = `https://api.github.com/repos/${GITHUB_REPO}/contents/${repoFilePath}`;
    const headers = {
        'Authorization': `Bearer ${GITHUB_TOKEN}`,
        'Accept': 'application/vnd.github+json',
        'Content-Type': 'application/json',
        'User-Agent': 'bus-monitor-app'
    };

    try {
        let sha;
        const getRes = await fetch(`${apiUrl}?ref=${encodeURIComponent(GITHUB_BRANCH)}`, { headers });
        if (getRes.ok) {
            const existing = await getRes.json();
            sha = existing.sha;
        }

        const putRes = await fetch(apiUrl, {
            method: 'PUT',
            headers,
            body: JSON.stringify({
                message: `Auto backup ${repoFilePath} (${formatTime(Date.now())})`,
                content: contentBase64,
                branch: GITHUB_BRANCH,
                ...(sha ? { sha } : {})
            })
        });

        if (!putRes.ok) {
            const errText = await putRes.text();
            lastGithubBackupError = `HTTP ${putRes.status} (${repoFilePath}): ${errText}`;
            console.error(`❌ GitHub 備份失敗 [${repoFilePath}] HTTP ${putRes.status}: ${errText}`);
            return false;
        }

        lastGithubBackupError = '';
        return true;
    } catch (err) {
        lastGithubBackupError = `連線異常 (${repoFilePath}): ${err.message}`;
        console.error(`❌ GitHub 備份連線異常 [${repoFilePath}]:`, err.message);
        return false;
    }
}

async function restoreFromGithub() {
    if (!GITHUB_TOKEN || !GITHUB_REPO) return;
    console.log(`☁️ 正在嘗試從 GitHub (${GITHUB_REPO} @ ${GITHUB_BRANCH}) 還原歷史備份...`);
    const bmRaw = await fetchFileFromGithub('backups/bookmarks.json');
    if (bmRaw) {
        try {
            const parsedBm = JSON.parse(bmRaw);
            if (Array.isArray(parsedBm) && parsedBm.length > 0) {
                bookmarks = parsedBm;
                saveBookmarks();
                console.log(`✅ 已從 GitHub 還原 ${bookmarks.length} 條監察路線`);
            }
        } catch (e) { /* ignore */ }
    }

    const tripsRaw = await fetchFileFromGithub('backups/trip_records_48h_rolling.json');
    if (tripsRaw) {
        try {
            const parsedTrips = JSON.parse(tripsRaw);
            if (parsedTrips && typeof parsedTrips === 'object') {
                trips = { ...parsedTrips, ...trips };
                pruneOldTrips(Date.now());
                console.log(`✅ 已從 GitHub 還原 ${Object.keys(trips).length} 筆 48 小時內班次記錄`);
            }
        } catch (e) { /* ignore */ }
    }
}

async function triggerGithubBackup(force = false) {
    if (!GITHUB_TOKEN || !GITHUB_REPO || isBackingUpGithub) return;
    const now = Date.now();
    if (!force && now - lastGithubBackupTs < GITHUB_BACKUP_INTERVAL_MS) return;

    isBackingUpGithub = true;
    try {
        const ok1 = await uploadFileToGithub(BOOKMARKS_FILE, 'backups/bookmarks.json');
        const ok2 = await uploadFileToGithub(ROLLING_48H_TRIPS_FILE, 'backups/trip_records_48h_rolling.json');
        const ok3 = await uploadFileToGithub(ROLLING_48H_EXCEL_FILE, `backups/${ROLLING_48H_EXCEL_FILENAME}`);
        const ok4 = await uploadFileToGithub(ROLLING_48H_CSV_FILE, `backups/${ROLLING_48H_CSV_FILENAME}`);
        if (ok1 && ok2 && ok3 && ok4) {
            lastGithubBackupTs = now;
            console.log(`☁️ [${formatTime(now)}] 已自動備份 Excel & CSV 記錄至 GitHub (${GITHUB_REPO} @ ${GITHUB_BRANCH})`);
        }
    } finally {
        isBackingUpGithub = false;
    }
}

async function resolveDefaultBookmarkNames() {
    let updated = false;
    for (const bm of bookmarks) {
        if (bm.origTc !== '(雲端預設)' && bm.destTc !== '(雲端預設)' && bm.origEn && bm.destEn) continue;
        try {
            if (bm.operator === 'KMB') {
                const json = await safeFetchJson(`${KMB_API}/route/${bm.route}/${bm.dirParam}/${bm.serviceType || '1'}`);
                if (json && json.data) {
                    bm.origTc = (json.data.orig_tc || '').trim() || bm.origTc;
                    bm.destTc = (json.data.dest_tc || '').trim() || bm.destTc;
                    bm.origEn = (json.data.orig_en || '').trim() || bm.origTc;
                    bm.destEn = (json.data.dest_en || '').trim() || bm.destTc;
                    updated = true;
                }
            } else {
                const json = await safeFetchJson(`${CTB_API}/route/ctb/${bm.route}`);
                if (json && json.data) {
                    const origTc = (json.data.orig_tc || '').trim();
                    const destTc = (json.data.dest_tc || '').trim();
                    const origEn = (json.data.orig_en || origTc).trim();
                    const destEn = (json.data.dest_en || destTc).trim();
                    bm.origTc = bm.dirParam === 'outbound' ? origTc : destTc;
                    bm.destTc = bm.dirParam === 'outbound' ? destTc : origTc;
                    bm.origEn = bm.dirParam === 'outbound' ? origEn : destEn;
                    bm.destEn = bm.dirParam === 'outbound' ? destEn : origEn;
                    updated = true;
                }
            }
        } catch (e) { /* ignore */ }
    }
    if (updated) {
        saveBookmarks();
        saveToExcel();
    }
}

function pruneOldTrips(nowTs) {
    const cutoff = nowTs - RETENTION_MS;
    Object.keys(trips).forEach(id => {
        if (trips[id].departureTs < cutoff) {
            delete trips[id];
        }
    });
}

function isHubOrExpresswayStop(stopName = '') {
    return /轉車站|轉乘站|西隧|紅隧|東隧|大橋|口岸/.test(stopName);
}

function isArrivalPhysicallyValid(trip, seq, arrivalTs) {
    if (seq > 1 && arrivalTs < trip.departureTs + 15000) return false;

    const recordedSeqs = Object.keys(trip.stops)
        .map(Number)
        .filter(s => s !== seq && trip.stops[s].actualArrivalTs !== null);

    const prevSeqs = recordedSeqs.filter(s => s < seq);
    if (prevSeqs.length > 0) {
        const maxPrevSeq = Math.max(...prevSeqs);
        const prevTs = trip.stops[maxPrevSeq].actualArrivalTs;
        const seqGap = seq - maxPrevSeq;
        const diffSecs = (arrivalTs - prevTs) / 1000;

        const minAllowedSecs = seqGap === 1 ? -60 : (seqGap - 1) * 25;
        if (diffSecs < minAllowedSecs) return false;
    }

    const nextSeqs = recordedSeqs.filter(s => s > seq);
    if (nextSeqs.length > 0) {
        const minNextSeq = Math.min(...nextSeqs);
        const nextTs = trip.stops[minNextSeq].actualArrivalTs;
        const seqGap = minNextSeq - seq;
        const diffSecs = (nextTs - arrivalTs) / 1000;

        const minAllowedSecs = seqGap === 1 ? -60 : (seqGap - 1) * 25;
        if (diffSecs < minAllowedSecs) return false;
    }

    return true;
}

function getRouteHeadwayMins(routeKeyPrefix) {
    const formalTrips = Object.values(trips)
        .filter(t => t.routeKeyPrefix === routeKeyPrefix && !t.isEstimatedDep)
        .sort((a, b) => b.departureTs - a.departureTs)
        .slice(0, 6);

    if (formalTrips.length < 2) return 20;

    const diffs = [];
    for (let i = 0; i < formalTrips.length - 1; i++) {
        const gap = (formalTrips[i].departureTs - formalTrips[i + 1].departureTs) / 60000;
        if (gap >= 2 && gap <= 60) diffs.push(gap);
    }
    if (diffs.length === 0) return 20;

    diffs.sort((a, b) => a - b);
    return diffs[Math.floor(diffs.length / 2)];
}

function getTripCurrentDelayMins(trip) {
    const arrivedWithSched = Object.values(trip.stops)
        .filter(st => st.actualArrivalTs !== null && st.scheduledEtaTs !== null)
        .sort((a, b) => b.seq - a.seq);

    if (arrivedWithSched.length === 0) return 0;

    const recent = arrivedWithSched.slice(0, 3);
    const sumDelay = recent.reduce((acc, st) => acc + (st.actualArrivalTs - st.scheduledEtaTs) / 60000, 0);
    return sumDelay / recent.length;
}

function isTripStillActive(trip, nowTs) {
    if (nowTs - trip.departureTs > 140 * 60000) return false;
    const seqs = Object.keys(trip.stops).map(Number);
    if (seqs.length === 0) return false;
    const maxSeq = Math.max(...seqs);
    const lastStop = trip.stops[maxSeq];
    if (lastStop && lastStop.actualArrivalTs !== null) return false;
    return true;
}

function stitchFragmentedTrips() {
    const tripIds = Object.keys(trips);
    const formalTrips = tripIds.map(id => trips[id]).filter(t => !t.isEstimatedDep);
    const estTrips = tripIds.map(id => trips[id]).filter(t => t.isEstimatedDep);

    for (const est of estTrips) {
        const estArrivedSeqs = Object.keys(est.stops)
            .map(Number)
            .filter(s => est.stops[s].actualArrivalTs !== null)
            .sort((a, b) => a - b);

        if (estArrivedSeqs.length === 0) continue;

        const headwayMins = getRouteHeadwayMins(est.routeKeyPrefix);
        const maxStitchWindowMins = Math.min(25, Math.max(5, headwayMins * 0.65));

        const candidates = formalTrips.filter(f =>
            f.routeKeyPrefix === est.routeKeyPrefix &&
            Math.abs(f.departureTs - est.departureTs) <= maxStitchWindowMins * 60000
        );

        for (const formal of candidates) {
            let canMerge = true;

            for (const seq of estArrivedSeqs) {
                const estTs = est.stops[seq].actualArrivalTs;

                if (formal.stops[seq]?.actualArrivalTs) {
                    if (Math.abs(formal.stops[seq].actualArrivalTs - estTs) > 90000) {
                        canMerge = false;
                        break;
                    }
                }

                if (!isArrivalPhysicallyValid(formal, seq, estTs)) {
                    canMerge = false;
                    break;
                }
            }

            if (canMerge) {
                for (const seq of estArrivedSeqs) {
                    if (!formal.stops[seq].actualArrivalTs) {
                        const arrTs = est.stops[seq].actualArrivalTs;
                        formal.stops[seq] = {
                            ...est.stops[seq],
                            scheduledEtaTs: formal.stops[seq].scheduledEtaTs || est.stops[seq].scheduledEtaTs,
                            cumulativeMins: Math.max(0, Math.round((arrTs - formal.departureTs) / 60000)),
                            status: '已到站'
                        };
                    }
                }
                delete trips[est.tripId];
                break;
            }
        }
    }
}

function saveToExcel() {
    const now = Date.now();
    pruneOldTrips(now);

    const sortedTrips = Object.values(trips).sort((a, b) => b.departureTs - a.departureTs);

    const detailRows = [];
    sortedTrips.forEach(trip => {
        const stopList = Object.values(trip.stops).sort((a, b) => a.seq - b.seq);
        stopList.forEach(st => {
            detailRows.push({
                '日期': trip.date,
                '營辦商': trip.operator === 'KMB' ? '九巴 KMB' : '城巴 CTB',
                '路線': trip.route,
                '方向': trip.dirParam === 'outbound' ? '去程' : '回程',
                '班次識別(總站開出時間)': trip.departureTimeStr + (trip.isEstimatedDep ? ' (途中估算)' : ''),
                '站序': st.seq,
                '車站名稱': st.stopName,
                '車站英文名稱': st.stopNameEn || '',
                '最初預計時間': st.scheduledEtaTs ? formatTime(st.scheduledEtaTs) : '',
                '實際到達時間': st.actualArrivalTs ? formatTime(st.actualArrivalTs) : '行駛中...',
                '由總站起累積時間(分鐘)': st.cumulativeMins !== null ? st.cumulativeMins : '',
                '狀態': st.status
            });
        });
    });

    const matrixRows = sortedTrips.map(trip => {
        const row = {
            '日期': trip.date,
            '路線': `${trip.route} (${trip.operator === 'KMB' ? '九巴' : '城巴'})`,
            '方向': trip.dirParam === 'outbound' ? '去程' : '回程',
            '班次(總站開出時間)': trip.departureTimeStr + (trip.isEstimatedDep ? ' (途中估算)' : '')
        };
        const stopList = Object.values(trip.stops).sort((a, b) => a.seq - b.seq);
        stopList.forEach(st => {
            const colName = `第${st.seq}站_${st.stopName}`;
            row[colName] = st.actualArrivalTs
                ? `${formatTime(st.actualArrivalTs)} (+${st.cumulativeMins}分)`
                : (st.scheduledEtaTs ? `預計 ${formatTime(st.scheduledEtaTs)}` : '');
        });
        return row;
    });

    const wb = XLSX.utils.book_new();

    const wsDetail = XLSX.utils.json_to_sheet(
        detailRows.length > 0 ? detailRows : [{ '提示': '尚無監察記錄，請在網頁將路線加入 Bookmark' }]
    );
    wsDetail['!cols'] = [
        { wch: 12 }, { wch: 10 }, { wch: 8 }, { wch: 8 }, { wch: 24 },
        { wch: 6 }, { wch: 26 }, { wch: 28 }, { wch: 18 }, { wch: 16 }, { wch: 22 }, { wch: 16 }
    ];
    XLSX.utils.book_append_sheet(wb, wsDetail, '班次行程及累積時間');

    const wsMatrix = XLSX.utils.json_to_sheet(
        matrixRows.length > 0 ? matrixRows : [{ '提示': '尚無班次矩陣資料' }]
    );
    XLSX.utils.book_append_sheet(wb, wsMatrix, '班次累積時間矩陣');

    const bmRows = bookmarks.map(b => ({
        '營辦商': b.operator === 'KMB' ? '九巴 KMB' : '城巴 CTB',
        '路線': b.route,
        '方向': b.dirParam === 'outbound' ? '去程' : '回程',
        '起點': b.origTc || '',
        '終點': b.destTc || '',
        '加入監察時間': b.addedAt
    }));
    const wsBm = XLSX.utils.json_to_sheet(bmRows.length > 0 ? bmRows : [{ '提示': '尚無收藏路線' }]);
    XLSX.utils.book_append_sheet(wb, wsBm, '已Bookmark路線');

    try {
        XLSX.writeFile(wb, ROLLING_48H_EXCEL_FILE);
    } catch (err) {
        console.error('⚠️ 寫入 48 小時滾動 Excel 失敗:', err.message);
    }

    try {
        const csvContent = '\uFEFF' + XLSX.utils.sheet_to_csv(wsDetail);
        fs.writeFileSync(ROLLING_48H_CSV_FILE, csvContent, 'utf8');
    } catch (err) {
        console.error('⚠️ 寫入 48 小時 CSV 失敗:', err.message);
    }

    try {
        XLSX.writeFile(wb, SESSION_EXCEL_FILE);
    } catch (err) {
        console.error('⚠️ 寫入 Session Excel 失敗:', err.message);
    }

    try {
        fs.writeFileSync(ROLLING_48H_TRIPS_FILE, JSON.stringify(trips, null, 2));
    } catch (err) {
        console.error('⚠️ 寫入 48 小時 JSON 失敗:', err.message);
    }
}

async function ensureKmbStops() {
    if (Object.keys(kmbStopsMap).length > 0) return;
    const json = await safeFetchJson(`${KMB_API}/stop`);
    if (!json || !json.data) return;
    json.data.forEach(s => {
        const tc = (s.name_tc || s.name_en || s.stop).trim();
        const en = (s.name_en || s.name_tc || s.stop).trim();
        kmbStopsMap[s.stop] = { tc, en };
    });
}

async function getRouteStops(operator, route, dirParam, serviceType = '1') {
    const key = `${operator}_${route}_${dirParam}_${serviceType}`;
    if (routeStopsCache[key]) return routeStopsCache[key];

    if (operator === 'KMB') {
        await ensureKmbStops();
        const json = await safeFetchJson(`${KMB_API}/route-stop/${route}/${dirParam}/${serviceType}`);
        if (!json || !json.data) return [];
        const list = json.data.map(item => {
            const info = kmbStopsMap[item.stop] || { tc: item.stop, en: item.stop };
            return {
                seq: parseInt(item.seq, 10),
                stopId: item.stop,
                stopName: info.tc,
                stopNameEn: info.en
            };
        });
        routeStopsCache[key] = list;
        return list;
    } else {
        const json = await safeFetchJson(`${CTB_API}/route-stop/ctb/${route}/${dirParam}`);
        if (!json || !json.data) return [];
        const list = await Promise.all(json.data.map(async item => {
            if (!ctbStopsMap[item.stop]) {
                const sJson = await safeFetchJson(`${CTB_API}/stop/${item.stop}`);
                const tc = (sJson?.data?.name_tc || sJson?.data?.name_en || item.stop).trim();
                const en = (sJson?.data?.name_en || sJson?.data?.name_tc || item.stop).trim();
                ctbStopsMap[item.stop] = { tc, en };
            }
            return {
                seq: parseInt(item.seq, 10),
                stopId: item.stop,
                stopName: ctbStopsMap[item.stop].tc,
                stopNameEn: ctbStopsMap[item.stop].en
            };
        }));
        routeStopsCache[key] = list;
        return list;
    }
}

function isEtaConsistentWithTrip(trip, seq, etaTs) {
    const diffFromDepMins = (etaTs - trip.departureTs) / 60000;

    const minFromDepMins = seq <= 4 ? (seq - 1) * 0.2 : (seq - 1) * 0.45;
    const maxFromDepMins = seq * 8.5;
    if (diffFromDepMins < minFromDepMins || diffFromDepMins > maxFromDepMins) {
        return false;
    }

    const existingSched = trip.stops[seq]?.scheduledEtaTs;
    if (existingSched) {
        const currentDelayMins = getTripCurrentDelayMins(trip);
        const expectedTs = existingSched + currentDelayMins * 60000;
        const headwayMins = getRouteHeadwayMins(trip.routeKeyPrefix);
        const maxEtaDriftMins = Math.min(8, Math.max(2.5, headwayMins * 0.45));
        if (Math.abs(etaTs - expectedTs) > maxEtaDriftMins * 60000) {
            return false;
        }
    }

    const prevSeqs = Object.keys(trip.stops)
        .map(Number)
        .filter(s => s < seq && (trip.stops[s].actualArrivalTs || trip.stops[s].scheduledEtaTs));

    if (prevSeqs.length > 0) {
        const latestPrevSeq = Math.max(...prevSeqs);
        const prevRefTs = trip.stops[latestPrevSeq].actualArrivalTs || trip.stops[latestPrevSeq].scheduledEtaTs;
        const stepDiffSecs = (etaTs - prevRefTs) / 1000;
        const seqGap = seq - latestPrevSeq;

        const minStepSecs = seqGap === 1 ? 10 : seqGap * 20;
        const maxStepSecs = Math.max(seqGap * 15 * 60, 35 * 60);
        if (stepDiffSecs < minStepSecs || stepDiffSecs > maxStepSecs) {
            return false;
        }
    }

    const nextSeqs = Object.keys(trip.stops)
        .map(Number)
        .filter(s => s > seq && (trip.stops[s].actualArrivalTs || trip.stops[s].scheduledEtaTs));

    if (nextSeqs.length > 0) {
        const earliestNextSeq = Math.min(...nextSeqs);
        const nextRefTs = trip.stops[earliestNextSeq].actualArrivalTs || trip.stops[earliestNextSeq].scheduledEtaTs;
        const stepDiffSecs = (nextRefTs - etaTs) / 1000;
        const seqGap = earliestNextSeq - seq;

        const minStepSecs = seqGap === 1 ? 10 : seqGap * 20;
        if (stepDiffSecs < minStepSecs) {
            return false;
        }
    }

    return true;
}

function findOrCreateTripForStopArrival(routeKeyPrefix, meta, seq, arrivalTs, scheduledEtaTs, stopsList, currentStopEtasSnapshot = {}) {
    if (seq === 1) {
        if (drainMode) return null;

        const preciseDepTs = (scheduledEtaTs && Math.abs(arrivalTs - scheduledEtaTs) <= 90000)
            ? scheduledEtaTs
            : arrivalTs;
        const depTimeStr = formatTime(preciseDepTs);
        const tripId = `${routeKeyPrefix}_DEP_${preciseDepTs}`;

        if (trips[tripId]) return trips[tripId];

        const newTrip = {
            tripId,
            routeKeyPrefix,
            date: formatDate(preciseDepTs),
            operator: meta.operator,
            route: meta.route,
            dirParam: meta.dirParam,
            destTc: meta.destTc,
            destEn: meta.destEn || meta.destTc,
            departureTs: preciseDepTs,
            departureTimeStr: depTimeStr,
            isEstimatedDep: false,
            stops: {}
        };
        stopsList.forEach(s => {
            newTrip.stops[s.seq] = {
                seq: s.seq,
                stopName: s.stopName,
                stopNameEn: s.stopNameEn || s.stopName,
                scheduledEtaTs: s.seq === 1 ? (scheduledEtaTs || preciseDepTs) : null,
                actualArrivalTs: s.seq === 1 ? preciseDepTs : null,
                cumulativeMins: s.seq === 1 ? 0 : null,
                status: s.seq === 1 ? '已由總站開出' : '待抵達'
            };
        });

        stopsList.forEach(s => {
            if (s.seq === 1) return;
            const etaCandidates = (currentStopEtasSnapshot[s.seq] || []).sort((a, b) => a.eta_seq - b.eta_seq);
            const matched = etaCandidates.find(item => isEtaConsistentWithTrip(newTrip, s.seq, item.etaTs));
            if (matched) {
                newTrip.stops[s.seq].scheduledEtaTs = matched.etaTs;
            }
        });

        Object.keys(trips).forEach(oldId => {
            const oldT = trips[oldId];
            if (oldT.routeKeyPrefix === routeKeyPrefix && oldT.isEstimatedDep && Math.abs(oldT.departureTs - preciseDepTs) <= 3.5 * 60000) {
                Object.values(oldT.stops).forEach(st => {
                    if (st.seq > 1 && st.seq <= 5 && st.actualArrivalTs) {
                        const fixedArrival = Math.max(preciseDepTs, st.actualArrivalTs);
                        if (isArrivalPhysicallyValid(newTrip, st.seq, fixedArrival)) {
                            newTrip.stops[st.seq] = {
                                ...st,
                                scheduledEtaTs: newTrip.stops[st.seq]?.scheduledEtaTs || st.scheduledEtaTs,
                                actualArrivalTs: fixedArrival,
                                cumulativeMins: Math.max(0, Math.round((fixedArrival - preciseDepTs) / 60000))
                            };
                        }
                    }
                });
                delete trips[oldId];
            }
        });

        trips[tripId] = newTrip;
        return newTrip;
    }

    const candidateTrips = Object.values(trips)
        .filter(t => t.routeKeyPrefix === routeKeyPrefix)
        .sort((a, b) => Number(a.isEstimatedDep) - Number(b.isEstimatedDep));

    const duplicateStopTrip = candidateTrips.find(t => {
        const existing = t.stops[seq];
        return existing && existing.actualArrivalTs && Math.abs(arrivalTs - existing.actualArrivalTs) <= 2.5 * 60000;
    });
    if (duplicateStopTrip) {
        return duplicateStopTrip;
    }

    const headwayMins = getRouteHeadwayMins(routeKeyPrefix);
    const allowedDriftMins = Math.min(7.5, Math.max(2.5, headwayMins * 0.45));

    let bestTrip = null;
    let smallestScore = Infinity;

    for (const trip of candidateTrips) {
        if (!isArrivalPhysicallyValid(trip, seq, arrivalTs)) continue;

        const elapsedMins = (arrivalTs - trip.departureTs) / 60000;
        if (elapsedMins > 150) continue;

        const minTotalTripMins = seq <= 4 ? (seq - 1) * 0.2 : (seq - 1) * 0.45;
        if (elapsedMins < minTotalTripMins) continue;

        const allRecordedSeqs = Object.keys(trip.stops)
            .map(Number)
            .filter(s => trip.stops[s].actualArrivalTs !== null);
        const highestSeqSoFar = allRecordedSeqs.length > 0 ? Math.max(...allRecordedSeqs) : 0;

        if (seq <= highestSeqSoFar) {
            const prevHighestTs = trip.stops[highestSeqSoFar].actualArrivalTs;
            const isRecentBackfill =
                !trip.stops[seq]?.actualArrivalTs &&
                (highestSeqSoFar - seq) <= 2 &&
                Math.abs(arrivalTs - prevHighestTs) <= 120000;

            if (!isRecentBackfill) continue;
        }

        const baselineSchedTs = trip.stops[seq]?.scheduledEtaTs;
        if (baselineSchedTs) {
            const delayMins = getTripCurrentDelayMins(trip);
            const expectedArrivalTs = baselineSchedTs + delayMins * 60000;
            const driftMins = Math.abs(arrivalTs - expectedArrivalTs) / 60000;

            if (driftMins <= allowedDriftMins) {
                const seqDistance = highestSeqSoFar > 0 ? Math.abs(seq - highestSeqSoFar) : seq;
                const score = driftMins + seqDistance * 0.15;
                if (score < smallestScore) {
                    smallestScore = score;
                    bestTrip = trip;
                }
                continue;
            }
        }

        if (highestSeqSoFar > 0) {
            const prevStop = trip.stops[highestSeqSoFar];
            const currStop = trip.stops[seq];
            const prevArrivalTs = prevStop.actualArrivalTs;
            const stepDiffSecs = (arrivalTs - prevArrivalTs) / 1000;
            const stepDiffMins = stepDiffSecs / 60;
            const seqGap = Math.abs(seq - highestSeqSoFar);

            let skippedHub = isHubOrExpresswayStop(prevStop?.stopName) || isHubOrExpresswayStop(currStop?.stopName);
            const minS = Math.min(highestSeqSoFar, seq);
            const maxS = Math.max(highestSeqSoFar, seq);
            for (let k = minS + 1; k < maxS; k++) {
                if (isHubOrExpresswayStop(trip.stops[k]?.stopName)) skippedHub = true;
            }

            const maxAllowedGap = skippedHub ? 10 : 8;
            if (seqGap > maxAllowedGap) continue;

            let minStepSecs = seq <= highestSeqSoFar ? -60 : (seqGap === 1 ? -60 : (seqGap - 1) * 25);
            let maxAllowedMins = Math.max(seqGap * 15, skippedHub ? 30 : 20);

            const currRefEta = currStop?.scheduledEtaTs || scheduledEtaTs;
            if (prevStop?.scheduledEtaTs && currRefEta && currRefEta > prevStop.scheduledEtaTs && seq > highestSeqSoFar) {
                const expectedLegMins = (currRefEta - prevStop.scheduledEtaTs) / 60000;
                minStepSecs = Math.max(minStepSecs, expectedLegMins * 60 * 0.6);
                maxAllowedMins = Math.max(maxAllowedMins, expectedLegMins + 14);
            }

            if (stepDiffSecs >= minStepSecs && stepDiffMins <= maxAllowedMins) {
                const score = 10 + Math.abs(stepDiffMins) + seqGap * 0.5;
                if (score < smallestScore) {
                    smallestScore = score;
                    bestTrip = trip;
                }
            }
        } else {
            if (seq <= 8 && elapsedMins <= seq * 10) {
                const score = 10 + elapsedMins;
                if (score < smallestScore) {
                    smallestScore = score;
                    bestTrip = trip;
                }
            }
        }
    }

    if (!bestTrip) {
        const activeFormalTrips = candidateTrips.filter(t => !t.isEstimatedDep && isTripStillActive(t, arrivalTs));
        if (activeFormalTrips.length === 1) {
            const soleActive = activeFormalTrips[0];
            if (isArrivalPhysicallyValid(soleActive, seq, arrivalTs)) {
                const recordedSeqs = Object.keys(soleActive.stops)
                    .map(Number)
                    .filter(s => soleActive.stops[s].actualArrivalTs !== null);
                const highestSeq = recordedSeqs.length > 0 ? Math.max(...recordedSeqs) : 0;

                if (seq > highestSeq || (!soleActive.stops[seq]?.actualArrivalTs && highestSeq - seq <= 2)) {
                    const prevTs = highestSeq > 0 ? soleActive.stops[highestSeq].actualArrivalTs : soleActive.departureTs;
                    const gapMins = (arrivalTs - prevTs) / 60000;
                    const seqGap = Math.max(0, seq - highestSeq);

                    if (gapMins >= (seqGap - 1) * 0.3 - 1 && gapMins <= 35) {
                        bestTrip = soleActive;
                    }
                }
            }
        }
    }

    const stopObj = stopsList.find(s => s.seq === seq);
    const sNameTc = stopObj?.stopName || `第${seq}站`;
    const sNameEn = stopObj?.stopNameEn || sNameTc;

    if (!bestTrip) {
        const nearFormalTrip = candidateTrips.find(t => {
            if (t.isEstimatedDep) return false;
            if (!isArrivalPhysicallyValid(t, seq, arrivalTs)) return false;
            return Object.values(t.stops).some(st =>
                st.actualArrivalTs !== null &&
                Math.abs(st.seq - seq) <= 3 &&
                Math.abs(st.actualArrivalTs - arrivalTs) <= 3.5 * 60000
            );
        });
        if (nearFormalTrip) {
            if (!nearFormalTrip.stops[seq]?.actualArrivalTs) {
                nearFormalTrip.stops[seq] = {
                    seq,
                    stopName: sNameTc,
                    stopNameEn: sNameEn,
                    scheduledEtaTs: nearFormalTrip.stops[seq]?.scheduledEtaTs || scheduledEtaTs || null,
                    actualArrivalTs: arrivalTs,
                    cumulativeMins: Math.max(0, Math.round((arrivalTs - nearFormalTrip.departureTs) / 60000)),
                    status: '已到站'
                };
            }
            return nearFormalTrip;
        }

        if (drainMode) return null;

        const refTs = scheduledEtaTs || arrivalTs;
        const estDepTs = refTs - (seq - 1) * 1.6 * 60000;

        const conflictsWithFormalTrip = candidateTrips.some(
            t => !t.isEstimatedDep && Math.abs(t.departureTs - estDepTs) <= Math.min(12, headwayMins * 0.5) * 60000
        );
        if (conflictsWithFormalTrip) {
            return null;
        }

        const estDepStr = formatTime(estDepTs);
        const estBucket = Math.round(estDepTs / (8 * 60000));
        const tripId = `${routeKeyPrefix}_EST_${estBucket}`;

        if (!trips[tripId]) {
            trips[tripId] = {
                tripId,
                routeKeyPrefix,
                date: formatDate(arrivalTs),
                operator: meta.operator,
                route: meta.route,
                dirParam: meta.dirParam,
                destTc: meta.destTc,
                destEn: meta.destEn || meta.destTc,
                departureTs: estDepTs,
                departureTimeStr: estDepStr,
                isEstimatedDep: true,
                stops: {}
            };
            stopsList.forEach(s => {
                trips[tripId].stops[s.seq] = {
                    seq: s.seq,
                    stopName: s.stopName,
                    stopNameEn: s.stopNameEn || s.stopName,
                    scheduledEtaTs: null,
                    actualArrivalTs: null,
                    cumulativeMins: null,
                    status: '途中切入追蹤'
                };
            });
        }
        bestTrip = trips[tripId];
    }

    if (!isArrivalPhysicallyValid(bestTrip, seq, arrivalTs)) {
        return null;
    }

    const cumMins = Math.max(0, Math.round((arrivalTs - bestTrip.departureTs) / 60000));
    const existingSched = bestTrip.stops[seq]?.scheduledEtaTs;
    const validSched = existingSched || ((scheduledEtaTs && isEtaConsistentWithTrip(bestTrip, seq, scheduledEtaTs)) ? scheduledEtaTs : null);

    bestTrip.stops[seq] = {
        seq,
        stopName: sNameTc,
        stopNameEn: sNameEn,
        scheduledEtaTs: validSched,
        actualArrivalTs: arrivalTs,
        cumulativeMins: cumMins,
        status: '已到站'
    };

    return bestTrip;
}

async function pollBookmarkedRoutes() {
    if (isPolling || bookmarks.length === 0) return;
    isPolling = true;
    const now = Date.now();

    try {
        pruneOldTrips(now);

        for (const bm of bookmarks) {
            const serviceType = bm.serviceType || '1';
            const routeKeyPrefix = `${bm.operator}_${bm.route}_${bm.dirParam}_${serviceType}`;
            const stops = await getRouteStops(bm.operator, bm.route, bm.dirParam, serviceType);
            if (stops.length === 0) continue;

            const targetDirCode = bm.dirParam === 'outbound' ? 'O' : 'I';
            const currentSeenTrackerKeys = new Set();
            const stopAllEtasMap = {};

            if (bm.operator === 'KMB') {
                const json = await safeFetchJson(`${KMB_API}/route-eta/${bm.route}/${serviceType}`);
                if (!json || !json.data) continue;
                json.data.forEach(item => {
                    if (item.dir === targetDirCode && item.eta) {
                        const seq = parseInt(item.seq, 10);
                        if (!stopAllEtasMap[seq]) stopAllEtasMap[seq] = [];
                        stopAllEtasMap[seq].push({
                            eta_seq: item.eta_seq,
                            etaTs: new Date(item.eta).getTime(),
                            rmk: (item.rmk_tc || '').trim()
                        });
                    }
                });
            } else {
                const batchSize = 6;
                for (let i = 0; i < stops.length; i += batchSize) {
                    const batch = stops.slice(i, i + batchSize);
                    await Promise.all(batch.map(async s => {
                        const json = await safeFetchJson(`${CTB_API}/eta/ctb/${s.stopId}/${bm.route}`);
                        if (!json || !json.data) return;
                        const matched = json.data
                            .filter(item => item.dir === targetDirCode && item.eta)
                            .map(item => ({
                                eta_seq: item.eta_seq,
                                etaTs: new Date(item.eta).getTime(),
                                rmk: (item.rmk_tc || '').trim()
                            }));
                        if (matched.length > 0) {
                            stopAllEtasMap[s.seq] = matched;
                        }
                    }));
                }
            }

            const sortedStops = [...stops].sort((a, b) => a.seq - b.seq);
            const activeTrips = Object.values(trips)
                .filter(t => t.routeKeyPrefix === routeKeyPrefix && !t.isEstimatedDep)
                .sort((a, b) => a.departureTs - b.departureTs);

            sortedStops.forEach(s => {
                const etaList = (stopAllEtasMap[s.seq] || []).sort((a, b) => a.eta_seq - b.eta_seq);
                const existingTrackersForStop = Object.values(activeStopTrackers)
                    .filter(tr => tr.routeKeyPrefix === routeKeyPrefix && tr.seq === s.seq);

                const matchedTrackerKeys = new Set();

                etaList.forEach(item => {
                    let matchedTr = existingTrackersForStop.find(tr =>
                        !matchedTrackerKeys.has(tr.trackerKey) &&
                        Math.abs(tr.latestEtaTs - item.etaTs) <= 3.5 * 60000
                    );

                    if (matchedTr) {
                        matchedTr.latestEtaTs = item.etaTs;
                        matchedTr.lastSeenTs = now;
                        matchedTr.missCount = 0;
                        matchedTr.lastEtaSeq = item.eta_seq;
                        matchedTrackerKeys.add(matchedTr.trackerKey);
                        currentSeenTrackerKeys.add(matchedTr.trackerKey);
                    } else if (item.eta_seq <= 2) {
                        const newKey = `${routeKeyPrefix}_SEQ_${s.seq}_${item.etaTs}_${Math.random().toString(36).slice(2, 6)}`;
                        activeStopTrackers[newKey] = {
                            trackerKey: newKey,
                            routeKeyPrefix,
                            seq: s.seq,
                            scheduledEtaTs: item.etaTs,
                            latestEtaTs: item.etaTs,
                            lastSeenTs: now,
                            lastEtaSeq: item.eta_seq,
                            missCount: 0
                        };
                        currentSeenTrackerKeys.add(newKey);
                    }
                });

                activeTrips.forEach(t => {
                    if (t.stops[s.seq] && !t.stops[s.seq].actualArrivalTs && !t.stops[s.seq].scheduledEtaTs) {
                        const matchingEta = etaList.find(item => isEtaConsistentWithTrip(t, s.seq, item.etaTs));
                        if (matchingEta) {
                            t.stops[s.seq].scheduledEtaTs = matchingEta.etaTs;
                        }
                    }
                });
            });

            const disappearingKeys = Object.keys(activeStopTrackers)
                .filter(tKey => activeStopTrackers[tKey].routeKeyPrefix === routeKeyPrefix && !currentSeenTrackerKeys.has(tKey))
                .sort((a, b) => activeStopTrackers[a].seq - activeStopTrackers[b].seq);

            disappearingKeys.forEach(tKey => {
                const tr = activeStopTrackers[tKey];
                tr.missCount = (tr.missCount || 0) + 1;

                const isWithinArrivalWindow = now >= tr.latestEtaTs - 45000 && now <= tr.latestEtaTs + 15 * 60000;

                if (isWithinArrivalWindow && tr.lastEtaSeq === 1) {
                    const realisticArrivalTs = Math.min(now, Math.max(tr.latestEtaTs, now - 12000));
                    findOrCreateTripForStopArrival(
                        routeKeyPrefix,
                        bm,
                        tr.seq,
                        realisticArrivalTs,
                        tr.scheduledEtaTs,
                        stops,
                        stopAllEtasMap
                    );
                    delete activeStopTrackers[tKey];
                } else if (tr.missCount >= 2 || now > tr.latestEtaTs + 15 * 60000) {
                    delete activeStopTrackers[tKey];
                }
            });
        }

        stitchFragmentedTrips();

        lastPollTimeStr = formatTime(now);
        saveToExcel();
        triggerGithubBackup(false);
    } catch (err) {
        console.error('背景監察錯誤:', err.message);
    } finally {
        isPolling = false;
    }
}

async function scheduleNextPoll() {
    await pollBookmarkedRoutes();

    const now = Date.now();
    const hasImminentArrival = Object.values(activeStopTrackers).some(
        tr => tr.lastEtaSeq === 1 && (tr.latestEtaTs - now) <= 90000 && (tr.latestEtaTs - now) >= -30000
    );

    const nextIntervalMs = hasImminentArrival ? 15000 : 25000;
    setTimeout(scheduleNextPoll, nextIntervalMs);
}

// ==========================================
// REST API
// ==========================================
app.get('/health', (req, res) => {
    res.status(200).json({ ok: true, time: formatTime(Date.now()), drainMode });
});

app.get('/api/backup-now', async (req, res) => {
    if (!GITHUB_TOKEN || !GITHUB_REPO) {
        return res.status(400).json({
            ok: false,
            error: '尚未設定 GITHUB_TOKEN 或 GITHUB_REPO 環境變數',
            hasToken: Boolean(GITHUB_TOKEN),
            repo: GITHUB_REPO || '(未設定)',
            branch: GITHUB_BRANCH
        });
    }
    saveBookmarks();
    saveToExcel();
    await triggerGithubBackup(true);
    res.json({
        ok: !lastGithubBackupError,
        repo: GITHUB_REPO,
        branch: GITHUB_BRANCH,
        lastGithubBackupTime: lastGithubBackupTs ? formatTime(lastGithubBackupTs) : '尚未成功',
        error: lastGithubBackupError || null
    });
});

app.get('/api/status', (req, res) => {
    res.json({
        running: true,
        drainMode,
        rollingExcelFileName: ROLLING_48H_EXCEL_FILENAME,
        rollingCsvFileName: ROLLING_48H_CSV_FILENAME,
        excelFileName: ROLLING_48H_EXCEL_FILENAME,
        sessionExcelFileName: SESSION_EXCEL_FILENAME,
        excelPath: ROLLING_48H_EXCEL_FILE,
        lastPollTime: lastPollTimeStr,
        bookmarkCount: bookmarks.length,
        tripCount: Object.keys(trips).length,
        githubBackup: {
            configured: Boolean(GITHUB_TOKEN && GITHUB_REPO),
            repo: GITHUB_REPO || null,
            branch: GITHUB_BRANCH,
            lastBackupTime: lastGithubBackupTs ? formatTime(lastGithubBackupTs) : null,
            lastError: lastGithubBackupError || null
        },
        bookmarks
    });
});

app.post('/api/auth/verify', (req, res) => {
    const { password } = req.body || {};
    if (password === ADMIN_PASSWORD) {
        return res.json({ valid: true });
    }
    return res.status(401).json({ valid: false, error: '密碼錯誤' });
});

app.post('/api/drain', (req, res) => {
    drainMode = true;
    const activeCount = Object.values(trips).filter(t => isTripStillActive(t, Date.now())).length;
    res.json({ drainMode: true, remainingActiveTrips: activeCount });
});

app.get('/api/route-shape', (req, res) => {
    const operator = (req.query.operator || '').toUpperCase();
    const route = (req.query.route || '').toUpperCase();
    const dirParam = req.query.dirParam || 'outbound';
    const targetSeq = dirParam === 'outbound' ? '1' : '2';

    let matched = busShapesFeatures.find(f =>
        f.properties.route === route &&
        f.properties.operator.includes(operator) &&
        f.properties.routeSeq === targetSeq
    );

    if (!matched) {
        matched = busShapesFeatures.find(f =>
            f.properties.route === route &&
            f.properties.operator.includes(operator)
        ) || busShapesFeatures.find(f => f.properties.route === route);
    }

    if (matched) {
        res.json({ found: true, feature: matched });
    } else {
        res.json({ found: false });
    }
});

app.post('/api/bookmarks/toggle', (req, res) => {
    const { password, operator, route, dirParam, serviceType = '1', origTc = '', destTc = '', origEn = '', destEn = '' } = req.body;

    if (password !== ADMIN_PASSWORD) {
        return res.status(401).json({ error: '管理密碼錯誤或未輸入密碼，無權更改監察路線' });
    }

    const id = `${operator}_${route}_${dirParam}_${serviceType}`;
    const idx = bookmarks.findIndex(b => b.id === id);

    let bookmarked = false;
    if (idx >= 0) {
        bookmarks.splice(idx, 1);
    } else {
        bookmarks.push({
            id,
            operator,
            route,
            dirParam,
            serviceType,
            origTc,
            destTc,
            origEn: origEn || origTc,
            destEn: destEn || destTc,
            addedAt: `${formatDate(Date.now())} ${formatTime(Date.now())}`
        });
        bookmarked = true;
    }

    saveBookmarks();
    saveToExcel();
    triggerGithubBackup(true);
    if (bookmarked) pollBookmarkedRoutes();

    res.json({ bookmarked, bookmarks });
});

app.get('/api/trips', (req, res) => {
    const list = Object.values(trips)
        .sort((a, b) => b.departureTs - a.departureTs)
        .slice(0, 100);
    res.json({ trips: list, lastPollTime: lastPollTimeStr });
});

app.get('/api/download-excel', (req, res) => {
    saveToExcel();
    if (fs.existsSync(ROLLING_48H_EXCEL_FILE)) {
        res.download(ROLLING_48H_EXCEL_FILE, ROLLING_48H_EXCEL_FILENAME);
    } else if (fs.existsSync(SESSION_EXCEL_FILE)) {
        res.download(SESSION_EXCEL_FILE, SESSION_EXCEL_FILENAME);
    } else {
        res.status(404).send('尚未產生 Excel 檔案');
    }
});

app.get('/api/download-csv', (req, res) => {
    saveToExcel();
    if (fs.existsSync(ROLLING_48H_CSV_FILE)) {
        res.download(ROLLING_48H_CSV_FILE, ROLLING_48H_CSV_FILENAME);
    } else {
        res.status(404).send('尚未產生 CSV 檔案');
    }
});

async function bootstrap() {
    await restoreFromGithub();
    saveBookmarks();
    saveToExcel();

    app.listen(PORT, HOST, () => {
        console.log('====================================================');
        console.log(`✅ 雲端巴士班次與累積時間監察系統已啟動：http://${HOST}:${PORT}`);
        console.log(`📊 48 小時滾動 Excel 主檔：${ROLLING_48H_EXCEL_FILENAME}`);
        console.log(`📄 48 小時滾動 CSV 預覽檔：${ROLLING_48H_CSV_FILENAME}`);
        console.log('====================================================');
        if (bookmarks.length > 0) {
            resolveDefaultBookmarkNames().then(() => scheduleNextPoll());
        } else {
            scheduleNextPoll();
        }
    });
}

bootstrap();