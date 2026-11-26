const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;
const HOST = '0.0.0.0';
const TIMEZONE = 'Asia/Hong_Kong';

const KMB_API = 'https://data.etabus.gov.hk/v1/transport/kmb';
const CTB_API = 'https://rt.data.gov.hk/v1/transport/citybus-nwfb';

const DATA_DIR = process.env.DATA_DIR || __dirname;
if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
}

const BOOKMARKS_FILE = path.join(DATA_DIR, 'bookmarks.json');

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
const EXCEL_FILENAME = `bus_arrival_records_${SESSION_TAG}.xlsx`;
const EXCEL_FILE = path.join(DATA_DIR, EXCEL_FILENAME);
const TRIPS_JSON_FILE = path.join(DATA_DIR, `trip_records_${SESSION_TAG}.json`);

app.use(cors());
app.use(express.json());
app.use(express.static(__dirname));

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
            addedAt: nowStr
        };
    }).filter(Boolean);
}

let bookmarks = loadInitialBookmarks();
let trips = {};
let activeStopTrackers = {};

let routeStopsCache = {};
let kmbStopsMap = {};
let ctbStopsMap = {};
let isPolling = false;
let lastPollTimeStr = '尚未執行';

function saveBookmarks() {
    fs.writeFileSync(BOOKMARKS_FILE, JSON.stringify(bookmarks, null, 2));
}

async function resolveDefaultBookmarkNames() {
    let updated = false;
    for (const bm of bookmarks) {
        if (bm.origTc !== '(雲端預設)' && bm.destTc !== '(雲端預設)') continue;
        try {
            if (bm.operator === 'KMB') {
                const res = await fetch(`${KMB_API}/route/${bm.route}/${bm.dirParam}/${bm.serviceType || '1'}`);
                const json = await res.json();
                if (json.data) {
                    bm.origTc = (json.data.orig_tc || '').trim() || bm.origTc;
                    bm.destTc = (json.data.dest_tc || '').trim() || bm.destTc;
                    updated = true;
                }
            } else {
                const res = await fetch(`${CTB_API}/route/ctb/${bm.route}`);
                const json = await res.json();
                if (json.data) {
                    const orig = (json.data.orig_tc || '').trim();
                    const dest = (json.data.dest_tc || '').trim();
                    bm.origTc = bm.dirParam === 'outbound' ? orig : dest;
                    bm.destTc = bm.dirParam === 'outbound' ? dest : orig;
                    updated = true;
                }
            }
        } catch (e) { /* ignore */ }
    }
    if (updated) {
        saveBookmarks();
        saveToExcel();
        console.log('✅ 已自動補全雲端預設路線的中文起訖點名稱');
    }
}

function pruneOldTrips(nowTs) {
    const cutoff = nowTs - 8 * 3600 * 1000;
    Object.keys(trips).forEach(id => {
        if (trips[id].departureTs < cutoff) {
            delete trips[id];
        }
    });
}

// 識別轉車站或無中繼站的高速公路／隧道節點（處理分站牌無 ETA 跳站與長距離直達路段）
function isHubOrExpresswayStop(stopName = '') {
    return /轉車站|轉乘站|西隧|紅隧|東隧|大橋|口岸/.test(stopName);
}

// 事後清洗：將因 ETA 桶位邊界或轉車站跳站而分裂出的「(途中估算)」碎片，精準縫合回主班次
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

        const candidates = formalTrips.filter(f =>
            f.routeKeyPrefix === est.routeKeyPrefix &&
            Math.abs(f.departureTs - est.departureTs) <= 25 * 60000
        );

        for (const formal of candidates) {
            let canMerge = true;

            for (const seq of estArrivedSeqs) {
                const estTs = est.stops[seq].actualArrivalTs;

                // 若正式班次在該站已有到站紀錄且時間差超過 90 秒，代表是不同班次（例如 10 分鐘一班的後車）
                if (formal.stops[seq]?.actualArrivalTs) {
                    if (Math.abs(formal.stops[seq].actualArrivalTs - estTs) > 90000) {
                        canMerge = false;
                        break;
                    }
                }

                // 檢查時間序列是否嚴格單調遞增（前一站 <= 本站 <= 後一站）
                const prevSeqs = Object.keys(formal.stops)
                    .map(Number)
                    .filter(s => s < seq && formal.stops[s].actualArrivalTs !== null);
                if (prevSeqs.length > 0) {
                    const maxPrevSeq = Math.max(...prevSeqs);
                    if (estTs < formal.stops[maxPrevSeq].actualArrivalTs - 15000) {
                        canMerge = false;
                        break;
                    }
                }

                const nextSeqs = Object.keys(formal.stops)
                    .map(Number)
                    .filter(s => s > seq && formal.stops[s].actualArrivalTs !== null);
                if (nextSeqs.length > 0) {
                    const minNextSeq = Math.min(...nextSeqs);
                    if (estTs > formal.stops[minNextSeq].actualArrivalTs + 15000) {
                        canMerge = false;
                        break;
                    }
                }
            }

            if (canMerge) {
                for (const seq of estArrivedSeqs) {
                    if (!formal.stops[seq].actualArrivalTs) {
                        const arrTs = est.stops[seq].actualArrivalTs;
                        formal.stops[seq] = {
                            ...est.stops[seq],
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
        { wch: 6 }, { wch: 26 }, { wch: 18 }, { wch: 16 }, { wch: 22 }, { wch: 16 }
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
        XLSX.writeFile(wb, EXCEL_FILE);
    } catch (err) {
        console.error('⚠️ 寫入 Excel 失敗:', err.message);
    }
    fs.writeFileSync(TRIPS_JSON_FILE, JSON.stringify(trips, null, 2));
}

async function ensureKmbStops() {
    if (Object.keys(kmbStopsMap).length > 0) return;
    const res = await fetch(`${KMB_API}/stop`);
    const json = await res.json();
    (json.data || []).forEach(s => {
        kmbStopsMap[s.stop] = (s.name_tc || s.name_en || s.stop).trim();
    });
}

async function getRouteStops(operator, route, dirParam, serviceType = '1') {
    const key = `${operator}_${route}_${dirParam}_${serviceType}`;
    if (routeStopsCache[key]) return routeStopsCache[key];

    if (operator === 'KMB') {
        await ensureKmbStops();
        const res = await fetch(`${KMB_API}/route-stop/${route}/${dirParam}/${serviceType}`);
        const json = await res.json();
        const list = (json.data || []).map(item => ({
            seq: parseInt(item.seq, 10),
            stopId: item.stop,
            stopName: kmbStopsMap[item.stop] || item.stop
        }));
        routeStopsCache[key] = list;
        return list;
    } else {
        const res = await fetch(`${CTB_API}/route-stop/ctb/${route}/${dirParam}`);
        const json = await res.json();
        const list = await Promise.all((json.data || []).map(async item => {
            if (!ctbStopsMap[item.stop]) {
                try {
                    const sRes = await fetch(`${CTB_API}/stop/${item.stop}`);
                    const sJson = await sRes.json();
                    ctbStopsMap[item.stop] = (sJson.data?.name_tc || sJson.data?.name_en || item.stop).trim();
                } catch {
                    ctbStopsMap[item.stop] = item.stop;
                }
            }
            return {
                seq: parseInt(item.seq, 10),
                stopId: item.stop,
                stopName: ctbStopsMap[item.stop]
            };
        }));
        routeStopsCache[key] = list;
        return list;
    }
}

function isEtaConsistentWithTrip(trip, seq, etaTs) {
    const diffFromDepMins = (etaTs - trip.departureTs) / 60000;

    const minFromDepMins = seq <= 4 ? (seq - 1) * 0.2 : (seq - 1) * 0.45;
    const maxFromDepMins = seq * 8;
    if (diffFromDepMins < minFromDepMins || diffFromDepMins > maxFromDepMins) {
        return false;
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
        const maxStepSecs = Math.max(seqGap * 15 * 60, 30 * 60);
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

function findOrCreateTripForStopArrival(routeKeyPrefix, meta, seq, arrivalTs, scheduledEtaTs, stopsList) {
    // 1. 第 1 站（總站）正式開出
    if (seq === 1) {
        const preciseDepTs = (scheduledEtaTs && Math.abs(arrivalTs - scheduledEtaTs) <= 90000)
            ? scheduledEtaTs
            : arrivalTs;
        const depTimeStr = formatTime(preciseDepTs);
        const tripId = `${routeKeyPrefix}_DEP_${preciseDepTs}`;

        const newTrip = {
            tripId,
            routeKeyPrefix,
            date: formatDate(preciseDepTs),
            operator: meta.operator,
            route: meta.route,
            dirParam: meta.dirParam,
            destTc: meta.destTc,
            departureTs: preciseDepTs,
            departureTimeStr: depTimeStr,
            isEstimatedDep: false,
            stops: {}
        };
        stopsList.forEach(s => {
            newTrip.stops[s.seq] = {
                seq: s.seq,
                stopName: s.stopName,
                scheduledEtaTs: s.seq === 1 ? (scheduledEtaTs || preciseDepTs) : null,
                actualArrivalTs: s.seq === 1 ? preciseDepTs : null,
                cumulativeMins: s.seq === 1 ? 0 : null,
                status: s.seq === 1 ? '已由總站開出' : '待抵達'
            };
        });

        Object.keys(trips).forEach(oldId => {
            const oldT = trips[oldId];
            if (oldT.routeKeyPrefix === routeKeyPrefix && oldT.isEstimatedDep && Math.abs(oldT.departureTs - preciseDepTs) <= 3.5 * 60000) {
                Object.values(oldT.stops).forEach(st => {
                    if (st.seq > 1 && st.seq <= 5 && st.actualArrivalTs) {
                        const fixedArrival = Math.max(preciseDepTs, st.actualArrivalTs);
                        newTrip.stops[st.seq] = {
                            ...st,
                            actualArrivalTs: fixedArrival,
                            cumulativeMins: Math.max(0, Math.round((fixedArrival - preciseDepTs) / 60000))
                        };
                    }
                });
                delete trips[oldId];
            }
        });

        trips[tripId] = newTrip;
        return newTrip;
    }

    // 2. 尋找現有班次（正式總站班次優先）
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

    let bestTrip = null;
    let smallestScore = Infinity;

    for (const trip of candidateTrips) {
        if (arrivalTs < trip.departureTs - 15000) continue;

        const elapsedMins = (arrivalTs - trip.departureTs) / 60000;
        if (elapsedMins > 150) continue;

        const minTotalTripMins = seq <= 4 ? (seq - 1) * 0.2 : (seq - 1) * 0.45;
        if (elapsedMins < minTotalTripMins) continue;

        const recordedSched = trip.stops[seq]?.scheduledEtaTs;
        if (recordedSched && (recordedSched - arrivalTs) / 60000 > 6) {
            continue;
        }

        const allRecordedSeqs = Object.keys(trip.stops)
            .map(Number)
            .filter(s => trip.stops[s].actualArrivalTs !== null);
        const highestSeqSoFar = allRecordedSeqs.length > 0 ? Math.max(...allRecordedSeqs) : 0;

        if (seq <= highestSeqSoFar) continue;

        if (highestSeqSoFar > 0) {
            const prevStop = trip.stops[highestSeqSoFar];
            const currStop = trip.stops[seq];
            const prevArrivalTs = prevStop.actualArrivalTs;
            const stepDiffSecs = (arrivalTs - prevArrivalTs) / 1000;
            const stepDiffMins = stepDiffSecs / 60;
            const seqGap = seq - highestSeqSoFar;

            // 檢查跳過的站點中是否包含轉車站（分站牌無 ETA 跳站）
            let skippedHub = isHubOrExpresswayStop(prevStop?.stopName) || isHubOrExpresswayStop(currStop?.stopName);
            for (let k = highestSeqSoFar + 1; k < seq; k++) {
                if (isHubOrExpresswayStop(trip.stops[k]?.stopName)) skippedHub = true;
            }

            const maxAllowedGap = skippedHub ? 8 : 5;
            if (seqGap > maxAllowedGap) continue;

            // 動態計算合理車程：若有官方預計時間，允許 [預計車程 * 0.65, 預計車程 + 12 分鐘]（適應屯門公路/西隧等 20 分鐘直達路段）
            let minStepSecs = seqGap === 1 ? -5 : (seqGap - 1) * 20;
            let maxAllowedMins = seqGap * 15;

            const currRefEta = currStop?.scheduledEtaTs || scheduledEtaTs;
            if (prevStop?.scheduledEtaTs && currRefEta && currRefEta > prevStop.scheduledEtaTs) {
                const expectedLegMins = (currRefEta - prevStop.scheduledEtaTs) / 60000;
                minStepSecs = Math.max(minStepSecs, expectedLegMins * 60 * 0.65);
                maxAllowedMins = Math.max(maxAllowedMins, expectedLegMins + 12);
            } else if (skippedHub) {
                maxAllowedMins = Math.max(maxAllowedMins, 28);
            }

            if (stepDiffSecs >= minStepSecs && stepDiffMins <= maxAllowedMins) {
                const score = stepDiffMins + seqGap * 0.5;
                if (score < smallestScore) {
                    smallestScore = score;
                    bestTrip = trip;
                }
            }
        } else {
            if (seq <= 6 && elapsedMins <= seq * 10) {
                if (elapsedMins < smallestScore) {
                    smallestScore = elapsedMins;
                    bestTrip = trip;
                }
            }
        }
    }

    // 3. 途中切入估算班次（用於捕捉系統重啟時已在半路的前車）
    if (!bestTrip) {
        const refTs = scheduledEtaTs || arrivalTs;
        const estDepTs = refTs - (seq - 1) * 1.6 * 60000;

        const conflictsWithFormalTrip = candidateTrips.some(
            t => !t.isEstimatedDep && Math.abs(t.departureTs - estDepTs) <= 5 * 60000
        );
        if (conflictsWithFormalTrip) {
            return null;
        }

        const estDepStr = formatTime(estDepTs);
        const estBucket = Math.round(estDepTs / (6 * 60000));
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
                departureTs: estDepTs,
                departureTimeStr: estDepStr,
                isEstimatedDep: true,
                stops: {}
            };
            stopsList.forEach(s => {
                trips[tripId].stops[s.seq] = {
                    seq: s.seq,
                    stopName: s.stopName,
                    scheduledEtaTs: null,
                    actualArrivalTs: null,
                    cumulativeMins: null,
                    status: '途中切入追蹤'
                };
            });
        }
        bestTrip = trips[tripId];
    }

    const cumMins = Math.max(0, Math.round((arrivalTs - bestTrip.departureTs) / 60000));
    const existingSched = bestTrip.stops[seq]?.scheduledEtaTs;
    const validSched = (scheduledEtaTs && isEtaConsistentWithTrip(bestTrip, seq, scheduledEtaTs))
        ? scheduledEtaTs
        : (existingSched || null);

    bestTrip.stops[seq] = {
        seq,
        stopName: stopsList.find(s => s.seq === seq)?.stopName || `第${seq}站`,
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
                const res = await fetch(`${KMB_API}/route-eta/${bm.route}/${serviceType}`);
                const json = await res.json();
                (json.data || []).forEach(item => {
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
                const batchSize = 8;
                for (let i = 0; i < stops.length; i += batchSize) {
                    const batch = stops.slice(i, i + batchSize);
                    await Promise.all(batch.map(async s => {
                        try {
                            const res = await fetch(`${CTB_API}/eta/ctb/${s.stopId}/${bm.route}`);
                            const json = await res.json();
                            const matched = (json.data || [])
                                .filter(item => item.dir === targetDirCode && item.eta)
                                .map(item => ({
                                    eta_seq: item.eta_seq,
                                    etaTs: new Date(item.eta).getTime(),
                                    rmk: (item.rmk_tc || '').trim()
                                }));
                            if (matched.length > 0) {
                                stopAllEtasMap[s.seq] = matched;
                            }
                        } catch (e) { /* ignore */ }
                    }));
                }
            }

            const sortedStops = [...stops].sort((a, b) => a.seq - b.seq);
            const activeTrips = Object.values(trips)
                .filter(t => t.routeKeyPrefix === routeKeyPrefix && !t.isEstimatedDep)
                .sort((a, b) => a.departureTs - b.departureTs);

            sortedStops.forEach(s => {
                const etaList = (stopAllEtasMap[s.seq] || []).sort((a, b) => a.eta_seq - b.eta_seq);
                if (etaList.length === 0) return;

                const firstInfo = etaList[0];
                const bucket = Math.round(firstInfo.etaTs / (10 * 60000));
                const trackerKey = `${routeKeyPrefix}_SEQ_${s.seq}_${bucket}`;
                currentSeenTrackerKeys.add(trackerKey);

                if (!activeStopTrackers[trackerKey]) {
                    activeStopTrackers[trackerKey] = {
                        trackerKey,
                        routeKeyPrefix,
                        seq: s.seq,
                        scheduledEtaTs: firstInfo.etaTs,
                        latestEtaTs: firstInfo.etaTs,
                        lastSeenTs: now,
                        missCount: 0
                    };
                } else {
                    activeStopTrackers[trackerKey].latestEtaTs = firstInfo.etaTs;
                    activeStopTrackers[trackerKey].lastSeenTs = now;
                    activeStopTrackers[trackerKey].missCount = 0;
                }

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

                if (isWithinArrivalWindow) {
                    const realisticArrivalTs = Math.min(now, Math.max(tr.latestEtaTs, now - 15000));
                    findOrCreateTripForStopArrival(
                        routeKeyPrefix,
                        bm,
                        tr.seq,
                        realisticArrivalTs,
                        tr.scheduledEtaTs,
                        stops
                    );
                    delete activeStopTrackers[tKey];
                } else if (tr.missCount >= 2 || now > tr.latestEtaTs + 15 * 60000) {
                    delete activeStopTrackers[tKey];
                }
            });
        }

        // 寫入 Excel 前執行碎片自動縫合清洗
        stitchFragmentedTrips();

        lastPollTimeStr = formatTime(now);
        saveToExcel();
        console.log(`[${lastPollTimeStr}] 監察完成，已寫入 ${EXCEL_FILENAME}（追蹤 ${Object.keys(trips).length} 個班次）`);
    } catch (err) {
        console.error('背景監察錯誤:', err.message);
    } finally {
        isPolling = false;
    }
}

// ==========================================
// REST API
// ==========================================
app.get('/health', (req, res) => {
    res.status(200).json({ ok: true, time: formatTime(Date.now()) });
});

app.get('/api/status', (req, res) => {
    res.json({
        running: true,
        excelFileName: EXCEL_FILENAME,
        excelPath: EXCEL_FILE,
        lastPollTime: lastPollTimeStr,
        bookmarkCount: bookmarks.length,
        tripCount: Object.keys(trips).length,
        bookmarks
    });
});

app.post('/api/bookmarks/toggle', (req, res) => {
    const { operator, route, dirParam, serviceType = '1', origTc = '', destTc = '' } = req.body;
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
            addedAt: `${formatDate(Date.now())} ${formatTime(Date.now())}`
        });
        bookmarked = true;
    }

    saveBookmarks();
    saveToExcel();
    if (bookmarked) pollBookmarkedRoutes();

    res.json({ bookmarked, bookmarks });
});

app.get('/api/trips', (req, res) => {
    const list = Object.values(trips)
        .sort((a, b) => b.departureTs - a.departureTs)
        .slice(0, 50);
    res.json({ trips: list, lastPollTime: lastPollTimeStr });
});

app.get('/api/download-excel', (req, res) => {
    saveToExcel();
    if (fs.existsSync(EXCEL_FILE)) {
        res.download(EXCEL_FILE, EXCEL_FILENAME);
    } else {
        res.status(404).send('尚未產生 Excel 檔案');
    }
});

saveBookmarks();
saveToExcel();
setInterval(pollBookmarkedRoutes, 30000);

app.listen(PORT, HOST, () => {
    console.log('====================================================');
    console.log(`✅ 雲端巴士班次與累積時間監察系統已啟動：http://${HOST}:${PORT}`);
    console.log(`📊 本次啟動新建 Excel 記錄檔：${EXCEL_FILENAME}`);
    console.log('====================================================');
    if (bookmarks.length > 0) {
        resolveDefaultBookmarkNames().then(() => pollBookmarkedRoutes());
    }
});