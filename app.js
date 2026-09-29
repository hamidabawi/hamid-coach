/* Hamid Coach - PWA front end. The Google Sheet "Hamid Coach" in Google Drive is the backend: a small Apps Script web app
   inside it answers with the logs and the computed dashboard, and writes the phone's entries into the log sheets. The phone keeps
   a copy, works offline, and computes the same dashboard locally between syncs. Demo mode needs no sheet at all. */
(() => {
'use strict';
const APP_VERSION = '3.1.0';
const APP_URL = new URL('./', location.href).href;

/* ---------- Excel serial dates (1899-12-30 epoch), local-calendar based ---------- */
const EPOCH = Date.UTC(1899, 11, 30);
const serialOf = (y, m, d) => Math.round((Date.UTC(y, m, d) - EPOCH) / 86400000);
const serialFromDate = dt => serialOf(dt.getFullYear(), dt.getMonth(), dt.getDate());
const dateFromSerial = s => { const t = new Date(EPOCH + s * 86400000); return new Date(t.getUTCFullYear(), t.getUTCMonth(), t.getUTCDate()); };
const START = serialOf(2026, 8, 28);
const fmtDate = (s, opts) => dateFromSerial(s).toLocaleDateString('en-CA', opts || { weekday: 'short', day: 'numeric', month: 'short' });
const weekOf = s => Math.max(1, Math.min(33, Math.floor((s - START) / 7) + 1));

/* ---------- column maps (0-based) — identical to the workbook tables ---------- */
const D = { date: 0, week: 1, day: 2, type: 3, session: 4, kcalT: 5, protT: 6, retaPlan: 7, tesaPlan: 8, weight: 9, waist: 10, rhr: 11, sleep: 12, steps: 13, kcalM: 14, protM: 15,
  multi: 16, fish: 17, d3: 18, biotin: 19, creatine: 20, mg: 21, collagen: 22, spine: 23, sauna: 24, retaU: 25, retaSite: 26, tesaU: 27, tesaSite: 28, glucose: 29, notes: 30,
  foodK: 31, foodP: 32, kcal: 33, prot: 34, supps: 35, logged: 36, sessLogged: 37 };
const DAILY_FIELDS = ['weight', 'waist', 'rhr', 'sleep', 'steps', 'kcalM', 'protM', 'multi', 'fish', 'd3', 'biotin', 'creatine', 'mg', 'collagen', 'spine', 'sauna', 'retaU', 'retaSite', 'tesaU', 'tesaSite', 'glucose', 'notes'];   // columns J..AE
const W = { date: 0, week: 1, day: 2, session: 3, n: 4, ex: 5, sets: 6, reps: 7, effort: 8, load: 9, repsTop: 10, setsDone: 11, rir: 12, note: 13, e1rm: 14, family: 15, done: 16 };
const WFIELDS = { 9: 'load', 10: 'reps', 11: 'sets', 12: 'rir', 13: 'note' };
const F = { date: 0, meal: 1, item: 2, kcal: 3, prot: 4, note: 5, id: 6 };
const BENCH_FIELDS = ['sorensen', 'sideR', 'sideL', 'flexor', 'plank', 'bridgeR', 'bridgeL', 'deadbug', 'thomas', 'pain', 'bw', 'backext', 'carry', 'notes'];   // columns C..P
const SUPPS = [
  { k: 'multi', n: 'Multivitamin', i: '💊', sheet: 'Multivitamin' }, { k: 'fish', n: 'Fish oil', i: '🐟', sheet: 'Fish oil' }, { k: 'd3', n: 'D3 + K2', i: '☀️', sheet: 'Vitamin D3 + K2' },
  { k: 'biotin', n: 'Biotin', i: '⚠️', warn: true, sheet: 'Biotin (plan says STOP; logged so lab timing is visible)' }, { k: 'creatine', n: 'Creatine 5 g', i: '⚡', sheet: 'Creatine 5 g' },
  { k: 'mg', n: 'Magnesium', i: '🌙', sheet: 'Magnesium (bedtime)' }, { k: 'collagen', n: 'Collagen + C', i: '🦴', sheet: 'Collagen + vit C (Tue/Thu/Fri)' }];
const SITES = ['Abdomen L', 'Abdomen R', 'Arm L (back)', 'Arm R (back)', 'Thigh L', 'Thigh R'];
const LABS_DEFAULT = [[serialOf(2026, 9, 1), 'Baseline bloods (7-10 am, fasted)'], [serialOf(2026, 9, 7), '2nd testosterone morning'], [serialOf(2026, 10, 18), 'Tesamorelin reference draw (if it runs)'], [serialOf(2026, 10, 21), 'Benchmark test #2'],
  [serialOf(2026, 11, 16), 'Main repeat panel'], [serialOf(2027, 0, 13), 'IGF-1 + glucose (if tesamorelin)'], [serialOf(2027, 0, 23), 'Benchmark test #3'], [serialOf(2027, 0, 29), 'DXA #2 (to Feb 1)'],
  [serialOf(2027, 1, 10), 'IGF-1 + glucose (conditional)'], [serialOf(2027, 2, 10), 'Full repeat panel + ECG'], [serialOf(2027, 2, 17), '2nd testosterone morning'], [serialOf(2027, 2, 20), 'Benchmark test #4'], [serialOf(2027, 2, 27), 'DXA #3']];
const STATUS_KEYS = ['RATE STATUS', 'WAIST STATUS', 'HR STATUS', 'CALORIE STATUS', 'PROTEIN STATUS', 'RETA STATUS', 'TESA STATUS', 'GLUCOSE STATUS', 'TRAINING STATUS', 'SPINE ROUTINE STATUS', 'SAUNA CHECK', 'STRENGTH OVERRIDE', 'BENCHMARK STATUS'];
const STATUS_TITLES = { 'RATE STATUS': 'Rate of loss', 'WAIST STATUS': 'Waist', 'HR STATUS': 'Resting heart rate', 'CALORIE STATUS': 'Calories', 'PROTEIN STATUS': 'Protein', 'RETA STATUS': 'Retatrutide', 'TESA STATUS': 'Tesamorelin', 'GLUCOSE STATUS': 'Fasting glucose', 'TRAINING STATUS': 'Training', 'SPINE ROUTINE STATUS': 'Spine routine', 'SAUNA CHECK': 'Sauna', 'STRENGTH OVERRIDE': 'Strength', 'BENCHMARK STATUS': 'Benchmarks' };
const FAMILIES = ['Squat', 'Press', 'Pull', 'RDL', 'Trap bar', 'Row', 'Back ext', 'Carry', 'Hip thrust'];

/* ---------- storage + state ---------- */
const store = {
  get(k, d) { try { const v = localStorage.getItem('coach.' + k); return v == null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem('coach.' + k, JSON.stringify(v)); } catch {} },
  del(k) { try { localStorage.removeItem('coach.' + k); } catch {} },
};
const settings = Object.assign({ api: '', key: '', demo: false, name: 'Hamid' }, store.get('settings', {}));
const saveSettings = () => store.set('settings', settings);
const state = { screen: store.get('screen', 'today'), data: store.get('cache', null), queue: store.get('queue', []), rejected: null, foodDraft: { item: '', kcal: '', prot: '', note: '' }, busy: false, apiDown: null, mode: '', logDate: null, trainDate: null, foodDate: null, meal: null, error: null };
const deviceId = store.get('device', null) || (() => { const id = (crypto.randomUUID ? crypto.randomUUID() : String(Math.random()).slice(2)).slice(0, 8); store.set('device', id); return id; })();
const uuid = () => crypto.randomUUID ? crypto.randomUUID() : (Date.now().toString(36) + Math.random().toString(36).slice(2, 10));
const todaySerial = () => settings.demo ? serialOf(2026, 10, 11) : serialFromDate(new Date());
const num = v => { if (v === '' || v == null) return null; const n = +String(v).replace(',', '.'); return isNaN(n) ? null : n; };
const isBlank = v => v === '' || v == null;
const asText = s => /^[=+\-@]/.test(s) ? ' ' + s : s;   // a leading = + - @ would be read as a formula by a spreadsheet
const fmt = (v, d = 0, min = d) => { const n = num(v); return n == null ? '—' : n.toLocaleString('en', { minimumFractionDigits: min, maximumFractionDigits: d }); };
const colL = i => { let s = ''; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
const colIdx = letters => letters.split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1;
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const r1 = n => Math.round(n * 10) / 10;
const LABS = () => (state.data && state.data.labs) || LABS_DEFAULT;

/* ---------- the dashboard, computed on the phone (mirror of the workbook formulas) ---------- */
function computeDashboard(d, today) {
  const avg = a => a.length ? a.reduce((x, y) => x + y, 0) / a.length : null;
  const byDate = {}; d.daily.forEach((r, i) => byDate[r[D.date]] = i);
  const foodByDate = {}; d.food.forEach(f => { (foodByDate[f[F.date]] = foodByDate[f[F.date]] || []).push(f); });
  d.daily.forEach(r => { const k = foodByDate[r[D.date]] || []; r[D.foodK] = k.reduce((a, f) => a + (num(f[F.kcal]) || 0), 0); r[D.foodP] = k.reduce((a, f) => a + (num(f[F.prot]) || 0), 0); r[D.kcal] = num(r[D.kcalM]) ?? (r[D.foodK] > 0 ? r[D.foodK] : ''); r[D.prot] = num(r[D.protM]) ?? (r[D.foodP] > 0 ? r[D.foodP] : ''); r[D.supps] = SUPPS.filter(s => r[D[s.k]] === 'Y').length; r[D.logged] = (r[D.weight] !== '' || r[D.kcal] !== '' || r[D.supps] > 0 || r[D.spine] === 'Y' || !isBlank(r[D.retaU]) || !isBlank(r[D.tesaU])) ? 1 : 0; });
  d.workout.forEach(r => { r[W.e1rm] = num(r[W.load]) != null && num(r[W.repsTop]) != null ? r[W.load] * (1 + r[W.repsTop] / 30) : ''; r[W.done] = (!isBlank(r[W.load]) || !isBlank(r[W.setsDone])) ? 1 : 0; });
  const doneDates = new Set(d.workout.filter(r => r[W.done] === 1).map(r => r[W.date])); d.daily.forEach(r => r[D.sessLogged] = doneDates.has(r[D.date]) ? 1 : 0);
  d.bench.forEach(b => { const s = num(b[2]); b[16] = s && num(b[5]) != null ? b[5] / s : ''; b[17] = s && num(b[3]) != null ? b[3] / s : ''; b[18] = s && num(b[4]) != null ? b[4] / s : ''; b[19] = num(b[3]) != null && num(b[4]) ? b[3] / b[4] : ''; });
  const fcTable = d.forecast || null;
  const fc = w => { if (fcTable && fcTable[w - 1]) return fcTable[w - 1][2]; const pts = [[0, 101], [3, 98.5], [12, 93.5], [26, 89], [33, 88.5]]; for (let i = 0; i < pts.length - 1; i++) { const [w1, v1] = pts[i], [w2, v2] = pts[i + 1]; if (w >= w1 && w <= w2) return v1 + (v2 - v1) * (w - w1) / (w2 - w1); } return 88.5; };
  const base = avg(d.daily.slice(0, 4).map(r => num(r[D.rhr])).filter(x => x != null));
  const weekly = []; let prev = d.startWeight || 101;
  for (let w = 1; w <= 33; w++) {
    const rows = d.daily.filter(r => weekOf(r[D.date]) === w), ws = rows.map(r => num(r[D.weight])).filter(x => x != null), aw = avg(ws);
    const kc = rows.map(r => num(r[D.kcal])).filter(x => x), pr = rows.map(r => num(r[D.prot])).filter(x => x), hr = rows.map(r => num(r[D.rhr])).filter(x => x);
    const ph = (fcTable && fcTable[w - 1] && fcTable[w - 1][1]) || ([9, 18, 26].includes(w) ? 'Deload' : w === 25 ? 'Retest' : w >= 27 ? 'Taper / maintenance' : 'Month ' + (w <= 5 ? 1 : w <= 9 ? 2 : w <= 14 ? 3 : w <= 18 ? 4 : w <= 22 ? 5 : 6));
    const elapsed = Math.max(0, Math.min(7, today - (START + 7 * (w - 1)))), ycount = rows.filter(r => r[D.date] < today).reduce((a, r) => a + SUPPS.filter(s => r[D[s.k]] === 'Y').length, 0);   // closed days only
    const floor = aw != null && aw < 92 ? 1900 : 2000, band = w <= 12 ? 1.5 : 2;
    const lo = fcTable && fcTable[w - 1] ? fcTable[w - 1][3] : r1(fc(w) - band), hi = fcTable && fcTable[w - 1] ? fcTable[w - 1][4] : r1(fc(w) + band);
    weekly.push([w, START + 7 * (w - 1), START + 7 * (w - 1) + 6, ph, aw == null ? '' : r1(aw), aw == null ? '' : r1(aw - prev), rows.filter(r => num(r[D.waist]) != null).map(r => r[D.waist]).pop() ?? '', hr.length ? r1(avg(hr)) : '', hr.length && base != null ? r1(avg(hr) - base) : '', kc.length ? Math.round(avg(kc)) : '', rows.length ? Math.round(avg(rows.map(r => num(r[D.kcalT]) || 0))) : '', kc.filter(x => x < floor).length, pr.length ? Math.round(avg(pr)) : '', pr.filter(x => x < 180).length, elapsed ? ycount / (elapsed * SUPPS.length) : '', rows.filter(r => r[D.spine] === 'Y').length, rows.filter(r => r[D.sauna] === 'Y').length, rows.filter(r => r[D.sessLogged]).length, rows.filter(r => r[D.session] !== 'Rest').length, d.workout.filter(r => r[W.week] === w && r[W.done]).length, d.workout.filter(r => r[W.week] === w).length, rows.reduce((a, r) => a + (num(r[D.retaU]) || 0), 0), rows.filter(r => num(r[D.tesaU]) != null).length, rows.filter(r => num(r[D.tesaPlan]) != null).length, r1(fc(w)), lo, hi, (d.weeklyNotes && d.weeklyNotes[String(w)]) || '']);
    if (aw != null) prev = aw;
  }
  const last = n => d.daily.filter(r => r[D.date] > today - n && r[D.date] <= today);          // morning measures: includes today
  const closed = n => d.daily.filter(r => r[D.date] > today - n - 1 && r[D.date] <= today - 1);  // intake, supplements, routine: closed days, yesterday back
  const w7 = avg(last(7).map(r => num(r[D.weight])).filter(x => x != null)), w7p = avg(d.daily.filter(r => r[D.date] > today - 14 && r[D.date] <= today - 7).map(r => num(r[D.weight])).filter(x => x != null));
  const wk = weekOf(today), chg = w7 != null && w7p != null ? w7 - w7p : null, calib = !(wk <= 3 || [9, 18, 26].includes(wk)), ceil = wk <= 12 ? 1 : 0.75;
  const rhr7 = avg(last(7).map(r => num(r[D.rhr])).filter(x => x != null)), rhrmax = Math.max(0, ...last(7).map(r => num(r[D.rhr]) || 0)), rhr10 = base != null ? last(7).filter(r => num(r[D.rhr]) != null && r[D.rhr] >= base + 10).length : 0;
  const c7 = closed(7), kcLogged = c7.filter(r => num(r[D.kcal]));
  const kc7 = avg(kcLogged.map(r => num(r[D.kcal]))), kT7 = avg(kcLogged.map(r => num(r[D.kcalT]) || 0)), kdays = kcLogged.length;
  const minWeekly = Math.min(...weekly.map(r => num(r[4])).filter(x => x != null && x > 0)), floor = isFinite(minWeekly) && minWeekly < 92 ? 1900 : 2000;   // latches once under 92 kg
  const floordays = kcLogged.filter(r => r[D.kcal] < floor).length, pr7 = avg(c7.map(r => num(r[D.prot])).filter(x => x)), protdays = c7.filter(r => num(r[D.prot]) && r[D.prot] < 180).length;
  const map = {}; const S = (k, v) => map[k] = v == null ? '' : v;
  S('Today', today); S('Tracker week', wk); S('Phase', weekly[wk - 1][3]); S('Days logged so far', d.daily.filter(r => r[D.date] <= today && r[D.logged]).length); S('Days elapsed', today - START + 1);
  const wl = d.daily.filter(r => num(r[D.weight]) != null).pop(); S('Latest weight (kg)', wl ? wl[D.weight] : ''); S('7-day average (kg)', w7 == null ? '' : r1(w7)); S('Previous 7-day average (kg)', w7p == null ? '' : r1(w7p)); S('Change this week (kg)', chg == null ? '' : r1(chg)); S('Total lost since 101 kg', w7 == null ? '' : r1((d.startWeight || 101) - w7));
  S('Forecast this week (kg)', weekly[wk - 1][24]); S('Forecast band', (+weekly[wk - 1][25]).toFixed(1) + ' - ' + (+weekly[wk - 1][26]).toFixed(1)); S('Rate ceiling this week (kg/week)', ceil); S('Calibration week?', calib ? 'Yes' : 'No (excluded week)');
  S('RATE STATUS', chg == null ? 'WATCH: need 14 days of weights' : !calib ? 'OK: excluded week, no change' : -chg > 1.5 ? 'ACTION: lost more than 1.5 kg this week - tell the doctor, no escalation' : -chg > ceil ? 'ACTION: above the ceiling - add 150-200 kcal of carbs around training' : -chg < 0.4 ? 'WATCH: under 0.4 kg/week - if 3 weeks running with intake logged, cut 150 kcal from training-day carbs (2,200 to 2,050; rest days sit on the floor)' : 'OK: inside the expected 0.5-0.8 kg/week');
  const waist = d.daily.filter(r => num(r[D.waist]) != null).pop(); S('Waist (latest, cm)', waist ? waist[D.waist] : ''); S('Waist-to-height', waist ? waist[D.waist] / (d.heightCm || 190) : ''); S('WAIST STATUS', !waist ? 'WATCH: no tape yet' : waist[D.waist] < 95 ? 'OK: under 0.5 (waist under 95 cm)' : 'WATCH: ' + r1(waist[D.waist] - 95).toFixed(1) + ' cm above the 95 cm target');
  S('Baseline (mean Sep 28 - Oct 1)', base == null ? '' : r1(base)); S('7-day mean', rhr7 == null ? '' : r1(rhr7)); S('7-day mean vs baseline (bpm)', rhr7 != null && base != null ? r1(rhr7 - base) : ''); S('Highest reading, last 7 days', rhrmax || ''); S('Days at +10 or more, last 7', rhr10);
  const diff = rhr7 != null && base != null ? Math.round(rhr7 - base) : null;
  S('HR STATUS', diff == null ? 'WATCH: log resting HR every morning' : (rhrmax >= 100 || rhr10 >= 3) ? 'ACTION: cut sets by a third, no sauna or HIIT, call the doctor before the next dose, ECG this week, no escalation' : diff >= 10 ? 'ACTION: +10 sustained - apply the rule' : diff >= 6 ? 'WATCH: +6 to +9 bpm - drop the technique sets this week and mention it at the next check-in' : 'OK: within 5 bpm of baseline');
  S('7-day average calories', kc7 == null ? '' : Math.round(kc7)); S('7-day average target', kT7 == null ? '' : Math.round(kT7)); S('Days with calories logged, last 7', kdays); S('Calorie floor now', floor); S('Days under the floor, last 7', floordays);
  S('CALORIE STATUS', kdays === 0 ? 'WATCH: no food logged in the last 7 days' : (kc7 < 1800 && kdays >= 7) ? 'ACTION: 7-day average under 1,800 - dose conversation with the written 3 mg fallback' : floordays >= 3 ? 'ACTION: under the floor on 3+ days - fixed liquid meal every day next week and tell the doctor before the next injection' : kc7 > kT7 + 200 ? 'WATCH: averaging ' + Math.round(kc7 - kT7) + ' kcal over target' : 'OK: ' + Math.round(kc7) + ' kcal vs ' + Math.round(kT7) + ' target');
  S('7-day average protein (g)', pr7 == null ? '' : Math.round(pr7)); S('Days under 180 g, last 7', protdays); S('PROTEIN STATUS', pr7 == null ? 'WATCH: no protein logged' : protdays >= 2 ? 'ACTION: under 180 g on ' + protdays + ' of the last 7 days - shakes are not optional' : pr7 < 180 ? 'WATCH: averaging under 180 g' : 'OK: ' + Math.round(pr7) + ' g average');
  const elapsedAll = Math.max(1, today - START), supp = {}, nextLab = LABS().find(l => l[0] >= today), c28 = closed(28);
  SUPPS.forEach(s => { const c = D[s.k], y7 = c7.filter(r => r[c] === 'Y').length, p7 = y7 / Math.min(7, elapsedAll), p28 = c28.filter(r => r[c] === 'Y').length / Math.min(28, elapsedAll); supp[s.sheet] = { d7: p7, d28: p28, tracked: 'Y', status: s.k === 'biotin' ? (y7 > 0 ? 'WATCH: the plan says stop biotin; none within 72 h of a blood draw (next: ' + (nextLab ? fmtDate(nextLab[0], { day: 'numeric', month: 'short' }) : '-') + ')' : 'OK: not taken') : s.k === 'collagen' ? (y7 >= 2 ? 'OK' : 'WATCH: aim for Tue, Thu and Fri') : p7 >= 6 / 7 ? 'OK' : p7 >= 4 / 7 ? 'WATCH: ' + Math.round(p7 * 100) + '% this week' : 'ACTION: ' + Math.round(p7 * 100) + '% this week' }; });
  const rp = d.daily.filter(r => num(r[D.retaPlan]) && r[D.date] <= today), rt = rp.filter(r => num(r[D.retaU])); S('Retatrutide doses planned to date', rp.length); S('Retatrutide doses logged', rt.length); const ld = d.daily.filter(r => num(r[D.retaU])).pop(); S('Last dose (date)', ld ? ld[D.date] : ''); S('Last dose (units)', ld ? ld[D.retaU] : '');
  const nx = d.daily.find(r => num(r[D.retaPlan]) && r[D.date] >= today); S('Next planned dose', nx ? fmtDate(nx[D.date]) + '  (' + nx[D.retaPlan] + ' mg = ' + nx[D.retaPlan] * 10 + ' units)' : 'taper: doctor sets'); S('RETA STATUS', rp.length === 0 ? 'OK: first dose Thu Oct 1' : rp.length - rt.length > 0 ? 'ACTION: ' + (rp.length - rt.length) + ' planned dose(s) not logged - log it or apply the missed-dose rule' : 'OK: every planned dose logged');
  const tp = d.daily.filter(r => num(r[D.tesaPlan]) && r[D.date] <= today), tt = tp.filter(r => num(r[D.tesaU])); S('Tesamorelin nights planned to date', tp.length); S('Tesamorelin nights logged', tt.length); S('Nights logged, last 7', c7.filter(r => num(r[D.tesaU])).length);
  const tr = d.daily[byDate[today]]; S("Tonight's tesamorelin", tr && num(tr[D.tesaPlan]) ? tr[D.tesaPlan] + ' mg = ' + tr[D.tesaPlan] * 5 + ' units (conditional on the gates)' : 'none planned'); S('TESA STATUS', tp.length === 0 ? 'OK: not started (conditional start Nov 19)' : tp.length - tt.length >= 3 ? 'ACTION: ' + (tp.length - tt.length) + ' planned nights not logged' : tp.length - tt.length > 0 ? 'WATCH: ' + (tp.length - tt.length) + ' night(s) missed - skip, never double' : 'OK: every planned night logged');
  const gl = d.daily.filter(r => num(r[D.glucose]) != null).pop(), g14 = d.daily.filter(r => r[D.date] > today - 14 && r[D.date] <= today && num(r[D.glucose]) != null); S('Last fasting glucose (mmol/L)', gl ? gl[D.glucose] : ''); const g61 = g14.filter(r => r[D.glucose] >= 6.1 && r[D.glucose] < 7).length, g7 = g14.filter(r => r[D.glucose] >= 7).length; S('Readings 6.1-6.9, last 14 days', g61); S('Readings 7.0 or more, last 14 days', g7);
  S('GLUCOSE STATUS', tp.length === 0 ? 'n/a until tesamorelin runs' : g7 > 0 ? 'ACTION: a reading of 7.0+ - lab venous glucose this week; only a lab value changes the dose' : g61 >= 2 ? 'WATCH: two readings 6.1-6.9 - lab venous glucose within the week' : !gl ? 'WATCH: log fasting glucose 3x per week' : 'OK');
  const wkrow = weekly[wk - 1], remaining = d.daily.filter(r => r[D.date] > today && r[D.date] <= wkrow[2] && r[D.session] !== 'Rest').length;
  S('Sessions done this week', wkrow[17]); S('Sessions planned this week', wkrow[18]); S('Sessions done, cumulative', d.daily.filter(r => r[D.date] <= today && r[D.sessLogged]).length); S('Sessions planned to date', d.daily.filter(r => r[D.date] <= today && r[D.session] !== 'Rest').length);
  const wlog = d.workout.filter(r => r[W.date] <= today); S('Exercises logged / planned to date', wlog.filter(r => r[W.done]).length + ' / ' + wlog.length); S('Volume shortfall (planned rows not logged, last 7 days)', d.workout.filter(r => r[W.date] > today - 8 && r[W.date] <= today - 1 && !r[W.done]).length);
  const sp7 = c7.filter(r => r[D.spine] === 'Y').length, sa7 = c7.filter(r => r[D.sauna] === 'Y').length; S('Spine routine days, last 7', sp7); S('Sauna days, last 7', sa7);
  S('TRAINING STATUS', wkrow[17] + remaining < wkrow[18] ? 'WATCH: ' + (wkrow[18] - wkrow[17] - remaining) + ' session(s) missed so far this week - never zero: B session rule' : 'OK: on track this week');
  S('SPINE ROUTINE STATUS', sp7 >= 6 ? 'OK: ' + sp7 + '/7' : sp7 >= 4 ? 'WATCH: ' + sp7 + '/7 - it never skips' : 'ACTION: ' + sp7 + '/7 days - the single most important line in the plan');
  S('SAUNA CHECK', sa7 > 4 ? 'WATCH: more than 4 sauna sessions in 7 days' : /^ACTION/.test(map['HR STATUS']) ? 'ACTION: no sauna this week (heart-rate rule)' : 'OK');
  const strength = {}; let downs = 0;
  FAMILIES.forEach(f => { const rows = d.workout.filter(r => r[W.family] === f && num(r[W.e1rm]) != null); const l = rows[rows.length - 1]; const m1 = Math.max(0, ...rows.filter(r => r[W.week] <= 5).map(r => r[W.e1rm])); const l14 = Math.max(0, ...rows.filter(r => r[W.date] > today - 14 && r[W.date] <= today).map(r => r[W.e1rm])); const ch = m1 && l14 ? l14 / m1 - 1 : ''; if (ch !== '' && ch < -0.05) downs++; strength[f] = { top: l ? l[W.load] + ' kg x ' + l[W.repsTop] : '', e1rm: l ? Math.round(l[W.e1rm]) : '', m1: m1 ? Math.round(m1) : '', last14: l14 ? Math.round(l14) : '', change: ch }; });
  S('STRENGTH OVERRIDE', downs >= 2 ? 'ACTION: two families down 5%+ vs Month 1 - if it holds two weeks, add 150 kcal regardless of the scale' : downs === 1 ? 'WATCH: one family down 5%+' : 'OK');
  const b = d.bench.filter(r => num(r[2]) != null).pop(); S('Latest test', b ? b[1] : 'none logged yet'); S('Sorensen (s)', b ? b[2] : ''); S('Side plank R / L (s)', b ? b[3] + ' / ' + b[4] : ''); S('Flexion:extension', b && b[16] !== '' ? r1(b[16] * 100) / 100 : ''); const nt = d.bench.find(r => r[0] >= today); S('Next test', nt ? fmtDate(nt[0]) : '');
  S('BENCHMARK STATUS', !b ? 'WATCH: no benchmark logged yet' : b[2] >= 150 ? 'OK: Sorensen ' + b[2] + ' s meets the Jan/Mar target' : b[2] >= 120 ? 'OK: Sorensen ' + b[2] + ' s passes the Nov gate; 150 s next' : 'WATCH: Sorensen ' + b[2] + ' s - under the 120 s gate, heavy 5s stay at RIR 3');
  const actions = STATUS_KEYS.map(k => map[k]).filter(v => /^(ACTION|WATCH)/.test(String(v)));
  Object.entries(supp).forEach(([n, s]) => { const m = String(s.status).match(/^(ACTION|WATCH):\s*(.*)$/); if (m) actions.push(`${m[1]}: ${n.replace(/ \(.*\)/, '')} - ${m[2]}`); });
  return { dashboard: { map, supp, strength, actions }, weekly };
}

/* ---------- GOOGLE SHEETS backend: the Apps Script web app inside the sheet ---------- */
const parsePairing = s => {   // "URL#KEY", "URL KEY", or an #api=...&key=... fragment
  s = String(s || '').trim(); if (!s) return null;
  let m = s.match(/api=([^&\s]+)&key=([^&\s]+)/); if (m) { try { return { api: decodeURIComponent(m[1]), key: decodeURIComponent(m[2]) }; } catch { return null; } }
  m = s.match(/^(https?:\/\/[^\s#]+?)\s*[#\s|,;]+\s*([A-Za-z0-9_-]{16,})$/); if (m) return { api: m[1].replace(/\/+$/, ''), key: m[2] };
  return null;
};
const looksLikeApi = u => /^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(u) || /^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?\//.test(u);
const shapeRoutine = rows => (rows || []).map(r => { const dl = String(r[5] == null ? '' : r[5]).trim().toLowerCase(); return { id: String(r[0]), name: String(r[1] || ''), meal: r[2] || 'Other', kcal: num(r[3]) || 0, prot: num(r[4]) || 0, days: dl === 'training' ? 'Training' : dl === 'rest' ? 'Rest' : 'All', active: !(r[6] === false || /^(n|no|false|0|off)$/i.test(String(r[6] == null ? '' : r[6]).trim())), order: num(r[7]) ?? 0, note: String(r[8] || '') }; }).sort((a, b) => a.order - b.order);
const DEMO_ROUTINE = [["r01", "Kirkland protein shake (ready-to-drink, chocolate, 325 mL)", "Wake-up", 160, 30, "All", "Y", 1, ""], ["r02", "Pre-workout (1 scoop)", "Wake-up", 10, 0, "Training", "Y", 2, ""], ["r03", "Espresso shot (only when drowsy)", "Wake-up", 3, 0, "All", "Y", 3, ""], ["r04", "Post-gym shake: 1.5 scoops LeanFit + 250 mL Natrel Plus + 1/4 cup oats + banana", "Post-gym", 558, 58, "Training", "Y", 4, ""], ["r05", "Rest-day shake: 1 scoop LeanFit + 250 mL Natrel Plus + banana (no oats)", "Post-gym", 405, 43, "Rest", "Y", 5, ""], ["r06", "Egg roll: 2 eggs + 2/3 cup egg whites + large tortilla", "Lunch", 500, 34, "All", "Y", 6, ""], ["r07", "Grenade protein bar, chocolate (the 'Granite' bar)", "Snack", 240, 21, "All", "Y", 7, ""], ["r08", "Fed meal", "Dinner", 750, 48, "All", "Y", 8, ""], ["r09", "Karak chai, Regular (Chaiiwala, Robson)", "Snack", 144, 5, "All", "Y", 9, ""], ["r10", "Greek yogurt 0%, 175 g", "Snack", 100, 18, "All", "Y", 10, ""], ["r11", "Pre-sleep protein: 2nd Kirkland carton or 1 cup 1% cottage cheese", "Pre-sleep", 160, 30, "All", "Y", 11, ""]];
const rtFoodId = (s, id) => `rt|${s}|${id}`;
class SheetsStore {
  constructor(api, key) { this.api = api.replace(/\/+$/, ''); this.key = key; }
  async call(params, body) {
    const ctrl = new AbortController(), timer = setTimeout(() => ctrl.abort(), 60000);
    const u = new URL(this.api); if (!body) Object.entries(Object.assign({ key: this.key }, params || {})).forEach(([k, v]) => u.searchParams.set(k, v));
    let res;
    try { res = await fetch(u.toString(), body ? { method: 'POST', body: JSON.stringify(Object.assign({ key: this.key }, body)), headers: { 'Content-Type': 'text/plain;charset=utf-8' }, redirect: 'follow', signal: ctrl.signal } : { redirect: 'follow', signal: ctrl.signal }); }
    catch (e) { throw new Error(e.name === 'AbortError' ? 'Google Sheets did not answer within 60 s' : 'No connection to Google Sheets'); }
    finally { clearTimeout(timer); }
    const text = await res.text(); let j;
    try { j = JSON.parse(text); } catch { throw new Error(/<html/i.test(text) ? 'The web app URL answered with a web page instead of data: check the deployment (Execute as: Me, Who has access: Anyone, the URL ends with /exec) and create a new version after code changes.' : 'Unexpected answer from Google Sheets (HTTP ' + res.status + ')'); }
    if (!j.ok) { const e = new Error(j.error || 'Google Sheets refused the request'); e.code = j.code || ''; throw e; }
    return j;
  }
  ping() { return this.call({ op: 'ping' }); }
  async readAll() { return SheetsStore.shape(await this.call({ op: 'read' })); }
  async writeOps(ops) { const j = await this.call(null, { ops, read: true }); return { data: SheetsStore.shape(j), applied: j.applied, rejected: j.rejected || [] }; }
  static shape(j) {
    const pad = (r, n) => { const o = r.slice(0, n); while (o.length < n) o.push(''); return o.map(v => v == null ? '' : v); };
    const meta = j.meta || {};
    return { daily: (j.daily || []).map(r => pad(r, 38)), workout: (j.workout || []).map(r => pad(r, 17)), food: (j.food || []).map(r => pad(r, 7)), bench: (j.bench || []).map(r => pad(r, 21)), weekly: (j.weekly || []).map(r => pad(r, 28)),
      dashboard: parseDashboard(j.dashboard || []), labs: meta.labs, forecast: meta.forecast, weeklyNotes: meta.weeklyNotes, startWeight: meta.startWeight, heightCm: meta.heightCm,
      routine: shapeRoutine(j.routine), routineSupported: Array.isArray(j.routine), sheetToday: j.today, sheetName: j.name, tz: j.tz, ts: Date.now(), source: 'sheets', dashboardSource: 'sheets' };
  }
}
function parseDashboard(rows) {
  const map = {}, supp = {}, strength = {}, actions = []; let mode = '';
  rows.forEach(r => {
    const a = r[0] == null ? '' : String(r[0]).trim(), b = r[1]; if (!a) return;
    if (/^SUPPLEMENTS/.test(a)) { mode = 'supp'; return; } if (/^PEPTIDES/.test(a)) mode = ''; if (/^STRENGTH BY/.test(a)) { mode = 'str'; return; } if (/^BENCHMARKS/.test(a)) mode = ''; if (/^ACTION LIST/.test(a)) { mode = 'act'; return; }
    if (mode === 'act') { if (/^(ACTION|WATCH)/.test(a)) actions.push(a); return; }
    if (mode === 'supp' && a !== 'Supplement') supp[a] = { d7: r[1], d28: r[2], tracked: r[3], status: r[4] };
    if (mode === 'str' && a !== 'Family' && !/^STRENGTH OVERRIDE/.test(a)) strength[a] = { top: r[1], e1rm: r[2], m1: r[3], last14: r[4], change: r[5] };
    map[a] = b;
  });
  return { map, supp, strength, actions };
}

/* ---------- DEMO backend ---------- */
class Mock {
  constructor() { this.data = store.get('mock', null) || this.gen(); }
  gen() {
    const daily = [], workout = [], food = [], today = serialOf(2026, 10, 11); let w = 101;
    const sess = ['Upper A', 'Lower A', 'Upper B', 'Lower B', 'Stacked', 'Rest', 'Rest'];
    const plan = { 'Upper A': ['Seated DB Shoulder Press|4|6-8|RIR 1-2, last set RIR 1', 'Incline DB Press (30 deg)|3|8-10|RIR 1, deep stretch', 'Cable Lateral Raise|4|12-15|Last set to failure + drop', 'Machine Chest Press|3|10-12|RIR 0-1', 'Overhead Cable Triceps Extension|3|12-15|Last set failure + drop', 'Rope Pushdown|3|10-12|Last set failure + drop', 'CORE: Side Plank (feet)|3 each side|20-30 s|RIR 2', 'CORE: Cable Crunch|2|12-15|RIR 2', 'CORE: Pallof Press (half-kneeling)|3 each side|10 with 2-s hold|RIR 2', 'CARDIO: 10 min jog + 10 min walk|1|20 min|Zone 2'],
      'Lower A': ['Barbell Back Squat|4|5-8|RIR 2-3 (weeks 1-2), RIR 1-2 (weeks 3-4)', 'Leg Press|3|10-12|RIR 0-1', 'DB Walking Lunge|3|10 ea|RIR 1', 'Leg Extension|3|12-15|Last set failure + drop', 'Seated Leg Curl|3|10-12|Last set failure', 'Standing Calf Raise|4|10-12|Failure, pause at stretch', 'CORE: Dead Bug (heel slides)|3|8 each side|Quality only, RIR 3', 'CORE/HIP: Copenhagen Plank (short lever)|3 each side|15-20 s|RIR 2', 'CORE: Ab Wheel (knees, short range)|2|8-10|RIR 2-3'],
      'Upper B': ['Pull-Up or Lat Pulldown|4|6-10|RIR 1', 'Chest-Supported Row|4|8-10|RIR 0-1', 'Single-Arm Cable Row|3|10-12|RIR 0-1', 'Reverse Pec Deck|3|12-15|Last set failure + drop', 'DB Lateral Raise|3|12-15|Last set failure', 'Incline DB Curl|3|10-12|Last set failure', 'EZ Bar Curl|3|8-12|RIR 0-1', 'Face Pull|3|15-20|RIR 1', 'CORE: Suitcase Carry|3 each side|30 m|RIR 3, 20-24 kg', 'CORE: Cable Chop (half-kneeling)|3 each side|10|RIR 2', 'CARDIO: 10 min jog + 10 min walk|1|20 min|Zone 2'],
      'Lower B': ['Romanian Deadlift|4|6-8|RIR 2', 'Barbell or Machine Hip Thrust|3|8-12|RIR 1', 'Bulgarian Split Squat (DB)|3|8-10 ea|RIR 1', 'Lying Leg Curl|3|10-12|Last set failure + partials', 'Standing Calf Raise|2|12-15|Failure', 'SUPERSET: Cable Curl + Rope Pushdown|3|12 + 12|RIR 0 on both', 'CORE/BACK: 45-degree Back Extension (bodyweight)|3|12-15|RIR 3, 2-s pause at the top', 'CORE/BACK: Bird Dog (3-s holds)|2 each side|6|RIR 4, bodyweight'],
      'Stacked': ['Stacked HIIT + yoga class|1|class|Friday rules: no loaded flexion until the Nov 21 gate'] };
    const fam = n => /Back Squat/.test(n) ? 'Squat' : /Shoulder Press|Overhead Press|Bench Press/.test(n) ? 'Press' : /Pull-Up|Pulldown/.test(n) ? 'Pull' : /Romanian/.test(n) ? 'RDL' : /Trap Bar/.test(n) ? 'Trap bar' : /Barbell Row|T-Bar|Landmine/.test(n) ? 'Row' : /Back Extension|Reverse Hyper/.test(n) ? 'Back ext' : /Suitcase/.test(n) ? 'Carry' : /Hip Thrust/.test(n) ? 'Hip thrust' : '';
    const rnd = (a, b) => a + Math.random() * (b - a);
    for (let i = 0; i < 231; i++) {
      const s = START + i, dt = dateFromSerial(s), wd = (dt.getDay() + 6) % 7, wk = weekOf(s), thu = wd === 3;
      const row = new Array(38).fill('');
      row[D.date] = s; row[D.week] = wk; row[D.day] = dt.toLocaleDateString('en', { weekday: 'short' }); row[D.type] = [9, 18, 26].includes(wk) ? 'Deload' : wk === 25 ? 'Retest' : (s >= serialOf(2026, 9, 29) && s <= serialOf(2026, 10, 11)) ? 'Escalation' : wd <= 4 ? 'Training' : 'Rest';
      row[D.session] = sess[wd]; row[D.kcalT] = wd <= 4 ? 2200 : 2000; row[D.protT] = wd <= 4 ? 190 : 180;
      if (thu && s >= serialOf(2026, 9, 1) && s <= serialOf(2027, 2, 25)) row[D.retaPlan] = s <= serialOf(2026, 9, 22) ? 2 : 4;
      if (s >= serialOf(2026, 10, 19) && s <= serialOf(2027, 3, 30)) row[D.tesaPlan] = s <= serialOf(2026, 11, 16) ? 1 : 2;
      if (s < today) {
        w -= rnd(0.02, 0.16); row[D.weight] = r1(w + rnd(-0.3, 0.3)); row[D.rhr] = Math.round(58 + rnd(-2, 4) + (s > today - 6 ? 3 : 0)); row[D.sleep] = r1(rnd(6.4, 8.1)); row[D.steps] = Math.round(rnd(6000, 11500));
        if (i % 7 === 2) row[D.waist] = r1(108 - i * 0.12);
        [D.multi, D.fish, D.d3, D.creatine].forEach(c => row[c] = Math.random() < 0.9 ? 'Y' : 'N'); row[D.mg] = Math.random() < 0.7 ? 'Y' : 'N'; row[D.biotin] = 'N'; row[D.collagen] = [1, 3, 4].includes(wd) && Math.random() < 0.8 ? 'Y' : 'N';
        row[D.spine] = Math.random() < 0.85 ? 'Y' : 'N'; row[D.sauna] = wd <= 3 && Math.random() < 0.5 ? 'Y' : 'N';
        if (row[D.retaPlan]) { row[D.retaU] = row[D.retaPlan] * 10; row[D.retaSite] = ['Abdomen L', 'Abdomen R', 'Arm L (back)'][i % 3]; }
        if (row[D.tesaPlan] && Math.random() < 0.93) { row[D.tesaU] = row[D.tesaPlan] * 5; row[D.tesaSite] = 'Abdomen R'; row[D.glucose] = r1(rnd(4.8, 5.9)); }
        [['Wake-up', 'Shake + oats + banana', 560, 35], ['Post-gym', 'Whey + protein milk', 380, 54], ['Lunch', 'Eggs + whites + tortilla', 450, 42], ['Snack', 'Greek yogurt + berries', 300, 30], ['Dinner', 'Fed meal (chicken, rice, veg)', 600, 45], ['Pre-sleep', 'Cottage cheese 200 g', 200, 32]].forEach(m => { if (Math.random() < 0.9) food.push([s, m[0], m[1], Math.round(m[2] + rnd(-50, 50)), m[3], '', uuid()]); });
      }
      if (s === today) { row[D.weight] = r1(w - 0.1); row[D.rhr] = 61; row[D.sleep] = 7.4; row[D.multi] = 'Y'; row[D.creatine] = 'Y'; row[D.d3] = 'Y'; row[D.spine] = 'Y'; food.push([s, 'Wake-up', 'Shake + oats + banana', 552, 35, '', uuid()]); food.push([s, 'Post-gym', 'Whey + protein milk', 381, 54, '', uuid()]); }
      daily.push(row);
      const list = plan[sess[wd]]; if (!list) continue;
      list.forEach((p, k) => {
        const [ex, sets, reps, eff] = p.split('|'); const r = new Array(17).fill('');
        r[W.date] = s; r[W.week] = wk; r[W.day] = row[D.day]; r[W.session] = sess[wd]; r[W.n] = k + 1; r[W.ex] = ex; r[W.sets] = sets; r[W.reps] = reps; r[W.effort] = eff; r[W.family] = fam(ex);
        if (s < today && Math.random() < 0.92 && !/^CARDIO|^Stacked/.test(ex)) { const base = { Squat: 100, Press: 30, Pull: 0, RDL: 90, 'Trap bar': 120, Row: 55, 'Back ext': 0, Carry: 24, 'Hip thrust': 100 }[r[W.family]] ?? 30; r[W.load] = base + Math.round(wk * 1.2) + [0, 2.5, 5][k % 3]; r[W.repsTop] = 5 + (k % 6); r[W.setsDone] = +String(sets).replace(/\D.*/, '') || 3; r[W.rir] = 1 + (k % 2); }
        else if (s < today && /^CARDIO|^Stacked/.test(ex) && Math.random() < 0.8) r[W.setsDone] = 1;
        workout.push(r);
      });
    }
    const bench = [[serialOf(2026, 8, 29), 'Baseline', 96, 41, 44, 118, 74, 12, 11, 'Y', 'N', 2, 101, 'BW x 15', 24, '', '', '', '', '', 'record'], [serialOf(2026, 10, 21), '#2 (gate for Month 3 heavy 5s)', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 120], [serialOf(2027, 0, 23), '#3 (gate for 175 kg trap bar)', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 150], [serialOf(2027, 2, 20), '#4 (retest week)', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 150]];
    const d = { daily, workout, food, bench, weekly: [], dashboard: null, routineRows: DEMO_ROUTINE.map(r => r.slice()), ts: Date.now(), source: 'demo', dashboardSource: 'phone' }; this.recompute(d); return d;
  }
  recompute(d) { const r = computeDashboard(d, serialOf(2026, 10, 11)); d.dashboard = r.dashboard; d.weekly = r.weekly; d.routineRows = d.routineRows || DEMO_ROUTINE.map(x => x.slice()); d.routine = shapeRoutine(d.routineRows); d.routineSupported = true; d.ts = Date.now(); store.set('mock', d); return d; }
  async writeRoutine(op) { const rows = this.data.routineRows; if (op.action === 'delete') { const i = rows.findIndex(r => String(r[0]) === String(op.id)); if (i >= 0) rows.splice(i, 1); } else { const it = op.item, i = rows.findIndex(r => String(r[0]) === String(it.id)); const row = i >= 0 ? rows[i].slice() : [it.id, 'Item', 'Other', 0, 0, 'All', 'Y', rows.length + 1, '']; if ('name' in it) row[1] = it.name; if ('meal' in it) row[2] = it.meal; if ('kcal' in it) row[3] = it.kcal; if ('prot' in it) row[4] = it.prot; if ('days' in it) row[5] = it.days; if ('active' in it) row[6] = it.active === false ? 'N' : 'Y'; if ('order' in it) row[7] = it.order; if ('note' in it) row[8] = it.note || ''; if (i >= 0) rows[i] = row; else rows.push(row); } this.recompute(this.data); }
  async readAll() { this.recompute(this.data); return JSON.parse(JSON.stringify(this.data)); }
  async writeCells(sheet, cells) { const tbl = sheet === 'Daily Log' ? this.data.daily : this.data.workout; cells.forEach(c => { const m = c.addr.match(/^([A-Z]+)(\d+)$/); tbl[+m[2] - 2][colIdx(m[1])] = c.value == null ? '' : c.value; }); this.recompute(this.data); }
  async addFoodRow(v) { const i = this.data.food.findIndex(r => r[6] && r[6] === v[6]); if (i >= 0) this.data.food[i] = v.slice(); else this.data.food.push(v.slice()); this.recompute(this.data); }
  async deleteFoodRow(i, expect, id) { const j = this.data.food.findIndex(r => r[6] === id); if (j >= 0) this.data.food.splice(j, 1); this.recompute(this.data); }
  async writeBench(date, fields) { const b = this.data.bench.find(x => x[0] === date); if (!b) return; Object.entries(fields).forEach(([k, v]) => { const ci = BENCH_FIELDS.indexOf(k); if (ci >= 0) b[2 + ci] = v == null ? '' : v; }); this.recompute(this.data); }
}

/* ---------- backend selection ---------- */
let backend = null;
function pairFromHash() {
  if (!location.hash || location.hash.length < 2) return false;
  const p = parsePairing(location.hash.slice(1)); if (!p) return false;
  settings.api = p.api; settings.key = p.key; settings.demo = false; saveSettings(); ['cache', 'mock'].forEach(k => store.del(k)); state.data = null;
  try { history.replaceState(null, '', location.pathname + location.search); } catch { }
  return true;
}
async function initBackend() {
  if (settings.demo) { backend = new Mock(); state.mode = 'demo'; return true; }
  if (!settings.api || !settings.key) return false;
  backend = new SheetsStore(settings.api, settings.key); state.mode = 'sheets'; return true;
}
function disconnect() { ['cache', 'queue', 'mock', 'probe', 'plan', 'planEtag', 'logs', 'logsEtag'].forEach(k => store.del(k)); settings.api = ''; settings.key = ''; settings.demo = false; saveSettings(); location.reload(); }

/* ---------- data helpers ---------- */
const dailyIdx = s => state.data ? state.data.daily.findIndex(r => r[D.date] === s) : -1;
const dailyRow = s => { const i = dailyIdx(s); return i >= 0 ? state.data.daily[i] : null; };
const sheetRow = i => i + 2;
const dash = () => (state.data && state.data.dashboard) || { map: {}, supp: {}, strength: {}, actions: [] };
const dv = k => { const v = dash().map[k]; return v == null ? '' : v; };
const statusCls = t => /^ACTION/.test(t) ? 'action' : /^WATCH/.test(t) ? 'watch' : /^OK/.test(t) ? 'ok' : '';
const statusText = t => String(t).replace(/^(OK|WATCH|ACTION):\s*/, '');
const routineFor = s => { const row = dailyRow(s) || [], rest = row[D.session] === 'Rest'; return (state.data.routine || []).filter(it => it.active && (it.days === 'All' || (it.days === 'Rest') === rest)); };
const routineTicked = (s, id) => state.data.food.find(r => r[F.id] === rtFoodId(s, id)) || null;
const foodFor = s => state.data.food.map((r, i) => ({ r, i })).filter(x => x.r[F.date] === s && !(+x.r[F.kcal] === 0 && /example row/i.test(x.r[F.item])));
const workoutFor = s => state.data.workout.map((r, i) => ({ r, i })).filter(x => x.r[W.date] === s);
function recomputeLocal() { if (!state.data) return; const r = computeDashboard(state.data, todaySerial()); state.data.dashboard = r.dashboard; if (state.data.source !== 'sheets' || state.apiDown) state.data.weekly = r.weekly; state.data.dashboardSource = 'phone'; }

/* ---------- sync + queue ---------- */
const sheetDateOff = d => d && d.source === 'sheets' && num(d.sheetToday) != null && d.sheetToday !== todaySerial();   // the sheet's TODAY() is not the phone's date: compute locally
async function refresh(silent) {
  if (!backend) return; state.busy = true; syncUI();
  try { const d = await backend.readAll(); state.data = d; state.queue.forEach(applyLocal); if (state.queue.length || sheetDateOff(d)) recomputeLocal(); store.set('cache', state.data); state.error = null; state.apiDown = null; if (!silent) toast(settings.demo ? 'Demo data loaded' : 'Synced with Google Sheets'); }
  catch (e) { state.error = e.message; state.apiDown = e.message; if (state.data) recomputeLocal(); if (!silent) toast('Could not sync: ' + e.message, true); }
  state.busy = false; syncUI(); render();
}
async function apply(op) {   // {kind:'cells', sheet, cells} | {kind:'food', values} | {kind:'delfood', index, expect, id} | {kind:'bench', date, fields}
  if (op.kind === 'cells' && state.data) { const tbl = op.sheet === 'Daily Log' ? state.data.daily : state.data.workout; op.cells.forEach(c => { const m = c.addr.match(/^([A-Z]+)(\d+)$/), row = tbl[+m[2] - 2]; if (row) c.key = op.sheet === 'Daily Log' ? String(row[D.date]) : row[W.date] + '|' + row[W.n]; }); }
  applyLocal(op); recomputeLocal(); store.set('cache', state.data);
  state.queue.push(op); store.set('queue', state.queue); render();
  await flush();
}
function applyLocal(op) {
  if (!state.data) return;
  if (op.kind === 'cells') { const tbl = op.sheet === 'Daily Log' ? state.data.daily : state.data.workout; op.cells.forEach(c => { const m = c.addr.match(/^([A-Z]+)(\d+)$/); const row = tbl[+m[2] - 2]; if (row) row[colIdx(m[1])] = c.value == null ? '' : c.value; }); }
  if (op.kind === 'food') { const i = state.data.food.findIndex(r => r[6] && r[6] === op.values[6]); if (i >= 0) state.data.food[i] = op.values.slice(); else state.data.food.push(op.values.slice()); }
  if (op.kind === 'delfood') { const i = state.data.food.findIndex(r => (op.id && r[6] === op.id) || (!op.id && r[F.date] === op.expect[F.date] && String(r[F.item]) === String(op.expect[F.item]))); if (i >= 0) state.data.food.splice(i, 1); }
  if (op.kind === 'bench') { const b = state.data.bench.find(x => x[0] === op.date); if (b) Object.entries(op.fields).forEach(([k, v]) => { const ci = BENCH_FIELDS.indexOf(k); if (ci >= 0) b[2 + ci] = v == null ? '' : v; }); }
  if (op.kind === 'routine') { const list = state.data.routine = state.data.routine || []; if (op.action === 'delete') { const i = list.findIndex(x => x.id === String(op.id)); if (i >= 0) list.splice(i, 1); } else if (op.item) { const i = list.findIndex(x => x.id === String(op.item.id)); const it = Object.assign(i >= 0 ? list[i] : { name: 'Item', meal: 'Other', kcal: 0, prot: 0, active: true, days: 'All', order: list.length + 1, note: '' }, op.item, { id: String(op.item.id) }); if (i >= 0) list[i] = it; else list.push(it); list.sort((a, b) => a.order - b.order); } }
}
let flushing = false;
async function flush() {
  if (flushing || !backend || !state.queue.length) return;
  if (!navigator.onLine && !settings.demo) { syncUI(); return; }
  flushing = true; state.busy = true; syncUI(); let failed = false; const rejected = [];
  try {
    if (backend.writeOps) {
      while (state.queue.length) {
        const ops = state.queue.slice();
        const res = await backend.writeOps(ops);
        state.queue = state.queue.slice(ops.length); store.set('queue', state.queue);
        (res.rejected || []).forEach(rj => { const op = ops[rj.index]; rejected.push('Not saved (' + (op ? op.kind : '?') + '): ' + rj.error); });
        state.data = res.data; state.queue.forEach(applyLocal); if (state.queue.length || sheetDateOff(state.data)) recomputeLocal();
        if (rejected.some(m => /Routine sheet missing/.test(m))) state.data.routineSupported = false;
        store.set('cache', state.data); state.error = null; state.apiDown = null;
      }
      state.rejected = rejected.length ? rejected : null;
      if (rejected.length) toast(rejected[0] + (rejected.length > 1 ? ' (+' + (rejected.length - 1) + ' more)' : ''), true); else toast('Saved to Google Sheets');
    } else {
      while (state.queue.length) {
        const op = state.queue[0];
        if (op.kind === 'cells') await backend.writeCells(op.sheet, op.cells);
        else if (op.kind === 'food') await backend.addFoodRow(op.values);
        else if (op.kind === 'delfood') await backend.deleteFoodRow(op.index, op.expect, op.id);
        else if (op.kind === 'bench') await backend.writeBench(op.date, op.fields);
        else if (op.kind === 'routine') await backend.writeRoutine(op);
        state.queue.shift(); store.set('queue', state.queue);
      }
      await refresh(true); toast('Saved');
    }
  } catch (e) { failed = true; state.error = e.message; state.apiDown = e.message; toast('Saved on the phone, will sync later: ' + e.message, true); }
  flushing = false; state.busy = false; syncUI(); render();
  if (state.queue.length && !failed) flush();
}
function syncUI() {
  const el = document.getElementById('sync'); if (!el) return;
  const t = state.data ? new Date(state.data.ts).toLocaleTimeString('en', { hour: '2-digit', minute: '2-digit' }) : '';
  el.className = 'sync' + (state.busy ? ' busy' : (!backend || state.apiDown || !navigator.onLine || state.queue.length) ? ' off' : '');
  el.textContent = settings.demo ? 'Demo mode' : !backend ? 'Not connected' : state.busy ? 'Syncing…' : state.queue.length ? `${state.queue.length} pending` : state.apiDown ? 'Sheet unreachable' : t ? 'Synced ' + t : 'Not synced';
  el.onclick = backend ? () => { flush(); refresh(); } : null;
}

/* ---------- UI primitives ---------- */
const $ = s => document.querySelector(s);
function toast(msg, err) { const t = $('#toast'); t.textContent = msg; t.className = 'toast show' + (err ? ' err' : ''); clearTimeout(toast._t); toast._t = setTimeout(() => t.className = 'toast', err ? 4500 : 1800); }
function openSheet(html) { $('#sheet').innerHTML = '<div class="handle"></div>' + html; $('#modal').classList.add('open'); }
function closeSheet() { $('#modal').classList.remove('open'); }
function ring(pct, label, val, color) {
  const p = Math.max(0, Math.min(1, pct || 0)), R = 27, C = 2 * Math.PI * R;
  return `<div><div class="ring"><svg viewBox="0 0 66 66"><circle cx="33" cy="33" r="${R}" stroke="rgba(255,255,255,.08)" stroke-width="7" fill="none"/><circle cx="33" cy="33" r="${R}" stroke="${color}" stroke-width="7" fill="none" stroke-linecap="round" stroke-dasharray="${C}" stroke-dashoffset="${C * (1 - p)}" style="transition:stroke-dashoffset .8s cubic-bezier(.2,.8,.2,1)"/></svg><div class="val">${val}<span>${Math.round(p * 100)}%</span></div></div><div class="ringlbl">${label}</div></div>`;
}
const greet = () => { const h = new Date().getHours(); return h < 12 ? 'Good morning' : h < 18 ? 'Good afternoon' : 'Good evening'; };
const dateNav = (which, s) => `<div class="datenav"><button class="iconbtn" data-nav="${which}:-1">‹</button><div class="d">${fmtDate(s, { weekday: 'long', day: 'numeric', month: 'long' })}${s === todaySerial() ? ' <span class="chip teal" style="margin-left:6px">Today</span>' : ''}</div><button class="iconbtn" data-nav="${which}:1" ${s >= todaySerial() ? 'disabled style="opacity:.3"' : ''}>›</button></div>`;
const banner = () => (state.rejected ? `<div class="status action" style="margin-bottom:12px">The sheet refused ${state.rejected.length} entr${state.rejected.length > 1 ? 'ies' : 'y'}: ${esc(state.rejected[0])}<br><span class="xs">They were not saved. This notice clears at the next successful save.</span></div>` : '') + (state.apiDown && !settings.demo ? `<div class="status watch" style="margin-bottom:12px">Google Sheets unreachable: ${esc(state.apiDown)}<br><span class="xs">Showing the last synced data plus your entries. Everything is kept on the phone and uploads when sync works again.${/Key rejected/.test(state.apiDown) ? ' Open Settings and paste the pairing code again.' : ''}</span></div>` : '');

/* ---------- screens ---------- */
function renderSetup() {
  return `<div class="setup">
    <img class="logo" src="icons/icon-192.png" alt="">
    <h1 style="text-align:center;font-size:30px;font-weight:800;letter-spacing:-.02em;margin:0 0 6px">Hamid Coach</h1>
    <p class="muted" style="text-align:center;margin:0 0 22px">Your plan, your logs and your dashboard, kept in the Google Sheet "Hamid Coach" in your Google Drive.</p>
    <div class="card">
      <div class="step"><div class="n">1</div><div><b>Set up the Google Sheet once</b> on the laptop (the App Setup sheet or the README names every click): upload the plan workbook, add the script, run setup, deploy the web app.</div></div>
      <div class="step"><div class="n">2</div><div><b>Copy the pairing code</b> from cell B7 of the App Setup sheet. In the Google Sheets app on the phone: tap the cell, tap it again, Copy.</div></div>
      <div class="step"><div class="n">3</div><div><b>Paste it below and tap Connect.</b> That is all: no account, no sign-in, nothing expires.</div></div>
    </div>
    <div class="card">
      <div class="field"><label>Pairing code (web app URL # key)</label><textarea class="inp" id="s-pair" rows="3" placeholder="https://script.google.com/macros/s/…/exec#…" autocapitalize="off" autocorrect="off" spellcheck="false">${esc(settings.api ? settings.api + '#' + settings.key : '')}</textarea></div>
      <button class="btn primary" id="s-connect">Connect</button>
      <button class="btn ghost" id="s-demo" style="margin-top:8px">Try the demo with sample data</button>
      ${state.error ? `<div class="status action" style="margin-top:12px">${esc(state.error)}</div>` : ''}
    </div>
    <p class="xs dim" style="text-align:center">Version ${APP_VERSION}. Nothing is stored anywhere except your Google Sheet and this phone.</p>
  </div>`;
}
function renderToday() {
  const t = todaySerial(), row = dailyRow(t) || [], d = dash(), items = foodFor(t), suppList = SUPPS.filter(s => !s.warn);
  const kcal = num(row[D.kcalM]) ?? items.reduce((a, x) => a + (num(x.r[F.kcal]) || 0), 0), kT = num(row[D.kcalT]) || 2200;
  const prot = num(row[D.protM]) ?? items.reduce((a, x) => a + (num(x.r[F.prot]) || 0), 0), pT = num(row[D.protT]) || 190;
  const tracked = suppList.length, supps = suppList.filter(s => row[D[s.k]] === 'Y').length;
  const wo = workoutFor(t), done = wo.filter(x => x.r[W.done] === 1 || !isBlank(x.r[W.load]) || !isBlank(x.r[W.setsDone])).length;
  const w7 = num(dv('7-day average (kg)')), wlatest = num(dv('Latest weight (kg)')), chg = num(dv('Change this week (kg)')), weightLogged = num(row[D.weight]);
  const reta = num(row[D.retaPlan]), tesa = num(row[D.tesaPlan]), nextLab = LABS().find(l => l[0] >= t), actions = d.actions || [];
  return `<div class="hdr"><div><h1>${greet()}, ${esc(settings.name)}</h1><div class="sub">${fmtDate(t, { weekday: 'long', day: 'numeric', month: 'long' })}</div></div><button class="iconbtn" data-open="settings">⚙︎</button></div>
  <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px"><span class="chip violet">Week ${dv('Tracker week') || weekOf(t)}</span><span class="chip">${esc(dv('Phase') || '')}</span><span class="chip">${esc(row[D.session] || '')}${row[D.type] && row[D.type] !== 'Training' && row[D.type] !== 'Rest' ? ' · ' + esc(row[D.type]) : ''}</span></div>
  ${banner()}
  ${!row.length ? '<div class="status watch">Today is outside the plan window (Sep 28 2026 to May 16 2027).</div>' : ''}
  ${sheetDateOff(state.data) && !state.apiDown ? `<div class="status watch">The Google Sheet thinks today is ${fmtDate(state.data.sheetToday)}, not ${fmtDate(t)}: its time zone is ${esc(state.data.tz || 'different')}. The dashboard below is computed on the phone for ${fmtDate(t)}; set the sheet's time zone under File, Settings.</div>` : ''}
  <div class="card hero">
    <div class="row"><div><h3 style="margin-bottom:6px">Weight</h3><div class="big">${w7 != null ? w7.toFixed(1) : (wlatest != null ? wlatest.toFixed(1) : '—')}<small>kg · 7-day avg</small></div>
      <div class="small muted" style="margin-top:6px">${chg != null ? `<span class="delta ${chg < 0 ? 'down' : 'up'}">${chg > 0 ? '+' : ''}${chg.toFixed(1)} kg this week</span> · ` : ''}forecast ${esc(dv('Forecast band') || '—')}${num(dv('Total lost since 101 kg')) != null ? ` · lost ${fmt(dv('Total lost since 101 kg'), 1)} kg` : ''}</div></div>
    </div>
    <div class="inline" style="margin-top:14px;align-items:flex-end"><div class="field rel" style="margin:0;flex:1"><label>${weightLogged != null ? 'Logged this morning' : 'Log this morning\'s weight'}</label><input class="inp num" id="t-weight" type="text" inputmode="decimal" placeholder="0.0" value="${weightLogged != null ? weightLogged : ''}"><span class="unit">kg</span></div><button class="btn primary" id="t-weight-save" style="width:auto;padding:14px 18px">Save</button></div>
  </div>
  <div class="card"><div class="grid4">${ring(kcal / kT, 'Calories', `${fmt(kcal)}`, 'var(--teal)')}${ring(prot / pT, 'Protein', `${fmt(prot)}g`, 'var(--violet)')}${ring(supps / tracked, 'Supps', `${supps}/${tracked}`, 'var(--blue)')}${ring(wo.length ? done / wo.length : (row[D.session] === 'Rest' ? 1 : 0), row[D.session] === 'Rest' ? 'Rest day' : 'Session', wo.length ? `${done}/${wo.length}` : '—', 'var(--green)')}</div>
    <div class="grid2" style="margin-top:12px"><button class="btn" data-go="log">Daily log</button><button class="btn" data-go="food">Add food</button></div></div>
  ${(reta || tesa) ? `<div class="card"><h3>Tonight</h3>
    ${reta ? `<div class="item"><div><div class="t">Retatrutide ${reta} mg = ${reta * 10} units</div><div class="s">Evening, after the session and the sauna. ${num(row[D.retaU]) ? `<b style="color:var(--green)">Logged: ${row[D.retaU]} units · ${esc(row[D.retaSite] || '')}</b>` : 'Not logged yet.'}</div></div><button class="btn sm ${num(row[D.retaU]) ? '' : 'primary'}" data-dose="reta">${num(row[D.retaU]) ? 'Edit' : 'Log dose'}</button></div>` : ''}
    ${tesa ? `<div class="item"><div><div class="t">Tesamorelin ${tesa} mg = ${tesa * 5} units</div><div class="s">Bedtime, 2 h after the pre-sleep protein, abdomen. ${num(row[D.tesaU]) ? `<b style="color:var(--green)">Logged: ${row[D.tesaU]} units · ${esc(row[D.tesaSite] || '')}</b>` : 'Only if the gates passed.'}</div></div><button class="btn sm ${num(row[D.tesaU]) ? '' : 'primary'}" data-dose="tesa">${num(row[D.tesaU]) ? 'Edit' : 'Log dose'}</button></div>` : ''}
  </div>` : ''}
  ${wo.length ? `<div class="card"><div class="row"><div><h3 style="margin-bottom:2px">${esc(row[D.session])}</h3><div class="small muted">${wo.length} planned · ${done} logged${row[D.type] === 'Deload' ? ' · deload: half sets, RIR 3' : ''}${row[D.type] === 'Escalation' ? ' · escalation window: hinges RIR 2, no loaded extension' : ''}</div></div><button class="btn sm primary" data-go="train">${done ? 'Continue' : 'Start'}</button></div></div>` : ''}
  <div class="card"><h3>Coach says</h3>${actions.length ? actions.map(a => `<div class="status ${statusCls(a)}">${esc(a)}</div>`).join('') : '<div class="status ok">Nothing triggered. Keep logging.</div>'}
    <div class="xs dim" style="margin-top:6px">${state.data.dashboardSource === 'sheets' ? 'Computed by your Google Sheet from your logs.' : 'Computed from your logs with the plan\'s rules.'} Full detail under Progress.</div></div>
  <div class="card"><h3>Coming up</h3><div class="kv">${nextLab ? `<div>${esc(nextLab[1])}</div><b>${fmtDate(nextLab[0])}</b>` : ''}<div>Next retatrutide</div><b>${esc(dv('Next planned dose') || '—')}</b><div>Next benchmark</div><b>${esc(dv('Next test') || '—')}</b></div></div>`;
}
function renderLog() {
  const s = state.logDate || todaySerial(), row = dailyRow(s);
  if (!row) return `<div class="hdr"><h1>Daily log</h1></div>${dateNav('log', s)}<div class="card">No plan row for this date.</div>`;
  const f = (id, label, col, unit) => `<div class="field rel"><label>${label}</label><input class="inp num" id="${id}" data-col="${col}" type="text" inputmode="decimal" value="${isBlank(row[col]) ? '' : row[col]}"><span class="unit">${unit}</span></div>`;
  return `<div class="hdr"><div><h1>Daily log</h1><div class="sub">Weight, heart rate, supplements, notes</div></div></div>
  ${dateNav('log', s)}${banner()}
  <div class="card"><h3>Morning</h3><div class="inline">${f('l-weight', 'Weight', D.weight, 'kg')}${f('l-rhr', 'Resting HR', D.rhr, 'bpm')}</div><div class="inline">${f('l-waist', 'Waist', D.waist, 'cm')}${f('l-glucose', 'Fasting glucose', D.glucose, 'mmol/L')}</div><div class="inline">${f('l-sleep', 'Sleep', D.sleep, 'h')}${f('l-steps', 'Steps', D.steps, 'steps')}</div></div>
  <div class="card"><h3>Supplements</h3><div class="toggles">${SUPPS.map(sp => `<button class="tg ${sp.warn ? 'warn' : ''} ${row[D[sp.k]] === 'Y' ? 'on' : ''}" data-tg="${sp.k}"><span class="dot">${row[D[sp.k]] === 'Y' ? '✓' : ''}</span>${sp.i} ${sp.n}</button>`).join('')}</div>
    <div class="xs dim" style="margin-top:8px">Biotin is on the plan's stop list: it corrupts the blood tests. Logging it keeps the lab-timing check honest.</div></div>
  <div class="card"><h3>Routine</h3><div class="toggles"><button class="tg ${row[D.spine] === 'Y' ? 'on' : ''}" data-tg="spine"><span class="dot">${row[D.spine] === 'Y' ? '✓' : ''}</span>🧘 Spine routine</button><button class="tg ${row[D.sauna] === 'Y' ? 'on' : ''}" data-tg="sauna"><span class="dot">${row[D.sauna] === 'Y' ? '✓' : ''}</span>🔥 Sauna</button></div></div>
  <div class="card"><h3>Peptides</h3><div class="kv"><div>Retatrutide</div><b>${num(row[D.retaPlan]) ? `${row[D.retaPlan]} mg planned · ${num(row[D.retaU]) ? row[D.retaU] + ' units logged' : 'not logged'}` : 'not today'}</b><div>Tesamorelin</div><b>${num(row[D.tesaPlan]) ? `${row[D.tesaPlan]} mg planned · ${num(row[D.tesaU]) ? row[D.tesaU] + ' units logged' : 'not logged'}` : 'not planned'}</b></div>
    <div class="grid2" style="margin-top:10px">${num(row[D.retaPlan]) ? '<button class="btn sm" data-dose="reta" style="width:100%">Retatrutide dose</button>' : ''}${num(row[D.tesaPlan]) ? '<button class="btn sm" data-dose="tesa" style="width:100%">Tesamorelin dose</button>' : ''}</div></div>
  <div class="card"><div class="field"><label>Notes, symptoms, energy</label><textarea class="inp" id="l-notes" rows="3" data-col="${D.notes}">${esc(String(row[D.notes] || '').trim())}</textarea></div></div>
  <button class="btn primary" id="l-save">Save day</button>`;
}
function renderTrain() {
  const s = state.trainDate || todaySerial(), wo = workoutFor(s), row = dailyRow(s) || [];
  const prevFor = ex => { const rows = state.data.workout.filter(r => r[W.ex] === ex && r[W.date] < s && num(r[W.load]) != null); const l = rows[rows.length - 1]; const best = rows.reduce((a, r) => Math.max(a, num(r[W.e1rm]) || 0), 0); return l ? `Last: ${l[W.load]} kg × ${l[W.repsTop] || '?'} · ${l[W.setsDone] || '?'} sets${best ? ` · best e1RM ${Math.round(best)}` : ''}` : ''; };
  return `<div class="hdr"><div><h1>Train</h1><div class="sub">${esc(row[D.session] || '')}${row[D.type] && !['Training', 'Rest'].includes(row[D.type]) ? ' · ' + esc(row[D.type]) : ''}</div></div></div>
  ${dateNav('train', s)}${banner()}
  ${!wo.length ? `<div class="card"><b>Rest day.</b><div class="muted small" style="margin-top:4px">Spine routine, steps, sleep. Log them under Daily log.</div></div>` : ''}
  ${wo.map(({ r, i }) => { const core = /^CORE|^SUPERSET|^GIANT/.test(r[W.ex]), cardio = /^CARDIO|^Stacked/.test(r[W.ex]), isDone = r[W.done] === 1 || !isBlank(r[W.load]) || !isBlank(r[W.setsDone]); return `<div class="ex ${core ? 'core' : ''} ${isDone ? 'done' : ''}" data-row="${i}">
    <div class="row"><div><div class="name">${esc(r[W.ex])}</div><div class="plan">${esc(r[W.sets])} × ${esc(r[W.reps])}${r[W.effort] ? ' · ' + esc(r[W.effort]) : ''}</div>${!cardio ? `<div class="prev">${esc(prevFor(r[W.ex]))}</div>` : ''}</div>${isDone ? '<span class="pill ok">done</span>' : ''}</div>
    ${cardio ? `<div style="margin-top:10px"><button class="btn sm ${r[W.setsDone] ? '' : 'primary'}" data-cardio="${i}">${r[W.setsDone] ? 'Done ✓ (tap to undo)' : 'Mark done'}</button></div>` : `<div class="inputs"><div><label>Load kg</label><input class="inp" type="text" inputmode="decimal" data-wcol="${W.load}" value="${isBlank(r[W.load]) ? '' : r[W.load]}"></div><div><label>Top reps</label><input class="inp" type="text" inputmode="numeric" data-wcol="${W.repsTop}" value="${isBlank(r[W.repsTop]) ? '' : r[W.repsTop]}"></div><div><label>Sets</label><input class="inp" type="text" inputmode="numeric" data-wcol="${W.setsDone}" value="${isBlank(r[W.setsDone]) ? '' : r[W.setsDone]}"></div><div><label>RIR</label><input class="inp" type="text" inputmode="numeric" data-wcol="${W.rir}" value="${isBlank(r[W.rir]) ? '' : r[W.rir]}"></div></div>`}
  </div>`; }).join('')}
  ${wo.length ? `<button class="btn primary" id="w-save">Save workout</button>` : ''}`;
}
function renderFood() {
  const s = state.foodDate || todaySerial(), row = dailyRow(s) || [], items = foodFor(s);
  const kcal = items.reduce((a, x) => a + (num(x.r[F.kcal]) || 0), 0), prot = items.reduce((a, x) => a + (num(x.r[F.prot]) || 0), 0), kT = num(row[D.kcalT]) || 2200, pT = num(row[D.protT]) || 190;
  const manual = num(row[D.kcalM]), shownK = manual ?? kcal, shownP = num(row[D.protM]) ?? prot, floor = num(dv('Calorie floor now')) || 2000;
  const recent = []; const seen = new Set(); [...state.data.food].reverse().forEach(r => { const k = String(r[F.item]).trim().toLowerCase(); if (!k || seen.has(k) || /example row/i.test(k)) return; seen.add(k); recent.push(r); }); recent.splice(12);
  const hour = new Date().getHours(); const meal = state.meal || (hour < 10 ? 'Wake-up' : hour < 13 ? 'Lunch' : hour < 17 ? 'Snack' : hour < 21 ? 'Dinner' : 'Pre-sleep');
  return `<div class="hdr"><div><h1>Food</h1><div class="sub">Log each item with its calories; the day is summed for you</div></div></div>
  ${dateNav('food', s)}${banner()}
  <div class="card"><div class="row"><div><h3 style="margin-bottom:4px">Calories</h3><div class="big" style="font-size:34px">${fmt(shownK)}<small>/ ${fmt(kT)}</small></div></div><div style="text-align:right"><h3 style="margin-bottom:4px">Protein</h3><div class="big" style="font-size:34px">${fmt(shownP)}<small>/ ${pT} g</small></div></div></div>
    <div class="bar ${shownK > kT + 200 ? 'warn' : ''}" style="margin-top:10px"><i style="width:${Math.min(100, shownK / kT * 100)}%"></i></div>
    <div class="bar" style="margin-top:6px"><i style="width:${Math.min(100, shownP / pT * 100)}%;background:var(--violet)"></i></div>
    <div class="xs dim" style="margin-top:8px">${shownK < floor && s === todaySerial() ? `Floor is ${fmt(floor)} kcal: below it a higher fraction of the loss is lean tissue.` : 'Protein first at every meal. Shakes count, collagen does not.'} ${manual != null ? '<br>A manual calorie total is set for this day (overrides the item sum).' : ''}</div></div>
  ${renderRoutineCard(s, row)}
  <div class="card"><h3>Anything else</h3>
    <div class="seg" style="margin-bottom:10px">${['Wake-up', 'Post-gym', 'Lunch', 'Snack', 'Dinner', 'Pre-sleep'].map(m => `<button class="${m === meal ? 'on' : ''}" data-meal="${m}">${m.replace('Wake-up', 'AM').replace('Post-gym', 'Gym').replace('Pre-sleep', 'Bed')}</button>`).join('')}</div>
    ${recent.length ? `<div class="quick">${recent.map(r => `<button data-quick="${esc(r[F.item])}|${r[F.kcal]}|${r[F.prot]}">${esc(r[F.item])}<small>${fmt(r[F.kcal])} kcal · ${fmt(r[F.prot])} g</small></button>`).join('')}</div>` : ''}
    <div class="field"><label>Item</label><input class="inp" id="f-item" placeholder="e.g. Whey + protein milk" autocapitalize="sentences" value="${esc(state.foodDraft.item)}"></div>
    <div class="inline"><div class="field rel"><label>Calories</label><input class="inp num" id="f-kcal" type="text" inputmode="numeric" placeholder="0" value="${esc(state.foodDraft.kcal)}"><span class="unit">kcal</span></div><div class="field rel"><label>Protein</label><input class="inp num" id="f-prot" type="text" inputmode="numeric" placeholder="0" value="${esc(state.foodDraft.prot)}"><span class="unit">g</span></div></div>
    <div class="field"><label>Note (optional)</label><input class="inp" id="f-note" placeholder="e.g. shared a dessert, half portion" autocapitalize="sentences" value="${esc(state.foodDraft.note)}"></div>
    <button class="btn primary" id="f-add">Add to ${esc(meal)}</button>
    <div class="row" style="margin-top:10px"><span class="xs muted">Or set the day's total by hand</span><button class="btn sm ghost" id="f-manual">${manual != null ? 'Edit total' : 'Enter total'}</button></div></div>
  <div class="card"><h3>${fmtDate(s)} · ${items.length} items</h3>${items.length ? items.map(x => `<div class="item"><div><div class="t">${esc(x.r[F.item])}</div><div class="s">${esc(x.r[F.meal])} · ${fmt(x.r[F.kcal])} kcal · ${fmt(x.r[F.prot])} g${/^rt\|/.test(String(x.r[F.id])) ? ' · routine' : ''}${x.r[F.note] && x.r[F.note] !== 'routine' ? ' · ' + esc(x.r[F.note]) : ''}</div></div><button class="iconbtn" data-delfood="${x.i}" style="width:34px;height:34px">✕</button></div>`).join('') : '<div class="muted small">Nothing logged yet.</div>'}</div>`;
}
function renderRoutineCard(s, row) {
  if (state.data.routineSupported === false) return `<div class="card"><h3>My routine</h3><div class="status watch">The Google Sheet needs the newer script: in the script editor paste the new HamidCoach.gs, choose "upgrade" and Run, then Deploy, Manage deployments, pencil, Version: New version, Deploy.</div></div>`;
  const list = routineFor(s), rest = row[D.session] === 'Rest';
  if (!list.length) return `<div class="card"><h3>My routine</h3><div class="muted small">No routine items yet. Add your usual foods under Settings, "Edit my routine".</div></div>`;
  let k = 0, p = 0, n = 0;
  const rows = list.map(it => { const f = routineTicked(s, it.id); const on = !!f; const kc = on ? num(f[F.kcal]) || 0 : it.kcal, pr = on ? num(f[F.prot]) || 0 : it.prot; if (on) { k += kc; p += pr; n++; }
    return `<div class="item rt ${on ? 'on' : ''}" data-rt="${esc(it.id)}"><span class="dot">${on ? '✓' : ''}</span><div style="flex:1;min-width:0"><div class="t">${esc(it.name)}</div><div class="s">${esc(it.meal)} · ${fmt(kc)} kcal · ${fmt(pr)} g${on && (kc !== it.kcal || pr !== it.prot) ? ' (adjusted)' : ''}</div></div><button class="iconbtn" data-rtadj="${esc(it.id)}" title="Adjust today's amount" style="width:34px;height:34px;flex:none">±</button></div>`; }).join('');
  return `<div class="card"><div class="row"><h3>My routine · ${rest ? 'rest day' : 'training day'}</h3><span class="xs muted">${n}/${list.length} · ${fmt(k)} kcal · ${fmt(p)} g</span></div><div class="xs dim" style="margin:-4px 0 10px">Tap what you had. Tap again to undo. ± changes today's amount only.</div>${rows}</div>`;
}
function renderProgress() {
  const d = dash(), wk = state.data.weekly.filter(r => r[4] !== '' || r[0] <= weekOf(todaySerial()));
  const keys = STATUS_KEYS.map(k => [k, dv(k)]).filter(x => x[1] !== '');
  const str = Object.entries(d.strength || {}).filter(([, v]) => v.top);
  const b = state.data.bench.filter(r => num(r[2]) != null).pop(), nextB = state.data.bench.find(r => r[0] >= todaySerial() - 7 && num(r[2]) == null);
  const supp = Object.entries(d.supp || {});
  return `<div class="hdr"><div><h1>Progress</h1><div class="sub">${state.data.dashboardSource === 'sheets' ? 'What your Google Sheet computes from your logs' : 'Computed from your logs with the plan\'s rules'}</div></div><button class="iconbtn" id="p-refresh">↻</button></div>
  ${banner()}
  <div class="card"><h3>Weight vs forecast</h3><canvas id="c-weight" height="190"></canvas>
    <div class="kv" style="margin-top:10px"><div>Total lost</div><b>${fmt(dv('Total lost since 101 kg'), 1)} kg</b><div>Change this week</div><b>${fmt(dv('Change this week (kg)'), 1)} kg</b><div>Ceiling this week</div><b>${fmt(dv('Rate ceiling this week (kg/week)'), 2, 0)} kg/week</b><div>Waist</div><b>${num(dv('Waist (latest, cm)')) != null ? fmt(dv('Waist (latest, cm)'), 1) + ' cm (' + fmt(dv('Waist-to-height'), 2) + ' WtH)' : '—'}</b></div></div>
  <div class="card"><h3>Resting heart rate</h3><canvas id="c-rhr" height="150"></canvas><div class="kv" style="margin-top:10px"><div>Baseline</div><b>${fmt(dv('Baseline (mean Sep 28 - Oct 1)'), 1)} bpm</b><div>7-day mean</div><b>${fmt(dv('7-day mean'), 1)} bpm</b></div></div>
  <div class="card"><h3>Status</h3>${keys.length ? keys.map(([k, v]) => `<div class="item"><div><div class="t">${STATUS_TITLES[k]}</div><div class="s">${esc(statusText(v))}</div></div><span class="pill ${statusCls(v) || 'watch'}">${statusCls(v) || 'n/a'}</span></div>`).join('') : '<div class="muted small">No status yet. Log a few days.</div>'}</div>
  <div class="card"><h3>Adherence · last 7 / 28 days</h3>${supp.map(([k, v]) => `<div class="item"><div><div class="t">${esc(k.replace(/ \(.*\)/, ''))}</div><div class="s">${esc(statusText(v.status))}</div></div><b style="font-variant-numeric:tabular-nums">${num(v.d7) == null ? '—' : Math.round(v.d7 * 100) + '%'} <span class="muted">/ ${num(v.d28) == null ? '—' : Math.round(v.d28 * 100) + '%'}</span></b></div>`).join('')}<div class="item"><div class="t">Spine routine</div><b>${fmt(dv('Spine routine days, last 7'))}/7</b></div><div class="item"><div class="t">Sessions</div><b>${fmt(dv('Sessions done, cumulative'))} / ${fmt(dv('Sessions planned to date'))}</b></div><div class="item"><div class="t">Exercises logged</div><b>${esc(dv('Exercises logged / planned to date') || '—')}</b></div></div>
  <div class="card"><h3>Strength · est. 1RM by lift family</h3>${str.length ? `<table class="tbl"><tr><th>Lift</th><th>Latest</th><th>e1RM</th><th>Month 1</th><th>Change</th></tr>${str.map(([k, v]) => `<tr><td>${esc(k)}</td><td>${esc(v.top)}</td><td>${fmt(v.e1rm)}</td><td>${fmt(v.m1)}</td><td style="color:${num(v.change) == null ? 'inherit' : v.change < -0.05 ? 'var(--red)' : v.change > 0 ? 'var(--green)' : 'inherit'}">${num(v.change) == null ? '—' : (v.change > 0 ? '+' : '') + Math.round(v.change * 100) + '%'}</td></tr>`).join('')}</table>` : '<div class="muted small">Log a few sessions to see this.</div>'}</div>
  <div class="card"><h3>Peptides</h3><div class="kv"><div>Retatrutide doses</div><b>${fmt(dv('Retatrutide doses logged'))} / ${fmt(dv('Retatrutide doses planned to date'))}</b><div>Last dose</div><b>${num(dv('Last dose (date)')) != null ? fmtDate(num(dv('Last dose (date)'))) + ' · ' + fmt(dv('Last dose (units)')) + ' u' : '—'}</b><div>Tesamorelin nights</div><b>${fmt(dv('Tesamorelin nights logged'))} / ${fmt(dv('Tesamorelin nights planned to date'))}</b><div>Fasting glucose (last)</div><b>${num(dv('Last fasting glucose (mmol/L)')) != null ? fmt(dv('Last fasting glucose (mmol/L)'), 1) + ' mmol/L' : '—'}</b></div></div>
  <div class="card"><div class="row"><h3>Benchmarks</h3><button class="btn sm" id="b-log">Log a test</button></div>${b ? `<div class="kv"><div>${esc(b[1])}</div><b>${fmtDate(b[0])}</b><div>Sorensen</div><b>${fmt(b[2])} s <span class="muted">/ target ${b[20] || 120}</span></b><div>Side plank R / L</div><b>${fmt(b[3])} / ${fmt(b[4])} s</b><div>Flexion : extension</div><b>${num(b[16]) != null ? fmt(b[16], 2) : '—'} <span class="muted">/ under 1.0</span></b><div>Pain (0-10)</div><b>${fmt(b[11])}</b></div>` : '<div class="muted small">Baseline battery Sep 29 or 30, then Sat Nov 21, Sat Jan 23, Sat Mar 20.</div>'}<div class="xs dim" style="margin-top:8px">Next: ${esc(dv('Next test') || '—')}${nextB ? ' · ' + esc(nextB[1]) : ''}</div></div>
  <div class="card"><h3>Weekly</h3><table class="tbl"><tr><th>Wk</th><th>Weight</th><th>Δ</th><th>RHR</th><th>kcal</th><th>Prot</th><th>Supp</th><th>Sess</th></tr>${wk.slice(-10).map(r => `<tr><td>${r[0]} <span class="dim">${esc(String(r[3]).replace('Month ', 'M').replace('Taper / maintenance', 'Taper'))}</span></td><td>${fmt(r[4], 1)}</td><td>${fmt(r[5], 1)}</td><td>${fmt(r[7])}</td><td>${fmt(r[9])}</td><td>${fmt(r[12])}</td><td>${num(r[14]) == null ? '—' : Math.round(r[14] * 100) + '%'}</td><td>${fmt(r[17])}/${fmt(r[18])}</td></tr>`).join('')}</table></div>`;
}
const charts = {};
function drawCharts() {
  if (typeof Chart === 'undefined' || !state.data) return;
  const wk = state.data.weekly, labels = wk.map(r => 'W' + r[0]);
  const mk = (id, cfg) => { const el = document.getElementById(id); if (charts[id]) { charts[id].destroy(); delete charts[id]; } if (el) charts[id] = new Chart(el, cfg); };
  const grid = { color: 'rgba(255,255,255,.06)' }, tick = { color: '#8B98AD', font: { size: 10 } };
  mk('c-weight', { type: 'line', data: { labels, datasets: [
    { label: 'Forecast high', data: wk.map(r => num(r[26])), borderColor: 'rgba(139,92,246,.25)', backgroundColor: 'rgba(139,92,246,.10)', fill: '+1', pointRadius: 0, borderWidth: 1, tension: .3 },
    { label: 'Forecast low', data: wk.map(r => num(r[25])), borderColor: 'rgba(139,92,246,.25)', pointRadius: 0, borderWidth: 1, tension: .3 },
    { label: 'Weekly average', data: wk.map(r => num(r[4])), borderColor: '#2DD4BF', backgroundColor: '#2DD4BF', pointRadius: 3, borderWidth: 2.5, tension: .3, spanGaps: false }] },
    options: { plugins: { legend: { display: false } }, scales: { x: { grid, ticks: Object.assign({ maxTicksLimit: 8 }, tick) }, y: { grid, ticks: tick, suggestedMin: 84, suggestedMax: 102 } }, animation: { duration: 600 } } });
  mk('c-rhr', { type: 'line', data: { labels, datasets: [{ label: 'RHR', data: wk.map(r => num(r[7])), borderColor: '#F472B6', pointRadius: 2, borderWidth: 2, tension: .3 }] }, options: { plugins: { legend: { display: false } }, scales: { x: { grid, ticks: Object.assign({ maxTicksLimit: 8 }, tick) }, y: { grid, ticks: tick } } } });
}
function renderSettingsSheet() {
  openSheet(`<h2>Settings</h2>
    <div class="field"><label>Your name</label><input class="inp" id="st-name" value="${esc(settings.name)}"></div>
    <div class="field"><label>Pairing code (web app URL # key)</label><textarea class="inp" id="st-pair" rows="3" autocapitalize="off" autocorrect="off" spellcheck="false">${esc(settings.api ? settings.api + '#' + settings.key : '')}</textarea></div>
    <div class="grid2"><button class="btn" id="st-save">Save</button><button class="btn" id="st-refresh">Refresh data</button></div>
    <button class="btn" id="st-routine" style="margin-top:8px">Edit my routine (usual foods)</button>
    <div class="grid2" style="margin-top:8px"><button class="btn ${settings.demo ? 'primary' : ''}" id="st-demo">${settings.demo ? 'Leave demo mode' : 'Demo mode'}</button><button class="btn danger" id="st-signout">${settings.demo ? 'Reset demo data' : 'Disconnect'}</button></div>
    <div class="xs dim" style="margin-top:14px">${settings.demo ? 'Demo mode: sample data on this phone only.' : state.data && state.data.sheetName ? 'Connected to the Google Sheet "' + esc(state.data.sheetName) + '"' + (state.data.tz ? ' (time zone ' + esc(state.data.tz) + ')' : '') + '.' : 'Not connected yet.'} Version ${APP_VERSION}. Pending writes: ${state.queue.length}. App address: ${esc(APP_URL)}. ${state.error ? 'Last error: ' + esc(state.error) : ''}</div>`);
}
function doseSheet(kind, s) {
  const row = dailyRow(s); if (!row) return;
  const mg = num(row[kind === 'reta' ? D.retaPlan : D.tesaPlan]) || 0, units = kind === 'reta' ? mg * 10 : mg * 5;
  const cur = num(row[kind === 'reta' ? D.retaU : D.tesaU]), curSite = row[kind === 'reta' ? D.retaSite : D.tesaSite] || '';
  openSheet(`<h2>${kind === 'reta' ? 'Retatrutide' : 'Tesamorelin'} · ${fmtDate(s)}</h2>
    <div class="status ok" style="margin-bottom:12px">${mg} mg planned = <b>${units} units</b> on a ${kind === 'reta' ? '0.5 mL syringe (10 mg/mL)' : '0.3 mL syringe (20 mg/mL)'}. ${kind === 'reta' ? 'Evening, after the session and sauna, never the thigh on a Thursday.' : 'Bedtime, 2 h after the pre-sleep protein, abdomen only, rotate sites.'}</div>
    <div class="field rel"><label>Units drawn</label><input class="inp num" id="d-units" type="text" inputmode="numeric" value="${cur != null ? cur : units}"><span class="unit">units</span></div>
    <div class="field"><label>Site</label><div class="site" id="d-sites">${SITES.filter(x => kind === 'tesa' ? /Abdomen/.test(x) : true).map(x => `<button data-site="${x}" class="${x === curSite ? 'on' : ''}">${x}</button>`).join('')}</div></div>
    <div class="grid2" style="margin-top:8px"><button class="btn ghost" id="d-clear">Clear</button><button class="btn primary" id="d-save">Save dose</button></div>`);
  $('#d-sites').addEventListener('click', e => { const b = e.target.closest('button'); if (!b) return; [...$('#d-sites').children].forEach(x => x.classList.remove('on')); b.classList.add('on'); });
  const uCol = kind === 'reta' ? D.retaU : D.tesaU, sCol = kind === 'reta' ? D.retaSite : D.tesaSite, r = sheetRow(dailyIdx(s));
  $('#d-save').onclick = () => { const u = num($('#d-units').value); const on = $('#d-sites .on'); apply({ kind: 'cells', sheet: 'Daily Log', cells: [{ addr: colL(uCol) + r, value: u }, { addr: colL(sCol) + r, value: on ? on.dataset.site : null }] }); closeSheet(); };
  $('#d-clear').onclick = () => { apply({ kind: 'cells', sheet: 'Daily Log', cells: [{ addr: colL(uCol) + r, value: null }, { addr: colL(sCol) + r, value: null }] }); closeSheet(); };
}
function routineEditorSheet() {
  if (!state.data) return;
  const MEAL_OPTS = ['Wake-up', 'Pre-lift', 'Post-gym', 'Lunch', 'Snack', 'Dinner', 'Pre-sleep', 'Cheat meal', 'Other'];
  const list = (state.data.routine || []).map(x => Object.assign({}, x)); const removed = new Set();
  const rowHtml = (it, i) => `<div class="card rtedit" data-i="${i}" style="padding:12px">
      <div class="field" style="margin-bottom:8px"><label>Item</label><input class="inp" data-rf="name" value="${esc(it.name)}"></div>
      <div class="inline"><div class="field rel" style="margin-bottom:8px"><label>Calories</label><input class="inp num" data-rf="kcal" type="text" inputmode="numeric" value="${it.kcal}"><span class="unit">kcal</span></div><div class="field rel" style="margin-bottom:8px"><label>Protein</label><input class="inp num" data-rf="prot" type="text" inputmode="numeric" value="${it.prot}"><span class="unit">g</span></div></div>
      <div class="inline"><div class="field" style="margin-bottom:8px"><label>Meal</label><select class="inp" data-rf="meal">${(MEAL_OPTS.includes(it.meal) ? MEAL_OPTS : [it.meal, ...MEAL_OPTS]).map(m => `<option ${m === it.meal ? 'selected' : ''}>${esc(m)}</option>`).join('')}</select></div><div class="field" style="margin-bottom:8px"><label>Days</label><select class="inp" data-rf="days">${['All', 'Training', 'Rest'].map(m => `<option ${m === it.days ? 'selected' : ''}>${m}</option>`).join('')}</select></div></div>
      <div class="row"><label class="small"><input type="checkbox" data-rf="active" ${it.active ? 'checked' : ''}> Active</label><button class="btn sm ghost" data-rtdel="${i}">Remove</button></div></div>`;
  openSheet(`<h2>My routine</h2><div class="xs muted" style="margin-bottom:10px">Your usual foods with their calories and protein. Days: All, Training (gym days) or Rest. The same list lives on the Routine sheet.</div>
    <div id="rt-list">${list.map(rowHtml).join('')}</div>
    <button class="btn ghost" id="rt-add">+ Add an item</button>
    <button class="btn primary" id="rt-save" style="margin-top:8px">Save routine</button>`);
  const read = () => { document.querySelectorAll('#rt-list .rtedit').forEach(card => { const i = +card.dataset.i, it = list[i]; if (!it) return; it.name = card.querySelector('[data-rf=name]').value.trim() || it.name; it.kcal = num(card.querySelector('[data-rf=kcal]').value) ?? 0; it.prot = num(card.querySelector('[data-rf=prot]').value) ?? 0; it.meal = card.querySelector('[data-rf=meal]').value; it.days = card.querySelector('[data-rf=days]').value; it.active = card.querySelector('[data-rf=active]').checked; }); };
  const redraw = () => { $('#rt-list').innerHTML = list.map((it, i) => removed.has(i) ? '' : rowHtml(it, i)).join(''); };
  $('#rt-add').onclick = () => { read(); list.push({ id: 'u' + Date.now().toString(36).slice(-6), name: '', meal: 'Snack', kcal: 0, prot: 0, days: 'All', active: true, order: list.length + 1, note: '', isNew: true }); redraw(); const cards = document.querySelectorAll('#rt-list .rtedit'); const last = cards[cards.length - 1]; if (last) { last.scrollIntoView({ behavior: 'smooth' }); last.querySelector('[data-rf=name]').focus(); } };
  $('#rt-list').addEventListener('click', e => { const b = e.target.closest('[data-rtdel]'); if (!b) return; read(); removed.add(+b.dataset.rtdel); redraw(); });
  $('#rt-save').onclick = () => {
    read(); const before = state.data.routine || []; const ops = [];
    list.forEach((it, i) => { if (removed.has(i)) { if (!it.isNew) ops.push({ kind: 'routine', action: 'delete', id: it.id }); return; } if (!it.name) return; const old = before.find(x => x.id === it.id); const clean = { id: it.id, name: asText(it.name), meal: it.meal, kcal: it.kcal, prot: it.prot, days: it.days, active: it.active, order: i + 1 }; if (!old) { ops.push({ kind: 'routine', action: 'set', item: clean }); return; } const item = { id: it.id }; ['name', 'meal', 'kcal', 'prot', 'days', 'active', 'order'].forEach(k => { if (String(old[k]) !== String(clean[k])) item[k] = clean[k]; }); if (Object.keys(item).length > 1) ops.push({ kind: 'routine', action: 'set', item }); });
    closeSheet(); if (!ops.length) return toast('Nothing changed'); ops.forEach(op => apply(op)); go('food');
  };
}
function benchSheet() {
  const tests = state.data.bench; const t = todaySerial(); let sel = tests.filter(b => b[0] <= t + 7).pop() || tests[0];
  const fields = [['sorensen', 'Sorensen hold (s)'], ['sideR', 'Side plank right (s)'], ['sideL', 'Side plank left (s)'], ['flexor', '60-degree flexor hold (s)'], ['plank', 'Prone plank (s)'], ['bridgeR', 'Single-leg bridge R (reps)'], ['bridgeL', 'Single-leg bridge L (reps)'], ['pain', 'Worst back pain this week (0-10)'], ['bw', 'Bodyweight today (kg)'], ['carry', 'Suitcase carry kg (50 m level)']];
  const rec = k => { const ci = BENCH_FIELDS.indexOf(k); return isBlank(sel[2 + ci]) ? '' : sel[2 + ci]; };
  openSheet(`<h2>Benchmark test</h2>
    <div class="field"><label>Which test</label><select class="inp" id="b-test">${tests.map(b => `<option value="${b[0]}" ${b === sel ? 'selected' : ''}>${fmtDate(b[0])} · ${esc(b[1])}</option>`).join('')}</select></div>
    <div id="b-fields">${fields.map(([k, l]) => `<div class="field rel"><label>${l}</label><input class="inp num" type="text" inputmode="decimal" data-bf="${k}" value="${esc(rec(k))}"></div>`).join('')}
    <div class="toggles"><button class="tg ${rec('deadbug') === 'Y' ? 'on' : ''}" data-btg="deadbug"><span class="dot">${rec('deadbug') === 'Y' ? '✓' : ''}</span>Dead-bug pass</button><button class="tg ${rec('thomas') === 'Y' ? 'on' : ''}" data-btg="thomas"><span class="dot">${rec('thomas') === 'Y' ? '✓' : ''}</span>Thomas test flat</button></div></div>
    <div class="xs dim" style="margin:10px 0">Ratios and the gate decision are computed for you. Test the Friday-before yoga-only; postpone to Sunday if resting HR is 10+ bpm over baseline.</div>
    <button class="btn primary" id="b-save">Save test</button>`);
  $('#b-test').onchange = () => { sel = tests.find(b => b[0] === +$('#b-test').value); fields.forEach(([k]) => { $(`[data-bf="${k}"]`).value = rec(k); }); ['deadbug', 'thomas'].forEach(k => { const b = $(`[data-btg="${k}"]`); b.classList.toggle('on', rec(k) === 'Y'); b.querySelector('.dot').textContent = rec(k) === 'Y' ? '✓' : ''; }); };
  $('#b-fields').addEventListener('click', e => { const b = e.target.closest('[data-btg]'); if (!b) return; b.classList.toggle('on'); b.querySelector('.dot').textContent = b.classList.contains('on') ? '✓' : ''; });
  $('#b-save').onclick = () => { const out = {}; fields.forEach(([k]) => { const v = num($(`[data-bf="${k}"]`).value); if (v != null) out[k] = v; }); ['deadbug', 'thomas'].forEach(k => { out[k] = $(`[data-btg="${k}"]`).classList.contains('on') ? 'Y' : 'N'; }); apply({ kind: 'bench', date: +$('#b-test').value, fields: out }); closeSheet(); };
}

/* ---------- render + events ---------- */
function render() {
  const app = $('#app'), nav = $('#nav');
  if (!backend) { app.innerHTML = renderSetup(); nav.style.display = 'none'; bindSetup(); syncUI(); return; }
  nav.style.display = '';
  if (!state.data) { app.innerHTML = `<div class="screen active"><div class="hdr"><h1>Loading your plan…</h1></div>${[1, 2, 3].map(() => '<div class="card"><div class="skel" style="width:40%"></div><div class="skel" style="width:70%;margin-top:10px"></div><div class="skel" style="width:55%;margin-top:10px"></div></div>').join('')}${state.error ? `<div class="status action">${esc(state.error)}</div><button class="btn" id="retry">Retry</button><button class="btn ghost" data-open="settings" style="margin-top:8px">Settings</button>` : ''}</div>`; const r = $('#retry'); if (r) r.onclick = () => refresh(); syncUI(); return; }
  const scr = { today: renderToday, log: renderLog, train: renderTrain, food: renderFood, progress: renderProgress }[state.screen] || renderToday;
  app.innerHTML = `<div class="screen active">${scr()}</div>`;
  [...nav.querySelectorAll('button')].forEach(b => b.classList.toggle('on', b.dataset.go === state.screen));
  if (state.screen === 'progress') requestAnimationFrame(drawCharts);
  syncUI();
}
function go(s) { state.screen = s; store.set('screen', s); window.scrollTo(0, 0); render(); }
function bindSetup() {
  $('#s-connect').onclick = async () => {
    const p = parsePairing($('#s-pair').value); if (!p) { toast('Paste the whole pairing code: the web app URL, a #, and the key', true); return; }
    if (!looksLikeApi(p.api)) { toast('The URL should start with https://script.google.com/macros/s/ and end with /exec', true); return; }
    const btn = $('#s-connect'); btn.disabled = true; btn.textContent = 'Connecting…';
    try {
      const j = await new SheetsStore(p.api, p.key).ping();
      if (!j.name) throw new Error('Key rejected. Copy the pairing code from cell B7 of the App Setup sheet again.');
      if (j.setup === false) throw new Error('The sheet answered, but the tracker sheets are missing: run "setup" in the script editor first (Step 2).');
      settings.api = p.api; settings.key = p.key; settings.demo = false; saveSettings(); ['cache', 'mock'].forEach(k => store.del(k)); state.data = null; state.error = null;
      toast('Connected to ' + j.name); await initBackend(); render(); await flush(); refresh(true);
    } catch (e) { state.error = e.message; render(); }
  };
  $('#s-demo').onclick = async () => { settings.demo = true; saveSettings(); await initBackend(); render(); refresh(true); };
}
document.addEventListener('click', async e => {
  const t = e.target.closest('[data-rtadj],[data-rt],[data-go],[data-nav],[data-open],[data-dose],[data-tg],[data-meal],[data-quick],[data-delfood],[data-cardio],#t-weight-save,#l-save,#w-save,#f-add,#f-manual,#p-refresh,#b-log');
  if (!t) return;
  if (t.dataset.go) return go(t.dataset.go);
  if (t.dataset.open === 'settings') return renderSettingsSheet();
  if (t.dataset.nav) { const [which, dir] = t.dataset.nav.split(':'); const key = which + 'Date'; const cur = state[key] || todaySerial(); state[key] = Math.max(START, Math.min(todaySerial(), cur + (+dir))); return render(); }
  if (t.dataset.dose) return doseSheet(t.dataset.dose, (state.screen === 'log' ? state.logDate : null) || todaySerial());
  if (t.dataset.tg && !t.dataset.btg) { t.classList.toggle('on'); t.querySelector('.dot').textContent = t.classList.contains('on') ? '✓' : ''; return; }
  if (t.dataset.meal) { state.meal = t.dataset.meal; return render(); }
  if (t.dataset.quick) { const [item, k, p] = t.dataset.quick.split('|'); $('#f-item').value = item; $('#f-kcal').value = k; $('#f-prot').value = p; Object.assign(state.foodDraft, { item, kcal: k, prot: p }); return; }
  if (t.dataset.rtadj != null || t.dataset.rt != null) {
    const s = state.foodDate || todaySerial(), it = (state.data.routine || []).find(x => x.id === (t.dataset.rtadj != null ? t.dataset.rtadj : t.dataset.rt)); if (!it) return;
    const cur = routineTicked(s, it.id), id = rtFoodId(s, it.id);
    if (t.dataset.rtadj != null) {
      const k = prompt(`${it.name}\nCalories today:`, cur ? cur[F.kcal] : it.kcal); if (k === null) return; const pr = prompt('Protein today (g):', cur ? cur[F.prot] : it.prot); if (pr === null) return;
      apply({ kind: 'food', values: [s, it.meal, it.name, num(k) ?? it.kcal, num(pr) ?? it.prot, 'routine', id] }); return;   // same id: the sheet updates the row in place
    }
    if (cur) apply({ kind: 'delfood', index: state.data.food.indexOf(cur), expect: cur.slice(0, 5), id }); else apply({ kind: 'food', values: [s, it.meal, it.name, it.kcal, it.prot, 'routine', id] });
    return;
  }
  if (t.dataset.delfood != null) { const i = +t.dataset.delfood, row = state.data.food[i]; if (confirm('Remove this item?')) apply({ kind: 'delfood', index: i, expect: row.slice(0, 5), id: row[6] || null }); return; }
  if (t.dataset.cardio != null) { const i = +t.dataset.cardio; const r = state.data.workout[i]; apply({ kind: 'cells', sheet: 'Workout Log', cells: [{ addr: colL(W.setsDone) + sheetRow(i), value: r[W.setsDone] ? null : 1 }] }); return; }
  if (t.id === 'b-log') return benchSheet();
  if (t.id === 't-weight-save') { const v = num($('#t-weight').value); if (v == null) return toast('Enter a weight', true); const i = dailyIdx(todaySerial()); if (i < 0) return toast('Today is outside the plan window', true); apply({ kind: 'cells', sheet: 'Daily Log', cells: [{ addr: colL(D.weight) + sheetRow(i), value: v }] }); return; }
  if (t.id === 'l-save') {
    const s = state.logDate || todaySerial(), i = dailyIdx(s), row = state.data.daily[i], r = sheetRow(i), cells = [];
    document.querySelectorAll('#app [data-col]').forEach(inp => { const col = +inp.dataset.col; const v = inp.tagName === 'TEXTAREA' ? asText(inp.value.trim()) : num(inp.value); const old = row[col]; if ((v == null || v === '') && !isBlank(old)) cells.push({ addr: colL(col) + r, value: null }); else if (v != null && v !== '' && String(v) !== String(old)) cells.push({ addr: colL(col) + r, value: v }); });
    [...SUPPS.map(sp => sp.k), 'spine', 'sauna'].forEach(k => { const on = $(`[data-tg="${k}"]`).classList.contains('on'); const col = D[k]; const v = on ? 'Y' : 'N'; if (String(row[col] || '') !== v && !(isBlank(row[col]) && v === 'N')) cells.push({ addr: colL(col) + r, value: v }); });
    if (!cells.length) return toast('Nothing changed');
    apply({ kind: 'cells', sheet: 'Daily Log', cells }); return;
  }
  if (t.id === 'w-save') {
    const cells = [];
    document.querySelectorAll('#app .ex[data-row]').forEach(card => { const i = +card.dataset.row, row = state.data.workout[i], r = sheetRow(i); card.querySelectorAll('[data-wcol]').forEach(inp => { const col = +inp.dataset.wcol, v = num(inp.value), old = row[col]; if (v == null && !isBlank(old)) cells.push({ addr: colL(col) + r, value: null }); else if (v != null && String(v) !== String(old)) cells.push({ addr: colL(col) + r, value: v }); }); });
    if (!cells.length) return toast('Nothing changed');
    apply({ kind: 'cells', sheet: 'Workout Log', cells }); return;
  }
  if (t.id === 'f-add') {
    const item = asText($('#f-item').value.trim()), k = num($('#f-kcal').value), p = num($('#f-prot').value) || 0, note = asText(($('#f-note') ? $('#f-note').value : '').trim()); if (!item || k == null) return toast('Item and calories are needed', true);
    const s = state.foodDate || todaySerial(); const hour = new Date().getHours(); const meal = state.meal || (hour < 10 ? 'Wake-up' : hour < 13 ? 'Lunch' : hour < 17 ? 'Snack' : hour < 21 ? 'Dinner' : 'Pre-sleep');
    state.foodDraft = { item: '', kcal: '', prot: '', note: '' }; apply({ kind: 'food', values: [s, meal, item, k, p, note, uuid()] }); return;
  }
  if (t.id === 'f-manual') { const s = state.foodDate || todaySerial(), i = dailyIdx(s), row = state.data.daily[i]; const v = prompt('Total calories for the day (leave empty to use the item sum):', row[D.kcalM] || ''); if (v === null) return; const p = prompt('Total protein for the day (g), optional:', row[D.protM] || ''); const cells = [{ addr: colL(D.kcalM) + sheetRow(i), value: num(v) }]; if (p !== null) cells.push({ addr: colL(D.protM) + sheetRow(i), value: num(p) }); apply({ kind: 'cells', sheet: 'Daily Log', cells }); return; }
  if (t.id === 'p-refresh') return refresh();
});
document.addEventListener('click', e => {
  if (e.target.id === 'modal') closeSheet();
  if (e.target.id === 'st-save') {
    settings.name = $('#st-name').value.trim() || 'Hamid';
    const raw = $('#st-pair').value.trim(), p = parsePairing(raw); if (raw && !p) { toast('The pairing code is incomplete: URL, a #, and the key', true); return; }
    const changed = p && (p.api !== settings.api || p.key !== settings.key); if (changed) { settings.api = p.api; settings.key = p.key; settings.demo = false; }
    saveSettings(); closeSheet(); if (changed) { ['cache', 'mock'].forEach(k => store.del(k)); location.reload(); } else render();
  }
  if (e.target.id === 'st-refresh') { closeSheet(); refresh(); }
  if (e.target.id === 'st-routine') { closeSheet(); setTimeout(routineEditorSheet, 50); }
  if (e.target.id === 'st-demo') { settings.demo = !settings.demo; saveSettings(); store.del('cache'); location.reload(); }
  if (e.target.id === 'st-signout') { if (settings.demo) { ['mock', 'cache', 'queue'].forEach(k => store.del(k)); location.reload(); } else if (!state.queue.length || confirm(state.queue.length + ' entries have not reached the sheet yet and would be lost. Disconnect anyway?')) disconnect(); }
});
document.addEventListener('input', e => { const m = { 'f-item': 'item', 'f-kcal': 'kcal', 'f-prot': 'prot', 'f-note': 'note' }[e.target.id]; if (m) state.foodDraft[m] = e.target.value; });
window.addEventListener('online', () => { syncUI(); flush(); });
window.addEventListener('offline', syncUI);
document.addEventListener('visibilitychange', () => { if (document.hidden || !backend) return; if (state.queue.length) flush(); else if (state.data && Date.now() - state.data.ts > 5 * 60000) refresh(true); });

/* ---------- boot ---------- */
window.__coach = { SheetsStore, Mock, parsePairing, looksLikeApi, computeDashboard, parseDashboard, D, W, F, serialOf, version: APP_VERSION };   // for diagnostics and tests
(async () => {
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
  const paired = pairFromHash();
  let ok = false; try { ok = await initBackend(); } catch (e) { state.error = e.message; }
  if (state.data && !state.data.dashboard) recomputeLocal();
  render();
  if (paired) toast('Pairing code received');
  if (ok) { await flush(); await refresh(!!state.data); }
})();
})();
