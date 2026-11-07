/**
 * The Tribe — live notes backend.
 * Paste into the Sheet's Extensions → Apps Script, run setup() once, then Deploy → New deployment → Web app
 * (Execute as: Me · Who has access: Anyone). Copy the web-app URL into the site.
 *
 * Moderation: edit or delete rows directly in the Sheet. The site picks up changes within about 30 seconds.
 */

// Change this to the passcode you share with the group.
const PASSCODE = 'kestrel';

const SHEETS = {
  notes:   ['id', 'at', 'member', 'kind', 'title', 'body', 'link'],
  species: ['id', 'at', 'group', 'name', 'where'],
  spnotes: ['id', 'at', 'species', 'group', 'kind', 'text'],
};
const MEMBERS = ['Sargam', 'Arul', 'Natasha', 'Vibish', 'Vinitha', 'Irene', 'Joel', 'Vishwa', 'Surya', 'Sachin'];
const NOTE_KINDS = ['explore', 'observe', 'question', 'debate', 'plan'];
const SP_KINDS = ['note', 'pattern', 'family', 'unique'];
const GROUPS = ['birds', 'odonates', 'inverts', 'reptiles', 'amphibians', 'mammals', 'plants'];

function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(SHEETS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(SHEETS[name]);
      sh.setFrozenRows(1);
    }
  });
  const notes = ss.getSheetByName('notes');
  if (notes.getLastRow() === 1) {
    [
      ['seed-sargam-invasives', 1, 'Sargam', 'explore', 'Are invasive plants smarter, or just riding the Sapiens wave?',
       'Not smarter: better matched to land humans disturbed.\nEnemy release: they leave their pests behind.\nNovel weapons: Lantana releases chemicals that local plants never evolved to handle.\nBaker’s “ideal weed”: fast growth, huge seed output, cloning, tolerance of almost anything.\nCrosby’s Ecological Imperialism: settlers arrived with weeds, livestock and germs that took over together. Lantana and vilayati babul reached India the same way.\nPlanning to read: Crosby, Ecological Imperialism.', ''],
      ['seed-sargam-guides', 2, 'Sargam', 'explore', 'Making interactive, personal field guides like this one', 'This page is the first try. Tell me what to add.', ''],
      ['seed-arul-snake', 3, 'Arul', 'plan', 'Snake walk with the Irulas · 14 Nov', 'Vadanemmeli, Chennai, 6:30–9 am. With NatureED.', 'https://www.instagram.com/natureed_emden/'],
      ['seed-arul-mannar', 4, 'Arul', 'plan', 'Gulf of Mannar marine explorers · 11–13 Dec', 'From Tuticorin. Kayaking, snorkelling, mangrove walk. 18+.', 'https://www.instagram.com/natureed_emden/'],
    ].forEach(function (r) { notes.appendRow(r); });
  }
  // Remove the default empty "Sheet1" if it is still there.
  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
}

function doGet() {
  return json({ ok: true, notes: read('notes'), species: read('species'), spnotes: read('spnotes') });
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (String(body.pass || '').trim().toLowerCase() !== PASSCODE.toLowerCase()) return json({ ok: false, error: 'passcode' });
    if (!SHEETS[body.type]) return json({ ok: false, error: 'type' });
    if (tooMany()) return json({ ok: false, error: 'busy' });
    const d = body.data || {};
    let row;
    if (body.type === 'notes') {
      if (MEMBERS.indexOf(d.member) < 0 || NOTE_KINDS.indexOf(d.kind) < 0) return json({ ok: false, error: 'invalid' });
      const title = clean(d.title, 120);
      if (!title) return json({ ok: false, error: 'invalid' });
      row = [d.member, d.kind, title, clean(d.body, 1500), safeUrl(d.link)];
    } else if (body.type === 'species') {
      const name = clean(d.name, 80);
      if (GROUPS.indexOf(d.group) < 0 || !name) return json({ ok: false, error: 'invalid' });
      row = [d.group, name, clean(d.where, 200)];
    } else {
      const text = clean(d.text, 600), sp = clean(d.species, 80);
      if (GROUPS.indexOf(d.group) < 0 || SP_KINDS.indexOf(d.kind) < 0 || !text || !sp) return json({ ok: false, error: 'invalid' });
      row = [sp, d.group, d.kind, text];
    }
    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      SpreadsheetApp.getActive().getSheetByName(body.type).appendRow([Utilities.getUuid(), Date.now()].concat(row));
    } finally {
      lock.releaseLock();
    }
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: 'server' });
  }
}

function read(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getDataRange().getValues();
  const head = vals[0];
  return vals.slice(1).filter(function (r) { return r[0] !== ''; }).map(function (r) {
    const o = {};
    head.forEach(function (h, i) { o[h] = r[i]; });
    return o;
  });
}

// Keep strings plain and short. A leading = + - @ is neutralised so a note can never run as a Sheet formula.
function clean(v, max) {
  let s = String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}
function safeUrl(u) {
  const s = String(u || '').trim();
  return /^https?:\/\/[^\s]+$/i.test(s) ? s.slice(0, 500) : '';
}
// At most 30 posts a minute across everyone, to stop runaway spam.
function tooMany() {
  const c = CacheService.getScriptCache();
  const n = Number(c.get('posts') || 0) + 1;
  c.put('posts', String(n), 60);
  return n > 30;
}
function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
