/* global Office, PowerPoint, SlideFactory, SlideFactoryPlus */
'use strict';

const SF_BUILD = '1.3.1';
const TEXTUAL = new Set(['GeometricShape', 'TextBox', 'Placeholder', 'Callout', 'Freeform']);
let busy = false;

function showBuild() {
  const title = document.getElementById('build-title');
  if (title) title.textContent = 'SlideFactory Storage ' + SF_BUILD;
}
showBuild();

Office.onReady(info => {
  Office.actions.associate('sfVerify', sfVerify);
  Office.actions.associate('sfInitialize', sfInitialize);
  Office.actions.associate('sfGetState', sfGetState);
  Office.actions.associate('sfReplaceState', sfReplaceState);
  Office.actions.associate('sfPatchPassport', sfPatchPassport);
  Office.actions.associate('sfPatchStoryline', sfPatchStoryline);
  Office.actions.associate('sfFingerprints', sfFingerprints);
  bind('refresh', () => show(readState()));
  bind('verify', () => sfVerify());
  bind('init', () => sfInitialize());
  bind('patch', patchField);
  bind('paste', pasteApply);
  bind('copy-project', () => copyProjection('project'));
  bind('copy-section', () => copyProjection('section'));
  bind('copy-slide', () => copyProjection('current-slide'));
  bind('copy-full', () => copyProjection('full'));
  if (info.host === Office.HostType.PowerPoint) {
    if (window.SF_LIVE_TEST === true) runLiveNilTest();
    else show(readState());
  } else setStatus('Откройте панель в PowerPoint');
});

function bind(id, fn) {
  const el = document.getElementById(id);
  if (el) el.addEventListener('click', () => runUi(fn));
}

async function runUi(fn) {
  if (busy) return;
  busy = true;
  document.querySelectorAll('button').forEach(button => { button.disabled = true; });
  const started = performance.now();
  try {
    await fn();
    const elapsed = Math.round(performance.now() - started);
    const status = document.getElementById('status');
    if (status && !status.dataset.sticky) status.textContent = `${status.textContent} · ${elapsed} мс`;
  } catch (err) {
    log({ code: err.code || 'ERROR', message: err.message });
    setStatus('Ошибка');
  } finally {
    busy = false;
    document.querySelectorAll('button').forEach(button => { button.disabled = false; });
  }
}

async function sfVerify(event) {
  const xml = await readXml();
  const report = xml == null
    ? { ok: false, partCount: 0, issues: ['missing part'] }
    : SlideFactory.verifyXml(xml);
  if (report.ok) report.partCount = 1;
  showReport(report);
  return finish(event, report);
}

async function sfInitialize(event) {
  const result = await write(xml => {
    if (xml) return { wrote: false, result: Object.assign({}, SlideFactory.getStateFromXml(xml), { status: 'unchanged' }) };
    return SlideFactory.applyOperation(null, {
      type: 'replace',
      payload: {
        passport: {
          status: 'draft',
          project_stage: 'не установлено',
          goals: { business: 'не установлено', communication: 'не установлено', learning: 'не установлено' },
          audience: 'не установлено',
          jtbd: 'не установлено',
          hunt_stage: { current: 'не установлено', target: 'не установлено' },
          value_proposition: { topic: 'не установлено', presentation: 'не установлено' }
        },
        storyline: { status: 'not-built', slides: [] }
      }
    });
  });
  show(Promise.resolve(result));
  return finish(event, result);
}

async function sfGetState(event) {
  const state = await readState();
  show(Promise.resolve(state));
  return finish(event, state);
}

async function sfReplaceState(message) {
  const args = parseArgs(message);
  const expected = requireRevision(args);
  const result = await write(xml => SlideFactory.applyOperation(xml, {
    type: 'replace',
    payload: args.payload || {},
    expectedRevision: expected
  }));
  return JSON.stringify(result);
}

async function sfPatchPassport(message) {
  const args = parseArgs(message);
  const expected = requireRevision(args);
  const result = await write(xml => SlideFactory.applyOperation(xml, {
    type: 'patch_passport',
    path: args.path,
    value: args.value,
    expectedRevision: expected
  }));
  return JSON.stringify(result);
}

async function sfPatchStoryline(message) {
  const args = parseArgs(message);
  const expected = requireRevision(args);
  const result = await write(xml => SlideFactory.applyOperation(xml, {
    type: 'patch_storyline',
    slideId: args.slide_id,
    patch: args.patch || {},
    expectedRevision: expected
  }));
  return JSON.stringify(result);
}

async function sfFingerprints(event) {
  const rows = await slideFingerprints();
  log(rows);
  return finish(event, rows);
}

async function readState() {
  const xml = await readXml();
  const state = SlideFactory.getStateFromXml(xml);
  render(state);
  return state;
}

async function patchField() {
  const result = await write(xml => SlideFactoryPlus.applyPatch(xml, {
    format: 'SLIDEFACTORY_PATCH',
    version: 1,
    expected_revision: revisionFromField(),
    passport_patches: [{ path: document.getElementById('path').value.trim(), value: document.getElementById('value').value }]
  }));
  render(result);
}

async function pasteApply() {
  const result = await write(xml => SlideFactoryPlus.applyPatch(xml, document.getElementById('patch').value));
  render(result);
}

async function copyProjection(projection) {
  const state = SlideFactory.getStateFromXml(await readXml());
  let currentSlideId = null;
  let fingerprints = [];
  if (projection === 'current-slide' || projection === 'full') {
    currentSlideId = await selectedSlideId();
    try {
      fingerprints = (await slideFingerprints()).slides;
    } catch (err) {
      fingerprints = [];
    }
  } else if (projection === 'section') {
    currentSlideId = await selectedSlideId();
  }
  let section = '';
  if (projection === 'section') {
    const typed = document.getElementById('section-name');
    section = typed ? typed.value.trim() : '';
    if (!section && currentSlideId && state.storyline && state.storyline.slides) {
      const entry = state.storyline.slides.find(slide => slide.slide_id === currentSlideId);
      section = entry && entry.section ? entry.section : '';
      if (typed && section) typed.value = section;
    }
  }
  const context = SlideFactoryPlus.projectContext(state, { projection, currentSlideId, fingerprints, section });
  const text = SlideFactoryPlus.serializeContext(context);
  const box = document.getElementById('context');
  if (box) box.value = text;
  const copied = await copyText(text);
  setStatus(copied
    ? 'Контекст скопирован один раз на сессию. Один патч — на всю итерацию.'
    : 'Контекст в поле. Один патч — на всю итерацию.');
  log(context);
}

function revisionFromField() {
  const raw = document.getElementById('expected').value.trim();
  if (!/^\d+$/.test(raw)) throw new SlideFactory.StorageError('MISSING_REVISION', 'expected_revision is required');
  return Number(raw);
}

function requireRevision(args) {
  if (args.expected_revision === undefined || args.expected_revision === null || args.expected_revision === '') {
    throw new SlideFactory.StorageError('MISSING_REVISION', 'expected_revision is required');
  }
  return args.expected_revision;
}

async function selectedSlideId() {
  try {
    return await PowerPoint.run(async context => {
      const selected = context.presentation.getSelectedSlides();
      selected.load('items');
      await context.sync();
      if (!selected.items.length) return null;
      selected.items[0].load('id');
      await context.sync();
      return String(selected.items[0].id);
    });
  } catch (err) {
    return null;
  }
}

async function copyText(text) {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch (err) {
    const box = document.getElementById('context');
    if (!box) return false;
    box.focus();
    box.select();
    try {
      return document.execCommand('copy');
    } catch (copyErr) {
      return false;
    }
  }
}

function show(promise) {
  return promise.then(render).catch(err => {
    log({ code: err.code || 'ERROR', message: err.message });
    setStatus('Ошибка');
  });
}

function render(state) {
  if (!state) return;
  const summary = document.getElementById('summary');
  if (!summary) {
    log(state);
    return;
  }
  if (state.status === 'absent' || state.partCount === 0) {
    summary.innerHTML = '<p>Части SlideFactory в этой презентации нет.</p>';
    setStatus('Часть не создана');
    log(state);
    return;
  }
  const passport = state.passport || {};
  const goals = passport.goals || {};
  const slides = state.storyline && state.storyline.slides ? state.storyline.slides.length : 0;
  summary.innerHTML = `<dl>
    <dt>Ревизия</dt><dd>${escapeHtml(state.revision)}</dd>
    <dt>Статус</dt><dd>${escapeHtml(state.status)}</dd>
    <dt>Сумма</dt><dd>${escapeHtml(state.checksum || 'нет')}</dd>
    <dt>Аудитория</dt><dd>${escapeHtml(passport.audience || '')}</dd>
    <dt>JTBD</dt><dd>${escapeHtml(passport.jtbd || '')}</dd>
    <dt>Учебная</dt><dd>${escapeHtml(goals.learning || '')}</dd>
    <dt>Слайды карты</dt><dd>${slides}</dd>
  </dl>`;
  const expected = document.getElementById('expected');
  if (expected && document.activeElement !== expected) expected.value = String(state.revision);
  setStatus(state.status === 'unchanged' ? 'Без изменений' : state.status === 'saved' ? 'Сохранено и проверено' : 'Прочитано');
  log(state);
}

function showReport(report) {
  setStatus(report.ok ? 'Проверка пройдена' : 'Проверка не пройдена');
  log(report);
}

function setStatus(text) {
  const status = document.getElementById('status');
  if (status) status.textContent = text;
}

function log(value) {
  const node = document.getElementById('log');
  if (node) node.textContent = JSON.stringify(value, null, 2);
}

function escapeHtml(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function parseArgs(message) {
  if (typeof message === 'string') return JSON.parse(message);
  if (message && typeof message === 'object' && typeof message.completed !== 'function') return message;
  return {};
}

function finish(event, result) {
  if (event && typeof event.completed === 'function') {
    event.completed();
    return result;
  }
  return JSON.stringify(result);
}

async function readXml() {
  return PowerPoint.run(async context => {
    const scoped = context.presentation.customXmlParts.getByNamespace(SlideFactory.NS);
    const count = scoped.getCount();
    await context.sync();
    if (count.value > 1) throw new SlideFactory.StorageError('CONFLICT', `${count.value} Custom XML parts`);
    if (count.value === 0) return null;
    const xml = scoped.getOnlyItem().getXml();
    await context.sync();
    return xml.value;
  });
}

async function putXml(nextXml) {
  return PowerPoint.run(async context => {
    const collection = context.presentation.customXmlParts;
    const scoped = collection.getByNamespace(SlideFactory.NS);
    const count = scoped.getCount();
    await context.sync();
    if (count.value > 1) throw new SlideFactory.StorageError('CONFLICT', `${count.value} Custom XML parts`);
    if (count.value === 0) collection.add(nextXml);
    else scoped.getOnlyItem().setXml(nextXml);
    await context.sync();
  });
}

function readBackOk(applied, readBackXml) {
  if (!readBackXml) return false;
  const semanticOk = SlideFactory.same(SlideFactory.semanticFromXml(readBackXml), applied.semantic);
  const valuesOk = !applied.patch || SlideFactoryPlus.persistedMatches(applied.patch, readBackXml);
  return semanticOk && valuesOk;
}

async function write(mutator) {
  const original = await readXml();
  const applied = mutator(original);
  if (!applied.wrote) return applied.result;
  await putXml(applied.xml);
  const readBack = await readXml();
  if (!readBackOk(applied, readBack)) {
    if (original) await putXml(original);
    throw new SlideFactory.StorageError('VERIFY_FAILED', 'Requested value was not persisted. The previous revision was restored.');
  }
  return Object.assign({ status: 'saved' }, SlideFactory.getStateFromXml(readBack));
}

async function replaceByAdd(nextXml) {
  return PowerPoint.run(async context => {
    const collection = context.presentation.customXmlParts;
    const scoped = collection.getByNamespace(SlideFactory.NS);
    const count = scoped.getCount();
    await context.sync();
    if (count.value === 1) {
      scoped.getOnlyItem().delete();
      await context.sync();
    }
    collection.add(nextXml);
    await context.sync();
  });
}

async function runLiveNilTest() {
  const value = 'Клиентский менеджер умеет распознавать сигналы потребности клиента и выбирать следующий шаг разговора';
  const beforeXml = await readXml();
  const before = SlideFactory.getStateFromXml(beforeXml);
  const patch = {
    format: 'SLIDEFACTORY_PATCH',
    version: 1,
    expected_revision: before.revision,
    passport_patches: [{ path: 'goals.learning', value }]
  };
  const planned = SlideFactoryPlus.applyPatch(beforeXml, patch);
  let setStatusCode = 'err';
  let setError = '';
  try {
    const saved = await write(xml => SlideFactoryPlus.applyPatch(xml, patch));
    setStatusCode = saved.status;
  } catch (err) {
    setError = err.code || 'ERR';
  }
  let afterSet = SlideFactory.getStateFromXml(await readXml());
  let setOk = afterSet.passport && afterSet.passport.goals && afterSet.passport.goals.learning === value ? 1 : 0;
  let addOk = 0;
  let addError = '';
  if (!setOk && planned.wrote) {
    try {
      await replaceByAdd(planned.xml);
      const afterAdd = SlideFactory.getStateFromXml(await readXml());
      addOk = afterAdd.passport && afterAdd.passport.goals && afterAdd.passport.goals.learning === value ? 1 : 0;
      afterSet = afterAdd;
    } catch (err) {
      addError = err.code || 'ERR';
    }
  }
  const summary = [
    'b=' + SF_BUILD,
    'br=' + before.revision,
    'ar=' + afterSet.revision,
    'set=' + setOk,
    'add=' + addOk,
    'nil=' + ((beforeXml || '').indexOf('nil="true"') >= 0 ? 1 : 0),
    'se=' + setError,
    'ae=' + addError,
    'st=' + setStatusCode
  ].join(';').slice(0, 250);
  await PowerPoint.run(async context => {
    const props = context.presentation.properties.customProperties;
    const existing = props.getItemOrNullObject('SF_LIVE_RESULT');
    existing.load('key');
    await context.sync();
    if (!existing.isNullObject) existing.delete();
    props.add('SF_LIVE_RESULT', summary);
    await context.sync();
  });
  setStatus(summary);
  log(summary);
}

async function slideFingerprints() {
  return PowerPoint.run(async context => {
    const slides = context.presentation.slides;
    slides.load('items');
    await context.sync();
    for (const slide of slides.items) {
      slide.load('id');
      slide.shapes.load('items');
    }
    await context.sync();
    const shapes = [];
    for (const slide of slides.items) {
      for (const shape of slide.shapes.items) {
        shape.load('type');
        shapes.push({ slide, shape });
      }
    }
    await context.sync();
    const textual = shapes.filter(item => TEXTUAL.has(String(item.shape.type)));
    for (const item of textual) item.shape.textFrame.load('hasText');
    await context.sync();
    for (const item of textual) {
      if (item.shape.textFrame.hasText) item.shape.textFrame.textRange.load('text');
    }
    await context.sync();
    const bySlide = new Map();
    slides.items.forEach((slide, index) => {
      bySlide.set(slide.id, { slideId: String(slide.id), order: index + 1, text: [] });
    });
    for (const item of textual) {
      if (item.shape.textFrame.hasText) bySlide.get(item.slide.id).text.push(item.shape.textFrame.textRange.text || '');
    }
    return SlideFactory.fingerprintFromSlides([...bySlide.values()]);
  });
}
