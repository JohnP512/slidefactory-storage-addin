globalThis.__sfSha256 = (function () {
'use strict';

// Compact synchronous SHA-256. Used in Node and in the Office add-in so
// checksums match across hosts. Verified against Node's crypto in tests.

function rotr(x, n) {
  return (x >>> n) | (x << (32 - n));
}

function sha256(input) {
  const bytes = utf8(input);
  const bitLen = bytes.length * 8;
  const withOne = bytes.length + 1;
  const pad = (withOne % 64 <= 56) ? (56 - (withOne % 64)) : (56 + 64 - (withOne % 64));
  const data = new Uint8Array(bytes.length + 1 + pad + 8);
  data.set(bytes);
  data[bytes.length] = 0x80;
  const view = new DataView(data.buffer);
  const hi = Math.floor(bitLen / 0x100000000);
  const lo = bitLen >>> 0;
  view.setUint32(data.length - 8, hi);
  view.setUint32(data.length - 4, lo);

  const k = new Uint32Array([
    0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
    0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
    0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
    0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
    0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
    0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
    0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
    0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2
  ]);

  let h0 = 0x6a09e667, h1 = 0xbb67ae85, h2 = 0x3c6ef372, h3 = 0xa54ff53a;
  let h4 = 0x510e527f, h5 = 0x9b05688c, h6 = 0x1f83d9ab, h7 = 0x5be0cd19;
  const w = new Uint32Array(64);

  for (let offset = 0; offset < data.length; offset += 64) {
    for (let i = 0; i < 16; i++) w[i] = view.getUint32(offset + i * 4);
    for (let i = 16; i < 64; i++) {
      const s0 = rotr(w[i - 15], 7) ^ rotr(w[i - 15], 18) ^ (w[i - 15] >>> 3);
      const s1 = rotr(w[i - 2], 17) ^ rotr(w[i - 2], 19) ^ (w[i - 2] >>> 10);
      w[i] = (w[i - 16] + s0 + w[i - 7] + s1) >>> 0;
    }
    let a = h0, b = h1, c = h2, d = h3, e = h4, f = h5, g = h6, h = h7;
    for (let i = 0; i < 64; i++) {
      const S1 = rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25);
      const ch = (e & f) ^ (~e & g);
      const temp1 = (h + S1 + ch + k[i] + w[i]) >>> 0;
      const S0 = rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22);
      const maj = (a & b) ^ (a & c) ^ (b & c);
      const temp2 = (S0 + maj) >>> 0;
      h = g; g = f; f = e; e = (d + temp1) >>> 0;
      d = c; c = b; b = a; a = (temp1 + temp2) >>> 0;
    }
    h0 = (h0 + a) >>> 0; h1 = (h1 + b) >>> 0; h2 = (h2 + c) >>> 0; h3 = (h3 + d) >>> 0;
    h4 = (h4 + e) >>> 0; h5 = (h5 + f) >>> 0; h6 = (h6 + g) >>> 0; h7 = (h7 + h) >>> 0;
  }
  return [h0, h1, h2, h3, h4, h5, h6, h7].map(hex8).join('');
}

function hex8(n) {
  return n.toString(16).padStart(8, '0');
}

function utf8(input) {
  if (typeof Buffer !== 'undefined') return Buffer.from(String(input), 'utf8');
  return new TextEncoder().encode(String(input));
}

return { sha256 };

})();

'use strict';

const sha256 = globalThis.__sfSha256.sha256;

const NS = 'urn:slide-factory:context:v1';
const ROOT = 'slideFactoryContext';

const STORY_FIELDS = ['section', 'message', 'role', 'transition', 'recommendation', 'status', 'fingerprint'];
const GOAL_FIELDS = ['business', 'communication', 'learning'];
const AUDIENCE_TEXT = ['identity', 'role', 'initialBeliefs', 'declaredMotives', 'barriers', 'jtbd'];

class StorageError extends Error {
  constructor(code, message) {
    super(message);
    this.name = 'StorageError';
    this.code = code;
  }
}

function xmlTools() {
  if (typeof DOMParser !== 'undefined' && typeof XMLSerializer !== 'undefined' && typeof module === 'undefined') {
    return browserXml();
  }
  if (typeof DOMParser === 'undefined') return nodeXml();
  return browserXml();
}

function nodeXml() {
  const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
  return {
    parse(xml) {
      const errors = [];
      let doc;
      try {
        doc = new DOMParser({
          onError(level, msg) {
            if (level === 'error' || level === 'fatalError') errors.push(String(msg));
          }
        }).parseFromString(String(xml), 'application/xml');
      } catch (err) {
        throw new StorageError('INVALID_XML', err.message || 'Invalid XML');
      }
      if (!doc || !doc.documentElement || errors.length) {
        throw new StorageError('INVALID_XML', errors[0] || 'Invalid XML');
      }
      return doc;
    },
    serialize(node) {
      return new XMLSerializer().serializeToString(node);
    }
  };
}

function browserXml() {
  return {
    parse(xml) {
      const doc = new DOMParser().parseFromString(String(xml), 'application/xml');
      if (!doc || !doc.documentElement || doc.getElementsByTagName('parsererror').length) {
        throw new StorageError('INVALID_XML', 'Invalid XML');
      }
      return doc;
    },
    serialize(node) {
      return new XMLSerializer().serializeToString(node);
    }
  };
}

const xml = xmlTools();

function parse(text) {
  return xml.parse(stripDecl(text));
}

function stripDecl(text) {
  return String(text).replace(/^\uFEFF?\s*<\?xml[^?]*\?>\s*/i, '');
}

function isElement(node, local) {
  return !!node && node.nodeType === 1 && node.localName === local && node.namespaceURI === NS;
}

function kids(parent, local) {
  const out = [];
  if (!parent) return out;
  for (let n = parent.firstChild; n; n = n.nextSibling) {
    if (isElement(n, local)) out.push(n);
  }
  return out;
}

function only(parent, local) {
  const list = kids(parent, local);
  if (list.length > 1) throw new StorageError('INVALID_XML', `Multiple <${local}> elements`);
  return list[0] || null;
}

function ensure(parent, local) {
  const found = only(parent, local);
  if (found) return found;
  const el = parent.ownerDocument.createElementNS(NS, local);
  parent.appendChild(el);
  return el;
}

function textOf(el) {
  if (!el) return '';
  let s = '';
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 3 || n.nodeType === 4) s += n.nodeValue;
    else if (n.nodeType === 1) s += textOf(n);
  }
  return s;
}

function setText(el, value) {
  while (el.firstChild) el.removeChild(el.firstChild);
  const text = value == null ? '' : String(value);
  if (text.length) el.appendChild(el.ownerDocument.createTextNode(text));
}

function rootOf(doc) {
  const root = doc.documentElement;
  if (!root || root.localName !== ROOT || root.namespaceURI !== NS) {
    throw new StorageError('INVALID_XML', 'Root must be slideFactoryContext in urn:slide-factory:context:v1');
  }
  const version = root.getAttribute('schemaVersion');
  if (version && version !== '1') {
    throw new StorageError('UNSUPPORTED_SCHEMA', `Unsupported schemaVersion ${version}`);
  }
  return root;
}

function revOf(el) {
  if (!el || !el.hasAttribute('revision')) return 0;
  const raw = el.getAttribute('revision');
  if (!/^\d+$/.test(raw)) throw new StorageError('INVALID_XML', 'revision must be a non-negative integer');
  return Number(raw);
}

function emptyDoc() {
  return parse(`<${ROOT} xmlns="${NS}" schemaVersion="1" revision="0"></${ROOT}>`);
}

function assertNamespace(namespace) {
  if (namespace !== NS) {
    throw new StorageError('NAMESPACE', `Only ${NS} is supported`);
  }
}

function stable(value) {
  if (Array.isArray(value)) return value.map(stable);
  if (value && typeof value === 'object') {
    const out = {};
    for (const key of Object.keys(value).sort()) {
      if (value[key] === undefined) continue;
      out[key] = stable(value[key]);
    }
    return out;
  }
  return value;
}

function canonical(value) {
  return JSON.stringify(stable(value));
}

function same(a, b) {
  return canonical(a) === canonical(b);
}

function readDecisions(el) {
  if (!el) return [];
  const nodes = kids(el, 'decision');
  if (nodes.length) {
    return nodes.map(node => {
      const item = { text: textOf(node) };
      if (node.hasAttribute('source')) item.source = node.getAttribute('source');
      if (node.hasAttribute('meetingDate')) item.meetingDate = node.getAttribute('meetingDate');
      if (item.source || item.meetingDate) return item;
      return item.text;
    });
  }
  const text = textOf(el);
  return text ? [text] : [];
}

function readSources(passport) {
  const box = only(passport, 'sources');
  if (!box) return [];
  return kids(box, 'source').map(node => {
    const item = {};
    for (const attr of ['file', 'stage', 'meetingDate']) {
      if (node.hasAttribute(attr)) item[attr] = node.getAttribute(attr);
    }
    const text = textOf(node);
    if (text) item.text = text;
    return item;
  });
}

function readAudience(node) {
  const audience = {};
  if (node.hasAttribute('id')) audience.id = node.getAttribute('id');
  if (node.hasAttribute('priority')) audience.priority = node.getAttribute('priority');
  for (const key of AUDIENCE_TEXT) {
    const child = only(node, key);
    if (child) audience[key] = textOf(child);
  }
  const motives = only(node, 'inferredMotives');
  if (motives) {
    const inferred = { text: textOf(motives) };
    for (const attr of ['status', 'confidence', 'basis']) {
      if (motives.hasAttribute(attr)) inferred[attr] = motives.getAttribute(attr);
    }
    audience.inferredMotives = inferred;
  }
  const hunt = only(node, 'hunt');
  if (hunt) {
    const stage = {};
    const initial = only(hunt, 'initial');
    const target = only(hunt, 'target');
    const rationale = only(hunt, 'rationale');
    if (initial) stage.current = textOf(initial);
    if (target) stage.target = textOf(target);
    if (rationale) stage.rationale = textOf(rationale);
    audience.hunt_stage = stage;
  }
  const topic = only(node, 'valueTopic');
  const presentation = only(node, 'valuePresentation');
  if (topic || presentation) {
    const value = {};
    if (topic) value.topic = textOf(topic);
    if (presentation) value.presentation = textOf(presentation);
    audience.value_proposition = value;
  }
  const change = only(node, 'change');
  if (change) {
    const mapped = {};
    for (const key of ['before', 'after', 'action']) {
      const child = only(change, key);
      if (child) mapped[key] = textOf(child);
    }
    audience.change = mapped;
  }
  return audience;
}

function readPassport(root) {
  const passport = only(root, 'passport');
  if (!passport) return null;
  const out = {};
  if (passport.hasAttribute('status')) out.status = passport.getAttribute('status');
  const title = only(passport, 'projectTitle');
  if (title) out.projectTitle = textOf(title);
  const stage = only(passport, 'currentStage');
  if (stage) out.project_stage = textOf(stage);
  const business = only(passport, 'businessContext');
  if (business) out.businessContext = textOf(business);
  const goals = only(passport, 'goals');
  if (goals) {
    out.goals = {};
    for (const key of GOAL_FIELDS) {
      const child = only(goals, key);
      if (child) out.goals[key] = textOf(child);
    }
  }
  const audiences = only(passport, 'audiences');
  if (audiences) out.audiences = kids(audiences, 'audience').map(readAudience);
  const message = only(passport, 'mainMessage');
  if (message) out.mainMessage = textOf(message);
  const decisions = only(passport, 'methodologistDecisions');
  if (decisions) out.methodological_decisions = readDecisions(decisions);
  const constraints = only(passport, 'constraints');
  if (constraints) out.constraints = textOf(constraints);
  const questions = only(passport, 'openQuestions');
  if (questions) out.openQuestions = textOf(questions);
  const sources = only(passport, 'sources');
  if (sources) out.source_manifest = readSources(passport);
  return out;
}

function readSlides(story) {
  return kids(story, 'slide').map(slide => {
    const item = { slide_id: slide.getAttribute('slideId') || slide.getAttribute('slide_id') || '' };
    if (slide.hasAttribute('order')) item.order = Number(slide.getAttribute('order'));
    for (const key of STORY_FIELDS) {
      const child = only(slide, key);
      if (child) item[key] = textOf(child);
    }
    return item;
  });
}

function readStoryline(root) {
  const story = only(root, 'storyline');
  if (!story) return null;
  return {
    status: story.getAttribute('status') || 'not-built',
    slides: readSlides(story)
  };
}

function readExtras(root) {
  const extras = {};
  for (let n = root.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1 || n.namespaceURI !== NS) continue;
    if (n.localName === 'passport' || n.localName === 'storyline') continue;
    if (Object.prototype.hasOwnProperty.call(extras, n.localName)) {
      throw new StorageError('INVALID_XML', `Multiple <${n.localName}> elements`);
    }
    extras[n.localName] = textOf(n);
  }
  return extras;
}

function semanticFromDoc(doc) {
  const root = rootOf(doc);
  return {
    passport: readPassport(root),
    storyline: readStoryline(root),
    extras: readExtras(root)
  };
}

function semanticFromXml(xmlText) {
  return semanticFromDoc(parse(xmlText));
}

function primaryOf(passport) {
  const list = passport && passport.audiences ? passport.audiences : [];
  return list.find(item => item.priority === 'primary') || list[0] || null;
}

function withConvenience(passport, element) {
  if (!passport) return null;
  const primary = primaryOf(passport);
  const hunt = primary && primary.hunt_stage ? primary.hunt_stage : {};
  const value = primary && primary.value_proposition ? primary.value_proposition : {};
  return Object.assign({}, passport, {
    revision: element ? revOf(element) : 0,
    updatedAt: element && element.getAttribute('updatedAt') ? element.getAttribute('updatedAt') : null,
    audience: primary && primary.identity != null ? primary.identity : '',
    jtbd: primary && primary.jtbd != null ? primary.jtbd : '',
    hunt_stage: {
      current: hunt.current || '',
      target: hunt.target || ''
    },
    value_proposition: {
      topic: value.topic || '',
      presentation: value.presentation || ''
    }
  });
}

function stateFromDoc(doc, partCount) {
  const root = rootOf(doc);
  const sem = semanticFromDoc(doc);
  const story = only(root, 'storyline');
  return {
    status: 'found',
    namespace: NS,
    partCount,
    revision: revOf(root),
    schemaVersion: root.getAttribute('schemaVersion') || null,
    updatedAt: root.getAttribute('updatedAt') || null,
    checksum: root.getAttribute('checksum') || null,
    passport: withConvenience(sem.passport, only(root, 'passport')),
    storyline: sem.storyline ? {
      revision: story ? revOf(story) : 0,
      status: sem.storyline.status,
      slides: sem.storyline.slides
    } : null,
    extras: sem.extras
  };
}

function absentState() {
  return {
    status: 'absent',
    namespace: NS,
    partCount: 0,
    revision: 0,
    schemaVersion: null,
    updatedAt: null,
    checksum: null,
    passport: null,
    storyline: null,
    extras: {}
  };
}

function getStateFromXml(xmlText) {
  if (xmlText == null || String(xmlText).trim() === '') return absentState();
  return stateFromDoc(parse(xmlText), 1);
}

function primaryElement(passport, create) {
  let box = only(passport, 'audiences');
  if (!box && !create) return null;
  if (!box) box = ensure(passport, 'audiences');
  const all = kids(box, 'audience');
  let primary = all.find(node => node.getAttribute('priority') === 'primary') || all[0] || null;
  if (!primary && create) {
    primary = passport.ownerDocument.createElementNS(NS, 'audience');
    primary.setAttribute('id', 'a1');
    primary.setAttribute('priority', 'primary');
    box.appendChild(primary);
  }
  return primary;
}

function setLeaf(parent, local, value) {
  setText(ensure(parent, local), value);
}

function setGoal(passport, key, value) {
  setText(ensure(ensure(passport, 'goals'), key), value);
}

function setAudienceLeaf(passport, local, value) {
  setText(ensure(primaryElement(passport, true), local), value);
}

function setHunt(passport, local, value) {
  const audience = primaryElement(passport, true);
  setText(ensure(ensure(audience, 'hunt'), local), value);
}

function setValue(passport, value) {
  const audience = primaryElement(passport, true);
  if (value && typeof value === 'object' && !Array.isArray(value)) {
    if (Object.prototype.hasOwnProperty.call(value, 'topic')) setText(ensure(audience, 'valueTopic'), value.topic);
    if (Object.prototype.hasOwnProperty.call(value, 'presentation')) setText(ensure(audience, 'valuePresentation'), value.presentation);
    return;
  }
  setText(ensure(audience, 'valuePresentation'), value);
}

function writeDecisions(passport, value) {
  const el = ensure(passport, 'methodologistDecisions');
  while (el.firstChild) el.removeChild(el.firstChild);
  const items = normalizeDecisions(value);
  if (items.length === 1 && typeof items[0] === 'string') {
    setText(el, items[0]);
    return;
  }
  for (const item of items) {
    const node = passport.ownerDocument.createElementNS(NS, 'decision');
    if (typeof item === 'string') setText(node, item);
    else {
      if (item.source) node.setAttribute('source', item.source);
      if (item.meetingDate) node.setAttribute('meetingDate', item.meetingDate);
      setText(node, item.text);
    }
    el.appendChild(node);
  }
}

function normalizeDecisions(value) {
  if (value == null || value === '') return [];
  if (typeof value === 'string') return [value];
  if (!Array.isArray(value)) throw new StorageError('BAD_PATH', 'methodological_decisions must be a string or an array');
  return value.map(item => {
    if (typeof item === 'string') return item;
    if (item && typeof item === 'object') {
      const out = { text: item.text == null ? '' : String(item.text) };
      if (item.source) out.source = String(item.source);
      if (item.meetingDate) out.meetingDate = String(item.meetingDate);
      return out;
    }
    throw new StorageError('BAD_PATH', 'A methodological decision must be a string or an object');
  });
}

function writeSources(passport, value) {
  if (!Array.isArray(value)) throw new StorageError('BAD_PATH', 'source_manifest must be an array');
  const box = ensure(passport, 'sources');
  for (const child of kids(box, 'source')) box.removeChild(child);
  for (const item of value) {
    if (!item || typeof item !== 'object') throw new StorageError('BAD_PATH', 'A source must be an object');
    const node = passport.ownerDocument.createElementNS(NS, 'source');
    for (const attr of ['file', 'stage', 'meetingDate']) {
      if (item[attr] != null && item[attr] !== '') node.setAttribute(attr, String(item[attr]));
    }
    if (item.text) setText(node, item.text);
    box.appendChild(node);
  }
}

function writeAudience(doc, data) {
  const node = doc.createElementNS(NS, 'audience');
  if (data.id) node.setAttribute('id', String(data.id));
  if (data.priority) node.setAttribute('priority', String(data.priority));
  for (const key of AUDIENCE_TEXT) {
    if (Object.prototype.hasOwnProperty.call(data, key)) setLeaf(node, key, data[key]);
  }
  if (data.inferredMotives && typeof data.inferredMotives === 'object') {
    const motives = ensure(node, 'inferredMotives');
    for (const attr of ['status', 'confidence', 'basis']) {
      if (data.inferredMotives[attr] != null) motives.setAttribute(attr, String(data.inferredMotives[attr]));
    }
    setText(motives, data.inferredMotives.text || '');
  }
  if (data.hunt_stage && typeof data.hunt_stage === 'object') {
    const hunt = ensure(node, 'hunt');
    if (Object.prototype.hasOwnProperty.call(data.hunt_stage, 'current')) setLeaf(hunt, 'initial', data.hunt_stage.current);
    if (Object.prototype.hasOwnProperty.call(data.hunt_stage, 'target')) setLeaf(hunt, 'target', data.hunt_stage.target);
    if (Object.prototype.hasOwnProperty.call(data.hunt_stage, 'rationale')) setLeaf(hunt, 'rationale', data.hunt_stage.rationale);
  }
  if (data.value_proposition && typeof data.value_proposition === 'object') {
    if (Object.prototype.hasOwnProperty.call(data.value_proposition, 'topic')) setLeaf(node, 'valueTopic', data.value_proposition.topic);
    if (Object.prototype.hasOwnProperty.call(data.value_proposition, 'presentation')) setLeaf(node, 'valuePresentation', data.value_proposition.presentation);
  }
  if (data.change && typeof data.change === 'object') {
    const change = ensure(node, 'change');
    for (const key of ['before', 'after', 'action']) {
      if (Object.prototype.hasOwnProperty.call(data.change, key)) setLeaf(change, key, data.change[key]);
    }
  }
  return node;
}

function replaceAudiences(passport, audiences) {
  if (!Array.isArray(audiences)) throw new StorageError('BAD_PATH', 'audiences must be an array');
  const box = ensure(passport, 'audiences');
  for (const child of kids(box, 'audience')) box.removeChild(child);
  for (const audience of audiences) {
    if (!audience || typeof audience !== 'object') throw new StorageError('BAD_PATH', 'An audience must be an object');
    box.appendChild(writeAudience(passport.ownerDocument, audience));
  }
}

function applyPassportFields(passport, data) {
  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    throw new StorageError('BAD_PATH', 'passport payload must be an object');
  }
  if (Object.prototype.hasOwnProperty.call(data, 'status')) passport.setAttribute('status', String(data.status));
  if (Object.prototype.hasOwnProperty.call(data, 'projectTitle')) setLeaf(passport, 'projectTitle', data.projectTitle);
  if (Object.prototype.hasOwnProperty.call(data, 'project_stage')) setLeaf(passport, 'currentStage', data.project_stage);
  if (Object.prototype.hasOwnProperty.call(data, 'businessContext')) setLeaf(passport, 'businessContext', data.businessContext);
  if (Object.prototype.hasOwnProperty.call(data, 'mainMessage')) setLeaf(passport, 'mainMessage', data.mainMessage);
  if (Object.prototype.hasOwnProperty.call(data, 'constraints')) setLeaf(passport, 'constraints', data.constraints);
  if (Object.prototype.hasOwnProperty.call(data, 'openQuestions')) setLeaf(passport, 'openQuestions', data.openQuestions);
  if (data.goals && typeof data.goals === 'object') {
    for (const key of GOAL_FIELDS) {
      if (Object.prototype.hasOwnProperty.call(data.goals, key)) setGoal(passport, key, data.goals[key]);
    }
  }
  if (Object.prototype.hasOwnProperty.call(data, 'methodological_decisions')) writeDecisions(passport, data.methodological_decisions);
  if (Object.prototype.hasOwnProperty.call(data, 'source_manifest')) writeSources(passport, data.source_manifest);
  if (Object.prototype.hasOwnProperty.call(data, 'audiences')) {
    replaceAudiences(passport, data.audiences);
    return;
  }
  if (Object.prototype.hasOwnProperty.call(data, 'audience')) setAudienceLeaf(passport, 'identity', data.audience);
  if (Object.prototype.hasOwnProperty.call(data, 'jtbd')) setAudienceLeaf(passport, 'jtbd', data.jtbd);
  if (data.hunt_stage && typeof data.hunt_stage === 'object') {
    if (Object.prototype.hasOwnProperty.call(data.hunt_stage, 'current')) setHunt(passport, 'initial', data.hunt_stage.current);
    if (Object.prototype.hasOwnProperty.call(data.hunt_stage, 'target')) setHunt(passport, 'target', data.hunt_stage.target);
  }
  if (Object.prototype.hasOwnProperty.call(data, 'value_proposition')) setValue(passport, data.value_proposition);
}

function ensurePassport(root) {
  let passport = only(root, 'passport');
  if (passport) return passport;
  passport = root.ownerDocument.createElementNS(NS, 'passport');
  passport.setAttribute('status', 'draft');
  const story = only(root, 'storyline');
  if (story) root.insertBefore(passport, story);
  else root.insertBefore(passport, root.firstChild);
  return passport;
}

function ensureStory(root) {
  let story = only(root, 'storyline');
  if (story) return story;
  story = root.ownerDocument.createElementNS(NS, 'storyline');
  story.setAttribute('status', 'not-built');
  story.setAttribute('revision', '0');
  const extras = [];
  for (let n = root.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1 && n.namespaceURI === NS && n.localName !== 'passport' && n.localName !== 'storyline') extras.push(n);
  }
  if (extras.length) root.insertBefore(story, extras[0]);
  else root.appendChild(story);
  return story;
}

function setSlideField(slide, key, value) {
  if (key === 'order') {
    slide.setAttribute('order', String(value));
    return;
  }
  if (!STORY_FIELDS.includes(key)) throw new StorageError('BAD_PATH', `Unsupported storyline field ${key}`);
  setText(ensure(slide, key), value);
}

function findSlide(story, slideId) {
  return kids(story, 'slide').find(slide => (slide.getAttribute('slideId') || slide.getAttribute('slide_id')) === String(slideId)) || null;
}

function patchStorylineDom(root, slideId, patch) {
  if (slideId == null || String(slideId).trim() === '') throw new StorageError('BAD_PATH', 'slide_id is required');
  if (!patch || typeof patch !== 'object' || Array.isArray(patch)) throw new StorageError('BAD_PATH', 'storyline patch must be an object');
  for (const key of Object.keys(patch)) {
    if (key !== 'order' && !STORY_FIELDS.includes(key)) throw new StorageError('BAD_PATH', `Unsupported storyline field ${key}`);
  }
  const story = ensureStory(root);
  let slide = findSlide(story, slideId);
  if (!slide) {
    slide = root.ownerDocument.createElementNS(NS, 'slide');
    slide.setAttribute('slideId', String(slideId));
    const orders = kids(story, 'slide').map(item => Number(item.getAttribute('order') || 0));
    slide.setAttribute('order', String(orders.length ? Math.max(...orders) + 1 : 1));
    story.appendChild(slide);
    if (!story.getAttribute('status') || story.getAttribute('status') === 'not-built') story.setAttribute('status', 'draft');
  }
  for (const [key, value] of Object.entries(patch)) setSlideField(slide, key, value);
}

function replaceStorylineDom(root, data) {
  if (!data || typeof data !== 'object') throw new StorageError('BAD_PATH', 'storyline payload must be an object');
  const story = ensureStory(root);
  if (Object.prototype.hasOwnProperty.call(data, 'status')) story.setAttribute('status', String(data.status));
  if (!Object.prototype.hasOwnProperty.call(data, 'slides')) return;
  if (!Array.isArray(data.slides)) throw new StorageError('BAD_PATH', 'storyline.slides must be an array');
  for (const child of kids(story, 'slide')) story.removeChild(child);
  data.slides.forEach((item, index) => {
    if (!item || item.slide_id == null || String(item.slide_id).trim() === '') {
      throw new StorageError('BAD_PATH', 'Every storyline slide needs a slide_id');
    }
    const slide = root.ownerDocument.createElementNS(NS, 'slide');
    slide.setAttribute('slideId', String(item.slide_id));
    slide.setAttribute('order', String(item.order == null ? index + 1 : item.order));
    for (const key of STORY_FIELDS) {
      if (Object.prototype.hasOwnProperty.call(item, key)) setSlideField(slide, key, item[key]);
    }
    story.appendChild(slide);
  });
  if (data.slides.length && (!story.getAttribute('status') || story.getAttribute('status') === 'not-built')) {
    story.setAttribute('status', 'draft');
  }
}

function applyExtras(root, extras) {
  if (!extras || typeof extras !== 'object') throw new StorageError('BAD_PATH', 'extras must be an object');
  for (const [name, value] of Object.entries(extras)) {
    if (!/^[A-Za-z_][\w.-]*$/.test(name) || name === 'passport' || name === 'storyline') {
      throw new StorageError('BAD_PATH', `Unsupported extra element ${name}`);
    }
    setText(ensure(root, name), value);
  }
}

const PATCH_PATHS = {
  'status': (passport, value) => passport.setAttribute('status', String(value)),
  'projecttitle': (passport, value) => setLeaf(passport, 'projectTitle', value),
  'project_stage': (passport, value) => setLeaf(passport, 'currentStage', value),
  'businesscontext': (passport, value) => setLeaf(passport, 'businessContext', value),
  'mainmessage': (passport, value) => setLeaf(passport, 'mainMessage', value),
  'constraints': (passport, value) => setLeaf(passport, 'constraints', value),
  'openquestions': (passport, value) => setLeaf(passport, 'openQuestions', value),
  'goals.business': (passport, value) => setGoal(passport, 'business', value),
  'goals.communication': (passport, value) => setGoal(passport, 'communication', value),
  'goals.learning': (passport, value) => setGoal(passport, 'learning', value),
  'audience': (passport, value) => setAudienceLeaf(passport, 'identity', value),
  'jtbd': (passport, value) => setAudienceLeaf(passport, 'jtbd', value),
  'hunt_stage.current': (passport, value) => setHunt(passport, 'initial', value),
  'hunt_stage.target': (passport, value) => setHunt(passport, 'target', value),
  'value_proposition': (passport, value) => setValue(passport, value),
  'value_proposition.topic': (passport, value) => setValue(passport, { topic: value }),
  'value_proposition.presentation': (passport, value) => setValue(passport, { presentation: value }),
  'methodological_decisions': (passport, value) => writeDecisions(passport, value),
  'source_manifest': (passport, value) => writeSources(passport, value)
};

function normalizePath(path) {
  return String(path || '').trim().replace(/\//g, '.').replace(/[A-Z]/g, ch => ch.toLowerCase());
}

function patchPassportDom(root, path, value) {
  const key = normalizePath(path);
  const fn = PATCH_PATHS[key];
  if (!fn) throw new StorageError('BAD_PATH', `Unsupported passport path ${path}`);
  if (value === undefined) throw new StorageError('BAD_PATH', 'A patch value is required');
  fn(ensurePassport(root), value);
}

function assertExpected(currentRevision, expected) {
  if (expected === undefined || expected === null || expected === '') return;
  if (typeof expected === 'string' && !/^\d+$/.test(expected)) {
    throw new StorageError('STALE_REVISION', 'expected_revision must be an integer');
  }
  const wanted = Number(expected);
  if (!Number.isInteger(wanted) || wanted < 0) {
    throw new StorageError('STALE_REVISION', 'expected_revision must be an integer');
  }
  if (wanted !== currentRevision) {
    throw new StorageError('STALE_REVISION', `expected revision ${wanted} but the current revision is ${currentRevision}`);
  }
}

function finalize(doc, before, now) {
  const after = semanticFromDoc(doc);
  if (same(before, after)) return { wrote: false, semantic: after };
  const root = rootOf(doc);
  const stamp = now();
  root.setAttribute('schemaVersion', '1');
  root.setAttribute('revision', String(revOf(root) + 1));
  root.setAttribute('updatedAt', stamp);
  const passport = only(root, 'passport');
  if (passport && !same(before.passport, after.passport)) {
    passport.setAttribute('revision', String(revOf(passport) + 1));
    passport.setAttribute('updatedAt', stamp);
    if (!passport.getAttribute('status')) passport.setAttribute('status', 'draft');
  }
  const story = only(root, 'storyline');
  if (story && !same(before.storyline, after.storyline)) {
    story.setAttribute('revision', String(revOf(story) + 1));
  }
  const semantic = semanticFromDoc(doc);
  root.setAttribute('checksum', sha256(canonical(semantic)));
  return { wrote: true, semantic };
}

function applyOperation(xmlText, operation, options = {}) {
  const now = options.now || (() => new Date().toISOString());
  const hasXml = xmlText != null && String(xmlText).trim() !== '';
  if (!hasXml && operation.type !== 'replace') {
    throw new StorageError('NOT_FOUND', 'No SlideFactory Custom XML part');
  }
  const currentRevision = hasXml ? revOf(rootOf(parse(xmlText))) : 0;
  assertExpected(currentRevision, operation.expectedRevision);
  const doc = hasXml ? parse(xmlText) : emptyDoc();
  const before = semanticFromDoc(doc);
  const root = rootOf(doc);
  if (operation.type === 'replace') {
    const payload = operation.payload || {};
    if (payload.passport) applyPassportFields(ensurePassport(root), payload.passport);
    if (payload.storyline) replaceStorylineDom(root, payload.storyline);
    if (payload.extras) applyExtras(root, payload.extras);
  } else if (operation.type === 'patch_passport') {
    patchPassportDom(root, operation.path, operation.value);
  } else if (operation.type === 'patch_storyline') {
    patchStorylineDom(root, operation.slideId, operation.patch || {});
  } else {
    throw new StorageError('BAD_PATH', `Unsupported operation ${operation.type}`);
  }
  const done = finalize(doc, before, now);
  const storedXml = done.wrote ? xml.serialize(doc) : String(xmlText);
  const view = done.wrote ? stateFromDoc(doc, 1) : (hasXml ? stateFromDoc(parse(xmlText), 1) : absentState());
  return {
    wrote: done.wrote,
    xml: storedXml,
    semantic: done.wrote ? done.semantic : before,
    result: Object.assign({}, view, {
      status: done.wrote ? 'saved' : 'unchanged',
      partCount: hasXml || done.wrote ? 1 : 0
    })
  };
}

function verifyXml(xmlText) {
  const issues = [];
  if (xmlText == null || String(xmlText).trim() === '') {
    return { ok: false, namespace: NS, partCount: 0, revision: 0, issues: ['missing part'], checksum: { status: 'absent', value: null } };
  }
  let doc;
  try {
    doc = parse(xmlText);
    rootOf(doc);
  } catch (err) {
    return {
      ok: false,
      namespace: NS,
      partCount: 1,
      issues: [err.message],
      code: err.code || 'INVALID_XML',
      checksum: { status: 'absent', value: null }
    };
  }
  const root = rootOf(doc);
  let revision = 0;
  try {
    revision = revOf(root);
    if (only(root, 'passport')) revOf(only(root, 'passport'));
    if (only(root, 'storyline')) revOf(only(root, 'storyline'));
  } catch (err) {
    issues.push(err.message);
  }
  let semantic;
  try {
    semantic = semanticFromDoc(doc);
  } catch (err) {
    issues.push(err.message);
  }
  const stored = root.getAttribute('checksum');
  let checksum = { status: 'absent', value: null };
  if (stored && semantic) {
    const actual = sha256(canonical(semantic));
    checksum = { status: stored === actual ? 'match' : 'mismatch', value: stored };
    if (stored !== actual) issues.push('checksum mismatch');
  }
  return {
    ok: issues.length === 0,
    namespace: NS,
    partCount: 1,
    revision,
    schemaVersion: root.getAttribute('schemaVersion') || null,
    updatedAt: root.getAttribute('updatedAt') || null,
    checksum,
    issues
  };
}

function normalizeSlideText(parts) {
  const list = Array.isArray(parts) ? parts : [parts];
  return list.map(part => String(part == null ? '' : part).replace(/\s+/g, ' ').trim()).filter(Boolean).join('\n');
}

function fingerprintFromSlides(slides) {
  const rows = (slides || []).map(slide => {
    const text = normalizeSlideText(slide.text == null ? slide.texts || '' : slide.text);
    const slideId = String(slide.slideId != null ? slide.slideId : slide.slide_id);
    return {
      slide_id: slideId,
      order: slide.order,
      fingerprint: sha256(`${slideId}\n${text}`),
      charCount: text.length
    };
  });
  return { count: rows.length, slides: rows };
}

function queue(store, fn) {
  const previous = store.__queue || Promise.resolve();
  const run = previous.then(fn, fn);
  store.__queue = run.then(() => undefined, () => undefined);
  return run;
}

async function readOne(store, namespace) {
  assertNamespace(namespace);
  const parts = await store.listParts(namespace);
  if (parts.length > 1) throw new StorageError('CONFLICT', `${parts.length} Custom XML parts found for ${namespace}`);
  return parts[0] || null;
}

async function commit(store, namespace, applied) {
  if (!applied.wrote) return applied.result;
  await store.savePart(namespace, applied.xml);
  const parts = await store.listParts(namespace);
  if (parts.length !== 1) throw new StorageError('VERIFY_FAILED', `Write verification failed: found ${parts.length} parts`);
  const back = semanticFromXml(parts[0].xml);
  if (!same(back, applied.semantic)) throw new StorageError('VERIFY_FAILED', 'Write verification failed: stored content differs');
  const stored = getStateFromXml(parts[0].xml);
  if (parts[0].id) stored.partId = parts[0].id;
  if (stored.revision !== applied.result.revision) throw new StorageError('VERIFY_FAILED', 'Write verification failed: revision differs');
  return Object.assign({}, stored, { status: 'saved' });
}

async function get_state(store, namespace) {
  return queue(store, async () => {
    const part = await readOne(store, namespace);
    if (!part) return absentState();
    const state = getStateFromXml(part.xml);
    if (part.id) state.partId = part.id;
    return state;
  });
}

async function replace_state(store, namespace, payload, expectedRevision, options) {
  return queue(store, async () => {
    const part = await readOne(store, namespace);
    const applied = applyOperation(part && part.xml, {
      type: 'replace',
      payload: payload || {},
      expectedRevision
    }, options);
    return commit(store, namespace, applied);
  });
}

async function patch_passport(store, namespace, path, value, expectedRevision, options) {
  return queue(store, async () => {
    const part = await readOne(store, namespace);
    const applied = applyOperation(part && part.xml, {
      type: 'patch_passport',
      path,
      value,
      expectedRevision
    }, options);
    return commit(store, namespace, applied);
  });
}

async function patch_storyline(store, namespace, slideId, patch, expectedRevision, options) {
  return queue(store, async () => {
    const part = await readOne(store, namespace);
    const applied = applyOperation(part && part.xml, {
      type: 'patch_storyline',
      slideId,
      patch,
      expectedRevision
    }, options);
    return commit(store, namespace, applied);
  });
}

async function verify_state(store, namespace) {
  return queue(store, async () => {
    assertNamespace(namespace);
    let parts;
    try {
      parts = await store.listParts(namespace);
    } catch (err) {
      if (err instanceof StorageError) return { ok: false, namespace, code: err.code, issues: [err.message], partCount: null };
      throw err;
    }
    if (parts.length !== 1) {
      return {
        ok: false,
        namespace,
        partCount: parts.length,
        revision: 0,
        issues: [parts.length === 0 ? 'missing part' : `${parts.length} parts found`],
        checksum: { status: 'absent', value: null }
      };
    }
    const report = verifyXml(parts[0].xml);
    if (parts[0].id) report.partId = parts[0].id;
    return report;
  });
}

async function get_slide_fingerprints(store) {
  return queue(store, async () => fingerprintFromSlides(await store.listSlides()));
}

globalThis.SlideFactory = {
  NS,
  ROOT,
  StorageError,
  sha256,
  canonical,
  same,
  getStateFromXml,
  semanticFromXml,
  verifyXml,
  applyOperation,
  normalizeSlideText,
  fingerprintFromSlides,
  get_state,
  replace_state,
  patch_passport,
  patch_storyline,
  verify_state,
  get_slide_fingerprints
};

(function () {
'use strict';

const SF = (typeof globalThis !== 'undefined' && globalThis.SlideFactory && globalThis.SlideFactory.applyOperation)
  ? globalThis.SlideFactory
  : require('./core');

const PASSPORT_PATHS = [
  'status',
  'projecttitle',
  'project_stage',
  'businesscontext',
  'mainmessage',
  'constraints',
  'openquestions',
  'goals.business',
  'goals.communication',
  'goals.learning',
  'audience',
  'jtbd',
  'hunt_stage.current',
  'hunt_stage.target',
  'value_proposition',
  'value_proposition.topic',
  'value_proposition.presentation',
  'methodological_decisions',
  'source_manifest'
];

const STORY_FIELDS = ['order', 'section', 'message', 'role', 'transition', 'recommendation', 'status', 'fingerprint'];
const PATCH_KEYS = ['format', 'version', 'expected_revision', 'passport_patches', 'storyline_patches'];
const PROJECTIONS = ['project', 'current-slide', 'full'];

function fail(code, message) {
  throw new SF.StorageError(code, message);
}

function clone(value) {
  if (value == null) return value;
  return JSON.parse(JSON.stringify(value));
}

function normalizePath(path) {
  return String(path == null ? '' : path).trim().replace(/\//g, '.').replace(/[A-Z]/g, ch => ch.toLowerCase());
}

function exactKeys(value, allowed, label) {
  const keys = Object.keys(value);
  for (const key of keys) {
    if (!allowed.includes(key)) fail('BAD_PATCH', `${label} contains unsupported field ${key}`);
  }
}

function summarizeStoryline(storyline) {
  if (!storyline) return null;
  return {
    revision: storyline.revision,
    status: storyline.status,
    slides: (storyline.slides || []).map(slide => {
      const item = { slide_id: slide.slide_id, order: slide.order };
      if (slide.section != null) item.section = slide.section;
      if (slide.status != null) item.status = slide.status;
      return item;
    })
  };
}

function findSlide(state, slideId) {
  const slides = state.storyline && state.storyline.slides ? state.storyline.slides : [];
  return slides.find(slide => slide.slide_id === slideId) || null;
}

function projectContext(state, options = {}) {
  if (!state || typeof state !== 'object') fail('BAD_PATCH', 'A stored state is required');
  const projection = options.projection;
  if (!PROJECTIONS.includes(projection)) fail('BAD_PATH', 'projection must be project, current-slide, or full');
  const context = {
    format: 'SLIDEFACTORY_CONTEXT',
    version: 1,
    projection,
    namespace: state.namespace || SF.NS,
    revision: state.revision == null ? 0 : state.revision,
    schemaVersion: state.schemaVersion || null,
    updatedAt: state.updatedAt || null,
    checksum: state.checksum || null,
    status: state.status || 'absent',
    passport: clone(state.passport),
    extras: clone(state.extras || {})
  };
  if (projection === 'project') {
    context.storyline = summarizeStoryline(state.storyline);
    return context;
  }
  const slideId = options.currentSlideId == null || options.currentSlideId === ''
    ? null
    : String(options.currentSlideId);
  const fingerprints = Array.isArray(options.fingerprints) ? options.fingerprints : [];
  const fingerprint = slideId ? fingerprints.find(item => String(item.slide_id) === slideId) || null : null;
  const entry = slideId ? findSlide(state, slideId) : null;
  context.currentSlide = slideId ? {
    slide_id: slideId,
    order: entry && entry.order != null ? entry.order : (fingerprint ? fingerprint.order : null),
    fingerprint: fingerprint ? fingerprint.fingerprint : (entry && entry.fingerprint ? entry.fingerprint : null)
  } : null;
  if (projection === 'current-slide') {
    context.storylineEntry = clone(entry);
    return context;
  }
  context.storyline = clone(state.storyline);
  context.slideFingerprints = clone(fingerprints);
  return context;
}

function serializeContext(context) {
  return JSON.stringify(context, null, 2);
}

function extractJson(text) {
  const raw = String(text == null ? '' : text).replace(/^\uFEFF/, '').trim();
  if (!raw) fail('BAD_PATCH', 'SLIDEFACTORY_PATCH is empty');
  const fences = [...raw.matchAll(/```(?:json|slidefactory-patch)?[^\n]*\n([\s\S]*?)```/gi)];
  if (fences.length > 1) fail('BAD_PATCH', 'SLIDEFACTORY_PATCH must contain one JSON block');
  const body = (fences.length === 1 ? fences[0][1] : raw).trim();
  if (!body.startsWith('{') || !body.endsWith('}')) fail('BAD_PATCH', 'SLIDEFACTORY_PATCH must be one JSON object');
  try {
    return JSON.parse(body);
  } catch (err) {
    fail('BAD_PATCH', 'SLIDEFACTORY_PATCH is not valid JSON');
  }
}

function parsePatch(input) {
  const data = typeof input === 'string' || input == null ? extractJson(input) : input;
  if (!data || typeof data !== 'object' || Array.isArray(data)) fail('BAD_PATCH', 'SLIDEFACTORY_PATCH must be an object');
  exactKeys(data, PATCH_KEYS, 'SLIDEFACTORY_PATCH');
  if (data.format !== 'SLIDEFACTORY_PATCH') fail('BAD_PATCH', 'format must be SLIDEFACTORY_PATCH');
  if (data.version !== 1) fail('BAD_PATCH', 'version must be 1');
  if (!Object.prototype.hasOwnProperty.call(data, 'expected_revision') || data.expected_revision === null || data.expected_revision === '') {
    fail('MISSING_REVISION', 'expected_revision is required');
  }
  if (typeof data.expected_revision !== 'number' || !Number.isInteger(data.expected_revision) || data.expected_revision < 0) {
    fail('BAD_PATCH', 'expected_revision must be a non-negative integer');
  }
  const passportPatches = data.passport_patches == null ? [] : data.passport_patches;
  const storylinePatches = data.storyline_patches == null ? [] : data.storyline_patches;
  if (!Array.isArray(passportPatches) || !Array.isArray(storylinePatches)) {
    fail('BAD_PATCH', 'passport_patches and storyline_patches must be arrays');
  }
  if (passportPatches.length + storylinePatches.length === 0) fail('BAD_PATCH', 'SLIDEFACTORY_PATCH has no changes');
  const seen = new Set();
  const passport = passportPatches.map(item => {
    if (!item || typeof item !== 'object' || Array.isArray(item)) fail('BAD_PATCH', 'A passport patch must be an object');
    exactKeys(item, ['path', 'value'], 'passport patch');
    if (typeof item.path !== 'string' || !Object.prototype.hasOwnProperty.call(item, 'value')) {
      fail('BAD_PATCH', 'A passport patch needs path and value');
    }
    const path = normalizePath(item.path);
    if (!PASSPORT_PATHS.includes(path)) fail('BAD_PATH', `Unsupported passport path ${item.path}`);
    if (seen.has(path)) fail('BAD_PATCH', `Duplicate passport path ${path}`);
    if ((path === 'value_proposition' && (seen.has('value_proposition.topic') || seen.has('value_proposition.presentation')))
      || (path.startsWith('value_proposition.') && seen.has('value_proposition'))) {
      fail('BAD_PATCH', 'Use either value_proposition or its topic and presentation fields');
    }
    seen.add(path);
    if (item.value === undefined || item.value === null) fail('BAD_PATCH', `Passport path ${path} needs a value`);
    validatePassportValue(path, item.value);
    return { path, value: item.value };
  });
  const storyline = storylinePatches.map(item => {
    if (!item || typeof item !== 'object' || Array.isArray(item)) fail('BAD_PATCH', 'A storyline patch must be an object');
    exactKeys(item, ['slide_id', 'patch'], 'storyline patch');
    if (typeof item.slide_id !== 'string' || item.slide_id.trim() === '') fail('BAD_PATCH', 'slide_id must be a non-empty string');
    if (!item.patch || typeof item.patch !== 'object' || Array.isArray(item.patch)) fail('BAD_PATCH', 'storyline patch must be an object');
    exactKeys(item.patch, STORY_FIELDS, 'storyline patch');
    if (Object.keys(item.patch).length === 0) fail('BAD_PATCH', 'storyline patch is empty');
    const patch = {};
    for (const key of Object.keys(item.patch)) {
      const value = item.patch[key];
      if (key === 'order') {
        if (typeof value !== 'number' || !Number.isInteger(value) || value < 0) fail('BAD_PATCH', 'order must be a non-negative integer');
      } else if (typeof value !== 'string') {
        fail('BAD_PATCH', `${key} must be a string`);
      }
      patch[key] = value;
    }
    return { slide_id: item.slide_id, patch };
  });
  return {
    format: 'SLIDEFACTORY_PATCH',
    version: 1,
    expected_revision: data.expected_revision,
    passport_patches: passport,
    storyline_patches: storyline
  };
}

function validatePassportValue(path, value) {
  if (path === 'source_manifest' && !Array.isArray(value)) fail('BAD_PATCH', 'source_manifest must be an array');
  if (path === 'methodological_decisions' && typeof value !== 'string' && !Array.isArray(value)) {
    fail('BAD_PATCH', 'methodological_decisions must be a string or an array');
  }
  if (path === 'value_proposition' && value && typeof value === 'object' && !Array.isArray(value)) {
    exactKeys(value, ['topic', 'presentation'], 'value_proposition');
  } else if (path === 'value_proposition' && typeof value !== 'string') {
    fail('BAD_PATCH', 'value_proposition must be a string or an object');
  } else if (path.startsWith('value_proposition.') && typeof value !== 'string') {
    fail('BAD_PATCH', `${path} must be a string`);
  } else if (!path.startsWith('value_proposition') && path !== 'source_manifest' && path !== 'methodological_decisions' && typeof value !== 'string') {
    fail('BAD_PATCH', `${path} must be a string`);
  }
}

function passportPayload(patches) {
  const passport = {};
  for (const item of patches) {
    const path = item.path;
    if (path === 'goals.business' || path === 'goals.communication' || path === 'goals.learning') {
      passport.goals = passport.goals || {};
      passport.goals[path.slice(6)] = item.value;
    } else if (path === 'projecttitle') passport.projectTitle = item.value;
    else if (path === 'businesscontext') passport.businessContext = item.value;
    else if (path === 'mainmessage') passport.mainMessage = item.value;
    else if (path === 'openquestions') passport.openQuestions = item.value;
    else if (path === 'value_proposition.topic' || path === 'value_proposition.presentation') {
      const current = passport.value_proposition && typeof passport.value_proposition === 'object'
        ? passport.value_proposition
        : {};
      passport.value_proposition = Object.assign({}, current, { [path.slice('value_proposition.'.length)]: item.value });
    } else if (path === 'hunt_stage.current' || path === 'hunt_stage.target') {
      passport.hunt_stage = passport.hunt_stage || {};
      passport.hunt_stage[path.slice(11)] = item.value;
    } else {
      passport[path] = item.value;
    }
  }
  return passport;
}

function applyPatch(xml, input, options = {}) {
  const patch = parsePatch(input);
  if (xml == null || String(xml).trim() === '') fail('NOT_FOUND', 'No SlideFactory Custom XML part. Initialize it in the task pane first.');
  let currentXml = String(xml);
  let wrote = false;
  let expected = patch.expected_revision;
  if (patch.passport_patches.length) {
    const applied = SF.applyOperation(currentXml, {
      type: 'replace',
      payload: { passport: passportPayload(patch.passport_patches) },
      expectedRevision: expected
    }, options);
    if (applied.wrote) {
      currentXml = applied.xml;
      wrote = true;
      expected = applied.result.revision;
    }
  }
  for (const item of patch.storyline_patches) {
    const applied = SF.applyOperation(currentXml, {
      type: 'patch_storyline',
      slideId: item.slide_id,
      patch: item.patch,
      expectedRevision: expected
    }, options);
    if (applied.wrote) {
      currentXml = applied.xml;
      wrote = true;
      expected = applied.result.revision;
    }
  }
  const view = SF.getStateFromXml(currentXml);
  return {
    wrote,
    xml: wrote ? currentXml : String(xml),
    semantic: wrote ? SF.semanticFromXml(currentXml) : SF.semanticFromXml(String(xml)),
    result: Object.assign({}, view, { status: wrote ? 'saved' : 'unchanged' })
  };
}

const api = {
  PROJECTIONS,
  PASSPORT_PATHS,
  projectContext,
  serializeContext,
  parsePatch,
  applyPatch
};

if (typeof module === 'object' && module.exports) module.exports = api;
if (typeof globalThis !== 'undefined') globalThis.SlideFactoryPlus = api;

})();
