'use strict';

/* ───────────────────────────────────────────────────────────────
   Portfolio built directly from WEBSITE_DESIGN_CONCEPT.pptx.
   window.SLIDES holds the exact geometry of every PowerPoint slide
   (1338 x 753 canvas); each tab / popup is rebuilt from it.
─────────────────────────────────────────────────────────────── */
const W = 1338, H = 753;
const stage = document.getElementById('stage');
const MEDIA = 'assets/pptx/';

/* ── scale the 1338x753 canvas to fit any window ── */
function fit() {
  const s = Math.min(innerWidth / W, innerHeight / H);
  stage.style.transform = `translate(${(innerWidth - W * s) / 2}px, ${(innerHeight - H * s) / 2}px) scale(${s})`;
  stage.dataset.scale = s;
}
addEventListener('resize', fit); fit();

/* ── notes copied from the design document ── */
const NOTES = {
  masters: ['This paper was the final report written for my research project that I completed as part of my MSc Geophysics degree.',
            'I created figures 9, 10, 11, 20, 21, 22, 23, 24, and 25 through MATLAB.'],
  isotope: ['This is an assignment where the professor gave us some old, unused data. We students were given freedom to make our own scientific hypotheses and report with this data.'],
  fieldguide: ['In May of 2026, I attended an Academic Field School to Arizona to look at Planetary Analogues. Each of the students was tasked with preparing a guide for a different stop on the trip.'],
  mpres: ['I presented this slideshow to the other students in my program at a final celebration event where we all shared our research.'],
  physics: ['I made this presentation for a class I took in graduate school called Physics of the Earth.',
            'I chose this specific topic as it combined the planetary science topics that were being studied in the class with my own personal research on seismic technology.'],
  personal: ['These projects were created purely for enjoyment and allowed me to flex my creative muscles in new ways!'],
  mothers: ['I created this card for Mother\'s Day for my Mom!'],
  christmas: ['In 2025, my partner and I created this card to send to our family and friends.'],
  norway: ['I filmed many parts of a trip I took to Norway to create these travel vlogs.'],
  dayinlife: ['I wanted to create an "influencer"-style video showcasing an average day for me.'],
  acland: ['My aunt and uncle are building a space on their farm to host weddings and other events. They asked me to create content for their social media page. I crafted up this piece to show them my vision for their online presence.'],
  wateraid: ['WaterAid is a club I was part of in graduate school. I created this poster to hang up around campus advertising our sticker sale.'],
  fiddle: ['I am very involved in the fiddle community in my city, and I created this poster advertising an event for a jam night a friend of mine is planning on hosting.']
};

/* ── tabs (open in front of slide one) ── */
const FOLDERS = {
  scientific:    { slide: 's2',  hot: { 'Picture 25': 'masters', 'Picture 27': 'isotope', 'Picture 29': 'fieldguide' } },
  presentations: { slide: 's4',  hot: { 'Picture 19': 'mpres', 'Picture 17': 'physics' } },
  gallery:       { slide: 's6',  gallery: true, drop: ['Rectangle 21', 'Rectangle 27', 'Rectangle 29', 'Picture 20', 'TextBox 23'] },
  resume:        { slide: 's7',  pane: { rect: [317, 168, 682, 551], src: 'assets/docs/resume.pdf' }, drop: ['Picture 28', 'Picture 26'] },
  freelance:     { slide: 's8',  hot: { 'Picture 8': 'acland', 'Picture 15': 'wateraid', 'Picture 23': 'fiddle' } },
  personal:      { slide: 's11', hot: { 'Picture 10': 'mothers', 'Picture 24': 'christmas', 'Picture 14': 'norway', 'Picture 26': 'dayinlife' }, notes: 'personal' }
};

/* ── files (pop up, drag, close) ── */
const pdfUrl = f => `${f}#toolbar=0&navpanes=0&view=FitH`;
const DOC = (title, src, key) => ({ tpl: 's3', title, key, pane: { rect: [371, 183, 597, 537], src }, drop: ['Picture 25', 'Picture 12'], titleName: 'TextBox 9', stickies: ['Picture 15', 'Picture 17'] });
const FILES = {
  masters:    DOC('Masters_Final_Report', 'assets/docs/masters_report.pdf', 'masters'),
  isotope:    DOC('Isotope_Project',      'assets/docs/isotope_project.pdf', 'isotope'),
  fieldguide: DOC('Arizona_Field_Guide',  'assets/docs/field_guide.pdf', 'fieldguide'),
  mpres:   { tpl: 's5', title: 'Masters_Research_Presentation', key: 'mpres', deck: ['assets/deck_masters/Slide', 14], rect: [257, 195, 824, 465], drop: ['Picture 19', 'Picture 14'], titleName: 'TextBox 3', stickies: ['Picture 21'] },
  physics: { tpl: 's5', title: 'Physics_Project', key: 'physics', deck: ['assets/deck_physics/Slide', 16], rect: [257, 195, 824, 465], drop: ['Picture 19', 'Picture 14'], titleName: 'TextBox 3', stickies: ['Picture 21'], extraSticky: [1092, 325, 224, 207, 7.3] },
  acland:  { tpl: 's9', title: 'Acland_Acres_Instagram_Post', key: 'acland', deck: ['assets/acland_pages/page_', 4], ext: '.png', rect: [340, 160, 680, 535], stickyPos: { x: 42, y: 155, w: 255, h: 235, r: 352.7 }, drop: ['Picture 8', 'Picture 3', 'Picture 12'], stickies: [] },
  wateraid:{ tpl: 's9', title: 'Wateraid_Sticker_Sale', key: 'wateraid', image: ['assets/images/wateraid_poster.png', [438, 160, 500, 535]], drop: ['Picture 8', 'Picture 3', 'Picture 12'], stickies: [] },
  fiddle:  { tpl: 's9', title: 'Fiddle_Jam_Session_Poster', key: 'fiddle', image: ['assets/images/fiddle_poster.png', [438, 160, 500, 535]], drop: ['Picture 8', 'Picture 3', 'Picture 12'], stickies: [] },
  mothers: { tpl: 's9', title: 'Mothers_Day_Card', key: 'mothers', stickyPos: { x: 42, y: 155, w: 255, h: 235, r: 352.7 }, image: ['assets/images/mothers_day_card.png', [315, 155, 880, 535]], drop: ['Picture 8', 'Picture 3', 'Picture 12'], stickies: [] },
  christmas:{ tpl: 's9', title: '2025_Christmas_Card', key: 'christmas', stickyPos: { x: 42, y: 155, w: 255, h: 235, r: 352.7 }, image: ['assets/images/christmas_card.jpg', [315, 155, 880, 535]], drop: ['Picture 8', 'Picture 3', 'Picture 12'], stickies: [] },
  norway:  { tpl: 's12', key: 'norway', video: [
    ['Picture 21', 'Picture 9',  'rKYf_84GaMw', 'https://youtube.com/shorts/rKYf_84GaMw?feature=share'],
    ['Picture 25', 'Picture 14', 'JSnGVYy4oj8', 'https://youtube.com/shorts/JSnGVYy4oj8?feature=share'],
    ['Picture 28', 'Picture 13', 'aHsmWYlE3jE', 'https://youtube.com/shorts/aHsmWYlE3jE?feature=share'],
    ['Picture 30', 'Picture 17', 'P5F6-aiXbsU', 'https://youtube.com/shorts/P5F6-aiXbsU?feature=share']
  ], stickies: ['Picture 32'] },
  dayinlife:{ tpl: 's13', key: 'dayinlife', video: [
    ['Picture 28', 'Picture 5',  'GFBa7IBeq6w', 'https://youtube.com/shorts/GFBa7IBeq6w']
  ], stickies: ['Picture 18'] }
};

/* ── element rendering ── */
const el = (cls, css, html) => { const d = document.createElement('div'); if (cls) d.className = cls; Object.assign(d.style, css || {}); if (html != null) d.innerHTML = html; return d; };
const px = (x, y, w, h, r) => ({ left: x + 'px', top: y + 'px', width: w + 'px', height: h + 'px', transform: r ? `rotate(${r}deg)` : '' });

function drawPic(e) {
  const img = new Image(); img.src = MEDIA + e.img; img.draggable = false;
  if (!e.crop || e.crop.every(v => !v)) {
    img.style.cssText = 'width:100%;height:100%';
    const d = el('', px(e.x, e.y, e.w, e.h, e.r)); d.append(img); d.dataset.name = e.name; return d;
  }
  const [l, t, r, b] = e.crop, iw = e.w / (1 - (l + r) / 100), ih = e.h / (1 - (t + b) / 100);
  Object.assign(img.style, { width: iw + 'px', height: ih + 'px', left: (-l / 100 * iw) + 'px', top: (-t / 100 * ih) + 'px' });
  const d = el('crop', px(e.x, e.y, e.w, e.h, e.r)); d.append(img); d.dataset.name = e.name; return d;
}

function drawText(e, over) {
  const run = e.runs[0], text = over || e.runs.map(r => r.text).join('<br>');
  const hand = run.font && /Dreaming/i.test(run.font);
  const d = el('txt' + (hand ? ' hand' : ''), px(e.x, e.y, e.w, e.h, e.r), text);
  d.style.fontSize = (run.sz || 16) + 'px';
  d.style.textAlign = ({ ctr: 'center', r: 'right', l: 'left' })[run.al] || 'left';
  if (run.b) d.style.fontWeight = '700';
  if (run.color) d.style.color = '#' + run.color;
  d.dataset.name = e.name; return d;
}

function stickyText(e, text) {
  const box = el('sticky-box', px(e.x, e.y, e.w, e.h, e.r));
  const inner = el('sticky-inner', {}, text);
  const maxFs = Math.min(25, Math.round(e.w * 0.088));
  inner.dataset.maxfs = maxFs;
  inner.style.fontSize = maxFs + 'px';
  box.append(inner);
  return box;
}

/* shrink sticky font only as much as needed so text stays centered inside the yellow paper */
function fitStickies(scope = stage) {
  scope.querySelectorAll('.sticky-inner').forEach(inner => {
    let fs = parseFloat(inner.dataset.maxfs) || 22;
    inner.style.fontSize = fs + 'px';
    while (fs > 9 && (inner.scrollHeight > inner.clientHeight + 1 || inner.scrollWidth > inner.clientWidth + 1)) {
      fs -= 0.5;
      inner.style.fontSize = fs + 'px';
    }
  });
}
if (document.fonts && document.fonts.ready) {
  document.fonts.ready.then(() => fitStickies(stage));
}

function drawRect(e) {
  const d = el('', px(e.x, e.y, e.w, e.h, e.r));
  if (e.fill) d.style.background = '#' + e.fill;
  if (e.line || e.name === 'Rectangle 14') d.style.border = '2px solid #1c3b52';
  d.dataset.name = e.name; return d;
}

function homeHot(root) {
  const h = el('home-hot', px(1234, 625, 91, 95)); h.title = 'Home'; h.onclick = goHome; root.append(h);
}

/* slide-text boxes that sit on a sticky note: render them through stickyText so they are centred */
const SMAP = { s1: { 'Picture 5': 'TextBox 7', 'Picture 15': 'TextBox 19' }, s6: { 'Picture 9': 'TextBox 11' } };

/* build a slide canvas from the spec; opts can drop / restyle elements */
function build(key, opts = {}) {
  const root = el('slide');
  const list = SLIDES[key].filter(e => !(opts.drop || []).includes(e.name) && !(e.t === 'pic' && e.img === 'image16.png'));
  const smap = SMAP[key] || {}, skipText = Object.values(smap);
  const stickies = [];
  list.forEach(e => {
    let node;
    if (e.t === 'pic') node = drawPic(e);
    else if (e.t === 'rect') node = drawRect(e);
    else if (e.t === 'text') {
      if (skipText.includes(e.name)) return;
      if (opts.titleName === e.name) {
        node = el('txt title', { left: '36px', top: '69px', fontSize: '23px', lineHeight: '28px' }, opts.title);
      }
      else node = drawText(e);
    }
    if (!node) return;
    root.append(node);
    if (e.t === 'pic' && e.img === 'image5.png') {
      const s = { ...e, node, used: false };
      stickies.push(s);
      if (smap[e.name]) { const t = list.concat(SLIDES[key]).find(x => x.name === smap[e.name]);
        root.append(stickyText(e, t.runs.map(r => r.text).join('<br>'))); s.used = true; }
    }
  });
  return { root, list, stickies };
}

/* sticky-note overlays with the design document's notes */
function addNotes(root, stickies, texts, only) {
  const use = only ? stickies.filter(s => only.includes(s.name)) : stickies;
  use.forEach((s, i) => { if (texts && texts[i]) { root.append(stickyText(s, texts[i])); s.used = true; } });
}

/* erase sticky notes that have no text */
function pruneStickies(stickies) { stickies.forEach(s => { if (!s.used && s.node) s.node.remove(); }); }

/* ── tabs & popups ── */
let tab = null, popups = [], zTop = 50;

function closeAll() {
  popups.forEach(p => p.remove()); popups = [];
  if (tab) { tab.remove(); tab = null; }
}
function goHome() { closeAll(); }

function openFolder(name) {
  closeAll();
  const f = FOLDERS[name];
  const { root, list, stickies } = build(f.slide, { drop: f.drop });
  root.classList.add('tab'); root.dataset.name = name;
  if (f.hot) Object.entries(f.hot).forEach(([n, target]) => {
    const e = list.find(x => x.name === n); if (!e) return;
    const pad = 58;
    const h = el('hot', { left: e.x - 10 + 'px', top: e.y - 6 + 'px', width: e.w + 20 + 'px', height: e.h + pad + 'px' });
    h.title = target; h.onclick = () => openFile(target); root.append(h);
  });
  if (f.pane) { root.append(el('patch', { left: '136px', top: '58px', width: '420px', height: '46px' })); addPdf(root, f.pane.rect, f.pane.src, 'Resume'); }
  if (f.gallery) addGallery(root);
  if (f.notes) { const g = { name: 'x', x: 50, y: 155, w: 245, h: 225, r: 352.7 };
    root.append(drawPic({ ...g, t: 'pic', img: 'image5.png' })); root.append(stickyText(g, NOTES[f.notes][0])); }
  homeHot(root); addDots(root, () => goHome());
  stage.append(root); fitStickies(root); tab = root;
}

function addDots(root, onClose) {
  // macOS close button (red dot centered at x:35, y:29) with crisp centered X on hover
  const d = el('dot-hot', { left: '24px', top: '18px', width: '22px', height: '22px' },
    '<svg class="close-x" viewBox="0 0 10 10" width="8" height="8"><line x1="1.5" y1="1.5" x2="8.5" y2="8.5" stroke="#480202" stroke-width="1.6" stroke-linecap="round"/><line x1="8.5" y1="1.5" x2="1.5" y2="8.5" stroke="#480202" stroke-width="1.6" stroke-linecap="round"/></svg>');
  d.title = 'Close';
  d.onclick = onClose;
  root.append(d);

  // Sleek arrow coming from middle of text and pointing directly to red button
  const hint = el('back-hint', { left: '48px', top: '18px' },
    `<svg class="back-arrow" viewBox="0 0 80 22" width="80" height="22" fill="none">
       <path d="M 79 11 C 56 4, 24 4, 3 11" stroke="#111" stroke-width="1.6" stroke-linecap="round"/>
       <path d="M 10 7 L 3 11 L 10 15" stroke="#111" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>
     </svg>
     <span class="back-text">Click here to leave/go back</span>`);
  hint.onclick = onClose;
  hint.title = 'Close';
  root.append(hint);
}

function addPdf(root, rect, src, title) {
  const p = el('pane', { left: rect[0] + 'px', top: rect[1] + 'px', width: rect[2] + 'px', height: rect[3] + 'px' });
  const f = document.createElement('iframe'); f.src = pdfUrl(src); f.title = 'document'; p.append(f); root.append(p);
  addExpand(root, rect, { type: 'pdf', src, title });
}

/* ── subtle "open in full screen viewer" button ── */
function addExpand(root, rect, spec) {
  const b = el('expand-btn', { left: rect[0] + rect[2] - 62 + 'px', top: rect[1] + 8 + 'px' },
    '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 3h6v6M9 21H3v-6M21 3l-7 7M3 21l7-7"/></svg>');
  b.title = 'Open in full screen viewer';
  b.onclick = e => { e.stopPropagation(); openViewer(spec); };
  root.append(b);
}

function openViewer(spec) {
  closeViewer();
  const v = document.createElement('div'); v.id = 'viewer';
  const bar = document.createElement('div'); bar.className = 'v-bar';
  bar.innerHTML = `<span>${(spec.title || '').replace(/_/g, ' ')}</span><button aria-label="Close full screen viewer" title="Close (Esc)">&times;</button>`;
  bar.querySelector('button').onclick = closeViewer;
  const body = document.createElement('div'); body.className = 'v-body';
  if (spec.type === 'pdf') {
    const f = document.createElement('iframe'); f.src = `${spec.src}#toolbar=1&navpanes=0&view=FitH`; f.title = spec.title || 'document'; body.append(f);
  } else {
    body.classList.add('v-scroll');
    (spec.srcs || [spec.src]).forEach(s => { const im = new Image(); im.src = s; im.alt = ''; body.append(im); });
  }
  v.append(bar, body);
  v.addEventListener('mousedown', e => { if (e.target === v || e.target === body) closeViewer(); });
  document.body.append(v);
  body.scrollTop = 0;
}
function closeViewer() { const v = document.getElementById('viewer'); if (v) v.remove(); }
addEventListener('keydown', e => { if (e.key === 'Escape') closeViewer(); });

function addGallery(root) {
  const photos = [
    { src: 'assets/gallery/Bahamas_day_one.jfif', cap: 'My first day of field school. I travelled with a group of students to San Salvador Island in the Bahamas to study modern carbonates.', w: 327, h: 245 },
    { src: 'assets/gallery/Bahamas_hill.jfif', cap: 'Checking out inland geologic features on San Salvador Island.', w: 327, h: 245 },
    { src: 'assets/gallery/benoit_Outcrop.jfif', cap: 'My professor standing in front of a large beach outcrop.', w: 327, h: 245 },
    { src: 'assets/gallery/Parke_463.jfif', cap: 'Measuring sandy outcrop features in Northern Calgary.', w: 327, h: 245 },
    { src: 'assets/gallery/Jenny_Watefall.jfif', cap: 'Taking measurements at a cool waterfall in the Rocky Mountains.', w: 202, h: 270 },
    { src: 'assets/gallery/BC_Outcrop.jfif', cap: 'Exploring large metamorphic rock outcrops in Trail, British Columbia.', w: 327, h: 245 },
    { src: 'assets/gallery/Nova_Scotia_Ocean.jfif', cap: 'Looking at some far-off rocks along the coast of Nova Scotia.', w: 152, h: 270 },
    { src: 'assets/gallery/Nova_Scotia__Rock.jfif', cap: 'A beautifully unique rock I found in Nova Scotia.', w: 202, h: 270 },
    { src: 'assets/gallery/Geophones.jfif', cap: 'Seismic surveys in Ontario.', w: 202, h: 270 },
    { src: 'assets/gallery/Arizona_NtnlPark.jfif', cap: 'The gorgeous landscape in Canyonlands National Park, Utah.', w: 202, h: 270 },
    { src: 'assets/gallery/Arizona_Campsite.jfif', cap: 'A rainbow over our amazing campsite in Utah.', w: 202, h: 270 },
    { src: 'assets/gallery/Sudbury_FieldNotes.jfif', cap: 'Taking field notes in Sudbury.', w: 202, h: 270 }
  ];
  const sc = el('hscroll', { left: '219px', top: '195px', width: '944px', height: '468px' });
  const row = el('row');
  photos.forEach(p => {
    const m = el('mat');
    const img = new Image();
    img.src = p.src;
    img.alt = p.cap;
    img.style.width = p.w + 'px';
    img.style.height = p.h + 'px';
    const cap = el('cap', { maxWidth: p.w + 'px', width: p.w + 'px' }, p.cap);
    m.append(img, cap);
    row.append(m);
  });
  sc.append(row); root.append(sc);
  addExpand(root, [219, 195, 944, 468], { type: 'imgs', title: 'Geoscience_Photo_Gallery', srcs: photos.map(p => p.src) });
}

/* real file dates (the chrome image has 'November 2026' baked in, so it is covered and redrawn) */
const DATES = { masters: 'August 2026', isotope: 'November 2025', fieldguide: 'May 2026', mpres: 'August 2026', physics: 'November 2025',
  mothers: 'May 2024', christmas: 'December 2025', norway: 'July 2025', dayinlife: 'November 2026',
  acland: 'August 2026', wateraid: 'January 2026', fiddle: 'October 2026' };
function setDate(root, text) {
  root.append(el('patch', { left: '30px', top: '101px', width: '210px', height: '24px' }));
  root.append(el('', { left: '36px', top: '100px', whiteSpace: 'nowrap', font: "400 19.5px/22px 'Montserrat', sans-serif", color: '#3c3c3c', pointerEvents: 'none' }, text));
}

function openFile(name) {
  const f = FILES[name]; if (!f) return;
  const ex = popups.find(p => p.dataset.file === name); if (ex) { ex.style.zIndex = ++zTop; return; }
  const { root, list, stickies } = build(f.tpl, { drop: f.drop, title: f.title, titleName: f.titleName });
  root.classList.add('popup'); root.dataset.file = name;
  const notes = NOTES[f.key] || [];
  if (f.stickies) addNotes(root, stickies, notes, f.stickies.length ? f.stickies.slice(0, notes.length) : []);
  if (f.tpl === 's9' && notes[0]) {
    const sp = f.stickyPos || { x: 64, y: 215, w: 280, h: 258, r: 352.7 };
    const g = { name: 'x', ...sp };
    root.append(drawPic({ ...g, t: 'pic', img: 'image5.png' })); root.append(stickyText(g, notes[0]));
  }
  if (f.extraSticky && notes[1]) { const [x, y, w, h, r] = f.extraSticky; const g = { name: 'x', x, y, w, h, r };
    const im = drawPic({ ...g, t: 'pic', img: 'image5.png' }); root.append(im); root.append(stickyText(g, notes[1])); }
  pruneStickies(stickies);
  if (DATES[f.key]) setDate(root, DATES[f.key]);

  if (f.pane) addPdf(root, f.pane.rect, f.pane.src, f.title || 'Resume');
  if (f.deck) {
    const r = f.rect, p = el('pane flat deck', { left: r[0] + 'px', top: r[1] + 'px', width: r[2] + 'px', height: r[3] + 'px' });
    const ext = f.ext || '.PNG';
    for (let i = 1; i <= f.deck[1]; i++) { const im = new Image(); im.className = 'page'; im.loading = 'lazy'; im.src = `${f.deck[0]}${i}${ext}`; im.style.pointerEvents = 'auto'; p.append(im); }
    root.append(p);
    addExpand(root, r, { type: 'imgs', title: f.title, srcs: Array.from({ length: f.deck[1] }, (_, i) => `${f.deck[0]}${i + 1}${ext}`) });
  }
  if (f.image) {
    const r = f.image[1], p = el('pane flat single', { left: r[0] + 'px', top: r[1] + 'px', width: r[2] + 'px', height: r[3] + 'px' });
    const im = new Image(); im.className = 'page'; im.src = f.image[0]; p.append(im); root.append(p);
    addExpand(root, r, { type: 'img', title: f.title, src: f.image[0] });
  }
  if (f.title && !f.titleName && f.tpl === 's9' && !f.keepTitle) {
    root.append(el('patch', { left: '20px', top: '58px', width: '500px', height: '46px' }));
    root.append(el('txt title', { left: '36px', top: '69px', fontSize: '23px', lineHeight: '28px' }, f.title));
  }
  if (f.video) f.video.forEach(([playName, thumbName, ytid, yturl]) => videoHot(root, list, playName, thumbName, ytid, yturl));
  if (name === 'norway') {
    const note = el('norway-note', { left: '175px', top: '672px' },
      '*Background music removed from some videos due to<br>YouTube copyright rules.');
    root.append(note);
  }

  homeHot(root);
  addDots(root, () => { root.remove(); popups = popups.filter(p => p !== root); });
  const dragbar = el('dragbar'); root.append(dragbar);
  Object.assign(root.style, { left: 60 + popups.length * 24 + 'px', top: 30 + popups.length * 18 + 'px', transform: 'scale(.9)', transformOrigin: '0 0', zIndex: ++zTop });
  root.style.setProperty('--ps', .9);
  makeDraggable(root, dragbar);
  root.addEventListener('mousedown', () => root.style.zIndex = ++zTop);
  stage.append(root); fitStickies(root); popups.push(root);
}

function videoHot(root, list, playName, thumbName, ytid, yturl) {
  const t = list.find(x => x.name === thumbName), p = list.find(x => x.name === playName); if (!t || !p) return;
  const h = el('play-hot', { left: t.x + 'px', top: t.y + 'px', width: t.w + 'px', height: t.h + 'px' });
  h.title = 'Play video';
  h.onclick = () => {
    const existing = root.querySelector(`[data-vid="${thumbName}"]`);
    if (existing) { existing.remove(); return; }
    root.querySelectorAll('.vid-player-wrap').forEach(w => w.remove());

    const wrap = el('vid-player-wrap', { left: t.x + 'px', top: t.y + 'px', width: t.w + 'px', height: t.h + 'px' });
    wrap.dataset.vid = thumbName;

    const closeBtn = el('vid-close-btn', {}, '&times;');
    closeBtn.title = 'Close video';
    closeBtn.onclick = e => { e.stopPropagation(); wrap.remove(); };

    if (location.protocol === 'file:') {
      // When opened locally via file://, YouTube embeds fail with Error 153 because file:// has no HTTP Referer.
      // Show clean in-frame card with one-click direct YouTube playback.
      const prompt = el('vid-file-prompt', {},
        `<div class="vfp-icon"><svg viewBox="0 0 24 24" width="42" height="42" fill="#ff0000"><path d="M23.498 6.186a3.016 3.016 0 0 0-2.122-2.136C19.505 3.545 12 3.545 12 3.545s-7.505 0-9.377.505A3.017 3.017 0 0 0 .502 6.186C0 8.07 0 12 0 12s0 3.93.502 5.814a3.016 3.016 0 0 0 2.122 2.136c1.871.505 9.376.505 9.376.505s7.505 0 9.377-.505a3.015 3.015 0 0 0 2.122-2.136C24 15.93 24 12 24 12s0-3.93-.502-5.814zM9.545 15.568V8.432L15.818 12l-6.273 3.568z"/></svg></div>
         <p class="vfp-title">Watch on YouTube</p>
         <p class="vfp-sub">YouTube blocks embedded playback on local file:// links (Error 153).<br><br>On GitHub Pages, it plays right inside this phone!</p>
         <button class="vfp-btn">Play Video &#8599;</button>`);
      prompt.onclick = () => window.open(yturl, '_blank');
      wrap.append(closeBtn, prompt);
    } else {
      // On GitHub Pages / HTTP server: native YouTube player with full referrerpolicy
      const iframe = document.createElement('iframe');
      iframe.src = `https://www.youtube.com/embed/${ytid}?autoplay=1&rel=0`;
      iframe.title = 'YouTube video';
      iframe.allow = 'accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share';
      iframe.referrerPolicy = 'strict-origin-when-cross-origin';
      iframe.allowFullscreen = true;

      const ytLink = el('vid-ext-link', {}, '<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6M15 3h6v6M10 14L21 3"/></svg> Watch on YouTube');
      ytLink.onclick = e => { e.stopPropagation(); window.open(yturl, '_blank'); };

      wrap.append(iframe, closeBtn, ytLink);
    }
    root.append(wrap);
  };
  root.append(h);
}

function makeDraggable(root, handle) {
  let sx, sy, ox, oy, on = false;
  handle.addEventListener('mousedown', e => {
    on = true; sx = e.clientX; sy = e.clientY; ox = parseFloat(root.style.left); oy = parseFloat(root.style.top);
    root.classList.add('dragging'); e.preventDefault();
  });
  addEventListener('mousemove', e => {
    if (!on) return; const s = +stage.dataset.scale;
    root.style.left = ox + (e.clientX - sx) / s + 'px'; root.style.top = oy + (e.clientY - sy) / s + 'px';
  });
  addEventListener('mouseup', () => { on = false; root.classList.remove('dragging'); });
}

/* ── slide one (main page) ── */
function initHome() {
  const { root, list, stickies } = build('s1');
  root.dataset.name = 'home'; root.style.zIndex = 1;
  // Fix PowerPoint typo "Personal_Pojects" -> "Personal_Projects"
  root.append(el('patch', { left: '880px', top: '435px', width: '160px', height: '22px', background: '#f2f2f2' }));
  root.append(el('', { left: '880px', top: '435px', width: '160px', height: '22px', display: 'flex', alignItems: 'center', justifyContent: 'center', font: "15px/1 'Aptos', 'Segoe UI', Arial, sans-serif", color: '#000', pointerEvents: 'none' }, 'Personal_Projects'));
  addNotes(root, stickies, [null, null]); // slide-one notes are real text boxes in the design
  const map = { 'Picture 12': 'scientific', 'Picture 11': 'presentations', 'Picture 14': 'gallery', 'Picture 16': 'resume', 'Picture 18': 'freelance', 'Picture 20': 'personal' };
  Object.entries(map).forEach(([n, target]) => {
    const e = list.find(x => x.name === n);
    const h = el('hot', { left: e.x - 30 + 'px', top: e.y - 6 + 'px', width: e.w + 60 + 'px', height: e.h + 46 + 'px' });
    h.title = target.replace(/_/g, ' '); h.onclick = () => openFolder(target); root.append(h);
  });
  homeHot(root); // the red/yellow/green circles on slide one intentionally do nothing
  stage.append(root); fitStickies(root);
}
initHome();

/* deep links for testing: #folder=scientific&file=masters */
const hp = new URLSearchParams(location.hash.slice(1));
if (hp.get('folder')) openFolder(hp.get('folder'));
if (hp.get('file')) openFile(hp.get('file'));

