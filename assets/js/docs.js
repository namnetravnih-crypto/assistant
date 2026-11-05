/* Ассистент — автозаполнение: интерфейс */
(function () {
  'use strict';
  const E = window.AssistantEngine, $ = id => document.getElementById(id);

  const GROUPS = [
    ['Организация', [['name_full', 'Полное наименование', 'wide'], ['name_short', 'Сокращённое наименование', 'wide'], ['inn', 'ИНН'], ['kpp', 'КПП'], ['ogrn', 'ОГРН / ОГРНИП'], ['okpo', 'ОКПО'], ['reg_date', 'Дата регистрации'], ['tax_reg_date', 'Дата постановки на налоговый учёт']]],
    ['Адреса и связь', [['addr_legal', 'Юридический адрес', 'wide'], ['addr_post', 'Почтовый адрес', 'wide'], ['addr_fact', 'Фактический адрес', 'wide'], ['phone', 'Телефон'], ['email', 'E-mail'], ['site', 'Сайт']]],
    ['Банк', [['bank_name', 'Банк', 'wide'], ['bik', 'БИК'], ['rs', 'Расчётный счёт'], ['ks', 'Корр. счёт'], ['bank_inn', 'ИНН банка'], ['bank_addr', 'Адрес банка']]],
    ['Руководитель и подписант', [['director_pos', 'Должность'], ['director_fio', 'ФИО'], ['director_pos_gen', 'Должность, род. падеж («в лице …»)', 0, 'gen'], ['director_fio_gen', 'ФИО, род. падеж («в лице …»)', 0, 'gen'], ['basis', 'Действует на основании', 'wide'], ['director_inn', 'ИНН руководителя'], ['accountant_fio', 'Главный бухгалтер']]],
    ['Контактное лицо и налогообложение', [['contact_fio', 'Контактное лицо (ФИО)'], ['contact_pos', 'Должность контактного лица'], ['contact_phone', 'Телефон контактного лица'], ['contact_email', 'E-mail контактного лица'], ['tax_system', 'Система налогообложения', 'sel'], ['branches', 'Филиалы'], ['holding', 'Холдинг / группа компаний', 'wide']]],
  ];
  const TAX = ['', 'ОСН (с НДС)', 'УСН (без НДС)', 'УСН (с НДС)', 'ПСН', 'НПД'];
  const ROLE = { customer: { title: 'Реквизиты Заказчика', badge: 'Заказчик', sub: 'Сторона, которая заказывает работы, покупает, арендует. Нужна для договоров и спецификаций.' },
                 executor: { title: 'Реквизиты Участника / Исполнителя', badge: 'Участник · Исполнитель', sub: 'Ваша сторона: участник закупки, подрядчик, поставщик. По этим данным заполняются заявки, анкеты и формы.' } };
  const DEMO = {
    customer: `Полное наименование: Общество с ограниченной ответственностью «Ромашка»\nИНН: 7700000001\nКПП: 770001001\nОГРН: 1027700000001\nЮридический адрес: 123112, г. Москва, Тестовая ул., д. 1\nБанк: АО «Тестбанк»\nБИК: 044525001\nРасчётный счёт: 40702810900000000001\nКорр. счёт: 30101810400000000001\nТелефон: +7 (495) 000-00-01\nE-mail: info@romashka.test\nГенеральный директор: Петров Пётр Петрович`,
    executor: `Полное наименование: Общество с ограниченной ответственностью «Василёк»\nИНН: 7800000002\nКПП: 780001001\nОГРН: 1027800000002\nОКПО: 12345678\nЮридический адрес: 190000, г. Санкт-Петербург, Образцовая ул., д. 2\nБанк: АО «Образецбанк»\nБИК: 044030002\nРасчётный счёт: 40702810100000000002\nКорр. счёт: 30101810500000000002\nТелефон: +7 (812) 000-00-02\nE-mail: office@vasilek.test\nГенеральный директор: Сидорова Анастасия Юрьевна`,
  };

  const cards = { customer: {}, executor: {} };
  const manualGen = { customer: new Set(), executor: new Set() };
  let forms = [], results = [], seq = 0;
  const ui = {};

  function save() { try { localStorage.setItem('assistant.cards.v2', JSON.stringify(cards)); localStorage.setItem('assistant.docinfo', JSON.stringify(docInfoRaw())); } catch (e) {} }
  function load() { try { const s = JSON.parse(localStorage.getItem('assistant.cards.v2') || 'null'); if (s) { cards.customer = s.customer || {}; cards.executor = s.executor || {}; } const d = JSON.parse(localStorage.getItem('assistant.docinfo') || 'null'); if (d) { $('docNumber').value = d.number || ''; $('docCity').value = d.city || ''; $('docSpec').value = d.specNo || ''; $('docAmount').value = d.amount || ''; $('docVat').value = d.vat || '22'; $('docVatMode').value = d.vatMode || 'incl'; } } catch (e) {} }
  const setStatus = (el, t, cls) => { el.textContent = t; el.className = 'status' + (cls ? ' ' + cls : ''); };

  /* ---------- окна реквизитов ---------- */
  function buildParty(role) {
    const root = $(role === 'customer' ? 'partyCustomer' : 'partyExecutor'), meta = ROLE[role], u = ui[role] = { inputs: {} };
    root.innerHTML = `<h2><span class="n">1</span> ${meta.title}<span class="badge-role">${meta.badge}</span></h2><div class="sub">${meta.sub}</div>
      <div class="drop" tabindex="0"><b>Перетащите карту сюда</b> или нажмите, чтобы выбрать<small>.docx, .xlsx, .pdf (с текстом), .txt</small><input type="file" accept=".docx,.xlsx,.pdf,.txt" hidden></div>
      <details class="paste"><summary>Или вставить текст реквизитов</summary><textarea rows="6" placeholder="Наименование: ООО «Ромашка»&#10;ИНН: 7701234567&#10;КПП: 770101001&#10;Генеральный директор: Иванов Иван Иванович"></textarea><div class="row"><button class="btn small" type="button" data-act="parse">Распознать</button></div></details>
      <div class="row"><button class="btn small ghost" type="button" data-act="demo">Подставить вымышленный пример</button><button class="btn small ghost" type="button" data-act="clear">Очистить</button></div>
      <div class="status"></div><div class="grid"></div>`;
    const drop = root.querySelector('.drop'), file = root.querySelector('input[type=file]'), ta = root.querySelector('textarea'); u.status = root.querySelector('.status'); u.grid = root.querySelector('.grid');
    drop.addEventListener('click', () => file.click());
    drop.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); file.click(); } });
    ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
    drop.addEventListener('drop', e => handleCardFile(role, e.dataTransfer.files[0]));
    file.addEventListener('change', () => { handleCardFile(role, file.files[0]); file.value = ''; });
    root.querySelector('[data-act=parse]').addEventListener('click', () => applyParsed(role, ta.value, 'Вставленный текст'));
    root.querySelector('[data-act=demo]').addEventListener('click', () => applyParsed(role, DEMO[role], 'Вымышленный пример'));
    root.querySelector('[data-act=clear]').addEventListener('click', () => { cards[role] = {}; manualGen[role].clear(); save(); renderGrid(role); setStatus(u.status, ''); updateRun(); });
    renderGrid(role);
  }

  function renderGrid(role) {
    const u = ui[role], card = cards[role]; u.grid.innerHTML = ''; u.inputs = {};
    for (const [gname, fields] of GROUPS) {
      const g = document.createElement('div'); g.className = 'grp'; g.textContent = gname; u.grid.append(g);
      for (const [key, label, cls, kind] of fields) {
        const l = document.createElement('label'); if (cls === 'wide') l.className = 'wide'; l.append(label);
        let inp;
        if (cls === 'sel') { inp = document.createElement('select'); TAX.forEach(t => { const o = document.createElement('option'); o.value = t; o.textContent = t || '— не указано —'; inp.append(o); }); inp.value = card[key] || ''; }
        else { inp = document.createElement('input'); inp.type = 'text'; inp.value = card[key] || ''; if (kind === 'gen') inp.classList.add('auto'); }
        inp.addEventListener('input', () => onEdit(role, key, inp));
        u.inputs[key] = inp; l.append(inp); u.grid.append(l);
      }
    }
    updateRun();
  }
  function onEdit(role, key, inp) {
    const c = cards[role]; c[key] = inp.value.trim();
    if (key === 'director_pos_gen' || key === 'director_fio_gen') manualGen[role].add(key);
    if (key === 'director_fio' && !manualGen[role].has('director_fio_gen')) { c.director_fio_gen = E.fioGen(c.director_fio); ui[role].inputs.director_fio_gen.value = c.director_fio_gen; }
    if (key === 'director_pos' && !manualGen[role].has('director_pos_gen')) { c.director_pos_gen = E.posGen(c.director_pos); ui[role].inputs.director_pos_gen.value = c.director_pos_gen; }
    save(); updateRun();
  }

  async function pdfToText(buf) {
    const pdfjs = window.pdfjsLib; if (!pdfjs) throw new Error('Не загрузилась библиотека PDF');
    pdfjs.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    const pdf = await pdfjs.getDocument({ data: buf }).promise, lines = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const content = await (await pdf.getPage(i)).getTextContent(), rows = [];
      for (const it of content.items) { const y = Math.round(it.transform[5] / 3), x = it.transform[4]; let row = rows.find(r => r.y === y); if (!row) rows.push(row = { y, parts: [] }); row.parts.push({ x, w: it.width, s: it.str }); }
      rows.sort((a, b) => b.y - a.y).forEach(r => {
        r.parts.sort((a, b) => a.x - b.x); let out = '', prevEnd = null;
        for (const p of r.parts) { if (!p.s.trim()) continue; out += (prevEnd !== null ? (p.x - prevEnd > 12 ? ' | ' : ' ') : '') + p.s; prevEnd = p.x + p.w; }
        if (out.trim()) lines.push(out.trim());
      });
    }
    return lines.join('\n');
  }
  async function readCardFile(file) {
    const ext = file.name.split('.').pop().toLowerCase(), buf = await file.arrayBuffer();
    if (ext === 'docx') return E.docxToText(buf); if (ext === 'xlsx') return E.xlsxToText(buf); if (ext === 'pdf') return pdfToText(buf);
    if (ext === 'txt') return new TextDecoder('utf-8').decode(buf);
    throw new Error('Формат .' + ext + ' не поддерживается. Сохраните карту как .docx, .xlsx, .pdf или .txt');
  }
  function applyParsed(role, text, source) {
    const parsed = E.parseCardText(text), n = Object.keys(parsed).length, st = ui[role].status;
    if (!n) { setStatus(st, 'Не удалось найти реквизиты. Проверьте файл или вставьте текст вручную.', 'err'); return; }
    const keep = {}; for (const k of ['tax_system', 'contact_fio', 'contact_pos', 'contact_phone', 'contact_email']) if (cards[role][k]) keep[k] = cards[role][k];
    cards[role] = Object.assign(parsed, keep); manualGen[role].clear(); save(); renderGrid(role);
    const miss = ['name_full', 'inn', 'addr_legal', 'director_fio'].filter(k => !cards[role][k]).map(k => E.FIELD_TITLES[k]);
    setStatus(st, source + ': распознано полей — ' + n + '.' + (miss.length ? ' Не найдено: ' + miss.join(', ') + '.' : '') + ' Проверьте данные ниже.', miss.length ? '' : 'ok');
  }
  async function handleCardFile(role, file) {
    if (!file) return; setStatus(ui[role].status, 'Читаю ' + file.name + '…');
    try { applyParsed(role, await readCardFile(file), file.name); } catch (e) { setStatus(ui[role].status, 'Ошибка: ' + e.message, 'err'); }
  }

  /* ---------- бланки ---------- */
  function renderForms() {
    const ul = $('formList'); ul.innerHTML = '';
    for (const f of forms) {
      const li = document.createElement('li'), nm = document.createElement('span'); nm.className = 'name'; nm.textContent = f.file.name; nm.title = f.file.name;
      const sel = document.createElement('select');
      [['default', 'Тип: как в настройках'], ['form', 'Форма участника'], ['contract', 'Договор / спецификация']].forEach(([v, t]) => { const o = document.createElement('option'); o.value = v; o.textContent = t; if (f.type === v) o.selected = true; sel.append(o); });
      sel.addEventListener('change', () => { f.type = sel.value; });
      const sel2 = document.createElement('select');
      [['default', 'Объём: как в настройках'], ['full', 'Заполнить всё'], ['min', 'Только наименование и подписант']].forEach(([v, t]) => { const o = document.createElement('option'); o.value = v; o.textContent = t; if (f.mode === v) o.selected = true; sel2.append(o); });
      sel2.addEventListener('change', () => { f.mode = sel2.value; });
      const x = document.createElement('button'); x.className = 'x'; x.type = 'button'; x.title = 'Убрать'; x.textContent = '×';
      x.addEventListener('click', () => { forms = forms.filter(q => q.id !== f.id); renderForms(); });
      li.append(nm, sel, sel2, x); ul.append(li);
    }
    updateRun();
  }
  function addForms(files) { for (const file of files) forms.push({ id: ++seq, file, mode: 'default', type: 'default' }); renderForms(); }
  function updateRun() {
    const hasCard = Object.keys(cards.executor).length || Object.keys(cards.customer).length;
    $('btnRun').disabled = !(forms.length && hasCard);
    $('runHint').textContent = !forms.length ? 'Добавьте бланки.' : !hasCard ? 'Заполните реквизиты хотя бы одной стороны.' : 'Готово к заполнению.';
  }

  /* ---------- запуск ---------- */
  function docInfoRaw() { return { number: $('docNumber').value.trim(), city: $('docCity').value.trim(), specNo: $('docSpec').value.trim(), amount: $('docAmount').value.trim(), vat: $('docVat').value, vatMode: $('docVatMode').value }; }
  const MIME = { docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', dotx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' };
  function uniqueName(name, used) { const dot = name.lastIndexOf('.'), base = dot > 0 ? name.slice(0, dot) : name, ext = dot > 0 ? name.slice(dot) : ''; let cand = base + ' (заполнено)' + ext, i = 2; while (used.has(cand)) cand = base + ' (заполнено ' + (i++) + ')' + ext; used.add(cand); return cand; }

  async function run() {
    const btn = $('btnRun'); btn.disabled = true; btn.textContent = 'Заполняю…'; results = [];
    const used = new Set(), info = Object.assign(docInfoRaw(), { date: $('optDate').value });
    const g = { highlight: $('optHighlight').checked, fallbackAddr: $('optFallback').checked, fallbackContact: $('optContact').checked, mode: $('optMode').value, type: $('optType').value };
    for (const f of forms) {
      const name = uniqueName(f.file.name, used);
      try {
        const buf = await f.file.arrayBuffer(), type = f.type === 'default' ? g.type : f.type;
        const out = await E.fillFile(f.file.name, buf, cards, info, { highlight: g.highlight, fallbackAddr: g.fallbackAddr, fallbackContact: g.fallbackContact, mode: f.mode === 'default' ? g.mode : f.mode, docType: type === 'form' ? 'form' : type });
        const ext = f.file.name.split('.').pop().toLowerCase();
        results.push({ name, blob: new Blob([out.data], { type: MIME[ext] || 'application/octet-stream' }), stats: out.stats, items: out.items, contract: out.contract });
      } catch (e) { results.push({ name: f.file.name, blob: f.file, error: e.message || String(e) }); }
    }
    renderResults(); btn.disabled = false; btn.textContent = 'Заполнить документы'; updateRun();
  }
  function download(blob, name) { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000); }
  function renderResults() {
    const ul = $('resultList'); ul.innerHTML = '';
    for (const r of results) {
      const li = document.createElement('li'); li.className = 'res';
      const line = document.createElement('div'); line.className = 'line';
      const nm = document.createElement('span'); nm.className = 'name'; nm.textContent = r.name; nm.style.flex = '1 1 220px'; line.append(nm);
      const pill = (t, c) => { const s = document.createElement('span'); s.className = 'pill ' + c; s.textContent = t; line.append(s); };
      if (r.error) pill('не обработан: ' + r.error, 'err');
      else { pill('заполнено: ' + r.stats.filled, 'ok'); if (r.stats.highlighted) pill('жёлтым выделено: ' + r.stats.highlighted, 'hl'); if (r.contract) pill('договор: две стороны', 'ok'); }
      const dl = document.createElement('button'); dl.className = 'btn small'; dl.type = 'button'; dl.textContent = r.error ? 'Скачать без изменений' : 'Скачать'; dl.addEventListener('click', () => download(r.blob, r.name)); line.append(dl); li.append(line);
      if (r.items && r.items.length) {
        const d = document.createElement('details'), s = document.createElement('summary'); s.textContent = 'Что именно сделано'; d.append(s);
        const u = document.createElement('ul');
        for (const it of r.items) {
          const l = document.createElement('li');
          if (it.role) { const t = document.createElement('span'); t.className = 'tag ' + (it.role === 'customer' ? 'c' : 'e'); t.textContent = it.role === 'customer' ? 'Заказчик' : 'Участник'; l.append(t); }
          if (it.kind === 'fill') l.append(it.key + ': ' + String(it.val).slice(0, 90));
          else if (it.kind === 'hl') { const m = document.createElement('mark'); m.textContent = it.key; l.append('Заполните вручную — ', m); }
          else l.append('Ошибка: ' + it.key);
          u.append(l);
        }
        d.append(u); li.append(d);
      }
      ul.append(li);
    }
    $('resultCard').hidden = false; $('resultCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
  async function downloadZip() { const zip = new JSZip(); for (const r of results) zip.file(r.name, r.blob); download(await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' }), 'Заполненные_документы.zip'); }

  /* ---------- старт ---------- */
  const dz = $('dropForms'), fin = $('fileForms');
  dz.addEventListener('click', () => fin.click());
  dz.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fin.click(); } });
  ['dragenter', 'dragover'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove('over'); }));
  dz.addEventListener('drop', e => addForms(Array.from(e.dataTransfer.files)));
  fin.addEventListener('change', () => { addForms(Array.from(fin.files)); fin.value = ''; });
  $('btnRun').addEventListener('click', run); $('btnZip').addEventListener('click', downloadZip);
  ['docNumber', 'docCity', 'docSpec', 'docAmount', 'docVat', 'docVatMode'].forEach(id => $(id).addEventListener('input', save));
  load(); buildParty('customer'); buildParty('executor'); updateRun();
  window.AssistantDocs = { cards, run, addForms, applyParsed };
})();
