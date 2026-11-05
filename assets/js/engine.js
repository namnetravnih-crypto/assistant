/* =====================================================================
   Ассистент — движок автозаполнения (DOCX / XLSX), версия 2
   Две стороны: Заказчик и Участник/Исполнитель. Всё в браузере.
   Структуру, шрифты и таблицы документа не меняет: правится только текст.
   ===================================================================== */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('jszip'), require('@xmldom/xmldom'));
  else root.AssistantEngine = factory(root.JSZip, { DOMParser: root.DOMParser, XMLSerializer: root.XMLSerializer });
})(typeof self !== 'undefined' ? self : this, function (JSZip, X) {
  'use strict';
  const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const NS_S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NS_XML = 'http://www.w3.org/XML/1998/namespace';

  /* ---------- XML ---------- */
  const parseXml = s => new X.DOMParser().parseFromString(s, 'application/xml');
  const serialize = d => { let s = new X.XMLSerializer().serializeToString(d); if (!/^<\?xml/.test(s)) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s; return s; };
  const lname = n => n.localName || String(n.nodeName).replace(/^.*:/, '');
  const kids = (n, name) => { const o = []; for (let c = n.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && lname(c) === name) o.push(c); return o; };
  const firstKid = (n, name) => kids(n, name)[0] || null;
  const byNS = (n, ns, name) => Array.from(n.getElementsByTagNameNS(ns, name));

  /* ---------- Поля сторон ---------- */
  const FIELD_TITLES = {
    name_full: 'Полное наименование', name_short: 'Сокращённое наименование', inn: 'ИНН', kpp: 'КПП', ogrn: 'ОГРН', okpo: 'ОКПО', okved: 'ОКВЭД', oktmo: 'ОКТМО',
    reg_date: 'Дата регистрации', tax_reg_date: 'Дата постановки на налоговый учёт', addr_legal: 'Юридический адрес', addr_post: 'Почтовый адрес', addr_fact: 'Фактический адрес',
    phone: 'Телефон', fax: 'Факс', email: 'E-mail', site: 'Сайт', bank_name: 'Банк', bank_addr: 'Адрес банка', bank_inn: 'ИНН банка', bank_phone: 'Телефон банка',
    bik: 'БИК', rs: 'Расчётный счёт', ks: 'Корр. счёт', director_pos: 'Должность руководителя', director_fio: 'ФИО руководителя', director_inn: 'ИНН руководителя',
    director_pos_gen: 'Должность (род. падеж)', director_fio_gen: 'ФИО (род. падеж)', basis: 'Действует на основании', tax_system: 'Система налогообложения',
    accountant_fio: 'Главный бухгалтер', contact_fio: 'Контактное лицо', contact_pos: 'Должность контактного лица', contact_phone: 'Телефон контактного лица', contact_email: 'E-mail контактного лица',
  };
  const T = k => FIELD_TITLES[k] || k;

  /* ---------- Разбор слов ---------- */
  const lc = s => String(s || '').toLowerCase().replace(/ё/g, 'е');
  const clean = s => String(s || '').replace(/[   ]/g, ' ').replace(/\s+/g, ' ').trim();
  const stripParens = s => { let o = s, p; do { p = o; o = o.replace(/\([^()]*\)/g, ' '); } while (o !== p); return o.replace(/\[[^\]]*\]/g, ' '); };

  const CUSTOMER_RE = /заказчик[а-яёa-z0-9_]*|покупател[а-яёa-z0-9_]*|арендатор[а-яёa-z0-9_]*|генподрядчик[а-яёa-z0-9_]*|наймодател[а-яёa-z0-9_]*(?![а-яёa-z0-9_]*не)/;
  const EXEC_RE = /исполнител[а-яёa-z0-9_]*|подрядчик[а-яёa-z0-9_]*|поставщик[а-яёa-z0-9_]*|продавец[а-яёa-z0-9_]*|продавц[а-яёa-z0-9_]*|арендодател[а-яёa-z0-9_]*|субподрядчик[а-яёa-z0-9_]*|участник[а-яёa-z0-9_]*|претендент[а-яёa-z0-9_]*|заявител[а-яёa-z0-9_]*/;
  function roleOfWord(w) {
    const t = lc(w);
    if (/^(?:суб)?подрядчик|^исполнител|^поставщик|^продав|^арендодател|^участник|^претендент|^заявител/.test(t)) return 'executor';
    if (/^заказчик|^покупател|^арендатор|^генподрядчик/.test(t)) return 'customer';
    return null;
  }
  function roleInText(text) {           // последняя упомянутая роль в тексте
    const t = lc(text); let best = null, pos = -1;
    for (const m of t.matchAll(new RegExp('(?<![а-я])(?:' + CUSTOMER_RE.source + '|' + EXEC_RE.source + ')', 'g'))) if (m.index >= pos) { pos = m.index; best = roleOfWord(m[0]); }
    return best;
  }
  function headingRole(text) {          // «ЗАКАЗЧИК:», «От Исполнителя», «Сторона: Подрядчик»
    const t = lc(clean(text)).replace(/[«»"':;.,\-–—_()]/g, ' ').replace(/\s+/g, ' ').trim();
    const m = /^(?:(?:от|за)\s+)?(?:сторона\s+)?(заказчик[а-яёa-z0-9_]*|покупател[а-яёa-z0-9_]*|арендатор[а-яёa-z0-9_]*|генподрядчик[а-яёa-z0-9_]*|исполнител[а-яёa-z0-9_]*|подрядчик[а-яёa-z0-9_]*|поставщик[а-яёa-z0-9_]*|продав[а-яёa-z0-9_]*|арендодател[а-яёa-z0-9_]*|субподрядчик[а-яёa-z0-9_]*)(?:\s+\d)?$/.exec(t);
    return m ? roleOfWord(m[1]) : null;
  }

  /* ---------- Деньги и числа ---------- */
  function parseNum(s) { const t = String(s == null ? '' : s).replace(/[\s ]/g, '').replace(',', '.'); if (!/^-?\d+(\.\d+)?$/.test(t)) return null; return parseFloat(t); }
  function fmtMoney(n) { const [a, b] = (Math.round(n * 100) / 100).toFixed(2).split('.'); return a.replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ',' + b; }
  const UN = [['', 'один', 'два', 'три', 'четыре', 'пять', 'шесть', 'семь', 'восемь', 'девять'], ['', 'одна', 'две', 'три', 'четыре', 'пять', 'шесть', 'семь', 'восемь', 'девять']];
  const TEN = ['десять', 'одиннадцать', 'двенадцать', 'тринадцать', 'четырнадцать', 'пятнадцать', 'шестнадцать', 'семнадцать', 'восемнадцать', 'девятнадцать'];
  const TENS = ['', '', 'двадцать', 'тридцать', 'сорок', 'пятьдесят', 'шестьдесят', 'семьдесят', 'восемьдесят', 'девяносто'];
  const HUN = ['', 'сто', 'двести', 'триста', 'четыреста', 'пятьсот', 'шестьсот', 'семьсот', 'восемьсот', 'девятьсот'];
  const plural = (n, f) => { const m = n % 100, d = n % 10; return (m > 10 && m < 20) ? f[2] : d === 1 ? f[0] : (d >= 2 && d <= 4) ? f[1] : f[2]; };
  function triad(n, fem) { const o = []; o.push(HUN[(n / 100) | 0]); const r = n % 100; if (r >= 10 && r < 20) o.push(TEN[r - 10]); else { o.push(TENS[(r / 10) | 0]); o.push(UN[fem ? 1 : 0][r % 10]); } return o.filter(Boolean).join(' '); }
  function numWords(n) {
    if (n === 0) return 'ноль';
    const parts = [[1e9, ['миллиард', 'миллиарда', 'миллиардов'], 0], [1e6, ['миллион', 'миллиона', 'миллионов'], 0], [1e3, ['тысяча', 'тысячи', 'тысяч'], 1]], out = [];
    for (const [d, f, fem] of parts) { const q = Math.floor(n / d) % 1000; if (q) out.push(triad(q, fem) + ' ' + plural(q, f)); }
    const r = n % 1000; if (r) out.push(triad(r, 0));
    return out.join(' ');
  }
  function rubWords(amount) {
    const rub = Math.floor(amount + 1e-9), kop = Math.round((amount - rub) * 100), w = numWords(rub);
    return w.charAt(0).toUpperCase() + w.slice(1) + ' ' + plural(rub, ['рубль', 'рубля', 'рублей']) + ' ' + String(kop).padStart(2, '0') + ' ' + plural(kop, ['копейка', 'копейки', 'копеек']);
  }

  /* ---------- Склонение ---------- */
  const POS_GEN = { 'генеральный директор': 'Генерального директора', 'исполнительный директор': 'Исполнительного директора', 'коммерческий директор': 'Коммерческого директора', 'директор': 'Директора', 'президент': 'Президента', 'руководитель': 'Руководителя', 'председатель': 'Председателя', 'управляющий': 'Управляющего', 'индивидуальный предприниматель': 'Индивидуального предпринимателя', 'и.о. директора': 'и.о. директора', 'главный врач': 'Главного врача', 'ректор': 'Ректора', 'начальник': 'Начальника' };
  function posGen(pos) {
    const t = clean(pos); if (!t) return '';
    const k = lc(t); if (POS_GEN[k]) return POS_GEN[k];
    return t.split(' ').map((w, i) => { const l = w.toLowerCase(); let r = l; if (/(ый|ой)$/.test(l)) r = l.replace(/(ый|ой)$/, 'ого'); else if (/ий$/.test(l)) r = l.replace(/ий$/, 'ого'); else if (/ь$/.test(l)) r = l.replace(/ь$/, 'я'); else if (/[бвгджзклмнпрстфхцчшщ]$/.test(l)) r = l + 'а'; return i === 0 ? r.charAt(0).toUpperCase() + r.slice(1) : r; }).join(' ');
  }
  function declWord(w, kind, fem) {  // родительный падеж одного слова ФИО
    const l = w.toLowerCase(); let r = null;
    const keepCase = x => w[0] === w[0].toUpperCase() ? x.charAt(0).toUpperCase() + x.slice(1) : x;
    if (kind === 'patr') { if (/ич$/.test(l)) r = l + 'а'; else if (/на$/.test(l)) r = l.replace(/а$/, 'ы'); else return w; return keepCase(r); }
    if (kind === 'sur') {
      if (fem) { if (/(ова|ева|ёва|ина|ына)$/.test(l)) r = l.replace(/а$/, 'ой'); else if (/ая$/.test(l)) r = l.replace(/ая$/, 'ой'); else return w; return keepCase(r); }
      if (/(ов|ев|ёв|ин|ын)$/.test(l)) r = l + 'а'; else if (/(ский|цкий|ской|ный|ной|вой|тый)$/.test(l)) r = l.replace(/(ий|ой|ый)$/, 'ого'); else if (/ий$/.test(l)) r = l.replace(/ий$/, 'ия'); else if (/ь$/.test(l)) r = l.replace(/ь$/, 'я'); else if (/й$/.test(l)) r = l.replace(/й$/, 'я'); else if (/(ко|о|е|у|ю|и|ых|их|а|я)$/.test(l) && !/(ка|ша)$/.test(l)) return w; else if (/(ка|га|ха)$/.test(l)) r = l.replace(/а$/, 'и'); else if (/[бвгджзклмнпрстфхцчшщ]$/.test(l)) r = l + 'а'; else return w;
      return keepCase(r);
    }
    if (fem) { if (/ия$/.test(l)) r = l.replace(/я$/, 'и'); else if (/[гкх]а$/.test(l) || /ша$/.test(l) || /жа$/.test(l) || /ча$/.test(l) || /ща$/.test(l)) r = l.replace(/а$/, 'и'); else if (/а$/.test(l)) r = l.replace(/а$/, 'ы'); else if (/я$/.test(l)) r = l.replace(/я$/, 'и'); else if (/ь$/.test(l)) r = l.replace(/ь$/, 'и'); else return w; return keepCase(r); }
    if (/ей$/.test(l) && l.length > 4) r = l.replace(/ей$/, 'ея'); else if (/ий$/.test(l)) r = l.replace(/ий$/, 'ия'); else if (/й$/.test(l)) r = l.replace(/й$/, 'я'); else if (/ь$/.test(l)) r = l.replace(/ь$/, 'я'); else if (/(а|я)$/.test(l)) r = /[гкх]а$/.test(l) || /[шжчщ]а$/.test(l) ? l.replace(/а$/, 'и') : l.replace(/а$/, 'ы').replace(/я$/, 'и'); else if (/(ев|ёв|лев)$/.test(l) && false) r = l; else if (/[бвгджзклмнпрстфхцчшщ]$/.test(l)) r = (l === 'лев' ? 'льва' : l === 'павел' ? 'павла' : l + 'а'); else return w;
    return keepCase(r);
  }
  function fioGen(fio) {
    const p = clean(fio).split(' '); if (p.length < 2) return clean(fio);
    const patr = p[2] || '', fem = /(на|зы)$/i.test(patr) || (!patr && /(а|я)$/i.test(p[1]) && /(ва|на|ая)$/i.test(p[0]));
    return [declWord(p[0], 'sur', fem), declWord(p[1], 'name', fem), patr ? declWord(patr, 'patr', fem) : ''].filter(Boolean).concat(p.slice(3)).join(' ');
  }
  function shortFio(fio) {
    const f = clean(fio); if (!f) return ''; if (/\./.test(f)) return f;
    const p = f.split(' '); if (p.length < 2) return f;
    return p[0] + ' ' + p.slice(1, 3).map(x => x[0].toUpperCase() + '.').join('');
  }
  const MONTHS = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
  const parseIso = iso => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || ''); return m ? { y: +m[1], mo: +m[2], d: +m[3] } : null; };
  const dateLong = iso => { const d = parseIso(iso); return d ? `«${String(d.d).padStart(2, '0')}» ${MONTHS[d.mo - 1]} ${d.y} г.` : ''; };
  const dateShort = iso => { const d = parseIso(iso); return d ? `${String(d.d).padStart(2, '0')}.${String(d.mo).padStart(2, '0')}.${d.y}` : ''; };
  const isIp = c => /^(ип|индивидуальный предприниматель)\s/i.test(c.name_full || c.name_short || '');

  /* =====================================================================
     Распознавание подписи поля
     ===================================================================== */
  const NLB = '(?<![а-яa-z0-9])', NRB = '(?![а-яa-z])';
  const ORGW = '(?:организации|юридического\\s+лица|участника(?:\\s+закупки)?|контрагента|претендента|заявителя|поставщика|подрядчика|исполнителя|заказчика|покупателя|компании|предприятия|партнера|субъекта|стороны|арендатора|арендодателя)';
  const RULES = [
    ['inn_kpp', new RegExp(NLB + 'инн\\s*[/,\\\\и]\\s*кпп' + NRB, 'g')],
    ['name_short', /сокращ[а-яёa-z0-9_]*\s+(?:фирменн[а-яёa-z0-9_]*\s+)?(?:наименован[а-яёa-z0-9_]*|назван[а-яёa-z0-9_]*)/g],
    ['name_full', new RegExp('(?:полн[а-яёa-z0-9_]*\\s+)?(?:фирменн[а-яёa-z0-9_]*\\s+)?наименован[а-яёa-z0-9_]*\\s+' + ORGW + '|полн[а-яёa-z0-9_]*\\s+(?:фирменн[а-яёa-z0-9_]*\\s+)?наименован[а-яёa-z0-9_]*|фирменн[а-яёa-z0-9_]*\\s+наименован[а-яёa-z0-9_]*|название\\s+(?:организации|компании)|участник\\s+(?:процедуры\\s+)?закупки|участник\\s+конкурса|лидер\\s+коллективного\\s+участника', 'g')],
    ['inn', new RegExp(NLB + 'инн' + NRB, 'g')],
    ['kpp', new RegExp(NLB + 'кпп' + NRB, 'g')],
    ['ogrn', new RegExp(NLB + 'огрн(?:ип)?' + NRB + '|основной\\s+государственный\\s+регистрационный\\s+номер', 'g')],
    ['okpo', new RegExp(NLB + 'окпо' + NRB, 'g')],
    ['oktmo', new RegExp(NLB + 'октмо' + NRB, 'g')],
    ['okved', new RegExp(NLB + 'оквэд[а-яёa-z0-9_]*', 'g')],
    ['tax_reg_date', /дата\s+постановки\s+на\s+(?:налоговый\s+)?учет/g],
    ['reg_date', /дата\s+(?:государственной\s+)?(?:регистрации|создания|внесения|присвоения)/g],
    ['addr_post', /почтов[а-яёa-z0-9_]*\s+адрес|(?<!электронн[а-яёa-z0-9_]*\s)адрес\s+для\s+(?:корреспонденции|направления|почтовой)/g],
    ['addr_fact', /фактическ[а-яёa-z0-9_]*\s+(?:адрес|местонахождение|место)|адрес\s+фактическ[а-яёa-z0-9_]*|место\s+осуществления\s+деятельности|складск[а-яёa-z0-9_]*\s+помещени/g],
    ['addr_legal', /юридическ[а-яёa-z0-9_]*\s+адрес|адрес\s*(?:места\s+нахождения|местонахождения|регистрации|юридическ[а-яёa-z0-9_]*)|место\s*нахождени[а-яёa-z0-9_]*|местонахождени[а-яёa-z0-9_]*|место\s+жительств[а-яёa-z0-9_]*|место\s+регистрац[а-яёa-z0-9_]*/g],
    ['email', new RegExp('e-?mail|' + NLB + 'эл[а-яёa-z0-9_]*\\.?\\s*почт[а-яёa-z0-9_]*|электронн[а-яёa-z0-9_]+\\s+(?:почт[а-яёa-z0-9_]*|адрес)|адрес[а-яёa-z0-9_]*\\s+электронной\\s+почты', 'g')],
    ['site', /сайт[а-яёa-z0-9_]*|веб-?[а-яёa-z0-9_]*|интернет-?[а-яёa-z0-9_]*\s*сайт/g],
    ['addr_any', new RegExp(NLB + 'адрес(?:а|у|ом)?' + NRB, 'g')],
    ['phone', new RegExp('телефон\\s*[/,]\\s*факс|телефон[а-яёa-z0-9_]*|' + NLB + 'тел\\.?' + NRB + '|контактн[а-яёa-z0-9_]*\\s+номер', 'g')],
    ['fax', /факс/g],
    ['bank_name', new RegExp('наименование\\s+банка|обслуживающ[а-яёa-z0-9_]+\\s+банк|' + NLB + 'банк' + NRB, 'g')],
    ['bik', new RegExp(NLB + 'бик' + NRB, 'g')],
    ['ks', new RegExp(NLB + 'к\\s*[/\\\\]\\s*с' + NRB + '|кор[а-яёa-z0-9_]*\\.?\\s*сч[её]т[а-яёa-z0-9_]*|корреспондентск[а-яёa-z0-9_]+', 'g')],
    ['rs', new RegExp(NLB + 'р\\s*[/\\\\]\\s*с' + NRB + '|р\\.\\s*с\\.|расч[её]тн[а-яёa-z0-9_]*\\s+сч[её]т|рас\\.?\\s*сч[её]т|номер\\s+сч[её]та', 'g')],
    ['tax_system', /систем[а-яёa-z0-9_]+\s+налогообложени[а-яёa-z0-9_]+|налогообложени[а-яёa-z0-9_]+/g],
    ['vat_payer', /плательщик[а-яёa-z0-9_]*\s+ндс|является\s+плательщиком/g],
    ['vat_text', new RegExp(NLB + 'ндс' + NRB, 'g')],
    ['signer_pos_fio', /должност[а-яёa-z0-9_]*\s*[,/и]\s*ф\.?\s*и\.?\s*о\.?|ф\.?\s*и\.?\s*о\.?\s*[,/и]\s*должност[а-яёa-z0-9_]*/g],
    ['signer_fio', new RegExp('ф\\.\\s*и\\.\\s*о\\.?|' + NLB + 'фио' + NRB + '|фамилия[, ]+имя[, ]+отчество', 'g')],
    ['signer_pos', new RegExp(NLB + 'должност[а-яёa-z0-9_]*', 'g')],
    ['inperson', /в\s+лице/g],
    ['basis', /действующ[а-яёa-z0-9_]+\s+на\s+основании|на\s+основании\s+(?:устава|доверенности)/g],
    ['doc_date', /дата\s+(?:документа|составления|заявки|подписания|договора)|(?<![а-я])дата(?![а-я])/g],
    ['branches', /филиал[а-яёa-z0-9_]*/g],
    ['holding', /холдинг[а-яёa-z0-9_]*|концерн[а-яёa-z0-9_]*|финансово-промышленн[а-яёa-z0-9_]*/g],
  ];
  const BARE_NAME = /^(?:наименование|организация|участник|исполнитель|подрядчик|поставщик|претендент|заявитель|контрагент|компания|предприятие|юридическое\s+лицо|фирма|заказчик|покупатель)[\s:,.*№\-–—]*$/;
  const BARE_BOSS = /^(?:руководитель(?:\s+организации)?|подписант|уполномоченное\s+лицо|представитель|генеральный\s+директор|директор|президент|управляющий)[\s:,.*\-–—]*$/;
  const CODE_KEYS = ['inn', 'ogrn', 'kpp', 'okpo', 'okved', 'oktmo'];

  function scanOrder(t, keys) {          // ключи, упомянутые в строке, в порядке появления
    const found = [];
    for (const [k, re] of RULES) if (keys.includes(k)) { re.lastIndex = 0; const m = re.exec(t); if (m) found.push([m.index, k]); }
    return found.sort((a, b) => a[0] - b[0]).map(x => x[1]);
  }
  const uniq = a => a.filter((x, i) => a.indexOf(x) === i);

  /* classify(label) → { items:[ключи], role, bare } | null.  opt.tail — подпись взята из хвоста длинного абзаца */
  function classify(raw, opt) {
    opt = opt || {};
    const original = lc(clean(raw)), t = clean(lc(stripParens(clean(raw)))).replace(/[:：*\s]+$/, '');
    if (!t && !original) return null;
    if (/подпис|(?<![а-я])м\.?\s?п\.?(?![а-я])/.test(t) && !/[а-я]{4,}\s+подпис/.test(t) && t.length < 40) return { items: [], skip: true };
    const role = roleInText(t.slice(-60));
    const out = items => ({ items, role, bare: false });

    if (!opt.tail) {
      // банковские реквизиты одной строкой
      if (/банковск[а-яёa-z0-9_]*\s+реквизит|реквизит[а-яёa-z0-9_]*\s+банка|платежн[а-яёa-z0-9_]*\s+реквизит/.test(t)) {
        const o = original, it = [];
        if (/наименовани[а-яёa-z0-9_]*(?:\s+и\s+адрес)?\s+банка|(?<![а-я])банк(?![а-я])/.test(o)) it.push('bank_name');
        if (/адрес[а-яёa-z0-9_]*\s+банка/.test(o)) it.push('bank_addr');
        if (/р\s*\/\s*с|расч[а-яёa-z0-9_]*\s*сч|номер\s+сч/.test(o)) it.push('rs');
        if (/к\s*\/\s*с|кор[а-яёa-z0-9_]*\.?\s*сч/.test(o)) it.push('ks');
        if (/(?<![а-я])бик(?![а-я])/.test(o)) it.push('bik');
        if (/инн\s+банка/.test(o)) it.push('bank_inn');
        if (/телефон[а-яёa-z0-9_]*\s+банка/.test(o)) it.push('bank_phone');
        const order = ['bank_name', 'bank_addr', 'rs', 'ks', 'bik', 'bank_inn', 'bank_phone'];
        return out(it.length ? order.filter(k => it.includes(k)) : ['bank_name', 'bik', 'rs', 'ks']);
      }
      const QUESTION = /наличи|судим|задолженн|решени|сведени[а-яёa-z0-9_]*\s+о|отсутств|не\s+находит|подтвержда|предусмотрен/.test(t);
      if (QUESTION) { /* вопрос анкеты — не поле реквизитов */ }
      else if (/бухгалтер/.test(t)) return out(['accountant_fio']);
      if (!QUESTION && /контактн[а-яёa-z0-9_]*\s+лиц|ответственн[а-яёa-z0-9_]*\s+лиц/.test(t)) {
        const it = [/должност/.test(original) ? 'contact_fio_pos' : 'contact_fio'];
        if (/телефон/.test(original)) it.push('contact_phone');
        if (/почт|e-?mail/.test(original)) it.push('contact_email');
        return out(it);
      }
      const DIR = /руковод|директор|единолич|подписант|уполномоченн[а-яёa-z0-9_]*\s+лиц|представител|первое\s+лицо/.test(t);
      if (DIR && !QUESTION && t.length <= 150 && !/наименовани|адрес|реквизит|телефон\s*$/.test(t)) {
        if (/инн/.test(t)) return out(['director_inn']);
        const FIOW = /ф\.?\s?и\.?\s?о|фамили|имя|отчеств/.test(t);
        const it = [(/должност/.test(original) || !FIOW) ? 'director_pos_fio' : 'director_fio'];
        if (/телефон/.test(original)) it.push('director_phone');
        return out(it);
      }
      // составные строки: «ИНН; ОГРН; КПП …», «Фактический; Почтовый адрес», «сайт, e-mail»
      const codes = scanOrder(t, CODE_KEYS);
      if (codes.length >= 2 && !/инн\s*[/,\\и]\s*кпп/.test(t)) return out(codes);
      const adr = [];
      if (/юридическ/.test(t)) adr.push(['addr_legal', t.search(/юридическ/)]);
      if (/фактическ/.test(t)) adr.push(['addr_fact', t.search(/фактическ/)]);
      if (/почтов/.test(t)) adr.push(['addr_post', t.search(/почтов/)]);
      if (adr.length >= 2) return out(adr.sort((a, b) => a[1] - b[1]).map(x => x[0]));
      const con = scanOrder(t, ['phone', 'email', 'site', 'fax']);
      if (con.length >= 2 && !/адрес\s+электронной\s+почты/.test(t.replace(/(?:сайт[а-яёa-z0-9_]*|веб[а-яёa-z0-9_]*)[^;]*/, ''))) return out(con);
      if (con.length >= 2) return out(con);
      if (!/^(?:дата|номер|№)/.test(t)) {
        if (BARE_NAME.test(t) && !opt.noBare) return { items: ['name_full'], role: roleInText(t), bare: true };
        if (BARE_BOSS.test(t) && !opt.noBare) return { items: ['director_pos_fio'], role: null, bare: true, boss: true };
      }
    }
    // одиночное поле: в обычном режиме — первое по приоритету, в «хвостовом» — самое близкое к концу
    let best = null;
    for (const [key, re] of RULES) {
      re.lastIndex = 0;
      for (const m of t.matchAll(re)) {
        const end = m.index + m[0].length, len = m[0].length;
        if (opt.tail) { if (!best || end > best.end || (end === best.end && len > best.len)) best = { key, end, len }; }
        else if (!best) best = { key, end, len };
      }
      if (best && !opt.tail) break;
    }
    if (!best) return null;
    if (!opt.tail && best.key === 'signer_pos' && /наименован|параметр|документ|продукц/.test(t)) return null;
    return out([best.key]);
  }

  /* =====================================================================
     Разбор карты стороны (текст → поля)
     ===================================================================== */
  const OPF = [[/Общество с ограниченной ответственностью/gi, 'ООО'], [/Публичное акционерное общество/gi, 'ПАО'], [/Акционерное общество/gi, 'АО'], [/Индивидуальный предприниматель/gi, 'ИП']];
  const POS_RE = /^(генеральный\s+директор|исполнительный\s+директор|коммерческий\s+директор|директор|президент|руководитель|председатель|управляющий|индивидуальный\s+предприниматель|и\.\s*о\.\s*директора|главный\s+врач|ректор|начальник)\s+(.+)$/i;
  const splitPosFio = v => { const m = POS_RE.exec(clean(v)); return m ? { pos: m[1].replace(/^./, c => c.toUpperCase()), fio: m[2].trim() } : null; };
  const looksFio = v => /^[А-ЯЁ][а-яё-]+\s+[А-ЯЁ][а-яё.-]+(\s+[А-ЯЁ][а-яё.-]+)?$/.test(clean(v));

  function parseCardText(text) {
    const card = {};
    const set = (k, v) => { v = clean(v); if (v && !card[k]) card[k] = v; };
    const lines = String(text || '').split(/\r?\n/).map(s => clean(s)).filter(Boolean);
    const pairs = [];
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (line.includes(' | ')) {
        const c = line.split(' | ').map(s => s.trim());
        if (c.length === 2) pairs.push([c[0], c[1]]);
        else if (c.length === 4) { pairs.push([c[0], c[1]]); pairs.push([c[2], c[3]]); }
        else if (c.length > 2) { pairs.push([c[0], c.slice(1).filter(Boolean).join(' ')]); pairs.push([c[c.length - 2], c[c.length - 1]]); }
        continue;
      }
      const ci = line.indexOf(':');
      if (ci > 0 && ci <= 80) {
        const lab = line.slice(0, ci), val = line.slice(ci + 1).trim();
        if (val) { pairs.push([lab, val]); continue; }
        if (lines[i + 1] && !classify(lines[i + 1], { noBare: true })) pairs.push([lab, lines[i + 1]]);
        continue;
      }
      if (line.length <= 60 && classify(line) && lines[i + 1] && !lines[i + 1].includes(' | ') && !classify(lines[i + 1])) pairs.push([line, lines[i + 1]]);
    }
    const digits = (v, re) => { const m = re.exec(String(v).replace(/[\s-]/g, '')); return m ? m[0] : ''; };
    for (const [label, value] of pairs) {
      if (!value) continue;
      const r = classify(label); if (!r || r.skip || !r.items.length) continue;
      if (r.items.includes('bank_name') && r.items.length > 1) { set('bank_name', String(value).split(/,\s*(?:р\/с|к\/с|бик|расч|корр|р\.\s*с)/i)[0].replace(/^банк[:\s]+/i, '')); continue; }
      if (r.items.length > 1) {          // составная строка карты: разложим числа по длине
        if (r.items.every(k => CODE_KEYS.includes(k))) for (const n of (String(value).match(/\d{8,15}/g) || [])) {
          if (n.length === 9) set('kpp', n); else if (n.length === 10 || n.length === 12) { if (!card.inn) set('inn', n); else set('okpo', n); } else if (n.length === 13 || n.length === 15) set('ogrn', n); else set('okpo', n);
        }
        continue;
      }
      const k = r.items[0];
      switch (k) {
        case 'inn_kpp': { const m = String(value).match(/\d{9,12}/g) || []; set('inn', m.find(x => x.length >= 10)); set('kpp', m.find(x => x.length === 9)); break; }
        case 'inn': set('inn', digits(value, /\d{10,12}/)); break;
        case 'kpp': set('kpp', digits(value, /\d{9}/)); break;
        case 'ogrn': set('ogrn', digits(value, /\d{13,15}/)); break;
        case 'okpo': set('okpo', digits(value, /\d{8,10}/)); break;
        case 'bik': set('bik', digits(value, /\d{9}/)); break;
        case 'rs': set('rs', digits(value, /\d{20}/)); break;
        case 'ks': set('ks', digits(value, /\d{20}/)); break;
        case 'director_inn': set('director_inn', digits(value, /\d{12}/)); break;
        case 'signer_fio': set('director_fio', value); break;
        case 'signer_pos': set('director_pos', value); break;
        case 'director_fio': set('director_fio', value); break;
        case 'director_pos_fio': case 'signer_pos_fio': {
          const sp = splitPosFio(value);
          if (sp) { set('director_pos', sp.pos); set('director_fio', sp.fio); }
          else if (looksFio(value)) { if (BARE_BOSS.test(lc(clean(label))) && !/^(руководитель|подписант|уполномоч|представитель)/i.test(label)) set('director_pos', label.replace(/[:\s]+$/, '').replace(/^./, c => c.toUpperCase())); set('director_fio', value); }
          break;
        }
        case 'addr_any': set('addr_legal', value); break;
        case 'contact_fio_pos': case 'contact_fio': { const sp = splitPosFio(value); if (sp) { set('contact_pos', sp.pos); set('contact_fio', sp.fio); } else set('contact_fio', value); break; }
        case 'inperson': case 'doc_date': case 'vat_text': case 'vat_payer': break;
        default: if (FIELD_TITLES[k] || k === 'branches' || k === 'holding') set(k, value);
      }
    }
    const all = lines.join('\n');
    const fb = (k, re, g) => { if (!card[k]) { const m = all.match(re); if (m) set(k, m[g || 1]); } };
    fb('inn', /ИНН[^\d]{0,15}(?<!\d)(\d{12}|\d{10})(?!\d)/i);
    fb('kpp', /КПП[^\d]{0,15}(?<!\d)(\d{9})(?!\d)/i);
    fb('ogrn', /ОГРН(?:ИП)?[^\d]{0,15}(?<!\d)(\d{15}|\d{13})(?!\d)/i);
    fb('bik', /БИК[^\d]{0,15}(?<!\d)(\d{9})(?!\d)/i);
    for (const m of all.matchAll(/(?<![\d])(\d{20})(?!\d)/g)) { if (m[1].startsWith('301')) set('ks', m[1]); else set('rs', m[1]); }
    fb('email', /[\w.+-]+@[\w-]+(?:\.[\w-]+)+/, 0);
    if (!card.phone) {          // телефон — только отдельное число, а не кусок длинного номера (счёта, ИНН, ОГРН)
      const m = all.match(/(?<![\d])(?:\+7|8)[\s(-]*\d{3}[\s)-]*\d{3}[\s-]*\d{2}[\s-]*\d{2}(?![\d])/);
      if (m) set('phone', m[0]);
    }
    if (!card.name_full) {
      const m = all.match(/(?:Общество с ограниченной ответственностью|Акционерное общество|Публичное акционерное общество)\s*[«"“][^»"”\n]+[»"”]|(?:ООО|АО|ПАО|ЗАО|ОАО|НАО|АНО|МУП|ГУП|ФГУП)\s*[«"“][^»"”\n]+[»"”]|Индивидуальный предприниматель\s+[А-ЯЁ][а-яё-]+(?:\s+[А-ЯЁ][а-яё-]+){1,2}|ИП\s+[А-ЯЁ][а-яё-]+(?:\s+[А-ЯЁ][а-яё.-]+){1,2}/);
      if (m) set('name_full', m[0]);
    }
    if (!card.director_fio) {
      const m = all.match(/(Генеральный директор|Исполнительный директор|Директор|Руководитель|Президент)[:\s–—-]+([А-ЯЁ][а-яё-]+(?:\s+[А-ЯЁ][а-яё.-]+){1,2})/);
      if (m) { set('director_pos', m[1]); set('director_fio', m[2]); }
    }
    return finishCard(card);
  }

  function finishCard(card) {
    const set = (k, v) => { v = clean(v); if (v && !card[k]) card[k] = v; };
    if (!card.name_short && card.name_full) { let s = card.name_full; for (const [re, to] of OPF) s = s.replace(re, to); set('name_short', s); }
    if (!card.name_full && card.name_short) set('name_full', card.name_short);
    const ip = isIp(card);
    if (ip && !card.director_fio) { const m = (card.name_full || card.name_short).match(/^(?:ИП|Индивидуальный предприниматель)\s+(.+)$/i); if (m) { set('director_fio', m[1]); set('director_pos', 'Индивидуальный предприниматель'); } }
    if (card.director_pos && !card.director_pos_gen) set('director_pos_gen', posGen(card.director_pos));
    if (card.director_fio && !card.director_fio_gen) set('director_fio_gen', fioGen(card.director_fio));
    if (!card.basis && card.director_fio) set('basis', ip ? 'свидетельства о государственной регистрации' + (card.ogrn ? ' (ОГРНИП ' + card.ogrn + ')' : '') : 'Устава');
    return card;
  }

  /* =====================================================================
     Контекст заполнения: стороны, документ, журнал
     ===================================================================== */
  const ALLOWED_MIN = new Set(['name_full', 'name_short', 'signer_fio', 'signer_fio_short', 'signer_pos', 'signer_pos_fio', 'director_pos_fio', 'director_fio', 'inperson']);
  const SHORT = { inn: 'ИНН', kpp: 'КПП', okpo: 'ОКПО', okved: 'ОКВЭД2', oktmo: 'ОКТМО', addr_fact: 'Фактический', addr_post: 'Почтовый', addr_legal: 'Юридический', phone: 'тел.', email: 'e-mail', site: 'сайт', fax: 'факс', rs: 'р/с', ks: 'к/с', bik: 'БИК', bank_inn: 'ИНН банка', bank_addr: 'адрес банка', bank_phone: 'тел. банка', contact_phone: 'тел.', contact_email: 'e-mail', director_phone: 'тел.' };

  function makeCtx(parties, doc, opts) {
    opts = Object.assign({ mode: 'full', highlight: true, fallbackAddr: true, fallbackContact: true, docType: 'auto', defaultRole: 'executor' }, opts || {});
    doc = Object.assign({ date: '', number: '', specNo: '', city: '', amount: '', vat: '22', vatMode: 'incl' }, doc || {});
    parties = { customer: (parties && parties.customer) || {}, executor: (parties && parties.executor) || {} };
    const ctx = {
      parties, doc, opts, handled: new Set(), items: [], stats: { filled: 0, highlighted: 0 }, role: null, contract: false,
      allowed(key) { return opts.mode !== 'min' || ALLOWED_MIN.has(key); },
      amountNum() { return parseNum(doc.amount); },
      vatRate() { return parseFloat(doc.vat) || 0; },
      single(key, role) {
        const c = parties[role] || {}; const d = doc;
        const dirFio = c.director_fio || '';
        switch (key) {
          case 'name_full': return c.name_full || c.name_short || '';
          case 'name_short': return c.name_short || c.name_full || '';
          case 'name_addr': return [c.name_full || c.name_short, c.addr_legal].filter(Boolean).join(', ');
          case 'inn_kpp': return c.inn ? (c.kpp ? c.inn + '/' + c.kpp : c.inn) : '';
          case 'ogrn': return c.ogrn || '';
          case 'addr_any': case 'addr_legal': return c.addr_legal || '';
          case 'addr_post': return c.addr_post || (opts.fallbackAddr ? c.addr_legal : '') || '';
          case 'addr_fact': return c.addr_fact || (opts.fallbackAddr ? c.addr_legal : '') || '';
          case 'signer_fio': case 'director_fio': return dirFio;
          case 'signer_fio_short': return shortFio(dirFio);
          case 'signer_pos': case 'director_pos': return c.director_pos || '';
          case 'signer_pos_fio': case 'director_pos_fio': case 'sig_title': return dirFio ? [c.director_pos, dirFio].filter(Boolean).join(' ') : '';
          case 'director_inn': return c.director_inn || (isIp(c) ? 'Отсутствует' : '');
          case 'director_phone': return c.phone || '';
          case 'inperson': return (c.director_pos_gen && c.director_fio_gen) ? c.director_pos_gen + ' ' + c.director_fio_gen : '';
          case 'contact_fio_pos': { const f = c.contact_fio || (opts.fallbackContact ? dirFio : ''); const p = c.contact_fio ? c.contact_pos : (opts.fallbackContact ? c.director_pos : ''); return f ? [p, f].filter(Boolean).join(' ') : ''; }
          case 'contact_fio': return c.contact_fio || (opts.fallbackContact ? dirFio : '');
          case 'contact_phone': return c.contact_phone || (opts.fallbackContact ? c.phone : '') || '';
          case 'contact_email': return c.contact_email || (opts.fallbackContact ? c.email : '') || '';
          case 'vat_payer': return c.tax_system ? (/осн|с\s*ндс/i.test(c.tax_system) ? 'Да' : 'Нет') : '';
          case 'vat_text': return c.tax_system ? (/осн|с\s*ндс/i.test(c.tax_system) ? 'Да, с НДС по ставке ' + (d.vat || 22) + ' %' : 'Без НДС') : '';
          case 'doc_date': return dateShort(d.date);
          case 'doc_no': return d.number || '';
          case 'spec_no': return d.specNo || '';
          case 'city': return d.city || '';
          case 'amount_num': { const n = parseNum(d.amount); return n == null ? '' : fmtMoney(n); }
          case 'amount_words': { const n = parseNum(d.amount); return n == null ? '' : rubWords(n); }
          case 'vat_amount': { const n = parseNum(d.amount), r = parseFloat(d.vat) || 0; if (n == null || !r) return ''; return fmtMoney(d.vatMode === 'excl' ? n * r / 100 : n * r / (100 + r)); }
          case 'amount': { const n = parseNum(d.amount); return n == null ? '' : fmtMoney(n) + ' руб. (' + rubWords(n) + ')'; }
          default: return c[key] || '';
        }
      },
      // значение для набора ключей: один ключ — как есть, несколько — «Подпись значение»
      render(items, role) {
        const vals = items.map(k => [k, this.single(k, role)]);
        if (items.length === 1) return { text: vals[0][1], miss: vals[0][1] ? [] : [items[0]] };
        const parts = [], miss = [];
        for (const [k, v] of vals) { if (!v) { miss.push(k); continue; } const sh = k === 'ogrn' ? (isIp(parties[role] || {}) ? 'ОГРНИП' : 'ОГРН') : SHORT[k]; parts.push((sh ? sh + (/^addr_/.test(k) ? ': ' : ' ') : '') + v); }
        const bankish = items.some(k => /^(bank_|rs$|ks$|bik$|contact_|director_)/.test(k));
        return { text: parts.join(bankish ? ', ' : '; '), miss };
      },
      dateLong() { return dateLong(doc.date); },
      filled(key, val, role) { ctx.stats.filled++; ctx.items.push({ kind: 'fill', key, val, role }); },
      hl(what, role) { ctx.stats.highlighted++; ctx.items.push({ kind: 'hl', key: what, role }); },
    };
    return ctx;
  }
  const titleOf = k => (FIELD_TITLES[k] || ({ director_pos_fio: 'Должность и ФИО руководителя', signer_fio: 'ФИО подписанта', signer_fio_short: 'ФИО (инициалы)', signer_pos: 'Должность подписанта', signer_pos_fio: 'Должность и ФИО', sig_title: 'Руководитель', inperson: 'В лице (род. падеж)', inn_kpp: 'ИНН/КПП', addr_any: 'Адрес', contact_fio_pos: 'Контактное лицо', vat_payer: 'Плательщик НДС', vat_text: 'НДС', doc_date: 'Дата документа', doc_no: '№ договора', spec_no: '№ спецификации', city: 'Город', amount: 'Сумма договора', branches: 'Филиалы', holding: 'Холдинг', director_phone: 'Телефон руководителя', name_addr: 'Наименование и адрес', amount_num: 'Сумма договора', amount_words: 'Сумма прописью', vat_amount: 'Сумма НДС', amount: 'Сумма договора' })[k] || k);

  /* =====================================================================
     DOCX: операции с прогонами (w:r)
     ===================================================================== */
  const mkW = (doc, name) => doc.createElementNS(NS_W, 'w:' + name);
  function runText(r) { let s = ''; for (let c = r.firstChild; c; c = c.nextSibling) { if (c.nodeType !== 1) continue; const n = lname(c); if (n === 't') s += c.textContent; else if (n === 'tab') s += '\t'; } return s; }
  function isSimpleRun(r) { for (let c = r.firstChild; c; c = c.nextSibling) { if (c.nodeType !== 1) continue; if (!['rPr', 't', 'tab'].includes(lname(c))) return false; } return true; }
  function setRunText(r, text) {
    const doc = r.ownerDocument;
    for (const c of Array.from(r.childNodes)) if (c.nodeType === 1 && ['t', 'tab'].includes(lname(c))) r.removeChild(c);
    text.split('\t').forEach((part, i) => {
      if (i > 0) r.appendChild(mkW(doc, 'tab'));
      if (part !== '') { const t = mkW(doc, 't'); t.setAttributeNS(NS_XML, 'xml:space', 'preserve'); t.appendChild(doc.createTextNode(part)); r.appendChild(t); }
    });
  }
  const RECURSE = new Set(['hyperlink', 'ins', 'smartTag', 'fldSimple', 'sdt', 'sdtContent', 'customXml']);
  function paraRuns(p) { const out = []; (function walk(n) { for (let c = n.firstChild; c; c = c.nextSibling) { if (c.nodeType !== 1) continue; const nm = lname(c); if (nm === 'r') out.push(c); else if (RECURSE.has(nm)) walk(c); } })(p); return out; }
  function mapParagraph(p) { let pos = 0; const items = paraRuns(p).map(r => { const t = runText(r); const o = { r, t, s: pos, e: pos + t.length }; pos += t.length; return o; }); return { items, text: items.map(i => i.t).join('') }; }
  function splitAt(p, k) {
    for (const it of mapParagraph(p).items) if (it.s < k && k < it.e) {
      if (!isSimpleRun(it.r)) return;
      const cl = it.r.cloneNode(true); setRunText(it.r, it.t.slice(0, k - it.s)); setRunText(cl, it.t.slice(k - it.s)); it.r.parentNode.insertBefore(cl, it.r.nextSibling); return;
    }
  }
  function rangeRuns(p, s, e) { splitAt(p, s); splitAt(p, e); return mapParagraph(p).items.filter(it => it.t.length > 0 && it.s < e && it.e > s).map(it => it.r); }
  const HL_BEFORE = ['u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'];
  function addHighlight(run) {
    const doc = run.ownerDocument; let rPr = firstKid(run, 'rPr');
    if (!rPr) { rPr = mkW(doc, 'rPr'); run.insertBefore(rPr, run.firstChild); }
    for (const h of kids(rPr, 'highlight')) rPr.removeChild(h);
    const hl = mkW(doc, 'highlight'); hl.setAttributeNS(NS_W, 'w:val', 'yellow');
    const before = Array.from(rPr.childNodes).find(c => c.nodeType === 1 && HL_BEFORE.includes(lname(c)));
    if (before) rPr.insertBefore(hl, before); else rPr.appendChild(hl);
  }
  const stripFmt = (rPr, names) => { for (const c of Array.from(rPr.childNodes)) if (c.nodeType === 1 && names.includes(lname(c))) rPr.removeChild(c); };
  function cleanRPr(rPr, extra) { if (!rPr) return null; const c = rPr.cloneNode(true); stripFmt(c, ['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange', 'highlight'].concat(extra || [])); return c; }
  function makeRun(doc, text, rPr) { const r = mkW(doc, 'r'); if (rPr) r.appendChild(rPr); setRunText(r, text); return r; }
  function fillRange(p, s, e, val) {
    const runs = rangeRuns(p, s, e); if (!runs.length) return false;
    setRunText(runs[0], val); const rp = firstKid(runs[0], 'rPr'); if (rp) stripFmt(rp, ['highlight', 'u']);
    for (let i = 1; i < runs.length; i++) runs[i].parentNode.removeChild(runs[i]);
    return true;
  }
  function highlightRange(p, s, e) { const runs = rangeRuns(p, s, e); runs.forEach(addHighlight); return runs.length > 0; }
  const NBSP6 = '      ';
  function lastTextRunRPr(p) { const runs = paraRuns(p).filter(r => runText(r).trim()); const r = runs[runs.length - 1]; return r ? firstKid(r, 'rPr') : null; }
  function appendToParagraph(doc, p, text, highlight) { const run = makeRun(doc, text, cleanRPr(lastTextRunRPr(p), ['b', 'bCs'])); if (highlight) addHighlight(run); p.appendChild(run); }

  /* ---------- пустые места в абзаце ---------- */
  const BLANK_RE = /_{3,}|\.{5,}|…{2,}/g;
  const PH_RE = /\[([^\]\n]{2,160})\]|\{\{([^}\n]{2,60})\}\}|(?<![\w/])<([^<>\n]{2,60})>/g;
  const DATE_RE = /[«"“]\s*_*\s*[»"”]\s*_{2,}\s*(?:20\d?_{1,}|20\d\d)\s*(?:г\.?|года)?/;
  function findBlanks(map) {
    const ranges = [];
    for (const m of map.text.matchAll(BLANK_RE)) ranges.push({ s: m.index, e: m.index + m[0].length });
    for (const it of map.items) if (it.t.length >= 3 && !it.t.trim()) {
      const rPr = firstKid(it.r, 'rPr'), u = rPr && firstKid(rPr, 'u');
      if (u && (u.getAttributeNS(NS_W, 'val') || u.getAttribute('w:val')) !== 'none') ranges.push({ s: it.s, e: it.e });
    }
    for (const m of map.text.matchAll(PH_RE)) ranges.push({ s: m.index, e: m.index + m[0].length, ph: true, inner: m[1] || m[2] || m[3] });
    ranges.sort((a, b) => a.s - b.s || b.e - a.e);
    const out = [];
    for (const r of ranges) { const last = out[out.length - 1]; if (last && r.s < last.e) continue; out.push(r); }
    return out;
  }
  const trimEnd = s => s.replace(/[\s ]+$/, ''), trimStart = s => s.replace(/^[\s ]+/, '');

  /* что стоит перед пустым местом → какие поля вписать */
  function resolveContext(before, after, prevText, nextText, ctx) {
    let b = trimEnd(before), a = trimStart(after);
    if (/^[\s.,;:_\-–—]*$/.test(a)) a = '';
    const lineIsBlankOnly = !b && !a;
    if (!b && prevText && prevText.length < 160 && !/именуем/.test(a)) b = trimEnd(prevText);
    if (lineIsBlankOnly && nextText) a = trimStart(nextText);
    const K = (items, role) => ({ act: 'key', items, role: role || null });
    if (/^[\s)\]/(]*(?:подпис|м\.\s*п)/i.test(a)) return { act: 'skip' };
    if (/(?:м\.\s*п\.?|подпись\s*[:\-–]?)[)\s]*$/i.test(b.slice(-30))) return { act: 'skip' };
    // название стороны договора: «_____, именуемое в дальнейшем «Заказчик»»
    if (/^[,\s]*(?:именуем|далее\s+именуем|в\s+дальнейшем)/i.test(a)) return K(['name_full']);
    const nm = /^[,\s]*(?:\(\s*)?(полное\s+)?(?:фирменное\s+)?наименование(\s+и\s+адрес)?/i.exec(a);
    if (nm && lineIsBlankOnly || (nm && !b)) return K([nm[2] ? 'name_addr' : 'name_full'], roleInText(a.slice(0, 80)));
    const hm = a.match(/^[)\]/]*\s*\(\s*([^)]{0,70})\)/);
    if (hm) {
      const h = lc(hm[1]);
      if (/должност/.test(h) && /ф\.?\s*и\.?\s*о|фамили/.test(h)) return K([/в\s+лице\s*$/i.test(b) ? 'inperson' : 'director_pos_fio'], roleInText(b.slice(-80)));
      if (/ф\.?\s*и\.?\s*о|фамили|расшифровк|инициал/.test(h)) return K([/расшифровк|инициал/.test(h) ? 'signer_fio_short' : 'signer_fio'], roleInText(b.slice(-80)));
      if (/должност/.test(h)) return K(['signer_pos'], roleInText(b.slice(-80)));
      if (/наименовани|организаци/.test(h)) return K(['name_full'], roleInText(b.slice(-80)));
      if (/подпис/.test(h)) return { act: 'skip' };
    }
    if (/\/\s*$/.test(b)) return K(['signer_fio_short'], roleInText(b.slice(-120)));
    if (/^\//.test(a)) return { act: 'skip' };
    const tail = b.slice(-120), lt = lc(tail);
    if (/(?:^|[\s«»"])г\.?\s*$|город\s*$/.test(lt) || /^г\.?$/.test(lt)) return K(['city']);
    if (/(?:№|n)\s*$/i.test(lt)) {
      const wide = lc(b);
      if (/спецификаци|приложени/.test(wide) && !/к\s+договор[а-яёa-z0-9_]*\s*$/.test(wide)) return K(['spec_no']);
      if (/договор|контракт/.test(wide) || ctx.contract) return K(['doc_no']);
      return { act: 'unknown', label: tail.slice(-40) };
    }
    if (/(?:в\s+том\s+числе|включая|из\s+них)\s+ндс[^.]{0,40}$|сумма\s+ндс[^.]{0,30}$/.test(lt) || /ндс[^.]{0,30}(?:составляет|—|-|:)\s*$/.test(lt)) return K(['vat_amount']);
    if (/(?:сумм[а-яёa-z0-9_]*|цен[а-яёa-z0-9_]*|стоимост[а-яёa-z0-9_]*)[^:.]{0,80}$/.test(lt) && !/ндс/.test(lt.slice(-30))) return K([/\(\s*$/.test(tail) ? 'amount_words' : 'amount_num']);
    if (BARE_NAME.test(lt.replace(/[:\s]+$/, ''))) return K(['name_full'], roleInText(lt));
    const r = classify(tail, { tail: true });
    if (!r) return { act: 'unknown', label: tail.slice(-40) };
    if (r.items[0] === 'director_pos_fio' && r.boss) return { act: 'skip' };
    if (r.items[0] === 'signer_pos_fio' && /:\s*$/.test(tail) === false && !a) return { act: 'skip' };
    return K(r.items, r.role);
  }
  const siblingParaText = (p, dir) => { let n = dir < 0 ? p.previousSibling : p.nextSibling; while (n && n.nodeType !== 1) n = dir < 0 ? n.previousSibling : n.nextSibling; return n && lname(n) === 'p' ? mapParagraph(n).text : ''; };

  function pickRole(ctx, explicit, extra) {
    return explicit || extra || ctx.role || (ctx.contract ? null : ctx.opts.defaultRole);
  }
  function applyResolution(p, b, r, ctx, role) {
    if (r.act === 'skip') return;
    if (r.act === 'unknown') { if (ctx.opts.mode !== 'min' && ctx.opts.highlight) { highlightRange(p, b.s, b.e); ctx.hl(r.label || 'unknown', null); } return; }
    const items = r.items.filter(k => ctx.allowed(k)); if (!items.length) return;
    const needsRole = items.some(k => !['doc_date', 'doc_no', 'spec_no', 'city', 'amount_num', 'amount_words', 'vat_amount'].includes(k));
    if (needsRole && !role) { if (ctx.opts.highlight) { highlightRange(p, b.s, b.e); ctx.hl('не определена сторона: ' + items.map(titleOf).join(', '), null); } return; }
    const out = ctx.render(items, role);
    if (out.text) { fillRange(p, b.s, b.e, out.text); ctx.filled(items.map(titleOf).join(' + '), out.text, needsRole ? role : null); }
    else if (ctx.opts.highlight) { highlightRange(p, b.s, b.e); ctx.hl(items.map(titleOf).join(' + '), needsRole ? role : null); }
  }

  function processParagraph(doc, p, ctx) {
    let map = mapParagraph(p);
    const text0 = map.text; if (!text0.trim()) return;
    const hr = headingRole(text0);
    if (hr) { ctx.role = hr; return; }
    // метки сторон в абзаце: «именуемое в дальнейшем «Заказчик»»
    const markers = []; for (const m of text0.matchAll(/именуем[а-яёa-z0-9_]*(?:\s+в\s+дальнейшем)?\s*[,(]?\s*[«"“]?\s*([А-Яа-яЁё]+)/gi)) { const rl = roleOfWord(m[1]); if (rl) markers.push({ idx: m.index, rl }); }
    let dateRange = null;
    const dm = map.text.match(DATE_RE);
    if (dm) {
      dateRange = { s: dm.index, e: dm.index + dm[0].length };
      const dv = ctx.dateLong();
      if (dv && ctx.allowed('doc_date')) { fillRange(p, dateRange.s, dateRange.e, dv); ctx.filled('Дата документа', dv, null); dateRange = null; }
      else if (ctx.opts.mode !== 'min' && ctx.opts.highlight) { highlightRange(p, dateRange.s, dateRange.e); ctx.hl('Дата документа', null); }
      map = mapParagraph(p);
    }
    let blanks = findBlanks(map);
    if (dateRange) blanks = blanks.filter(b => !(b.s >= dateRange.s && b.e <= dateRange.e));
    const prevText = siblingParaText(p, -1), nextText = siblingParaText(p, +1);
    const textNow = map.text;
    if (blanks.length) {
      // роли считаем по исходным позициям, правим с конца
      const plan = blanks.map((b, i) => {
        let r;
        if (b.ph) { const c = classify(b.inner, { noBare: true }); r = c && c.items.length ? { act: 'key', items: c.items, role: c.role } : { act: 'skip' }; }
        else r = resolveContext(textNow.slice(0, b.s), textNow.slice(b.e), prevText, nextText, ctx);
        let mr = null;
        const lastB = markers.filter(m => m.idx < b.s).pop(), nextA = markers.find(m => m.idx >= b.e);
        if (r.act === 'key' && r.items.includes('name_full') && nextA && !(blanks[i + 1] && blanks[i + 1].s < nextA.idx) && nextA.idx - b.e < 220) mr = nextA.rl;
        else if (lastB) mr = lastB.rl;
        return { b, r, role: pickRole(ctx, r.role, mr) };
      });
      for (let i = 0; i < plan.length - 1; i++) {      // «_____ [наименование участника]» — оставляем одно значение
        const a1 = plan[i], a2 = plan[i + 1];
        if (!a1.b.ph && a2.b.ph && !textNow.slice(a1.b.e, a2.b.s).trim()) {
          if (a1.r.act !== 'key' && a2.r.act === 'key') { a1.r = a2.r; a1.role = pickRole(ctx, a2.r.role, a1.role); }
          a2.drop = true;
        }
      }
      for (let i = plan.length - 1; i >= 0; i--) { if (plan[i].drop) fillRange(p, plan[i].b.s, plan[i].b.e, ''); else applyResolution(p, plan[i].b, plan[i].r, ctx, plan[i].role); }
    } else {
      const t = trimEnd(map.text);
      if (t.length <= 140 && /[:：]$/.test(t)) {
        const c = classify(t.slice(0, -1), { noBare: false });
        if (c && !c.skip && c.items.length && !c.boss) {
          const items = c.items.filter(k => ctx.allowed(k));
          const role = pickRole(ctx, c.role);
          if (items.length && role) {
            const out = ctx.render(items, role);
            if (out.text) { appendToParagraph(doc, p, ' ' + out.text, false); ctx.filled(items.map(titleOf).join(' + '), out.text, role); }
            else if (ctx.opts.highlight) { appendToParagraph(doc, p, ' ' + NBSP6, true); ctx.hl(items.map(titleOf).join(' + '), role); }
          }
        }
      }
    }
    if (markers.length) ctx.role = markers[markers.length - 1].rl;
  }

  /* =====================================================================
     DOCX: таблицы
     ===================================================================== */
  const cellText = tc => byNS(tc, NS_W, 't').map(t => t.textContent).join('');
  const isBlankText = s => /^[\s _.…]*$/.test(s);
  const hasObjects = tc => byNS(tc, NS_W, 'drawing').length > 0 || byNS(tc, NS_W, 'tbl').length > 0;
  const looksLikeField = l => /[:：]\s*$/.test(l) || (l.split(/\s+/).length >= 2 && !/\d{3,}/.test(l));
  const AMOUNT_LABEL = /(?:сумма|цена|стоимость)\s+(?:договора|контракта|заявки|предложения)|общая\s+(?:сумма|стоимость)\s+(?:договора|контракта)?/;

  function writeCell(doc, tc, val, o) {
    o = o || {};
    const p = kids(tc, 'p')[0]; if (!p) return;
    const runs = paraRuns(p); let rPr = null;
    if (runs.length) rPr = cleanRPr(firstKid(runs[0], 'rPr'));
    if (!rPr) { const pPr = firstKid(p, 'pPr'); const mark = pPr && firstKid(pPr, 'rPr'); if (mark) rPr = cleanRPr(mark); }
    if (!rPr && o.labelTc) { const lp = kids(o.labelTc, 'p')[0]; const lr = lp && paraRuns(lp).find(r => runText(r).trim()); const rp = lr && firstKid(lr, 'rPr'); if (rp) rPr = cleanRPr(rp, ['b', 'bCs', 'u']); }
    for (const r of runs) r.parentNode.removeChild(r);
    const run = makeRun(doc, o.highlight ? NBSP6 : val, rPr);
    if (o.highlight) addHighlight(run);
    p.appendChild(run);
  }

  function handleTarget(doc, ctx, tcLabel, tcTarget, label, noBare, cellRole) {
    const lab = lc(label);
    let items, role = null, explicit = null;
    if (AMOUNT_LABEL.test(lab) && !/ндс/.test(lab.replace(/с\s+учет[а-яёa-z0-9_]*[^,]*ндс/, ''))) { items = ['amount']; }
    else {
      const r = classify(label, { noBare });
      if (r && r.skip) return;
      if (r && r.items.length) { items = r.items; explicit = r.role; }
    }
    if (items) {
      items = items.filter(k => ctx.allowed(k)); if (!items.length) return;
      const needsRole = items.some(k => !['doc_date', 'doc_no', 'spec_no', 'city', 'amount', 'amount_num', 'amount_words', 'vat_amount'].includes(k));
      role = pickRole(ctx, explicit, cellRole);
      if (needsRole && !role) { if (ctx.opts.highlight) { writeCell(doc, tcTarget, '', { labelTc: tcLabel, highlight: true }); ctx.hl('не определена сторона: ' + items.map(titleOf).join(', '), null); } return; }
      const out = ctx.render(items, role);
      if (out.text) { writeCell(doc, tcTarget, out.text, { labelTc: tcLabel }); ctx.filled(items.map(titleOf).join(' + '), out.text, needsRole ? role : null); }
      else if (ctx.opts.highlight) { writeCell(doc, tcTarget, '', { labelTc: tcLabel, highlight: true }); ctx.hl(items.map(titleOf).join(' + '), needsRole ? role : null); }
    } else if (ctx.opts.mode !== 'min' && ctx.opts.highlight && looksLikeField(label)) {
      writeCell(doc, tcTarget, '', { labelTc: tcLabel, highlight: true });
      ctx.hl(clean(label).slice(0, 50), null);
    }
  }

  /* таблица-спецификация: суммы строк, ИТОГО, НДС */
  const gridCols = tr => { let c = 0; return kids(tr, 'tc').map(tc => { const pr = firstKid(tc, 'tcPr'), gs = pr && firstKid(pr, 'gridSpan'); const start = c; c += gs ? (parseInt(gs.getAttributeNS(NS_W, 'val') || gs.getAttribute('w:val'), 10) || 1) : 1; return { tc, start, span: c - start }; }); };
  function processSpec(doc, tbl, ctx) {
    const rows = kids(tbl, 'tr'); if (rows.length < 2) return false;
    let hi = -1, qty = -1, price = -1, sum = -1, priceNoVat = false;
    for (let i = 0; i < Math.min(rows.length, 3) && hi < 0; i++) {
      const g = gridCols(rows[i]); let q = -1, p = -1, s = -1, pn = false;
      g.forEach(c => { const t = lc(cellText(c.tc)); if (/кол-?\s?во|количеств/.test(t) && q < 0) q = c.start; else if (/цена/.test(t) && p < 0) { p = c.start; pn = /без\s+ндс/.test(t); } else if (/сумма|стоимост|всего/.test(t) && !/цена/.test(t) && s < 0) s = c.start; });
      if (q >= 0 && p >= 0 && s >= 0) { hi = i; qty = q; price = p; sum = s; priceNoVat = pn; }
    }
    if (hi < 0) return false;
    const cellAt = (tr, col) => gridCols(tr).find(c => c.start === col) || null;
    let total = 0, cnt = 0, missing = 0, totalRowIdx = -1, vatRowIdx = -1;
    const vatExcl = priceNoVat || ctx.doc.vatMode === 'excl';
    rows.slice(hi + 1).forEach((tr, k) => {
      const idx = hi + 1 + k, first = lc(cellText(kids(tr, 'tc')[0] || tr));
      if (/^\s*итого/.test(first)) { totalRowIdx = idx; return; }
      if (/ндс/.test(first) && /(?:в\s+том\s+числе|включая|сумма|ндс)/.test(first) && idx > hi + 1) { vatRowIdx = idx; return; }
      const cq = cellAt(tr, qty), cp = cellAt(tr, price), cs = cellAt(tr, sum); if (!cq || !cp || !cs) return;
      const q = parseNum(cellText(cq.tc)), pr = parseNum(cellText(cp.tc)), sv = parseNum(cellText(cs.tc));
      if (isBlankText(cellText(cp.tc)) && q != null) { if (ctx.opts.highlight && ctx.opts.mode !== 'min') { writeCell(doc, cp.tc, '', { highlight: true }); ctx.hl('Цена за единицу', null); } }
      if (q != null && pr != null) { const v = Math.round(q * pr * 100) / 100; if (isBlankText(cellText(cs.tc))) { writeCell(doc, cs.tc, fmtMoney(v), { labelTc: cp.tc }); ctx.filled('Сумма по позиции', fmtMoney(v), null); } total += v; cnt++; }
      else if (sv != null) { total += sv; cnt++; } else missing++;
    });
    const rate = ctx.vatRate();
    const fillTotal = (idx, val, name) => {
      if (idx < 0 || val == null) return; const c = cellAt(rows[idx], sum) || gridCols(rows[idx]).slice(-1)[0];
      if (c && isBlankText(cellText(c.tc))) { writeCell(doc, c.tc, fmtMoney(val), {}); ctx.filled(name, fmtMoney(val), null); }
    };
    let grand = null;
    if (cnt && !missing) grand = vatExcl && rate ? total * (1 + rate / 100) : total;
    if (grand == null) { const a = parseNum(ctx.doc.amount); if (a != null && cnt === 0) grand = a; }
    if (grand != null) {
      fillTotal(totalRowIdx, vatExcl && rate && cnt && !missing ? total : grand, 'ИТОГО');
      const vat = rate ? (vatExcl ? grand - grand / (1 + rate / 100) : grand * rate / (100 + rate)) : null;
      if (vat != null) fillTotal(vatRowIdx, vat, 'В том числе НДС');
    } else if (ctx.opts.highlight && ctx.opts.mode !== 'min') {
      for (const idx of [totalRowIdx, vatRowIdx]) if (idx >= 0) { const c = cellAt(rows[idx], sum) || gridCols(rows[idx]).slice(-1)[0]; if (c && isBlankText(cellText(c.tc))) { writeCell(doc, c.tc, '', { highlight: true }); ctx.hl(idx === totalRowIdx ? 'ИТОГО' : 'НДС', null); } }
    }
    return true;
  }

  function walkBlock(doc, node, ctx) {
    for (let c = node.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      const n = lname(c);
      if (n === 'p') { if (!insideHandled(c, ctx)) { try { processParagraph(doc, c, ctx); } catch (e) { ctx.items.push({ kind: 'err', key: String(e && e.message || e) }); } } }
      else if (n === 'tbl') processTable(doc, c, ctx);
      else if (n === 'sdt' || n === 'sdtContent') walkBlock(doc, c, ctx);
    }
  }
  function insideHandled(p, ctx) { for (let n = p.parentNode; n; n = n.parentNode) if (ctx.handled.has(n)) return true; return false; }

  function processTable(doc, tbl, ctx) {
    const rows = kids(tbl, 'tr');
    const saveRole = ctx.role;
    if (processSpec(doc, tbl, ctx)) { ctx.role = saveRole; return; }
    // роли колонок: «ЗАКАЗЧИК | ИСПОЛНИТЕЛЬ» в первых строках
    const colRole = {};
    rows.slice(0, 2).forEach(tr => gridCols(tr).forEach((c, i) => { const hr = headingRole(cellText(c.tc)); if (hr) colRole[i] = hr; }));
    const twoRole = Object.keys(colRole).length >= 2;
    for (let ri = 0; ri < rows.length; ri++) {
      const tr = rows[ri], cells = kids(tr, 'tc');
      for (let i = 0; i < cells.length - 1; i++) {
        const label = cellText(cells[i]).trim();
        if (!label || isBlankText(label) || label.length > 700 || /_{3,}|\.{5,}|…{2,}/.test(label)) continue;
        if (twoRole || headingRole(label)) continue;
        const next = cells[i + 1];
        if (isBlankText(cellText(next)) && !hasObjects(next)) { handleTarget(doc, ctx, cells[i], next, label, false, colRole[i + 1] || null); ctx.handled.add(cells[i]); ctx.handled.add(next); i++; }
        else if (classify(label) || /[:：]\s*$/.test(label)) ctx.handled.add(cells[i]);
      }
      const nextRow = rows[ri + 1];
      if (nextRow) {
        const below = kids(nextRow, 'tc');
        const allHeaders = cells.length >= 1 && cells.every(c => !isBlankText(cellText(c)));
        const belowEmpty = below.length === cells.length && below.every(c => { const t = cellText(c).trim(); return isBlankText(t) || /^\d+[.)]?$/.test(t); }) && below.some(c => isBlankText(cellText(c)));
        if (allHeaders && belowEmpty && !cells.some(c => ctx.handled.has(c)) && !twoRole) {
          cells.forEach((hc, j) => { const tcb = below[j]; if (!isBlankText(cellText(tcb)) || ctx.handled.has(tcb)) return; handleTarget(doc, ctx, hc, tcb, cellText(hc).trim(), true, null); ctx.handled.add(tcb); });
          cells.forEach(c => ctx.handled.add(c));
        }
      }
      // остальное содержимое ячеек — абзацы и вложенные таблицы, с ролью колонки
      cells.forEach((tc, i) => {
        if (ctx.handled.has(tc)) return;
        const fp = kids(tc, 'p')[0], hr = fp ? headingRole(mapParagraph(fp).text) : null;
        const prev = ctx.role; if (hr || colRole[i]) ctx.role = hr || colRole[i];
        walkBlock(doc, tc, ctx);
        ctx.role = prev;
      });
    }
    ctx.role = saveRole;
  }

  async function fillDocx(buf, parties, docInfo, opts) {
    const zip = await JSZip.loadAsync(buf);
    if (!zip.file('word/document.xml')) throw new Error('Не найден word/document.xml — это не DOCX');
    const ctx = makeCtx(parties, docInfo, opts);
    const parts = ['word/document.xml'].concat(Object.keys(zip.files).filter(n => /^word\/(header|footer)\d*\.xml$/.test(n)));
    for (const part of parts) {
      const doc = parseXml(await zip.file(part).async('string'));
      if (part === 'word/document.xml') {
        const all = lc(byNS(doc, NS_W, 't').map(t => t.textContent).join(' '));
        const t = ctx.opts.docType;
        ctx.contract = t === 'contract' || (t === 'auto' && (/именуем[а-яёa-z0-9_]*\s+(?:в\s+дальнейшем\s+)?[«"]?(?:заказчик|исполнитель|подрядчик|поставщик|покупатель)/.test(all) || /спецификаци/.test(all.slice(0, 600)) && /договор/.test(all.slice(0, 900))));
      }
      ctx.role = null;
      walkBlock(doc, byNS(doc, NS_W, 'body')[0] || doc.documentElement, ctx);
      zip.file(part, serialize(doc));
    }
    const data = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    return { data, stats: ctx.stats, contract: ctx.contract, items: ctx.items };
  }

  /* =====================================================================
     XLSX
     ===================================================================== */
  const colToNum = s => s.split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
  const numToCol = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const splitRef = ref => { const m = /^([A-Z]+)(\d+)$/.exec(ref); return m ? { col: colToNum(m[1]), row: +m[2] } : null; };
  async function readShared(zip) { const f = zip.file('xl/sharedStrings.xml'); if (!f) return []; const doc = parseXml(await f.async('string')); return byNS(doc, NS_S, 'si').map(si => byNS(si, NS_S, 't').map(t => t.textContent).join('')); }
  function xlCellText(c, sst) { const t = c.getAttribute('t'), v = firstKid(c, 'v'); if (t === 's') return v ? (sst[+v.textContent] || '') : ''; if (t === 'inlineStr') return byNS(c, NS_S, 't').map(x => x.textContent).join(''); return v ? v.textContent : ''; }
  function setXlText(doc, c, text) {
    for (const ch of Array.from(c.childNodes)) c.removeChild(ch);
    c.setAttribute('t', 'inlineStr');
    const is = doc.createElementNS(NS_S, 'is'), t = doc.createElementNS(NS_S, 't');
    t.setAttributeNS(NS_XML, 'xml:space', 'preserve'); t.appendChild(doc.createTextNode(text)); is.appendChild(t); c.appendChild(is);
  }
  function makeHighlightStyler(stylesDoc) {
    if (!stylesDoc) return null;
    const fills = byNS(stylesDoc, NS_S, 'fills')[0], xfsEl = byNS(stylesDoc, NS_S, 'cellXfs')[0]; if (!fills || !xfsEl) return null;
    let fillId = null; const cache = {};
    return function (oldS) {
      oldS = oldS || 0; if (cache[oldS] != null) return cache[oldS];
      if (fillId == null) {
        const fill = stylesDoc.createElementNS(NS_S, 'fill'), pf = stylesDoc.createElementNS(NS_S, 'patternFill'); pf.setAttribute('patternType', 'solid');
        const fg = stylesDoc.createElementNS(NS_S, 'fgColor'); fg.setAttribute('rgb', 'FFFFFF00'); const bg = stylesDoc.createElementNS(NS_S, 'bgColor'); bg.setAttribute('indexed', '64');
        pf.appendChild(fg); pf.appendChild(bg); fill.appendChild(pf); fills.appendChild(fill); fillId = kids(fills, 'fill').length - 1; fills.setAttribute('count', String(fillId + 1));
      }
      const xfs = kids(xfsEl, 'xf'), base = xfs[oldS] || xfs[0], xf = base ? base.cloneNode(true) : stylesDoc.createElementNS(NS_S, 'xf');
      xf.setAttribute('fillId', String(fillId)); xf.setAttribute('applyFill', '1'); xfsEl.appendChild(xf);
      const idx = kids(xfsEl, 'xf').length - 1; xfsEl.setAttribute('count', String(idx + 1)); return (cache[oldS] = idx);
    };
  }
  function processSheet(doc, sst, ctx, styler) {
    const merges = {};
    for (const m of byNS(doc, NS_S, 'mergeCell')) { const [a, b] = (m.getAttribute('ref') || '').split(':'); const A = splitRef(a), B = b && splitRef(b); if (A && B) merges[a] = B.col; }
    const colRole = {};
    for (const row of byNS(doc, NS_S, 'row')) {
      const rn = +row.getAttribute('r'), cellEls = kids(row, 'c'), byCol = {};
      cellEls.forEach(c => { const ref = splitRef(c.getAttribute('r') || ''); if (ref) byCol[ref.col] = c; });
      for (const c of cellEls) {
        const ref = splitRef(c.getAttribute('r') || ''); if (!ref) continue;
        const label = xlCellText(c, sst).trim();
        if (!label || !/[A-Za-zА-Яа-яЁё]/.test(label) || label.length > 700) continue;
        const hr = headingRole(label); if (hr) { colRole[ref.col] = hr; continue; }
        const tcol = (merges[c.getAttribute('r')] || ref.col) + 1;
        const cellRole = colRole[tcol] || colRole[ref.col] || null;
        const hlCell = (cell, what, role) => { if (ctx.opts.highlight && styler) { cell.setAttribute('s', String(styler(+cell.getAttribute('s')))); ctx.hl(what, role); } };
        const bm = /_{3,}\s*$/.exec(label);
        if (bm) {
          const r = classify(label.slice(0, bm.index), { tail: true });
          if (r && !r.skip && r.items.length && !r.boss) {
            const items = r.items.filter(k => ctx.allowed(k)), role = pickRole(ctx, r.role, cellRole); if (!items.length || !role) continue;
            const out = ctx.render(items, role);
            if (out.text) { setXlText(doc, c, label.slice(0, bm.index).replace(/\s*$/, ' ') + out.text); ctx.filled(items.map(titleOf).join(' + '), out.text, role); } else hlCell(c, items.map(titleOf).join(' + '), role);
          }
          continue;
        }
        let tgt = byCol[tcol];
        if (tgt && !isBlankText(xlCellText(tgt, sst))) continue;
        const lab = lc(label);
        let items = null, explicit = null;
        if (AMOUNT_LABEL.test(lab)) items = ['amount'];
        else { const r = classify(label); if (r && !r.skip && r.items.length) { items = r.items; explicit = r.role; } }
        const ensure = () => {
          if (!tgt) { tgt = doc.createElementNS(NS_S, 'c'); tgt.setAttribute('r', numToCol(tcol) + rn); const after = cellEls.find(x => { const q = splitRef(x.getAttribute('r') || ''); return q && q.col > tcol; }); if (after) row.insertBefore(tgt, after); else row.appendChild(tgt); byCol[tcol] = tgt; }
          return tgt;
        };
        if (items) {
          items = items.filter(k => ctx.allowed(k)); if (!items.length) continue;
          const role = pickRole(ctx, explicit, cellRole);
          const needsRole = items.some(k => !['doc_date', 'doc_no', 'spec_no', 'city', 'amount', 'amount_num', 'amount_words', 'vat_amount'].includes(k));
          if (needsRole && !role) continue;
          const out = ctx.render(items, role);
          if (out.text) { setXlText(doc, ensure(), out.text); ctx.filled(items.map(titleOf).join(' + '), out.text, needsRole ? role : null); }
          else if (tgt) hlCell(tgt, items.map(titleOf).join(' + '), needsRole ? role : null);
        } else if (ctx.opts.mode !== 'min' && tgt && looksLikeField(label)) hlCell(tgt, clean(label).slice(0, 50), null);
      }
    }
  }
  async function fillXlsx(buf, parties, docInfo, opts) {
    const zip = await JSZip.loadAsync(buf), ctx = makeCtx(parties, docInfo, opts), sst = await readShared(zip);
    const sf = zip.file('xl/styles.xml'), stylesDoc = sf ? parseXml(await sf.async('string')) : null, styler = makeHighlightStyler(stylesDoc);
    const sheets = Object.keys(zip.files).filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
    if (!sheets.length) throw new Error('В файле нет листов');
    for (const name of sheets) { const doc = parseXml(await zip.file(name).async('string')); ctx.role = null; processSheet(doc, sst, ctx, styler); zip.file(name, serialize(doc)); }
    if (stylesDoc) zip.file('xl/styles.xml', serialize(stylesDoc));
    return { data: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), stats: ctx.stats, items: ctx.items };
  }

  /* ---------- чтение карты: DOCX / XLSX → текст ---------- */
  async function docxToText(buf) {
    const zip = await JSZip.loadAsync(buf), f = zip.file('word/document.xml'); if (!f) throw new Error('Не найден word/document.xml');
    const doc = parseXml(await f.async('string')), body = byNS(doc, NS_W, 'body')[0], out = [];
    const ptext = p => byNS(p, NS_W, 't').map(t => t.textContent).join('');
    (function walk(node) {
      for (let c = node.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue; const n = lname(c);
        if (n === 'p') { const t = ptext(c).trim(); if (t) out.push(t); }
        else if (n === 'tbl') for (const tr of kids(c, 'tr')) { const cells = kids(tr, 'tc').map(tc => kids(tc, 'p').map(ptext).join(' ').replace(/\s+/g, ' ').trim()); if (cells.some(Boolean)) out.push(cells.join(' | ')); }
        else if (n === 'sdt' || n === 'sdtContent') walk(c);
      }
    })(body);
    return out.join('\n');
  }
  async function xlsxToText(buf) {
    const zip = await JSZip.loadAsync(buf), sst = await readShared(zip), out = [];
    for (const name of Object.keys(zip.files).filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))) {
      const doc = parseXml(await zip.file(name).async('string'));
      for (const row of byNS(doc, NS_S, 'row')) { const cells = kids(row, 'c').map(c => xlCellText(c, sst).trim()).filter(Boolean); if (cells.length) out.push(cells.join(' | ')); }
    }
    return out.join('\n');
  }
  async function fillFile(name, buf, parties, docInfo, opts) {
    const ext = (name.split('.').pop() || '').toLowerCase();
    if (ext === 'docx' || ext === 'dotx' || ext === 'docm') return fillDocx(buf, parties, docInfo, opts);
    if (ext === 'xlsx' || ext === 'xlsm') return fillXlsx(buf, parties, docInfo, opts);
    const e = new Error('Формат .' + ext + ' не поддерживается. Сохраните бланк как .docx или .xlsx'); e.unsupported = true; throw e;
  }

  return { parseCardText, finishCard, fillDocx, fillXlsx, fillFile, docxToText, xlsxToText, classify, shortFio, fioGen, posGen, rubWords, fmtMoney, FIELD_TITLES };
});
