import { parseMoneyToNumber, parseIntSafe } from './utils.js';
const MONTHS = {
    ENE: 1, FEB: 2, MAR: 3, ABR: 4, MAY: 5, JUN: 6,
    JUL: 7, AGO: 8, SEP: 9, OCT: 10, NOV: 11, DIC: 12,
};

function normalizeText(text) {
    return String(text || '')
            .replace(/\r/g, '\n')
            .replace(/[ \t]+/g, ' ')
            .replace(/\n{3,}/g, '\n\n')
            .replace(/[“”]/g, '"')
            .replace(/[’]/g, "'")
            .trim();
}

function normalizeUpper(text) {
    return normalizeText(text)
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .toUpperCase();
}

function buildLines(text) {
    return normalizeText(text)
            .split('\n')
            .map(s => s.trim())
            .filter(Boolean);
}
function firstPositive(...values) {
    for (const v of values) {
        const n = Number(v || 0);
        if (Number.isFinite(n) && n > 0)
            return n;
    }
    return 0;
}

function cleanLine(line) {
    return String(line || '')
            .replace(/\s+/g, ' ')
            .replace(/[|]/g, 'I')
            .trim();
}

function isStopLabel(line) {
    const up = normalizeUpper(line);
    return [
        'NO DE SERVICIO',
        'NO. DE SERVICIO',
        'CUENTA:',
        'LIMITE DE PAGO',
        'CORTE A PARTIR',
        'TARIFA:',
        'NO HILOS',
        'PERIODO FACTURADO',
        'LECTURA ACTUAL',
        'LECTURA ANTERIOR',
        'CONCEPTO',
        'DESGLOSE DEL IMPORTE',
        'SUBTOTAL',
        'CONSUMO HISTORICO',
        'DESCARGA NUESTRA APP',
        'APP AUTORIZADA'
    ].some(x => up.includes(x));
}

function looksLikeAddressLine(line) {
    const raw = cleanLine(line);
    const up = normalizeUpper(raw);

    if (!raw || raw.length < 6)
        return false;
    if (isStopLabel(raw))
        return false;
    if (/\$\s*\d/.test(raw))
        return false;
    if (up.includes('TOTAL A PAGAR'))
        return false;
    if (up.includes('ESTE GRAFICO'))
        return false;

    return true;
}
function parsePeriodo(text) {
    const upper = normalizeUpper(text);

    const m = upper.match(
            /PERIODO\s+FACTURADO[:\s]*([0-3]?\d)\s*(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)\s*(\d{2})\s*[-–]\s*([0-3]?\d)\s*(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)\s*(\d{2})/
            );

    if (!m)
        return {raw: '', start: null, end: null, days: 0};

    const [, sd, sm, sy, ed, em, ey] = m;
    const start = new Date(2000 + parseIntSafe(sy), (MONTHS[sm] || 1) - 1, parseIntSafe(sd));
    const end = new Date(2000 + parseIntSafe(ey), (MONTHS[em] || 1) - 1, parseIntSafe(ed));
    const days = Math.max(0, Math.round((end - start) / (1000 * 60 * 60 * 24)));

    return {
        raw: `${sd} ${sm} ${sy} - ${ed} ${em} ${ey}`,
        start,
        end,
        days,
    };
}

function parseFechaCortaLabel(rawPeriodo) {
    if (!rawPeriodo)
        return '';
    const m = rawPeriodo.match(/(\d{2})\s+([A-Z]{3})\s+(\d{2})\s+-\s+(\d{2})\s+([A-Z]{3})\s+(\d{2})/i);
    if (!m)
        return rawPeriodo;
    return `${m[1]} ${m[2]} ${m[3]} - ${m[4]} ${m[5]} ${m[6]}`;
}
function canonicalPeriodLabel(value) {
    return normalizeUpper(value)
            .replace(/\bDEL\b/g, ' ')
            .replace(/\bAL\b/g, ' ')
            .replace(/[-–]/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();
}

function parseNombreDireccion(text) {
    const lines = buildLines(text);
    const upperLines = lines.map(l => normalizeUpper(l));

    const blacklist = [
        'COMISION FEDERAL DE ELECTRICIDAD',
        'COMISION FEDERAL',
        'RFC:',
        'AV. PASEO DE LA REFORMA',
        'ALCALDIA',
        'CIUDAD DE MEXICO',
        'TOTAL A PAGAR',
        'LECTURA ACTUAL',
        'LECTURA ANTERIOR',
        'CONCEPTO',
        'PERIODO FACTURADO',
        'NO DE SERVICIO',
        'NO. DE SERVICIO',
        'CUENTA:',
        'TARIFA:',
        'DESCARGA NUESTRA APP',
        'APP AUTORIZADA'
    ];

    let nombre = '';
    let nombreIndex = -1;

    for (let i = 0; i < lines.length; i++) {
        const raw = cleanLine(lines[i]);
        const up = upperLines[i];

        if (isStopLabel(raw))
            break;
        if (blacklist.some(x => up.includes(x)))
            continue;
        if (/\d/.test(up))
            continue;

        const words = up.split(/\s+/).filter(Boolean);
        if (words.length >= 3 && words.length <= 8) {
            nombre = raw
                    .replace(/\$.*$/, '')
                    .replace(/\bPAS\b.*$/i, '')
                    .trim();
            nombreIndex = i;
            break;
        }
    }

    const addressLines = [];

    if (nombreIndex >= 0) {
        for (let i = nombreIndex + 1; i < Math.min(lines.length, nombreIndex + 6); i++) {
            const raw = cleanLine(lines[i]);
            if (!raw)
                continue;
            if (isStopLabel(raw))
                break;
            if (!looksLikeAddressLine(raw))
                continue;

            const cleaned = raw.replace(/\$.*$/, '').trim();
            if (cleaned)
                addressLines.push(cleaned);
        }
    }

    return {
        nombre,
        direccion: addressLines.join(' ').replace(/\s+/g, ' ').trim()
    };
}

function parseServicio(text) {
    const upper = normalizeUpper(text);
    return upper.match(/NO\.?\s*DE\s*SERVICIO[:\s]*([0-9]{8,15})/)?.[1] || '';
}

function parseTarifa(text) {
    const lines = buildLines(text);
    const allowed = ['1', '1A', '1B', '1C', '1D', '1E', '1F', 'DAC', 'PDBT', 'GDBT', 'GDMTO', 'GDMTH'];

    for (const rawLine of lines) {
        const line = normalizeUpper(rawLine);
        if (!line.includes('TARIFA'))
            continue;

        const m = line.match(/TARIFA[:\s]*\b(1[A-F]?|DAC|PDBT|GDBT|GDMTO|GDMTH)\b/);
        if (m && allowed.includes(m[1])) {
            return m[1];
        }
    }

    return '';
}

function parseHilos(text) {
    const upper = normalizeUpper(text);
    return upper.match(/NO\s+HILOS[:\s]*([0-9]+)/)?.[1] || '';
}

function parseMedidor(text) {
    const upper = normalizeUpper(text);
    return upper.match(/NO\.?\s*MEDIDOR[:\s]*([A-Z0-9-]+)/)?.[1] || '';
}

function parseLimitePago(text) {
    const upper = normalizeUpper(text);
    return upper.match(/LIMITE\s+DE\s+PAGO[:\s]*([0-3]?\d\s+[A-Z]{3}\s+\d{2})/)?.[1] || '';
}

function parseCorte(text) {
    const upper = normalizeUpper(text);
    return upper.match(/CORTE\s+A\s+PARTIR[:\s]*([0-3]?\d\s+[A-Z]{3}\s+\d{2})/)?.[1] || '';
}

function parseTotalAPagar(text) {
    const lines = buildLines(text);

    for (const rawLine of lines) {
        const line = cleanLine(rawLine);
        const up = normalizeUpper(line);

        if (up.includes('TOTAL A PAGAR')) {
            const m = line.match(/\$?\s*([\d,]+(?:\.\d{1,2})?)/);
            if (m) {
                const n = parseMoneyToNumber(m[1]);
                if (n > 0)
                    return n;
            }
        }
    }

    const full = normalizeText(text);

    let m = full.match(/DESGLOSE\s+DEL\s+IMPORTE\s+A\s+PAGAR[\s\S]{0,250}?\bTotal\b[^\d$]*\$?\s*([\d,]+(?:\.\d{1,2})?)/i);
    if (m) {
        const n = parseMoneyToNumber(m[1]);
        if (n > 0)
            return n;
    }

    m = full.match(/\bTotal\b[^\d$]*\$?\s*([\d,]+(?:\.\d{1,2})?)/i);
    if (m) {
        const n = parseMoneyToNumber(m[1]);
        if (n > 0)
            return n;
    }

    m = full.match(/TOTAL\s+A\s+PAGAR[\s\S]{0,80}?\$?\s*([\d,]+(?:\.\d{1,2})?)/i);
    if (m) {
        const n = parseMoneyToNumber(m[1]);
        if (n > 0)
            return n;
    }

    return 0;
}

function parseBaseCosts(text) {
    const upper = normalizeUpper(text);

    const suministro = (() => {
        const m = upper.match(/SUMINISTRO\s+([\d,.]+)/);
        return m ? parseMoneyToNumber(m[1]) : 0;
    })();

    const ivaPct = (() => {
        const m = upper.match(/IVA\s+(\d+)%/);
        return m ? parseIntSafe(m[1]) : 16;
    })();

    const dap = (() => {
        const m = upper.match(/DAP(?:\(\d+\))?\s+([\d,.]+)/);
        return m ? parseMoneyToNumber(m[1]) : 0;
    })();

    const costoBase = (suministro * (1 + ivaPct / 100)) + dap;

    return {suministro, iva: ivaPct, dap, costoBase};
}

function parseConsumoPeriodo(text, historial = [], periodoRaw = '') {
    const lines = buildLines(text);

    for (const rawLine of lines) {
        const line = cleanLine(rawLine);
        const up = normalizeUpper(line);

        if (up.includes('ENERGIA (KWH)') || up.includes('ENERGIA KWH')) {
            const nums = (line.match(/\d{1,3}(?:,\d{3})*/g) || [])
                    .map(v => parseIntSafe(v))
                    .filter(n => n > 0);

            if (nums.length >= 3) {
                return nums[2];
            }
        }
    }

    const upText = normalizeUpper(text);
    let m = upText.match(/ENERGIA\s*\(KWH\)[\s\S]{0,80}?(\d{1,3}(?:,\d{3})*)\s+(\d{1,3}(?:,\d{3})*)\s+(\d{1,3}(?:,\d{3})*)/);
    if (m) {
        return parseIntSafe(m[3]);
    }

    if (Array.isArray(historial) && historial.length) {
        const first = historial.find(x => Number(x?.kwh || 0) > 0);
        if (first)
            return Number(first.kwh);
    }

    return 0;
}

function parseEstadoFromDireccion(direccion) {
    if (!direccion)
        return '';
    const m = direccion.match(/,\s*([A-Za-zÁÉÍÓÚÑáéíóúñ. ]{2,})$/);
    return m ? m[1].replace(/\.+$/, '').trim().toUpperCase() : '';
}

function parseHistoricalTable(text) {
    const lines = buildLines(text);
    const rows = [];
    const month = '(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)';

    for (const rawLine of lines) {
        const line = normalizeUpper(cleanLine(rawLine))
                .replace(/\$/g, ' ')
                .replace(/\s+/g, ' ')
                .trim();
        const m = line.match(new RegExp(
                `(?:DEL\\s+)?(\\d{2})\\s+${month}\\s+(\\d{2})\\s+(?:AL\\s+)?(\\d{2})\\s+${month}\\s+(\\d{2}).*?(\\d{2,5})\\s+([\\d,.]+)\\s+([\\d,.]+)(?:\\s+([\\d,.]+))?`
                ));

        if (!m)
            continue;

        const periodo = `del ${m[1]} ${m[2]} ${m[3]} al ${m[4]} ${m[5]} ${m[6]}`;
        const kwh = parseIntSafe(m[7]);
        const importe = parseMoneyToNumber(m[8]);
        const pago = parseMoneyToNumber(m[9]);
        const pendiente = parseMoneyToNumber(m[10] || '0');

        if (kwh > 0) {
            rows.push({periodo, kwh, importe, pago, pendiente});
        }
    }

    return rows;
}

function chooseBestText(primary, secondary) {
    const p = normalizeText(primary);
    const s = normalizeText(secondary);

    const score = (txt) => {
        const up = normalizeUpper(txt);
        let n = 0;
        if (up.includes('NO DE SERVICIO') || up.includes('NO. DE SERVICIO'))
            n += 3;
        if (up.includes('TARIFA'))
            n += 2;
        if (up.includes('PERIODO FACTURADO'))
            n += 3;
        if (up.includes('CONSUMO HISTORICO'))
            n += 2;
        if (up.includes('TOTAL A PAGAR'))
            n += 2;
        if (up.includes('NO HILOS'))
            n += 1;
        return n;
    };

    return score(s) > score(p) ? s : p;
}

function looksCorruptedText(text) {
    const up = normalizeUpper(text);
    if (!up)
        return true;

    const hasKeyFields =
            up.includes('NO DE SERVICIO') ||
            up.includes('NO. DE SERVICIO') ||
            up.includes('PERIODO FACTURADO') ||
            up.includes('TOTAL A PAGAR');

    if (!hasKeyFields)
        return true;

    const strange = (text.match(/[�]/g) || []).length;
    return strange > 5;
}

function parseCfeReceiptTextFromPages(page1Text, page2Text, mergedText) {
    const bestPage1 = chooseBestText(page1Text, mergedText);
    const bestPage2 = chooseBestText(page2Text, mergedText);
    const fullText = normalizeText([bestPage1, bestPage2, mergedText].filter(Boolean).join('\n\n'));

    const page1 = normalizeText(bestPage1 || '');
    const page2 = normalizeText(bestPage2 || '');

    const tarifa = parseTarifa(page1);
    const servicio = parseServicio(page1);
    const periodo = parsePeriodo(page1);
    const {nombre, direccion} = parseNombreDireccion(page1);
    const historial = parseHistoricalTable(page2);
    const consumoPeriodo = parseConsumoPeriodo(page1, historial, periodo?.raw || '');
    const total = firstPositive(
            parseTotalAPagar(page1),
            parseTotalAPagar(fullText)
            );
    const {suministro, iva, dap, costoBase} = parseBaseCosts(fullText);
    const hilos = parseHilos(bestPage1 || fullText);
    const medidor = parseMedidor(bestPage1 || fullText);
    const limitePago = parseLimitePago(bestPage1 || fullText);
    const corteAPartir = parseCorte(bestPage1 || fullText);

    const pagosHist = historial.map(x => Number(x.pago || 0)).filter(n => n > 0);
    const pagoProm = pagosHist.length
            ? pagosHist.reduce((a, b) => a + b, 0) / pagosHist.length
            : total;

    const ahorroEstimado = Math.max(0, pagoProm - costoBase);
    const tipoPeriodo = (periodo?.days ?? 0) >= 45 ? 'Bimestral' : 'Mensual';
    const estado = parseEstadoFromDireccion(direccion);

    return {
        fuente: 'CFE',
        tarifa,
        servicio,
        totalAPagar: total,
        periodo,
        tipoPeriodo,
        nombre: nombre || '',
        direccion: direccion || '',
        consumoPeriodo: consumoPeriodo || 0,
        historial,
        suministro,
        iva,
        dap,
        costoBase,
        ahorroEstimado,
        hilos,
        medidor,
        limitePago,
        corteAPartir,
        estado,
        rawText: fullText,
    };
}

function groupPdfItemsToLines(items) {
    const rows = items
            .map(it => {
                const tr = it.transform || [];
                return {
                    str: (it.str || '').trim(),
                    x: Number(tr[4] || 0),
                    y: Number(tr[5] || 0),
                };
            })
            .filter(it => it.str);

    rows.sort((a, b) => {
        if (Math.abs(b.y - a.y) > 2)
            return b.y - a.y;
        return a.x - b.x;
    });

    const lines = [];
    for (const item of rows) {
        const last = lines[lines.length - 1];
        if (!last || Math.abs(last.y - item.y) > 3) {
            lines.push({y: item.y, items: [item]});
        } else {
            last.items.push(item);
        }
    }

    return lines.map(line =>
        line.items
                .sort((a, b) => a.x - b.x)
                .map(i => i.str)
                .join(' ')
                .replace(/\s{2,}/g, ' ')
                .trim()
    ).filter(Boolean);
}

async function renderPageToCanvas(page, scale = 1.8) {
    const viewport = page.getViewport({scale});
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d', {alpha: false});

    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);

    await page.render({
        canvasContext: ctx,
        viewport,
    }).promise;

    return canvas;
}

async function extractPdfTextByPage(pdf, onProgress) {
    const pages = [];

    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const content = await page.getTextContent();
        const lines = groupPdfItemsToLines(content.items || []);
        const pageText = lines.join('\n').trim();

        pages.push(pageText);

        if (onProgress) {
            onProgress({message: `Leyendo texto PDF… (${i}/${pdf.numPages})`});
        }
    }

    return pages;
}

async function runOcrOnCanvas(canvas, pageIndex, onProgress) {
    if (!window.Tesseract)
        return '';

    const psm = pageIndex === 1 ? '6' : '11';

    const result = await window.Tesseract.recognize(canvas, 'spa+eng', {
        logger: (msg) => {
            if (onProgress && msg?.status === 'recognizing text') {
                const pct = Math.round((msg.progress || 0) * 100);
                onProgress({message: `OCR página ${pageIndex}: ${pct}%`});
            }
        },
        tessedit_pageseg_mode: psm,
        preserve_interword_spaces: '1',
    });

    return normalizeText(result?.data?.text || '');
}

async function pdfToTextAndPreview(file, { onProgress } = {}){
    const pdfjsLib = window.pdfjsLib;
    if (!pdfjsLib)
        return {text: '', pageTexts: [], canvas: null};

    pdfjsLib.GlobalWorkerOptions.workerSrc =
            pdfjsLib.GlobalWorkerOptions.workerSrc ||
            'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({data: arrayBuffer}).promise;

    const pageTexts = await extractPdfTextByPage(pdf, onProgress);

    const firstPage = await pdf.getPage(1);
    const previewCanvas = await renderPageToCanvas(firstPage, 0.8);

    let mergedText = pageTexts.join('\n\n');

    const missingHistory = !normalizeUpper(mergedText).includes('CONSUMO HISTORICO');
    const poorText = looksCorruptedText(mergedText) || missingHistory;

    if (poorText && window.Tesseract) {
        const ocrTexts = [];

        const pagesToOcr = Math.min(pdf.numPages, 2);
        for (let i = 1; i <= pagesToOcr; i++) {
            const page = await pdf.getPage(i);
            const canvas = await renderPageToCanvas(page, 2.0);
            const ocrText = await runOcrOnCanvas(canvas, i, onProgress);
            ocrTexts.push(ocrText);
        }

        const combined = ocrTexts.join('\n\n');
        mergedText = normalizeText([mergedText, combined].filter(Boolean).join('\n\n'));

        for (let i = 0; i < ocrTexts.length; i++) {
            if ((ocrTexts[i] || '').length > (pageTexts[i] || '').length * 0.6) {
                pageTexts[i] = chooseBestText(pageTexts[i], ocrTexts[i]);
            }
        }
    }

    return {
        text: mergedText,
        pageTexts,
        canvas: previewCanvas,
    };
}

function seedFromString(str) {
    let h = 2166136261;
    for (let i = 0; i < str.length; i++) {
        h ^= str.charCodeAt(i);
        h = Math.imul(h, 16777619);
    }
    return (h >>> 0);
}

function mulberry32(a) {
    return function () {
        let t = a += 0x6D2B79F5;
        t = Math.imul(t ^ (t >>> 15), t | 1);
        t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
        return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
}

function fmtPeriodo(start, end) {
    const m = ['ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SEP', 'OCT', 'NOV', 'DIC'];
    const d2 = (n) => String(n).padStart(2, '0');
    const sd = d2(start.getDate());
    const sm = m[start.getMonth()];
    const sy = String(start.getFullYear()).slice(-2);
    const ed = d2(end.getDate());
    const em = m[end.getMonth()];
    const ey = String(end.getFullYear()).slice(-2);
    return `${sd} ${sm} ${sy} - ${ed} ${em} ${ey}`;
}

function buildMockReceipt( { selectedTariff, file }){
    const seed = seedFromString((file?.name || 'recibo') + '|' + (selectedTariff?.key || ''));
    const rnd = mulberry32(seed);

    const now = new Date();
    const isBim = selectedTariff?.periodo === 'Bimestral';
    const end = new Date(now.getFullYear(), now.getMonth(), Math.max(1, Math.min(28, 10 + Math.floor(rnd() * 15))));
    const start = new Date(end);
    start.setDate(end.getDate() - (isBim ? 60 : 30));

    const consumo = Math.round(220 + rnd() * 980);
    const total = Math.round(600 + rnd() * 4200);
    const servicio = String(300000000000 + Math.floor(rnd() * 900000000000));
    const hilos = String(1 + Math.floor(rnd() * 3));
    const estados = ['SON', 'BC', 'BCS', 'CHIH', 'NL', 'JAL', 'QRO', 'CDMX', 'MEX', 'GTO', 'SIN', 'PUE'];
    const estado = estados[Math.floor(rnd() * estados.length)];

    const historial = Array.from({length: 12}).map(() => {
        const kwh = Math.max(50, Math.round(consumo * (0.75 + rnd() * 0.55)));
        const pago = Math.max(100, Math.round(total * (0.75 + rnd() * 0.55)));
        return {kwh, pago, importe: pago, pendiente: 0};
    });

    const costoBase = Math.max(120, Math.round(total * 0.35));
    const pagoProm = historial.reduce((a, b) => a + b.pago, 0) / historial.length;
    const ahorroEstimado = Math.max(0, Math.round(pagoProm - costoBase));

    const tarifaDetectada = (() => {
        if (selectedTariff?.familia === 'Doméstica') {
            const opts = ['1', '1A', '1B', '1C', '1D', '1E', '1F', 'DAC'];
            return opts[Math.floor(rnd() * opts.length)];
        }
        if (selectedTariff?.familia === 'PDBT')
            return 'PDBT';
        if (selectedTariff?.familia === 'GDMTH')
            return 'GDMTH';
        if (selectedTariff?.familia === 'GDMTO')
            return 'GDMTO';
        return selectedTariff?.label || '';
    })();

    return {
        fuente: 'CFE',
        tarifa: tarifaDetectada,
        servicio,
        totalAPagar: total,
        periodo: {raw: fmtPeriodo(start, end), start, end, days: isBim ? 60 : 30},
        tipoPeriodo: isBim ? 'Bimestral' : 'Mensual',
        nombre: 'CLIENTE',
        direccion: 'Dirección del servicio',
        consumoPeriodo: consumo,
        historial,
        suministro: Math.round(costoBase * 0.65),
        iva: 16,
        dap: Math.round(costoBase * 0.15),
        costoBase,
        ahorroEstimado,
        hilos,
        estado,
        rawText: '',
    };
}

export function parseCfeReceiptText(text) {
    return parseCfeReceiptTextFromPages(text, text, text);
}

export async function analyzeReceiptFile(file, options = {}){
    const {selectedTariff = null, onProgress} = options;
    const isPdf = file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf');

    let text = '';
    let canvas = null;
    let pageTexts = [];

    if (onProgress)
        onProgress({message: 'Preparando lectura…'});

    try {
        if (isPdf) {
            const result = await pdfToTextAndPreview(file, {onProgress});
            text = result.text || '';
            pageTexts = Array.isArray(result.pageTexts) ? result.pageTexts : [];
            canvas = result.canvas || null;
        }
    } catch (err) {
        console.error('Error leyendo PDF:', err);
    }

    let parsed = null;

    if ((text || '').replace(/\s/g, '').length > 60) {
        if (onProgress)
            onProgress({message: 'Extrayendo datos del recibo…'});

        parsed = parseCfeReceiptTextFromPages(
                pageTexts[0] || text,
                pageTexts[1] || text,
                text
                );
    } else {
        if (onProgress)
            onProgress({message: 'No se pudo leer el PDF. Se habilitará captura manual.'});
        parsed = {
            fuente: 'CFE',
            tarifa: '',
            servicio: '',
            totalAPagar: 0,
            periodo: {raw: '', start: null, end: null, days: 0},
            tipoPeriodo: selectedTariff?.periodo || '',
            nombre: '',
            direccion: '',
            consumoPeriodo: 0,
            historial: [],
            suministro: 0,
            iva: 16,
            dap: 0,
            costoBase: 0,
            ahorroEstimado: 0,
            hilos: '',
            estado: '',
            rawText: text || '',
            instalacion: {},
            insumos: [],
            impuestosPct: 0.16
        };
    }

    if (selectedTariff?.periodo) {
        parsed.tipoPeriodo = selectedTariff.periodo;
        parsed.periodo = parsed.periodo || {raw: '', start: null, end: null, days: 0};
        parsed.periodo.days = selectedTariff.periodo === 'Bimestral' ? 60 : 30;
    }

    parsed.instalacion = parsed.instalacion || {};
    parsed.insumos = Array.isArray(parsed.insumos) ? parsed.insumos : [];
    parsed.impuestosPct = Number.isFinite(Number(parsed.impuestosPct)) ? Number(parsed.impuestosPct) : 0.16;

    if (!parsed.estado && parsed.direccion) {
        parsed.estado = parseEstadoFromDireccion(parsed.direccion);
    }

    if (!parsed.periodo?.raw && parsed.periodo) {
        parsed.periodo.raw = parseFechaCortaLabel(parsed.periodo.raw || '');
    }

    if (onProgress)
        onProgress({message: 'Listo.'});

    return {text, parsed, canvas};
}