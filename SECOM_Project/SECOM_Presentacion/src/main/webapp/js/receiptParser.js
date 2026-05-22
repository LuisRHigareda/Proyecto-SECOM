import { loadUserPreferences } from './preferences.js';
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
        if (Number.isFinite(n) && n > 0) {
            return n;
        }
    }
    return 0;
}

function cleanLine(line) {
    return String(line || '')
            .replace(/\s+/g, ' ')
            .replace(/[|]/g, 'I')
            .trim();
}

function removeAccentsUpper(line) {
    return String(line || '')
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .toUpperCase();
}

function stripRightColumnNoise(line) {
    return cleanLine(line)
            .replace(/\bTOTAL\s+A\s+PAGAR\b[\s\S]*$/i, '')
            .replace(/\b(?:DESCARGA|APP\s+AUTORIZADA|PAGO\s+EN\s+LINEA)\b[\s\S]*$/i, '')
            .replace(/\s{2,}/g, ' ')
            .trim();
}

function isMoneyTextLine(line) {
    const up = removeAccentsUpper(line);

    return (
        /\$\s*\d/.test(line || '') ||
        up.includes('PESOS M.N') ||
        up.includes('PESOS MN') ||
        up.includes('TOTAL A PAGAR') ||
        up.includes('CUATROCIENTOS') ||
        up.includes('QUINIENTOS') ||
        up.includes('SEISCIENTOS') ||
        up.includes('SETECIENTOS') ||
        up.includes('OCHOCIENTOS') ||
        up.includes('NOVECIENTOS') ||
        up.includes('MIL PESOS') ||
        up.includes('PESOS 00/100')
    );
}

function hasLongMixedIdentifier(line) {
    const raw = String(line || '');
    const up = removeAccentsUpper(raw);

    if (ADDRESS_HINTS && ADDRESS_HINTS.some(h => up.includes(h))) {
        return false;
    }
    if (/\b(CP|C\.P\.|CD|CIUDAD|COL|FRACC|BLVD|AV|CALLE)\b/.test(up)) {
        return false;
    }

    const tokens = raw.match(/[A-Z0-9]{10,}/gi) || [];

    return tokens.some(t => /[A-Z]/i.test(t) && /\d/.test(t));
}

const ADDRESS_HINTS = [
    'CALLE', 'AV', 'AV.', 'AVENIDA', 'BLVD', 'BOULEVARD', 'COL', 'COLONIA',
    'FRACC', 'FRACC.', 'SECCION', 'SECC', 'MANZANA', 'MZ', 'LOTE', 'LT', 'NUM',
    'NO', 'N°', 'CP', 'C.P.', 'INT', 'EXT', 'DEPTO', 'PRIV', 'PASEO', 'CARRETERA',
    'KM', 'RESIDENCIAL', 'VILLA', 'SECTOR', 'LOCALIDAD', 'MUNICIPIO'
];

function isLikelyCfeCustomerName(line) {
    const raw = stripRightColumnNoise(line);
    const up = removeAccentsUpper(raw);

    if (!looksLikeNameCandidate(raw)) {
        return false;
    }
    if (isMoneyTextLine(raw) || hasLongMixedIdentifier(raw)) {
        return false;
    }
    if (ADDRESS_HINTS.some(h => up.includes(h))) {
        return false;
    }
    if (/,/.test(raw) || /\b(CD|CIUDAD|OBREGON|SON|SONORA|HERMOSILLO|NAVOJOA|CAJEME|GUAYMAS|NOGALES)\b/.test(up)) {
        return false;
    }
    if (/\b(CFE|COMISION|FEDERAL|ELECTRICIDAD|SUCURSAL|CORREOS|AVISO|PAGO|PAGAR|CLIENTE|TARIFA|SERVICIO|RMU|CUENTA|REPARTIR|MENSUAL|BIMESTRAL)\b/.test(up)) {
        return false;
    }

    const letters = (raw.match(/[A-Za-zÁÉÍÓÚÑáéíóúñ]/g) || []).length;
    const digits = (raw.match(/\d/g) || []).length;

    return letters >= 8 && digits === 0;
}

function isLikelyCfeAddressLine(line) {
    const raw = stripRightColumnNoise(line);
    const up = removeAccentsUpper(raw);

    if (!raw || raw.length < 4) {
        return false;
    }
    if (isMoneyTextLine(raw) || isStopLabel(raw) || hasLongMixedIdentifier(raw)) {
        return false;
    }
    if (/\b(RMU|CUENTA|NO\.?\s*DE\s*SERVICIO|LIMITE|CORTE|TOTAL|PAGAR|COMISION|FEDERAL|ELECTRICIDAD|DESCARGA|APP|AUTORIZADA|REPARTIR)\b/.test(up)) {
        return false;
    }
    if (!/\s/.test(raw) && /^[A-Z0-9]{12,}/i.test(raw.replace(/\s+/g, ''))) {
        return false;
    }

    const hasDigit = /\d/.test(raw);
    const hasHint = ADDRESS_HINTS.some(h => up.includes(h));
    const hasCityState = /\b(CD|CIUDAD|OBREGON|SON|SONORA|HERMOSILLO|NAVOJOA|CAJEME|GUAYMAS|NOGALES|MEXICO|MEX)\b/.test(up);
    const enoughLetters = (raw.match(/[A-Za-zÁÉÍÓÚÑáéíóúñ]/g) || []).length >= 4;

    return enoughLetters && (hasDigit || hasHint || hasCityState || /,/.test(raw));
}

function normalizeServiceCandidate(raw) {
    return String(raw || '')
            .toUpperCase()
            .replace(/[OQ]/g, '0')
            .replace(/[IL]/g, '1')
            .replace(/S/g, '5')
            .replace(/B/g, '8')
            .replace(/Z/g, '2')
            .replace(/[^0-9]/g, '');
}

function extractBestServiceFromText(raw) {
    const candidates = (String(raw || '').match(/[0-9OQILSBZ\s-]{8,22}/g) || [])
            .map(normalizeServiceCandidate)
            .filter(v => v.length >= 8 && v.length <= 15)
            .sort((a, b) => b.length - a.length);

    return candidates[0] || '';
}

function looksLikeNameCandidate(line) {
    const raw = cleanLine(line);
    const up = normalizeUpper(raw);

    if (!raw || raw.length < 6 || raw.length > 80) {
        return false;
    }
    if (isStopLabel(raw)) {
        return false;
    }
    if (/\$\s*\d/.test(raw) || isMoneyTextLine(raw) || hasLongMixedIdentifier(raw)) {
        return false;
    }
    if (/(RFC|TOTAL|PAGO|PAGAR|PESOS|TARIFA|SERVICIO|PERIODO|MEDIDOR|HILOS|LECTURA|COMISION FEDERAL|CFE|CUENTA|RMU|RPU|REPARTIR)/.test(up)) {
        return false;
    }
    if ((raw.match(/\d/g) || []).length > 2) {
        return false;
    }

    const words = raw.split(/\s+/).filter(Boolean);
    return words.length >= 2 && words.length <= 8;
}

function scoreAddressLine(line) {
    const raw = cleanLine(line);
    const up = normalizeUpper(raw);

    if (!raw || isStopLabel(raw)) {
        return -100;
    }
    if (/\$\s*\d/.test(raw) || isMoneyTextLine(raw) || hasLongMixedIdentifier(raw)) {
        return -50;
    }

    let score = 0;

    if ((raw.match(/\d/g) || []).length >= 1) {
        score += 3;
    }
    if (ADDRESS_HINTS.some(h => up.includes(h))) {
        score += 3;
    }
    if (/\,/.test(raw)) {
        score += 1;
    }
    if (/\b\d{5}\b/.test(raw)) {
        score += 2;
    }
    if (raw.length >= 10) {
        score += 1;
    }

    return score;
}

function isStopLabel(line) {
    const up = normalizeUpper(line);

    return [
        'NO DE SERVICIO',
        'NO. DE SERVICIO',
        'NUMERO DE SERVICIO',
        'NÚMERO DE SERVICIO',
        'CUENTA:',
        'CUENTA',
        'RMU',
        'RPU',
        'REPARTIR',
        'LIMITE DE PAGO',
        'CORTE A PARTIR',
        'TARIFA:',
        'TARIFA',
        'NO HILOS',
        'PERIODO FACTURADO',
        'LECTURA ACTUAL',
        'LECTURA ANTERIOR',
        'CONCEPTO',
        'DESGLOSE DEL IMPORTE',
        'SUBTOTAL',
        'CONSUMO HISTORICO',
        'TOTAL A PAGAR',
        'DESCARGA NUESTRA APP',
        'APP AUTORIZADA',
        'COMISION FEDERAL',
        'COMISIÓN FEDERAL'
    ].some(x => up.includes(x));
}

function looksLikeAddressLine(line) {
    const raw = cleanLine(line);
    const up = normalizeUpper(raw);

    if (!raw || raw.length < 6) {
        return false;
    }
    if (isStopLabel(raw)) {
        return false;
    }
    if (/\$\s*\d/.test(raw) || isMoneyTextLine(raw) || hasLongMixedIdentifier(raw)) {
        return false;
    }
    if (up.includes('TOTAL A PAGAR') || up.includes('ESTE GRAFICO') || up.includes('REPARTIR') || up.includes('CUENTA') || up.includes('RMU')) {
        return false;
    }

    return scoreAddressLine(raw) >= 2;
}

function parsePeriodo(text) {
    const upper = normalizeUpper(text);

    const m = upper.match(
            /PERIODO\s+FACTURADO[:\s]*([0-3]?\d)\s*(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)\s*(\d{2})\s*[-–]\s*([0-3]?\d)\s*(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)\s*(\d{2})/
            );

    if (!m) {
        return {raw: '', start: null, end: null, days: 0};
    }

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
    if (!rawPeriodo) {
        return '';
    }

    const m = rawPeriodo.match(/(\d{2})\s+([A-Z]{3})\s+(\d{2})\s+-\s+(\d{2})\s+([A-Z]{3})\s+(\d{2})/i);

    if (!m) {
        return rawPeriodo;
    }

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


function parseCfeCustomerBlock(text) {
    const originalLines = buildLines(text).slice(0, 80);
    const lines = originalLines
            .map(stripRightColumnNoise)
            .map(cleanLine)
            .filter(Boolean);

    const serviceIndex = lines.findIndex(line => /NO\.?\s*DE\s*SERVICIO|NUMERO\s+DE\s+SERVICIO|NÚMERO\s+DE\s+SERVICIO/i.test(removeAccentsUpper(line)));
    const cut = serviceIndex > 0 ? serviceIndex : Math.min(lines.length, 45);
    const head = lines.slice(0, cut);

    let nameIndex = -1;

    for (let i = 0; i < head.length; i++) {
        const line = head[i];

        if (isLikelyCfeCustomerName(line)) {
            const following = head.slice(i + 1, Math.min(head.length, i + 7));
            const addressCount = following.filter(isLikelyCfeAddressLine).length;

            if (addressCount >= 1 || i < 12) {
                nameIndex = i;
                break;
            }
        }
    }

    if (nameIndex < 0) {
        return {nombre: '', direccion: ''};
    }

    const nombre = cleanLine(head[nameIndex]);
    const addressLines = [];

    for (let i = nameIndex + 1; i < head.length; i++) {
        const line = cleanLine(head[i]);
        const up = removeAccentsUpper(line);

        if (!line) {
            continue;
        }
        if (isLikelyCfeCustomerName(line) && addressLines.length) {
            break;
        }
        if (/NO\.?\s*DE\s*SERVICIO|NUMERO\s+DE\s+SERVICIO|RMU|CUENTA|TOTAL\s+A\s+PAGAR|LIMITE\s+DE\s+PAGO|CORTE\s+A\s+PARTIR/i.test(up)) {
            break;
        }
        if (isLikelyCfeAddressLine(line)) {
            addressLines.push(line);
        } else if (addressLines.length >= 2) {
            break;
        }
    }

    return {
        nombre,
        direccion: addressLines.join(' ').replace(/\s+/g, ' ').trim(),
    };
}

function parseNombreDireccion(text) {
    const cfeBlock = parseCfeCustomerBlock(text);

    if (cfeBlock.nombre || cfeBlock.direccion) {
        return cfeBlock;
    }

    const lines = buildLines(text).slice(0, 45);

    const labeledName = normalizeText(text).match(/(?:NOMBRE|CLIENTE|TITULAR)[:\s]+([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s.]{5,80})/i);

    if (labeledName && !isMoneyTextLine(labeledName[1])) {
        const possibleAddress = lines.filter(looksLikeAddressLine).slice(0, 3).join(' ').trim();
        return {nombre: cleanLine(labeledName[1]), direccion: possibleAddress};
    }

    let best = {score: -999, nombre: '', direccion: ''};

    for (let i = 0; i < lines.length; i++) {
        const current = cleanLine(lines[i]);

        if (!looksLikeNameCandidate(current)) {
            continue;
        }

        const addressLines = [];
        let score = 2;

        for (let j = i + 1; j < Math.min(lines.length, i + 6); j++) {
            const candidate = cleanLine(lines[j]);

            if (!candidate) {
                continue;
            }

            if (isStopLabel(candidate)) {
                break;
            }

            const lineScore = scoreAddressLine(candidate);

            if (lineScore >= 2) {
                addressLines.push(candidate.replace(/\$.*$/, '').trim());
                score += lineScore;
            } else if (addressLines.length) {
                break;
            }
        }

        if (addressLines.length >= 1) {
            score += addressLines.length * 2;
        }

        if ((current.match(/\b[A-ZÁÉÍÓÚÑ]{2,}\b/g) || []).length >= 2) {
            score += 2;
        }

        const direccion = addressLines.join(' ').replace(/\s+/g, ' ').trim();

        if (score > best.score) {
            best = {
                score,
                nombre: current.replace(/\$.*$/, '').trim(),
                direccion
            };
        }
    }

    if (best.score > 0) {
        return {nombre: best.nombre, direccion: best.direccion};
    }

    return {nombre: '', direccion: ''};
}

function parseServicio(text) {
    const lines = buildLines(text);
    const labels = ['NO DE SERVICIO', 'NO. DE SERVICIO', 'NUMERO DE SERVICIO', 'NÚMERO DE SERVICIO'];

    for (let i = 0; i < lines.length; i++) {
        const raw = cleanLine(lines[i]);
        const up = normalizeUpper(raw);

        if (!labels.some(label => up.includes(label))) {
            continue;
        }

        const sameLineText = raw.replace(/.*?(?:NO\.?\s*DE\s*SERVICIO|NUMERO\s+DE\s+SERVICIO)[:\s]*/i, '');
        const sameLine = extractBestServiceFromText(sameLineText);

        if (sameLine) {
            return sameLine;
        }

        for (let j = i + 1; j < Math.min(lines.length, i + 3); j++) {
            const nextCandidate = extractBestServiceFromText(lines[j]);

            if (nextCandidate) {
                return nextCandidate;
            }
        }
    }

    const upper = normalizeUpper(text);
    const nearLabel = upper.match(/(?:NO\.?\s*DE\s*SERVICIO|NUMERO\s+DE\s+SERVICIO)[:\s]*([0-9OQILSBZ\s-]{8,22})/);

    if (nearLabel) {
        const candidate = normalizeServiceCandidate(nearLabel[1]);

        if (candidate.length >= 8 && candidate.length <= 15) {
            return candidate;
        }
    }

    return extractBestServiceFromText(text) || '';
}

function parseTarifa(text) {
    const lines = buildLines(text);
    const allowed = ['1', '1A', '1B', '1C', '1D', '1E', '1F', 'DAC', 'PDBT', 'GDBT', 'GDMTO', 'GDMTH'];
    const tariffRegex = /\b(1[A-F]?|DAC|PDBT|GDBT|GDMTO|GDMTH)\b/;

    for (let i = 0; i < lines.length; i++) {
        const line = normalizeUpper(lines[i]);

        if (!line.includes('TARIFA')) {
            continue;
        }

        const afterLabel = line.replace(/^.*?TARIFA[:\s]*/i, ' ');
        let m = afterLabel.match(tariffRegex);

        if (m && allowed.includes(m[1])) {
            return m[1];
        }

        for (let j = i + 1; j < Math.min(lines.length, i + 3); j++) {
            m = normalizeUpper(lines[j]).match(tariffRegex);

            if (m && allowed.includes(m[1])) {
                return m[1];
            }
        }
    }

    const full = normalizeUpper(text);
    const m = full.match(/TARIFA[\s:.-]{0,20}(1[A-F]?|DAC|PDBT|GDBT|GDMTO|GDMTH)\b/);

    return m && allowed.includes(m[1]) ? m[1] : '';
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

                if (n > 0) {
                    return n;
                }
            }
        }
    }

    const full = normalizeText(text);

    let m = full.match(/DESGLOSE\s+DEL\s+IMPORTE\s+A\s+PAGAR[\s\S]{0,250}?\bTotal\b[^\d$]*\$?\s*([\d,]+(?:\.\d{1,2})?)/i);

    if (m) {
        const n = parseMoneyToNumber(m[1]);

        if (n > 0) {
            return n;
        }
    }

    m = full.match(/\bTotal\b[^\d$]*\$?\s*([\d,]+(?:\.\d{1,2})?)/i);

    if (m) {
        const n = parseMoneyToNumber(m[1]);

        if (n > 0) {
            return n;
        }
    }

    m = full.match(/TOTAL\s+A\s+PAGAR[\s\S]{0,80}?\$?\s*([\d,]+(?:\.\d{1,2})?)/i);

    if (m) {
        const n = parseMoneyToNumber(m[1]);

        if (n > 0) {
            return n;
        }
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

/**
 * Detecta el consumo actual del periodo.
 * Importante:
 * - NO debe usar el historial como fallback.
 * - El historial son periodos anteriores.
 * - consumoPeriodo es el periodo actual del recibo.
 */
function parseConsumoPeriodo(text, historial = [], periodoRaw = '') {
    const lines = buildLines(text);

    for (const rawLine of lines) {
        const line = cleanLine(rawLine);
        const up = normalizeUpper(line);

        if (up.includes('ENERGIA (KWH)') || up.includes('ENERGIA KWH')) {
            const nums = (line.match(/\d{1,5}(?:,\d{3})*/g) || [])
                    .map(v => parseIntSafe(v))
                    .filter(n => n > 0);

            if (nums.length >= 3) {
                return nums[2];
            }
        }
    }

    const upText = normalizeUpper(text);

    let m = upText.match(/ENERGIA\s*\(KWH\)[\s\S]{0,100}?(\d{1,5}(?:,\d{3})*)\s+(\d{1,5}(?:,\d{3})*)\s+(\d{1,5})/);

    if (m) {
        return parseIntSafe(m[3]);
    }

    m = upText.match(/LECTURA\s+ACTUAL[\s\S]{0,160}?LECTURA\s+ANTERIOR[\s\S]{0,160}?TOTAL\s+PERIODO[\s\S]{0,120}?(\d{1,5})/);

    if (m) {
        return parseIntSafe(m[1]);
    }

    return 0;
}

function parseEstadoFromDireccion(direccion) {
    if (!direccion) {
        return '';
    }

    const m = direccion.match(/,\s*([A-Za-zÁÉÍÓÚÑáéíóúñ. ]{2,})$/);
    return m ? m[1].replace(/\.+$/, '').trim().toUpperCase() : '';
}

/**
 * Parser más tolerante para la tabla de consumo histórico.
 *
 * Soporta líneas como:
 * del 20 ENE 26 al 18 FEB 26 185 $318.00 $318.00
 * DEL 20 ENE 26 AL 18 FEB 26 185 318.00 318.00
 *
 * También tolera comas, signos de pesos y espacios irregulares.
 */
function parseHistoricalTable(text) {
    const rows = [];
    const seen = new Set();
    const month = '(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)';

    const normalized = normalizeUpper(text)
            .replace(/\$/g, ' ')
            .replace(/,/g, '')
            .replace(/[|]/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();

    const regex = new RegExp(
            `(?:DEL\\s+)?(\\d{1,2})\\s+${month}\\s+(\\d{2})\\s+(?:AL\\s+)?(\\d{1,2})\\s+${month}\\s+(\\d{2})\\s+(\\d{2,5})\\s+([\\d]+(?:\\.\\d{1,2})?)\\s+([\\d]+(?:\\.\\d{1,2})?)(?:\\s+([\\d]+(?:\\.\\d{1,2})?))?`,
            'g'
            );

    let m;

    while ((m = regex.exec(normalized)) !== null) {
        const periodo = `del ${m[1]} ${m[2]} ${m[3]} al ${m[4]} ${m[5]} ${m[6]}`;
        const kwh = parseIntSafe(m[7]);
        const importe = parseMoneyToNumber(m[8]);
        const pago = parseMoneyToNumber(m[9]);
        const pendiente = parseMoneyToNumber(m[10] || '0');
        const key = canonicalPeriodLabel(periodo);

        if (kwh > 0 && !seen.has(key)) {
            rows.push({
                periodo,
                kwh,
                importe,
                pago,
                pendiente
            });
            seen.add(key);
        }
    }

    /*
     * Fallback por líneas:
     * Si por alguna razón el texto no se pudo normalizar completo,
     * se intenta detectar línea por línea.
     */
    if (!rows.length) {
        const lines = buildLines(text);

        for (const rawLine of lines) {
            const line = normalizeUpper(cleanLine(rawLine))
                    .replace(/\$/g, ' ')
                    .replace(/,/g, '')
                    .replace(/[|]/g, ' ')
                    .replace(/\s+/g, ' ')
                    .trim();

            const lineRegex = new RegExp(
                    `(?:DEL\\s+)?(\\d{1,2})\\s+${month}\\s+(\\d{2})\\s+(?:AL\\s+)?(\\d{1,2})\\s+${month}\\s+(\\d{2}).*?(\\d{2,5})\\s+([\\d]+(?:\\.\\d{1,2})?)\\s+([\\d]+(?:\\.\\d{1,2})?)(?:\\s+([\\d]+(?:\\.\\d{1,2})?))?`
                    );

            const lm = line.match(lineRegex);

            if (!lm) {
                continue;
            }

            const periodo = `del ${lm[1]} ${lm[2]} ${lm[3]} al ${lm[4]} ${lm[5]} ${lm[6]}`;
            const kwh = parseIntSafe(lm[7]);
            const importe = parseMoneyToNumber(lm[8]);
            const pago = parseMoneyToNumber(lm[9]);
            const pendiente = parseMoneyToNumber(lm[10] || '0');
            const key = canonicalPeriodLabel(periodo);

            if (kwh > 0 && !seen.has(key)) {
                rows.push({
                    periodo,
                    kwh,
                    importe,
                    pago,
                    pendiente
                });
                seen.add(key);
            }
        }
    }

    return rows;
}


function parseHistoricalPeriodDays(periodo = '') {
    const m = normalizeUpper(periodo).match(/(?:DEL\s+)?(\d{1,2})\s+(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)\s+(\d{2})\s+(?:AL\s+)?(\d{1,2})\s+(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)\s+(\d{2})/);
    if (!m) return 30;
    const start = new Date(2000 + parseIntSafe(m[3]), (MONTHS[m[2]] || 1) - 1, parseIntSafe(m[1]));
    const end = new Date(2000 + parseIntSafe(m[6]), (MONTHS[m[5]] || 1) - 1, parseIntSafe(m[4]));
    const days = Math.round((end - start) / (1000 * 60 * 60 * 24));
    return Number.isFinite(days) && days > 0 ? days : 30;
}

function buildMonthlyAverages({consumoPeriodo = 0, total = 0, periodo = {}, tipoPeriodo = 'Mensual', historial = []} = {}) {
    const consumosMensuales = [];
    const pagosMensuales = [];
    const consumosPeriodo = [];
    const pagosPeriodo = [];
    const currentIsBimestral = tipoPeriodo === 'Bimestral' || Number(periodo?.days || 0) >= 45;

    const consumoActual = Number(consumoPeriodo || 0);
    const pagoActual = Number(total || 0);
    if (consumoActual > 0) {
        consumosPeriodo.push(consumoActual);
        consumosMensuales.push(currentIsBimestral ? consumoActual / 2 : consumoActual);
    }
    if (pagoActual > 0) {
        pagosPeriodo.push(pagoActual);
        pagosMensuales.push(currentIsBimestral ? pagoActual / 2 : pagoActual);
    }

    (historial || []).forEach(row => {
        const days = parseHistoricalPeriodDays(row?.periodo || '');
        const isBimestral = days >= 45;
        const kwh = Number(row?.kwh || 0);
        const pago = Number(row?.pago || row?.importe || 0);
        if (kwh > 0) {
            consumosPeriodo.push(kwh);
            consumosMensuales.push(isBimestral ? kwh / 2 : kwh);
        }
        if (pago > 0) {
            pagosPeriodo.push(pago);
            pagosMensuales.push(isBimestral ? pago / 2 : pago);
        }
    });

    const sum = arr => arr.reduce((a, b) => a + b, 0);
    const avg = arr => arr.length ? sum(arr) / arr.length : 0;
    const round2 = value => Math.round(Number(value || 0) * 100) / 100;

    return {
        consumoPromedioMensual: round2(avg(consumosMensuales)),
        pagoPromedioMensual: round2(avg(pagosMensuales)),
        totalConsumoHistorico: round2(sum(consumosPeriodo)),
        totalPagoHistorico: round2(sum(pagosPeriodo)),
        consumoMaximoHistorico: round2(consumosPeriodo.length ? Math.max(...consumosPeriodo) : 0),
        periodosPromediados: Math.max(consumosMensuales.length, pagosMensuales.length),
    };
}

function chooseBestText(primary, secondary) {
    const p = normalizeText(primary);
    const s = normalizeText(secondary);

    const score = (txt) => {
        const up = normalizeUpper(txt);
        let n = 0;

        if (up.includes('NO DE SERVICIO') || up.includes('NO. DE SERVICIO')) {
            n += 3;
        }
        if (up.includes('TARIFA')) {
            n += 2;
        }
        if (up.includes('PERIODO FACTURADO')) {
            n += 3;
        }
        if (up.includes('CONSUMO HISTORICO')) {
            n += 2;
        }
        if (up.includes('TOTAL A PAGAR')) {
            n += 2;
        }
        if (up.includes('NO HILOS')) {
            n += 1;
        }

        return n;
    };

    return score(s) > score(p) ? s : p;
}

function looksCorruptedText(text) {
    const up = normalizeUpper(text);

    if (!up) {
        return true;
    }

    const hasKeyFields =
            up.includes('NO DE SERVICIO') ||
            up.includes('NO. DE SERVICIO') ||
            up.includes('PERIODO FACTURADO') ||
            up.includes('TOTAL A PAGAR');

    if (!hasKeyFields) {
        return true;
    }

    const strange = (text.match(/[ ]/g) || []).length;
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

    /*
     * Primero intentamos detectar el historial en página 2.
     * Si no se detecta, intentamos con todo el texto combinado.
     */
    let historial = parseHistoricalTable(page2);

    if (!historial.length) {
        historial = parseHistoricalTable(fullText);
    }

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

    const tipoPeriodo = (periodo?.days ?? 0) >= 45 ? 'Bimestral' : 'Mensual';
    const monthly = buildMonthlyAverages({consumoPeriodo, total, periodo, tipoPeriodo, historial});
    const pagoProm = monthly.pagoPromedioMensual || total;

    const ahorroEstimado = Math.max(0, pagoProm - costoBase);
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
        consumoPromedioMensual: monthly.consumoPromedioMensual || 0,
        pagoPromedioMensual: monthly.pagoPromedioMensual || 0,
        totalConsumoHistorico: monthly.totalConsumoHistorico || 0,
        totalPagoHistorico: monthly.totalPagoHistorico || 0,
        consumoMaximoHistorico: monthly.consumoMaximoHistorico || 0,
        periodosPromediados: monthly.periodosPromediados || 0,
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
        if (Math.abs(b.y - a.y) > 2) {
            return b.y - a.y;
        }

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

function createHighContrastCanvas(source) {
    const canvas = document.createElement('canvas');
    canvas.width = source.width;
    canvas.height = source.height;

    const ctx = canvas.getContext('2d', {alpha: false});
    ctx.drawImage(source, 0, 0);

    const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const data = img.data;

    for (let i = 0; i < data.length; i += 4) {
        const gray = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
        const val = gray > 175 ? 255 : gray < 90 ? 0 : gray;
        data[i] = data[i + 1] = data[i + 2] = val;
    }

    ctx.putImageData(img, 0, 0);

    return canvas;
}

function scoreRecognizedText(text) {
    const up = normalizeUpper(text);
    const parsedName = parseNombreDireccion(text);

    let score = 0;

    if (up.includes('NO DE SERVICIO') || up.includes('NO. DE SERVICIO')) {
        score += 5;
    }
    if (up.includes('PERIODO FACTURADO')) {
        score += 4;
    }
    if (up.includes('TOTAL A PAGAR')) {
        score += 4;
    }
    if (up.includes('TARIFA')) {
        score += 3;
    }
    if (parseServicio(text)) {
        score += 5;
    }
    if (parsedName.direccion) {
        score += 4;
    }
    if (parsedName.nombre) {
        score += 3;
    }

    return score + Math.min(6, Math.round(String(text || '').length / 250));
}

async function runBestOcrVariant(canvas, pageIndex, onProgress) {
    const prefs = loadUserPreferences();
    const variants = [canvas];

    if (prefs?.ocr?.preferHighContrast !== false) {
        variants.push(createHighContrastCanvas(canvas));
    }

    let bestText = '';
    let bestScore = -Infinity;

    for (const variant of variants) {
        const text = await runOcrOnCanvas(variant, pageIndex, onProgress);
        const score = scoreRecognizedText(text);

        if (score > bestScore) {
            bestScore = score;
            bestText = text;
        }

        if (bestScore >= 18 && prefs?.ocr?.aggressiveMode === false) {
            break;
        }
    }

    return bestText;
}

async function runOcrOnCanvas(canvas, pageIndex, onProgress) {
    if (!window.Tesseract) {
        return '';
    }

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

async function extractPdfTextWithPdfBox(file, onProgress) {
    if (onProgress) {
        onProgress({message: 'Leyendo texto real del PDF con PDFBox…'});
    }

    const formData = new FormData();
    formData.append('file', file);

    const response = await fetch('api/receipts/pdf-text', {
        method: 'POST',
        body: formData,
    });

    const payload = await response.json().catch(() => null);

    if (!response.ok || !payload?.ok) {
        throw new Error(payload?.message || 'No se pudo leer el PDF con PDFBox.');
    }

    return payload.data || {};
}

function hasEnoughPdfText(text) {
    const compact = String(text || '').replace(/\s/g, '');
    if (compact.length < 80) {
        return false;
    }

    const up = normalizeUpper(text);
    return (
        up.includes('NO DE SERVICIO') ||
        up.includes('NO. DE SERVICIO') ||
        up.includes('PERIODO FACTURADO') ||
        up.includes('TOTAL A PAGAR') ||
        up.includes('TARIFA')
    );
}

async function pdfToTextAndPreview(file, {onProgress} = {}) {
    let pdfBoxResult = null;
    let pdfBoxError = null;

    try {
        pdfBoxResult = await extractPdfTextWithPdfBox(file, onProgress);
    } catch (err) {
        pdfBoxError = err;
        console.warn('PDFBox no pudo extraer texto del PDF:', err?.message || err);
    }

    const pdfjsLib = window.pdfjsLib;
    let previewCanvas = null;
    let pdfJsPageTexts = [];

    if (pdfjsLib) {
        try {
            pdfjsLib.GlobalWorkerOptions.workerSrc =
                    pdfjsLib.GlobalWorkerOptions.workerSrc ||
                    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

            const arrayBuffer = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({data: arrayBuffer}).promise;

            const firstPage = await pdf.getPage(1);
            previewCanvas = await renderPageToCanvas(firstPage, 0.8);

            if (!hasEnoughPdfText(pdfBoxResult?.text || '')) {
                pdfJsPageTexts = await extractPdfTextByPage(pdf, onProgress);
            }
        } catch (err) {
            console.warn('No se pudo leer o previsualizar el PDF con PDF.js:', err?.message || err);
        }
    }

    const pdfBoxText = normalizeText(pdfBoxResult?.text || '');
    const pdfBoxPages = Array.isArray(pdfBoxResult?.pageTexts)
            ? pdfBoxResult.pageTexts.map(normalizeText)
            : [];

    if (hasEnoughPdfText(pdfBoxText)) {
        if (onProgress) {
            onProgress({message: 'PDF leído correctamente con texto real.'});
        }

        return {
            text: pdfBoxText,
            pageTexts: pdfBoxPages.length ? pdfBoxPages : [pdfBoxText],
            canvas: previewCanvas,
            source: 'PDFBox',
        };
    }

    const pdfJsText = normalizeText(pdfJsPageTexts.join('\n\n'));

    if (hasEnoughPdfText(pdfJsText)) {
        if (onProgress) {
            onProgress({message: 'PDF leído con texto interno del documento.'});
        }

        return {
            text: pdfJsText,
            pageTexts: pdfJsPageTexts,
            canvas: previewCanvas,
            source: 'PDF.js',
        };
    }

    /*
     * Para PDFs se evita OCR automático: si el archivo está escaneado o es una imagen
     * dentro de un PDF, se habilita captura manual para no llenar la cotización con
     * datos inventados o mal reconocidos. El OCR se mantiene solo para imágenes.
     */
    if (onProgress) {
        onProgress({
            message: pdfBoxError?.message
                    ? `${pdfBoxError.message} Se habilitará captura manual.`
                    : 'El PDF no contiene texto real suficiente. Se habilitará captura manual.'
        });
    }

    return {
        text: normalizeText([pdfBoxText, pdfJsText].filter(Boolean).join('\n\n')),
        pageTexts: pdfBoxPages.length ? pdfBoxPages : pdfJsPageTexts,
        canvas: previewCanvas,
        source: pdfBoxText ? 'PDFBox parcial' : 'Sin texto suficiente',
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

function buildMockReceipt({selectedTariff, file}) {
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

        if (selectedTariff?.familia === 'PDBT') {
            return 'PDBT';
        }

        if (selectedTariff?.familia === 'GDMTH') {
            return 'GDMTH';
        }

        if (selectedTariff?.familia === 'GDMTO') {
            return 'GDMTO';
        }

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

export function createEmptyReceiptData(selectedTariff = null, rawText = '') {
    return {
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
        rawText: rawText || '',
        instalacion: {},
        insumos: [],
        impuestosPct: 0.16
    };
}

async function imageFileToTextAndPreview(file, {onProgress} = {}) {
    const imageBitmap = await createImageBitmap(file);
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d', {alpha: false});

    const maxSide = 1800;
    const scale = Math.min(1, maxSide / Math.max(imageBitmap.width, imageBitmap.height));

    canvas.width = Math.max(1, Math.round(imageBitmap.width * scale));
    canvas.height = Math.max(1, Math.round(imageBitmap.height * scale));

    ctx.drawImage(imageBitmap, 0, 0, canvas.width, canvas.height);

    if (onProgress) {
        onProgress({message: 'Aplicando OCR a imagen…'});
    }

    const text = await runBestOcrVariant(canvas, 1, onProgress);

    return {
        text,
        pageTexts: [text],
        canvas,
    };
}

export async function analyzeReceiptFile(file, options = {}) {
    const {selectedTariff = null, onProgress} = options;
    const isPdf = file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf');
    const isImage = file.type.startsWith('image/') || /\.(png|jpe?g|webp|bmp)$/i.test(file.name || '');

    let text = '';
    let canvas = null;
    let pageTexts = [];
    let analysisOk = true;
    let manualReason = '';

    if (onProgress) {
        onProgress({message: 'Preparando lectura…'});
    }

    try {
        if (isPdf) {
            const result = await pdfToTextAndPreview(file, {onProgress});
            text = result.text || '';
            pageTexts = Array.isArray(result.pageTexts) ? result.pageTexts : [];
            canvas = result.canvas || null;
        } else if (isImage) {
            const result = await imageFileToTextAndPreview(file, {onProgress});
            text = result.text || '';
            pageTexts = Array.isArray(result.pageTexts) ? result.pageTexts : [];
            canvas = result.canvas || null;
        } else {
            throw new Error('Formato de archivo no compatible. Usa PDF, PNG o JPG.');
        }
    } catch (err) {
        console.error('Error leyendo recibo:', err);
        analysisOk = false;
        manualReason = err?.message || 'No se pudo leer el archivo.';

        if (onProgress) {
            onProgress({message: `${manualReason} Se habilitará captura manual.`});
        }
    }

    let parsed = null;

    if ((text || '').replace(/\s/g, '').length > 60) {
        if (onProgress) {
            onProgress({message: 'Extrayendo datos del recibo…'});
        }

        parsed = parseCfeReceiptTextFromPages(
                pageTexts[0] || text,
                pageTexts[1] || text,
                text
                );
    } else {
        const label = isPdf ? 'PDF' : 'imagen';
        analysisOk = false;
        manualReason = manualReason || `No se pudo extraer texto suficiente del ${label}.`;

        if (onProgress) {
            onProgress({message: `${manualReason} Se habilitará captura manual.`});
        }

        parsed = createEmptyReceiptData(selectedTariff, text || '');
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

    const hasCriticalData = Boolean(
            parsed.servicio ||
            parsed.nombre ||
            parsed.direccion ||
            Number(parsed.consumoPeriodo || 0) > 0 ||
            Number(parsed.totalAPagar || 0) > 0
            );

    if (analysisOk && !hasCriticalData) {
        analysisOk = false;
        manualReason = 'No se detectaron datos clave del recibo.';
    }

    const message = analysisOk
            ? 'Recibo analizado correctamente.'
            : `${manualReason || 'No se pudo analizar el recibo.'} Se habilitará captura manual.`;

    if (onProgress) {
        onProgress({message: analysisOk ? 'Listo.' : message});
    }

    return {
        ok: analysisOk,
        manual: !analysisOk,
        message,
        text,
        parsed,
        canvas
    };
}