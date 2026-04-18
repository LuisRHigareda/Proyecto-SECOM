export const $ = (sel, root=document) => root.querySelector(sel);
export const $$ = (sel, root=document) => Array.from(root.querySelectorAll(sel));

export function uuid(){
  // Short-ish unique id for UI purposes
  return 'Q' + Math.random().toString(16).slice(2, 8).toUpperCase() + '-' + Date.now().toString(36).toUpperCase();
}

export function clamp(n, min, max){
  return Math.max(min, Math.min(max, n));
}

export function formatCurrencyMXN(value){
  const n = Number(value || 0);
  return new Intl.NumberFormat('es-MX', { style:'currency', currency:'MXN', maximumFractionDigits: 0 }).format(n);
}

export function formatNumber(value){
  const n = Number(value || 0);
  return new Intl.NumberFormat('es-MX', { maximumFractionDigits: 0 }).format(n);
}

export function formatDateTime(ts){
  const d = new Date(ts);
  return d.toLocaleString('es-MX', { year:'numeric', month:'short', day:'2-digit', hour:'2-digit', minute:'2-digit' });
}

export function formatDate(ts){
  const d = new Date(ts);
  return d.toLocaleDateString('es-MX', { year:'numeric', month:'short', day:'2-digit' });
}

export function safeText(s){
  return String(s ?? '').replace(/[<>]/g, '');
}

export function parseMoneyToNumber(s){
  if (!s) return 0;
  const cleaned = String(s).replace(/[^0-9.,-]/g, '').replace(/\.(?=.*\.)/g, '');
  // If it has comma and dot, assume comma = thousands
  let num = cleaned;
  const hasComma = cleaned.includes(',');
  const hasDot = cleaned.includes('.');
  if (hasComma && hasDot){
    num = cleaned.replace(/,/g,'');
  } else if (hasComma && !hasDot){
    // could be thousands or decimal; in bills usually thousands
    num = cleaned.replace(/,/g,'');
  }
  const out = Number(num);
  return Number.isFinite(out) ? out : 0;
}

export function parseIntSafe(s){
  const n = parseInt(String(s).replace(/[^0-9]/g,''), 10);
  return Number.isFinite(n) ? n : 0;
}

export function pickNonEmpty(...vals){
  for (const v of vals){
    if (v !== undefined && v !== null && String(v).trim() !== '') return v;
  }
  return '';
}

export function setPillStatus(text, type='ok'){
  const pill = document.getElementById('pillStatus');
  const label = document.getElementById('pillStatusText');
  if (!pill || !label) return;
  label.textContent = text;
  const dot = pill.querySelector('.pill__dot');
  if (!dot) return;
  const map = {
    ok: 'var(--success)',
    busy: 'var(--warn)',
    error: 'var(--danger)',
  };
  dot.style.background = map[type] || map.ok;
  dot.style.boxShadow = type === 'busy' ? '0 0 0 5px rgba(245,158,11,0.12)' :
                         type === 'error' ? '0 0 0 5px rgba(239,68,68,0.12)' :
                         '0 0 0 5px rgba(34,197,94,0.10)';
}

export function toast({title, message, icon='check-circle'}){
  const wrap = document.getElementById('toastWrap');
  if (!wrap) return;
  const el = document.createElement('div');
  el.className = 'toast';
  el.innerHTML = `
    <div class="toast__icon"><i data-lucide="${icon}"></i></div>
    <div>
      <div class="toast__title">${safeText(title)}</div>
      <div class="toast__msg">${safeText(message)}</div>
    </div>
    <div style="margin-left:auto">
      <button class="icon-btn" aria-label="Cerrar"><i data-lucide="x"></i></button>
    </div>
  `;
  wrap.appendChild(el);
  if (window.lucide) window.lucide.createIcons();

  const close = el.querySelector('button');
  const kill = () => {
    el.style.opacity = '0';
    el.style.transform = 'translateY(4px)';
    setTimeout(()=> el.remove(), 160);
  };
  close?.addEventListener('click', kill);
  setTimeout(kill, 5200);
}

export function openModal({title='Detalle', subtitle='', bodyHtml='', footHtml=''}){
  const modal = document.getElementById('modal');
  if (!modal) return;
  modal.classList.add('is-open');
  modal.setAttribute('aria-hidden', 'false');
  document.getElementById('modalTitle').textContent = title;
  document.getElementById('modalSubtitle').textContent = subtitle;
  document.getElementById('modalBody').innerHTML = bodyHtml;
  document.getElementById('modalFoot').innerHTML = footHtml;
  if (window.lucide) window.lucide.createIcons();
}

export function closeModal(){
  const modal = document.getElementById('modal');
  if (!modal) return;
  modal.classList.remove('is-open');
  modal.setAttribute('aria-hidden', 'true');
}

export function debounce(fn, wait=200){
  let t;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), wait);
  };
}
