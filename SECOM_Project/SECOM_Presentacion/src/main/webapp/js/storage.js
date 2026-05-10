function parseJsonSafe(text, fallback = null){
  try{
    return text ? JSON.parse(text) : fallback;
  } catch {
    return fallback;
  }
}

function request(method, url, body){
  const xhr = new XMLHttpRequest();
  xhr.open(method, url, false);
  xhr.setRequestHeader('Accept', 'application/json');

  if (body !== undefined && body !== null){
    xhr.setRequestHeader('Content-Type', 'application/json;charset=UTF-8');
  }

  xhr.send(body !== undefined && body !== null ? JSON.stringify(body) : null);

  const payload = parseJsonSafe(xhr.responseText, { ok:false, message:'Respuesta no válida del servidor.' });

  if (xhr.status >= 200 && xhr.status < 300){
    return payload?.data ?? payload;
  }

  throw new Error(payload?.message || `Error HTTP ${xhr.status}`);
}

export function getQuotes(){
  return request('GET', 'api/quotes');
}

export function saveQuote(quote){
  return request('POST', 'api/quotes', quote);
}

export function updateQuote(id, patch){
  return request('PUT', `api/quotes/${encodeURIComponent(id)}`, patch);
}

export function removeQuote(id){
  return request('DELETE', `api/quotes/${encodeURIComponent(id)}`);
}

export function getInsumos(){
  return request('GET', 'api/insumos');
}

export function saveInsumo(insumo){
  return request('POST', 'api/insumos', insumo);
}

export function updateInsumo(id, patch){
  return request('PUT', `api/insumos/${encodeURIComponent(id)}`, patch);
}

export function removeInsumo(id){
  return request('DELETE', `api/insumos/${encodeURIComponent(id)}`);
}

export function getProjects(){
  return request('GET', 'api/projects');
}

export function saveProjectFromQuote(quote){
  return request('POST', `api/projects/from-quote/${encodeURIComponent(quote.id)}`, {});
}

export function saveProject(project){
  return request('POST', 'api/projects', project);
}

export function updateProject(id, patch){
  return request('PUT', `api/projects/${encodeURIComponent(id)}`, patch);
}

export function removeProject(id){
  return request('DELETE', `api/projects/${encodeURIComponent(id)}`);
}

export function resetAllData(){
  return request('POST', 'api/debug/reset', {});
}