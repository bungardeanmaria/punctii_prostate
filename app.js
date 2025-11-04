/***** NEW: inline base64 DOCX so we never fetch the template *****/
const EMBEDDED_DOCX_BASE64 =
  "UEsDBBQAAAAIAI9iZFsxpqS4/gAAADoCAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFVV9E7DMBB9Tn9wTyB5K96wiC5Q1pQq2l0fQ9w2z9b4yqN7kXJkYk+K8+zho2aB0dUC8tqGSMk3m2YJ2q9m8lMXjQH2Qt7lIh3Kk2g3b2lG1c9iYxE1A2u5P9Q4lZr7m9oJ7Sx3aA1m1c0c0Qk7O3n1l3X0q4q9k8e3YQ2mQ6sQd0mYc6m2f1m8b8i4o2+4mSg8P4iF9Oq2tq6b8bQd3Qv0k4U2bF9T2f9Wkq1o0p7qQkVwE7r8oGgFUEsDBBQAAAAIAIhiZFvvM7W9yQAALpEBAAARAAAAX3JlbHMvLnJlbHNVVAkAA1w6c2NcOnNjdXgLAAEE6AMAAAToAwAAdZDBDoIwEEX3fYqz1H2k2Yb0Tg9tO0uG0m2rYt4sKQv9v0URilFaI1y4z2+U2C5QZp0uM9+3F8qGx3fSziQJgdrm4l1R5m1o0n42Y7h7FqK3Yd8s7s1iYw1gq3gQx+R4H0T3bC4X9SOFJQe5w5iYl0o4k7hQq2Qf0v5N5H9cO4Vxw3I3Zk8Xb2bDk8p4fQ2kJQEAUEsDBBQAAAAIADFiZFsk4jS3YgAAAJUBAAAVAAAAd29yZC9zdHlsZXMueG1sVVQJAANcOnNjXDpzc3V4CwABBOgDAAAE6AMAAI2QwU7CMBBF73mK8mTqQWqE0d2m3J5C1hT8G3t7k8t2BEo0bq6BymY3Qf6ShFZVX1o9b8i4g3J2lJw0DWqM+2eZEe0g4n2WJQe3k6H3a7vZ0gqg4X0r2T2l3r7i1lH5k8T8V9oP2JbJg9pQaeV5P6cQpM1p5a8k8Ck7n4v5P0r8m3i0zpQY2x4nK3z0L+z0e1QSwMEFAAAAAgA5WJkW5GcgE2eAAAAWgEAAA4AAAB3b3JkL2RvY3VtZW50LnhtbFVUCQADXDpzY1w6c3N1eAsAAQToAwAABOgDAAC9UctOwzAQ3fYp5k1xk7l6Q6g8sWQmV2K3q4m6cWk5hYFS9Vnzcq0i4V+Kp2H8VbE2Sic9u7tH4h4p7+co9c4TqKf2i0w3b6a8k3AYzT0k8T2k8a5uQh0h3QdX2Vf4r3igv4qkEJp9T3mGQyNw6wqXy3bq6nQzQmQW+fQw8Y9m7+8M9Xb9k7L0kZ0s1p2d3t1Q0bXo6w+u3qQ6mJQ0r0yQnSg3FQ8Qh8YbQ2xqkYh2cU6i3k0Xb0dS5K8M5w0dX4VdYfVQXb4WZkQ0hGdQYk0cFZ0+J1IEdDb250ZW50PC94bTpzcGFjZT48L3c6dD48L3c6cj48L3c6cD48L3c6Ym9keT48L3c6ZG9jdW1lbnQ+UEsDBBQAAAAIAHViZFs7Cq8c/wAAABwCAAAPAAAAX3JlbHMvd29yZC9fcmVscy54bWxVVAkAA1w6c2NcOnNjdXgLAAEE6AMAAAToAwAAiZDRCoAgDETP8Vxj0hE0l6e9wC7F7CtU9b8Y8CU2x0Qb3p8WQ8w1E2QeO5hZp1Gk2wLyqzG3s1f1A2Gq8uM7Qn1q6dWwqR3sX3s3wFUEsBAhQAFAAAAAgAj2JkWzGmpLj+AAA AOgIAABMAAAAAAAAAAAAAAACAAQAAAABbQ29udGVudF9UeXBlc10ueG1sVVQFAANcOnNjc3V4CwABBOgDAAAE6AMAAFBLAQIUABQAAAAIAIhiZFvvM7W9yQAALpEBAAARAAAAAAAAAAAAAAAAAIABAAAAAF9yZWxzLy5yZWxzVVQFAANcOnNjc3V4CwABBOgDAAAE6AMAAFBLAQIUABQAAAAIADFiZFsk4jS3YgAAAJUBAAAVAAAAAAAAAAAAAAAAAIABAAAAd29yZC9zdHlsZXMueG1sVVQFAANcOnNjc3V4CwABBOgDAAAE6AMAAFBLAQIUABQAAAAIAOViZFuRnIBNngAAAFoBAAAOAAAAAAAAAAAAAAAAAIABAAAAd29yZC9kb2N1bWVudC54bWxVVAUAA1w6c2NzdXgLAAEE6AMAAAToAwAAUEsBAhQAFAAAAAgAdWJkWzsKrxz/AAAAHAAAABAAAAAAAAAAAAAAAAACAAQAAAF9yZWxzL3dvcmQvX3JlbHMueG1sVVQFAANcOnNjc3V4CwABBOgDAAAE6AMAAFBLBQYAAAAABAAEAP0BAABcAQAAAAA=";
/***** helper: base64 -> ArrayBuffer *****/
function b64ToArrayBuffer(b64) {
  const bin = atob(b64);
  const len = bin.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i);
  return bytes.buffer;
}
/***** END new block *****/

// (… keep your existing merged app code below …)
// IMPORTANT: wherever we previously loaded the template with fetch(),
// just do:
async function loadTemplateArrayBuffer() {
  // always use embedded docx to avoid CORS/file:// issues
  return b64ToArrayBuffer(EMBEDDED_DOCX_BASE64);
}

/* Then in your loadAll():
   replace:
     templateArrayBuffer = await fetchAB(PATHS.DOCX)
   with:
     templateArrayBuffer = await loadTemplateArrayBuffer();
*/

// And keep the rest of the app: JSON loader (or inline fallback),
// recursive editor with “Add field”, Body Sections, preview, and Generate.

// ---------- Paths ----------
const PATHS = {
  DOCX: "templates/template_base.docx"
};

// ---------- State ----------
let templateArrayBuffer = null;
let dataModel = {};
let sectionsUI = []; // [{key, include}]

// ---------- DOM ----------
const jsonChoice = document.getElementById("jsonChoice");
const reloadBtn = document.getElementById("reload");
const loadStatus = document.getElementById("loadStatus");
const btnGenerate = document.getElementById("btnGenerate");
const outputName = document.getElementById("outputName");
const genStatus = document.getElementById("genStatus");
const errorMsg = document.getElementById("errorMsg");

const toggleRaw = document.getElementById("toggleRaw");
const rawEditor = document.getElementById("rawEditor");
const simpleEditor = document.getElementById("simpleEditor");
const rawJsonTA = document.getElementById("rawJson");
const applyRawBtn = document.getElementById("applyRaw");

const kvContainer = document.getElementById("kvContainer");
const tablesContainer = document.getElementById("tablesContainer");

const sectionsList = document.getElementById("sections");
const bodyPreview = document.getElementById("bodyPreview");
const addSectionBtn = document.getElementById("addSection");

// ---------- Utils ----------
const isPrimitive = (v) => typeof v === "string" || typeof v === "number" || typeof v === "boolean";
const isPlainObject = (v) => v && typeof v === "object" && !Array.isArray(v);
const clone = (o) => JSON.parse(JSON.stringify(o));

async function fetchJSON(url) {
  const res = await fetch(url, { cache: "no-cache" });
  if (!res.ok) throw new Error(`HTTP ${res.status} for ${url}`);
  return res.json();
}
async function fetchAB(url) {
  const res = await fetch(url, { cache: "no-cache" });
  if (!res.ok) throw new Error(`HTTP ${res.status} for ${url}`);
  return res.arrayBuffer();
}

function refreshRawEditor() { rawJsonTA.value = JSON.stringify(dataModel, null, 2); }

// ---- Add Field controls (any object, including nested) ----
function addFieldControls(parentObj, containerEl) {
  const wrap = document.createElement("div");
  wrap.className = "inline";
  wrap.style.gap = "8px";

  const nameInput = document.createElement("input");
  nameInput.type = "text";
  nameInput.placeholder = "new field name";
  nameInput.style.minWidth = "160px";

  const typeSelect = document.createElement("select");
  ["string","number","boolean","object","array-objects","array-primitives"].forEach(t=>{
    const o=document.createElement("option"); o.value=t; o.textContent=t; typeSelect.appendChild(o);
  });

  const addBtn = document.createElement("button");
  addBtn.type = "button"; addBtn.className = "secondary"; addBtn.textContent = "Add field";

  addBtn.addEventListener("click", () => {
    const k = (nameInput.value || "").trim();
    if (!k) { alert("Field name required"); return; }
    if (k in parentObj) { alert(`Field "${k}" already exists`); return; }

    const t = typeSelect.value;
    switch (t) {
      case "string": parentObj[k] = ""; break;
      case "number": parentObj[k] = 0; break;
      case "boolean": parentObj[k] = false; break;
      case "object": parentObj[k] = {}; break;
      case "array-objects": parentObj[k] = [ {} ]; break;
      case "array-primitives": parentObj[k] = [ "" ]; break;
      default: parentObj[k] = "";
    }
    renderSimpleEditor(dataModel);
    refreshRawEditor();
    refreshPreview();
  });

  wrap.append(nameInput, typeSelect, addBtn);
  containerEl.appendChild(wrap);
}

// ---- Recursive data editor (with Add Field + seeded rows) ----
function renderNode(key, value, parent, path, container) {
  if (isPrimitive(value)) {
    const row = document.createElement("div"); row.className = "kv";
    const label = document.createElement("label"); label.textContent = key;
    const input = document.createElement("input");
    if (typeof value === "number") {
      input.type = "number"; input.step = "any"; input.value = String(value);
      input.addEventListener("input", () => { parent[key] = Number(input.value || 0); refreshRawEditor(); refreshPreview(); });
    } else if (typeof value === "boolean") {
      input.type="checkbox"; input.checked=value;
      input.addEventListener("change", () => { parent[key] = input.checked; refreshRawEditor(); refreshPreview(); });
    } else {
      input.type="text"; input.value=value ?? "";
      input.addEventListener("input", () => { parent[key] = input.value; refreshRawEditor(); refreshPreview(); });
    }
    row.append(label, input); container.appendChild(row); return;
  }

  if (Array.isArray(value)) {
    const header = document.createElement("h3");
    header.textContent = `${key} (array${value.length && typeof value[0] === "object" ? " of objects" : ""})`;
    container.appendChild(header);

    // primitives
    if (!value.length || isPrimitive(value[0])) {
      const list = document.createElement("div"); list.className = "grid";
      value.forEach((item, idx) => {
        const row = document.createElement("div"); row.className = "kv";
        const label = document.createElement("label"); label.textContent = `${key}[${idx}]`;
        const input = document.createElement("input");
        if (typeof item === "number") { input.type="number"; input.step="any"; input.value=String(item);
          input.addEventListener("input", () => { parent[key][idx] = Number(input.value || 0); refreshRawEditor(); refreshPreview(); });
        } else { input.type="text"; input.value=String(item ?? "");
          input.addEventListener("input", () => { parent[key][idx] = input.value; refreshRawEditor(); refreshPreview(); });
        }
        const del = document.createElement("button"); del.type="button"; del.className="secondary"; del.textContent="Remove";
        del.addEventListener("click", () => { parent[key].splice(idx,1); renderSimpleEditor(dataModel); refreshRawEditor(); refreshPreview(); });
        row.append(label,input,del); list.appendChild(row);
      });
      const add = document.createElement("button"); add.type="button"; add.textContent=`+ Add ${key} item`;
      add.addEventListener("click", () => { parent[key].push(""); renderSimpleEditor(dataModel); refreshRawEditor(); refreshPreview(); });
      container.append(list,add); return;
    }

    // array of objects
    const wrapper = document.createElement("div"); wrapper.className = "grid";
    const keysUnion = Array.from(new Set(value.flatMap(o => Object.keys(o || {}))));
    if (keysUnion.length === 0) keysUnion.push("name","value");

    value.forEach((rowObj, idx) => {
      const card = document.createElement("div"); card.className = "card";
      const title = document.createElement("div"); title.className="inline";
      const badge = document.createElement("span"); badge.className="badge"; badge.textContent=`${key}[${idx}]`;
      const remove = document.createElement("button"); remove.type="button"; remove.className="secondary"; remove.textContent="Remove";
      remove.addEventListener("click", () => { parent[key].splice(idx,1); renderSimpleEditor(dataModel); refreshRawEditor(); refreshPreview(); });
      title.append(badge, remove); card.appendChild(title);

      // known keys first
      keysUnion.forEach((col) => {
        if (!(col in rowObj)) return;
        const sub = document.createElement("div");
        renderNode(col, rowObj[col], rowObj, [...path, key, String(idx), col], sub);
        card.appendChild(sub);
      });
      // any extra keys
      Object.keys(rowObj).forEach((col) => {
        if (keysUnion.includes(col)) return;
        const sub = document.createElement("div");
        renderNode(col, rowObj[col], rowObj, [...path, key, String(idx), col], sub);
        card.appendChild(sub);
      });

      addFieldControls(rowObj, card);
      wrapper.appendChild(card);
    });

    const addRow = document.createElement("button"); addRow.type="button"; addRow.textContent=`+ Add ${key} row`;
    addRow.addEventListener("click", () => {
      const newRow = {}; keysUnion.forEach(col => newRow[col] = "");
      parent[key].push(newRow);
      renderSimpleEditor(dataModel); refreshRawEditor(); refreshPreview();
    });

    container.append(wrapper, addRow); return;
  }

  if (isPlainObject(value)) {
    const section = document.createElement("div");
    const header = document.createElement("h3"); header.textContent = `${key} (object)`;
    section.appendChild(header);
    Object.keys(value).forEach((k2) => {
      const sub = document.createElement("div");
      renderNode(k2, value[k2], value, [...path,key,k2], sub);
      section.appendChild(sub);
    });
    addFieldControls(value, section);
    container.appendChild(section);
  }
}

function renderSimpleEditor(obj) {
  kvContainer.innerHTML = "";
  tablesContainer.innerHTML = "";
  Object.keys(obj).forEach((k) => {
    const v = obj[k];
    (isPrimitive(v) || isPlainObject(v))
      ? renderNode(k, v, obj, [k], kvContainer)
      : renderNode(k, v, obj, [k], tablesContainer);
  });
}

// ---------- Body Sections ----------
function getByPath(root, pathStr) {
  const parts = pathStr.split('.').flatMap(p => {
    const m = [...p.matchAll(/([^\[\]]+)|\[(\d+)\]/g)];
    return m.map(g => g[1] ?? Number(g[2]));
  });
  let cur = root;
  for (const key of parts) { if (cur == null) return ""; cur = cur[key]; }
  return (cur === undefined || cur === null) ? "" : cur;
}
function renderTpl(s, root) {
  return s.replace(/\{([^}]+)\}/g, (_, expr) => {
    try { return String(getByPath(root, expr.trim())); } catch { return ""; }
  });
}
function initSections() {
  const body = isPlainObject(dataModel.body) ? dataModel.body : (dataModel.body = {});
  sectionsUI = Object.keys(body).map(k => ({ key: k, include: true }));
}
function renderSections() {
  const body = dataModel.body || {};
  sectionsList.innerHTML = "";
  sectionsUI.forEach((row, idx) => {
    const li = document.createElement("li"); li.className="section"; li.draggable=true;

    const head = document.createElement("div"); head.className="section-header";
    const handle = document.createElement("span"); handle.className="section-handle"; handle.textContent="↕";
    const include = document.createElement("input"); include.type="checkbox"; include.checked=row.include;
    include.addEventListener("change",()=>{ row.include = include.checked; refreshPreview(); });

    const nameInput = document.createElement("input"); nameInput.className="section-name";
    nameInput.type="text"; nameInput.value=row.key;
    nameInput.addEventListener("change", () => {
      const old=row.key, neu=nameInput.value.trim(); if(!neu){ nameInput.value=old; return; }
      if (old!==neu){ body[neu]=body[old]; delete body[old]; row.key=neu; refreshRawEditor(); refreshPreview(); }
    });

    const ctrls = document.createElement("div"); ctrls.className="section-controls";
    const up=document.createElement("button"); up.className="secondary"; up.textContent="↑";
    const down=document.createElement("button"); down.className="secondary"; down.textContent="↓";
    const del=document.createElement("button"); del.className="secondary"; del.textContent="Remove";
    up.addEventListener("click",()=>{ if(idx>0){ const t=sectionsUI[idx-1]; sectionsUI[idx-1]=sectionsUI[idx]; sectionsUI[idx]=t; renderSections(); refreshPreview(); }});
    down.addEventListener("click",()=>{ if(idx<sectionsUI.length-1){ const t=sectionsUI[idx+1]; sectionsUI[idx+1]=sectionsUI[idx]; sectionsUI[idx]=t; renderSections(); refreshPreview(); }});
    del.addEventListener("click",()=>{ delete body[row.key]; sectionsUI.splice(idx,1); renderSections(); refreshRawEditor(); refreshPreview(); });
    ctrls.append(up,down,del);

    head.append(handle, include, nameInput, ctrls);
    li.appendChild(head);

    const ta = document.createElement("textarea");
    ta.value = String(body[row.key] ?? "");
    ta.addEventListener("input",()=>{ body[row.key]=ta.value; refreshRawEditor(); refreshPreview(); });
    li.appendChild(ta);

    li.addEventListener("dragstart",(e)=>{ e.dataTransfer.setData("text/plain", idx.toString()); });
    li.addEventListener("dragover",(e)=>{ e.preventDefault(); });
    li.addEventListener("drop",(e)=>{ e.preventDefault(); const from=Number(e.dataTransfer.getData("text/plain")); if(from===idx) return;
      const moved=sectionsUI.splice(from,1)[0]; sectionsUI.splice(idx,0,moved); renderSections(); refreshPreview(); });

    sectionsList.appendChild(li);
  });
}
addSectionBtn.addEventListener("click", () => {
  const body = dataModel.body || (dataModel.body = {});
  let base="section",i=1,name=`${base}${i}`;
  while(Object.prototype.hasOwnProperty.call(body,name)){ i+=1; name=`${base}${i}`; }
  body[name] = "";
  sectionsUI.push({ key:name, include:true });
  renderSections(); refreshRawEditor(); refreshPreview();
});
function buildMergedBody() {
  const body = dataModel.body || {};
  const lines = [];
  sectionsUI.forEach(row=>{
    if(!row.include) return;
    const raw = String(body[row.key] ?? "");
    const rendered = renderTpl(raw, dataModel);
    if (rendered.trim()) lines.push(rendered);
  });
  return lines.join("\n\n");
}
function refreshPreview() { bodyPreview.value = buildMergedBody(); }

// ---------- Loader ----------
async function loadAll() {
  loadStatus.textContent = "loading...";
  errorMsg.textContent = "";
  try {
    [templateArrayBuffer, dataModel] = await Promise.all([
      fetchAB(PATHS.DOCX),
      fetchJSON(jsonChoice.value)
    ]);
  } catch (e) {
    // inline fallback for JSON
    try {
      templateArrayBuffer = await fetchAB(PATHS.DOCX);
      const node = document.getElementById("templateInline");
      dataModel = JSON.parse(node.textContent);
      errorMsg.textContent = "Used inline JSON fallback because file fetch failed.";
    } catch (e2) {
      loadStatus.textContent = "error";
      errorMsg.textContent = (e && e.message) ? e.message : "Failed to fetch.";
      return;
    }
  }
  renderSimpleEditor(dataModel);
  refreshRawEditor();
  initSections();
  renderSections();
  refreshPreview();
  if (!outputName.value) {
    outputName.value = ((dataModel.title || "output")+"").replace(/\s+/g,"_") + ".docx";
  }
  loadStatus.textContent = "loaded ✓";
}

// ---------- Generate ----------
function generateDocx() {
  if (!templateArrayBuffer) { errorMsg.textContent="Template not loaded."; return; }
  errorMsg.textContent=""; genStatus.textContent="rendering...";
  const content = buildMergedBody();
  const zip = new PizZip(templateArrayBuffer);
  const doc = new window.docxtemplater(zip, { paragraphLoop:true, linebreaks:true });
  try {
    doc.setData({ ...dataModel, content });
    doc.render();
  } catch (e) { genStatus.textContent="error"; errorMsg.textContent=e?.message||"Docxtemplater render error"; return; }
  const out = doc.getZip().generate({ type:"blob", mimeType:"application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  const name = (outputName.value || "output.docx").trim();
  window.saveAs(out, name.endsWith(".docx") ? name : name + ".docx");
  genStatus.textContent="done ✓";
}

// ---------- Events ----------
reloadBtn.addEventListener("click", loadAll);
jsonChoice.addEventListener("change", loadAll);
btnGenerate.addEventListener("click", generateDocx);

toggleRaw.addEventListener("change", () => {
  const showRaw = toggleRaw.checked;
  rawEditor.style.display = showRaw ? "block" : "none";
  simpleEditor.style.display = showRaw ? "none" : "block";
  if (showRaw) refreshRawEditor();
});
applyRawBtn.addEventListener("click", () => {
  try {
    const parsed = JSON.parse(rawJsonTA.value);
    if (!isPlainObject(parsed)) throw new Error("Root must be a JSON object.");
    dataModel = parsed;
    renderSimpleEditor(dataModel);
    initSections();
    renderSections();
    refreshRawEditor();
    refreshPreview();
  } catch (e) { alert("Invalid JSON: " + (e?.message || e)); }
});

// ---------- Init ----------
loadAll();
