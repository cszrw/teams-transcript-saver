(async function () {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  const list = document.getElementById('list');
  const status = document.getElementById('status');
  const items = tab ? await chrome.runtime.sendMessage({ type: 'list', tabId: tab.id }) : [];

  if (!items || !items.length) { document.getElementById('empty').hidden = false; return; }

  for (const t of items.slice().reverse()) {
    const li = document.createElement('li');
    const info = document.createElement('div');
    const name = document.createElement('div');
    name.className = 'name';
    name.textContent = t.name;
    const meta = document.createElement('div');
    meta.className = 'meta';
    meta.textContent = [t.language, t.transcriptType, 'captured ' + new Date(t.capturedAt).toLocaleTimeString()].filter(Boolean).join(' · ');
    info.append(name, meta);
    const btn = document.createElement('button');
    btn.textContent = 'Save';
    btn.addEventListener('click', async () => {
      btn.disabled = true;
      status.className = '';
      status.textContent = 'Saving…';
      const r = await chrome.runtime.sendMessage({ type: 'save', tabId: tab.id, id: t.id });
      status.className = r.ok ? '' : 'err';
      status.textContent = r.message;
      btn.disabled = false;
    });
    li.append(info, btn);
    list.append(li);
  }
})();
