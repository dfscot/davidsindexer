(function(){
  const statusEl = document.getElementById('status');
  const progressEl = document.getElementById('progress');
  const ids = ['excludeWords','includeWords','includePhrases'];
  function setStatus(msg){ statusEl.textContent = msg; }
  function setProgress(msg){ progressEl.textContent = msg || ''; }
  function getSettings(){ return Office.context.document.settings; }
  function loadCached(){
    const s = getSettings();
    document.getElementById('excludeWords').value = (s.get('excludeWords')||[]).join('
');
    document.getElementById('includeWords').value = (s.get('includeWords')||[]).join('
');
    document.getElementById('includePhrases').value = (s.get('includePhrases')||[]).join('
');
  }
  function parseLines(id){ return document.getElementById(id).value.split(/?
/).map(x=>x.trim()).filter(Boolean); }
  function saveCached(){
    const s = getSettings();
    s.set('excludeWords', parseLines('excludeWords'));
    s.set('includeWords', parseLines('includeWords'));
    s.set('includePhrases', parseLines('includePhrases'));
    s.saveAsync(function(asyncResult){
      setStatus(asyncResult.status === Office.AsyncResultStatus.Succeeded ? 'Cached settings saved.' : 'Failed to save cached settings.');
    });
  }
  async function reloadSample(){
    setProgress('Reloading sample config...');
    try{
      const res = await fetch('../../config.json');
      const cfg = await res.json();
      document.getElementById('excludeWords').value = (cfg.excludeWords||[]).join('
');
      document.getElementById('includeWords').value = (cfg.includeWords||[]).join('
');
      document.getElementById('includePhrases').value = (cfg.includePhrases||[]).join('
');
      saveCached();
      setStatus('Sample config loaded into cache.');
    }catch(e){
      setStatus('Unable to load sample config.');
    }finally{ setProgress(''); }
  }
  Office.onReady(function(){
    loadCached();
    document.getElementById('saveBtn').addEventListener('click', saveCached);
    document.getElementById('reloadBtn').addEventListener('click', reloadSample);
  });
})();
