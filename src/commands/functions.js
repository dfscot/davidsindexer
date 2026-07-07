/* DavidsIndexer command scaffold.
   This file contains the core matching logic and placeholder Office.js command handlers.
   It is intended as a buildable starter, not a guaranteed drop-in production implementation. */

function normalizeText(value, caseSensitive) {
  const v = caseSensitive ? value : value.toLowerCase();
  return v.replace(/[‘’]/g, "'").replace(/[“”]/g, '"');
}

function tokenizeWithPositions(text) {
  const regex = /[\p{L}\p{N}'][\p{L}\p{N}'-]*/gu;
  const out = [];
  let m;
  while ((m = regex.exec(text)) !== null) {
    out.push({ text: m[0], start: m.index, end: m.index + m[0].length });
  }
  return out;
}

function findPhraseMatches(text, phrases, caseSensitive) {
  const haystack = normalizeText(text, caseSensitive);
  const matches = [];
  for (const phrase of phrases) {
    const p = normalizeText(phrase, caseSensitive);
    if (!p) continue;
    let idx = 0;
    while (true) {
      idx = haystack.indexOf(p, idx);
      if (idx === -1) break;
      matches.push({ phrase, start: idx, end: idx + p.length, length: p.length });
      idx += 1;
    }
  }
  matches.sort((a,b)=> a.start - b.start || b.length - a.length);
  const accepted = [];
  let cursor = -1;
  for (const m of matches) {
    if (m.start >= cursor) {
      accepted.push(m);
      cursor = m.end;
    }
  }
  return accepted;
}

function buildIndexPlan(text, config) {
  const caseSensitive = !!config.caseSensitive;
  const exclude = new Set((config.excludeWords || []).map(w => normalizeText(w, caseSensitive)));
  const includeWords = new Set((config.includeWords || []).map(w => normalizeText(w, caseSensitive)));
  const phraseMatches = findPhraseMatches(text, config.includePhrases || [], caseSensitive);
  const covered = [];
  for (const m of phraseMatches) covered.push([m.start, m.end]);
  function isCovered(start, end) {
    return covered.some(([s,e]) => start < e && end > s);
  }
  const tokens = tokenizeWithPositions(text);
  const wordEntries = [];
  for (const t of tokens) {
    if (isCovered(t.start, t.end)) continue;
    const n = normalizeText(t.text, caseSensitive);
    if (includeWords.has(n) || !exclude.has(n)) {
      wordEntries.push({ term: t.text, start: t.start, end: t.end });
    }
  }
  return { phraseMatches, wordEntries };
}

async function showTaskpane(event) {
  try { await Office.addin.showAsTaskpane(); } catch(e) {}
  event.completed();
}

async function generateIndex(event) {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load('text');
      await context.sync();
      const s = Office.context.document.settings;
      const config = {
        caseSensitive: s.get('caseSensitive') || false,
        excludeWords: s.get('excludeWords') || [],
        includeWords: s.get('includeWords') || [],
        includePhrases: s.get('includePhrases') || []
      };
      const plan = buildIndexPlan(body.text || '', config);
      const info = 'DavidsIndexer scaffold generated a plan with ' + plan.phraseMatches.length + ' phrase matches and ' + plan.wordEntries.length + ' word entries.';
      body.insertParagraph('', Word.InsertLocation.end);
      body.insertParagraph('DavidsIndexer status: ' + info, Word.InsertLocation.end);
      await context.sync();
    });
  } catch(e) {
  }
  event.completed();
}

async function removeXeCodes(event) {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const results = body.search('XE', {matchCase:false, matchWholeWord:false});
      results.load('items');
      await context.sync();
      for (const r of results.items) {
        try { r.clear(); } catch(e) {}
      }
      await context.sync();
    });
  } catch(e) {
  }
  event.completed();
}

Office.actions.associate('showTaskpane', showTaskpane);
Office.actions.associate('generateIndex', generateIndex);
Office.actions.associate('removeXeCodes', removeXeCodes);
