/* Continu Studio – forgiving search.
   Understands everyday words (cut → crop), small typos (hperlink → hyperlink),
   plurals and -ing/-ed endings, and ranks the best matches first.
   To teach it a new word, add it to the right group in SYNONYMS. */
(function () {
  const SYNONYMS = [
    ["crop", "cut", "trim", "chop", "clip", "cropping", "cut off", "cut out"],
    ["picture", "image", "photo", "photograph", "pic", "pics", "img", "graphic", "jpeg", "jpg", "png"],
    ["link", "hyperlink", "url", "web address", "website", "web link", "weblink", "href", "clickable"],
    ["colour", "color", "colours", "colors", "shade", "tint"],
    ["green", "lime", "brand colour", "highlight", "accent"],
    ["resize", "size", "shrink", "smaller", "bigger", "enlarge", "scale", "stretch", "squash", "fit"],
    ["fill", "fit", "cover", "frame", "placeholder", "box"],
    ["replace", "swap", "change", "switch", "substitute", "update"],
    ["add", "insert", "new", "create", "put", "make", "place", "include"],
    ["delete", "remove", "get rid", "take out", "erase", "clear"],
    ["move", "reorder", "rearrange", "order", "drag", "shift", "sort", "reposition"],
    ["drag", "drop", "drag and drop", "pull in"],
    ["folder", "finder", "explorer", "file", "files", "desktop", "computer"],
    ["text", "words", "copy", "writing", "wording", "type", "font", "letters"],
    ["paste", "copy over", "plain text", "keep text only", "formatting", "unformatted"],
    ["format", "formatting", "style", "styles", "look"],
    ["painter", "brush", "match", "same style", "copy style"],
    ["overflow", "too long", "too much", "spill", "doesn't fit", "doesnt fit", "running over", "autofit"],
    ["slide", "page", "slides", "pages"],
    ["layout", "design", "columns", "column", "arrangement", "template layout"],
    ["background", "bg", "backdrop", "dark", "light", "navy", "white"],
    ["reset", "undo", "undo changes", "start again", "fix", "restore", "broken", "messed", "messed up", "moved", "back to normal", "gone wrong", "wonky"],
    ["table", "grid", "tables", "cells", "matrix"],
    ["row", "rows", "column", "columns", "cell"],
    ["chart", "graph", "charts", "graphs", "bar", "pie", "line chart", "plot", "diagram"],
    ["data", "numbers", "figures", "values", "stats"],
    ["excel", "spreadsheet", "sheet", "xls", "xlsx", "workbook"],
    ["save", "export", "store", "keep", "download"],
    ["template", "potx", "master", "reuse"],
    ["button", "shape", "icon", "object"],
    ["help", "how", "guide"]
  ];
  const STOP = new Set("a an the to of in on at for by with and or i my me we our it its this that is are be can do does how what where why when you your into from up out please make want need get just some can t cant wont isnt".split(" "));

  const norm = s => s.toLowerCase().replace(/[’']/g, "").replace(/[^a-z0-9 ]+/g, " ").replace(/\s+/g, " ").trim();
  const stem = w => {
    if (w.length <= 4) return w;
    if (/(ches|shes|xes|sses|zes)$/.test(w)) return w.slice(0, -2);
    return w.replace(/(ings|ing|ies|ied|ed|s)$/, m => (m === "ies" || m === "ied") ? "y" : "");
  };

  // word → set of related words
  const REL = new Map();
  SYNONYMS.forEach(g => {
    const words = g.map(norm);
    words.forEach(w => { const k = stem(w); if (!REL.has(k)) REL.set(k, new Set()); words.forEach(x => REL.get(k).add(x)); });
  });

  function dist(a, b, max) { // Damerau-Levenshtein, early exit
    if (Math.abs(a.length - b.length) > max) return max + 1;
    const d = []; for (let i = 0; i <= a.length; i++) { d[i] = [i]; }
    for (let j = 0; j <= b.length; j++) d[0][j] = j;
    for (let i = 1; i <= a.length; i++) {
      let best = 99;
      for (let j = 1; j <= b.length; j++) {
        const c = a[i - 1] === b[j - 1] ? 0 : 1;
        d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + c);
        if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) d[i][j] = Math.min(d[i][j], d[i - 2][j - 2] + 1);
        best = Math.min(best, d[i][j]);
      }
      if (best > max) return max + 1;
    }
    return d[a.length][b.length];
  }
  const allowed = w => w.length <= 3 ? 0 : w.length <= 5 ? 1 : 2;

  let VOCAB = null; // every known word, for typo correction
  function vocab(items) {
    if (VOCAB) return VOCAB;
    const v = new Set();
    items.forEach(it => norm(it._text).split(" ").forEach(w => w.length > 2 && v.add(w)));
    REL.forEach((set, k) => { v.add(k); set.forEach(x => x.split(" ").forEach(y => y.length > 2 && v.add(y))); });
    return (VOCAB = [...v]);
  }

  // expand one query word into weighted alternatives
  function expand(word, items) {
    const out = new Map([[word, 1]]);
    let w = word;
    if (w.length > 3 && !REL.has(stem(w)) && !vocab(items).includes(w)) { // likely a typo
      let best = null, bd = 9;
      for (const v of vocab(items)) { const dd = dist(w, v, allowed(w)); if (dd < bd) { bd = dd; best = v; } }
      if (best && bd <= allowed(w)) { out.set(best, 0.85); w = best; }
    }
    const rel = REL.get(stem(w));
    if (rel) rel.forEach(x => { if (!out.has(x)) out.set(x, 0.7); });
    return out;
  }

  function hit(field, term) {
    if (term.includes(" ")) return field.includes(term);
    const words = field.split(" "), t = stem(term);
    return words.some(x => x === term || stem(x) === t || (term.length >= 4 && x.startsWith(term)));
  }

  // items: [{title, area, keywords, steps}] → ranked [{item, score, all}]
  window.smartSearch = function (query, items) {
    items.forEach(it => { if (!it._text) { it._title = norm(it.title); it._text = norm([it.title, it.area, it.keywords || "", (it.steps || []).join(" ")].join(" ")); } });
    const q = norm(query);
    if (!q) return [];
    let words = q.split(" ").filter(w => !STOP.has(w));
    if (!words.length) words = q.split(" ");
    // also catch two-word phrases like "web address" or "too long"
    const phrases = []; for (let i = 0; i < words.length - 1; i++) { const p = words[i] + " " + words[i + 1]; if (REL.has(stem(p)) || REL.has(p)) phrases.push(p); }
    const terms = words.map(w => expand(w, items));
    phrases.forEach(p => terms.push(expand(p, items)));
    const res = items.map(it => {
      let score = 0, matched = 0;
      terms.forEach(alts => {
        let best = 0;
        alts.forEach((wt, term) => {
          if (hit(it._title, term)) best = Math.max(best, 3 * wt);
          else if (hit(it._text, term)) best = Math.max(best, 1 * wt);
        });
        if (best) { matched++; score += best; }
      });
      return { item: it, score, matched, all: matched >= words.length };
    }).filter(r => r.score > 0);
    res.sort((a, b) => (b.matched - a.matched) || (b.score - a.score));
    return res;
  };
})();
