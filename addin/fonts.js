/* Continu Studio – font check.
   1. Embedding: reads the whole file once (getFileAsync) and checks whether "Embed fonts in the file" is on.
   2. Unrecognised fonts: lists any font that isn't Effra (or a symbol font such as Wingdings), with the slides it's on.
      The whole deck is checked when the panel opens; after that the slide you're working on is re-checked as you click
      around, so pasted text or slides are picked up.
   3. Convert to Effra: rewrites the fonts on the affected slides and puts them back with the Continu theme.
      Bold or heavier → Effra Medium, light or thin → Effra Light, everything else → Effra Regular. */
(function () {
  const LIGHT = "EFFRA-LIGHT", REG = "Effra", MED = "EFFRA-MEDIUM";
  const PAN = { [LIGHT]: "02000306080000020004", [REG]: "02000506080000020004", [MED]: "02000506080000020004" };
  const SYMBOL = /^(wingdings( 2| 3)?|symbol|webdings|marlett|segoe ui symbol)$/i;
  const isEffra = tf => /^effra/i.test(tf);
  const isTheme = tf => !tf || tf.charAt(0) === "+";
  const unrecognised = tf => !isTheme(tf) && !isEffra(tf) && !SYMBOL.test(tf);

  function canon(tf) {
    const t = tf.toLowerCase().replace(/-/g, " ");
    if (/light|thin|hairline/.test(t)) return LIGHT;
    if (/medium|semi ?bold|demi|bold|black|heavy/.test(t)) return MED;
    return REG;
  }
  const fontEl = (tag, face) => `<a:${tag} typeface="${face}" panose="${PAN[face]}" pitchFamily="2" charset="0"/>`;

  // ---------- finding fonts ----------
  function fontsInXml(xml, into) {
    const re = /<a:(latin|buFont) typeface="([^"]*)"/g; let m;
    while ((m = re.exec(xml))) if (unrecognised(m[2])) into.add(m[2]);
    return into;
  }
  function themeFonts(xml, into) {
    const re = /<a:(majorFont|minorFont)><a:latin typeface="([^"]*)"/g; let m;
    while ((m = re.exec(xml))) if (unrecognised(m[2])) into.add(m[2]);
    return into;
  }
  const dirOf = p => p.slice(0, p.lastIndexOf("/") + 1);
  const relsOf = p => dirOf(p) + "_rels/" + p.slice(p.lastIndexOf("/") + 1) + ".rels";
  function resolve(base, target) {
    if (target.charAt(0) === "/") return target.slice(1);
    const parts = (dirOf(base) + target).split("/"), out = [];
    parts.forEach(x => { if (x === "..") out.pop(); else if (x !== ".") out.push(x); });
    return out.join("/");
  }
  async function readText(zip, path) { const f = zip.file(path); return f ? f.async("string") : ""; }
  async function rels(zip, path) {
    const xml = await readText(zip, relsOf(path)), list = [];
    const re = /<Relationship\b[^>]*>/g; let m;
    while ((m = re.exec(xml))) {
      const tag = m[0], g = k => (tag.match(new RegExp(k + '="([^"]*)"')) || [])[1] || "";
      if (g("TargetMode") === "External") continue;
      list.push({ id: g("Id"), type: g("Type").split("/").pop(), target: resolve(path, g("Target")) });
    }
    return list;
  }
  // fonts used by one slide: its own text, its charts and SmartArt, and the theme it sits on
  async function slideFonts(zip, slidePath, cache) {
    const found = new Set();
    fontsInXml(await readText(zip, slidePath), found);
    for (const r of await rels(zip, slidePath)) {
      if (/^(chart|diagramData|diagramDrawing)$/.test(r.type) || /drawing$/i.test(r.type)) fontsInXml(await readText(zip, r.target), found);
      if (r.type === "slideLayout") {
        const key = r.target;
        if (!cache[key]) {
          const t = new Set();
          for (const lr of await rels(zip, r.target)) if (lr.type === "slideMaster")
            for (const mr of await rels(zip, lr.target)) if (mr.type === "theme") themeFonts(await readText(zip, mr.target), t);
          cache[key] = t;
        }
        cache[key].forEach(f => found.add(f));
      }
    }
    return found;
  }

  // the whole file, as it is now, in 4 MB slices
  function fileBytes() {
    return new Promise((res, rej) => {
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, r => {
        if (r.status !== Office.AsyncResultStatus.Succeeded) return rej(r.error);
        const file = r.value, n = file.sliceCount, parts = [];
        const next = i => file.getSliceAsync(i, s => {
          if (s.status !== Office.AsyncResultStatus.Succeeded) { file.closeAsync(); return rej(s.error); }
          parts[i] = s.value.data;
          if (i + 1 < n) return next(i + 1);
          file.closeAsync();
          let len = 0; parts.forEach(p => len += p.length);
          const u = new Uint8Array(len); let o = 0; parts.forEach(p => { u.set(p, o); o += p.length; });
          res(u);
        });
        next(0);
      });
    });
  }

  // whole-deck check: { embedded, perSlide: [Set of fonts for slide 1, 2, …] }
  async function scanFile() {
    const zip = await JSZip.loadAsync(await fileBytes());
    const pres = await readText(zip, "ppt/presentation.xml");
    const embedded = /\bembedTrueTypeFonts="(1|true)"/.test(pres);
    const subset = /\bsaveSubsetFonts="(1|true)"/.test(pres);
    const presRels = await rels(zip, "ppt/presentation.xml");
    const order = []; const re = /<p:sldId\b[^>]*r:id="([^"]*)"/g; let m;
    while ((m = re.exec(pres))) { const r = presRels.find(x => x.id === m[1]); if (r) order.push(r.target); }
    const cache = {}, perSlide = [];
    for (const p of order) perSlide.push(await slideFonts(zip, p, cache));
    return { embedded, subset, perSlide };
  }
  // one slide, exported on its own
  async function scanSlide(b64) {
    const zip = await JSZip.loadAsync(b64, { base64: true });
    const p = Object.keys(zip.files).find(f => /^ppt\/slides\/slide\d+\.xml$/.test(f));
    return p ? slideFonts(zip, p, {}) : new Set();
  }

  // ---------- converting ----------
  function fixRun(attrs, inner) {
    const lat = inner.match(/<a:latin typeface="([^"]*)"[^>]*?(?:\/>|>\s*<\/a:latin>)/);
    const bold = /\sb="(1|true)"/.test(attrs);
    if (lat && !isTheme(lat[1])) {
      let face = isEffra(lat[1]) ? canon(lat[1].replace(/^effra/i, "")) : canon(lat[1]);
      if (bold && face !== LIGHT) face = MED;
      inner = inner.replace(lat[0], fontEl("latin", face));
    } else if (bold) {
      inner = fontEl("latin", MED) + inner;
    }
    if (bold) attrs = attrs.replace(/\sb="(1|true)"/, "");
    inner = inner.replace(/<a:(ea|cs|sym) typeface="([^"]*)"[^>]*?(?:\/>|>\s*<\/a:\1>)/g,
      (all, tag, tf) => isTheme(tf) || (tag !== "sym" && isEffra(tf)) ? all : "");
    return [attrs, inner];
  }
  function toEffra(xml) {
    const re = /<a:(rPr|defRPr|endParaRPr)(\s[^>]*?)?(\/?)>/g; let out = "", last = 0, m;
    while ((m = re.exec(xml))) {
      const tag = m[1], attrs = m[2] || "", selfClose = m[3] === "/";
      out += xml.slice(last, m.index);
      if (selfClose) {
        if (/\sb="(1|true)"/.test(attrs)) { const [a] = fixRun(attrs, ""); out += `<a:${tag}${a}>${fontEl("latin", MED)}</a:${tag}>`; }
        else out += m[0];
        last = re.lastIndex; continue;
      }
      const close = `</a:${tag}>`, end = xml.indexOf(close, re.lastIndex);
      if (end < 0) { out += m[0]; last = re.lastIndex; continue; }
      const [a, inner] = fixRun(attrs, xml.slice(re.lastIndex, end));
      out += `<a:${tag}${a}>${inner}${close}`;
      last = end + close.length; re.lastIndex = last;
    }
    out += xml.slice(last);
    return out.replace(/<a:buFont typeface="([^"]*)"[^>]*\/>/g, (all, tf) => unrecognised(tf) ? fontEl("buFont", REG) : all);
  }
  async function convertSlide(b64) {
    const zip = await JSZip.loadAsync(b64, { base64: true });
    for (const p of Object.keys(zip.files)) {
      if (!/^ppt\/(slides\/slide\d+|charts\/chart\d+|diagrams\/(data|drawing)\d+)\.xml$/.test(p)) continue;
      const xml = await zip.file(p).async("string"), fixed = toEffra(xml);
      if (fixed !== xml) zip.file(p, fixed);
    }
    return zip.generateAsync({ type: "base64" });
  }

  window.continuFonts = { scanFile, scanSlide, convertSlide, toEffra, unrecognised };
})();
