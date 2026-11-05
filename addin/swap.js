/* Continu Studio – swap a picture but keep its shape.
   Exports the slide, replaces just that picture's image inside it, crops the new
   image to fill the same frame without stretching, and puts the slide back. */
(function () {
  const NS = {
    p: "http://schemas.openxmlformats.org/presentationml/2006/main",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  };

  // slideB64: one-slide .pptx from exportAsBase64; match: {id, name}; img: {bytes, ext, w, h}; frame: {w, h} in points (fallback)
  async function swapPictureInSlide(slideB64, match, img, frame) {
    const zip = await JSZip.loadAsync(slideB64, { base64: true });
    const slidePath = Object.keys(zip.files).find(f => /^ppt\/slides\/slide\d+\.xml$/.test(f));
    if (!slidePath) throw new Error("couldn't read the slide");
    const relsPath = slidePath.replace("slides/", "slides/_rels/") + ".rels";
    const doc = new DOMParser().parseFromString(await zip.file(slidePath).async("string"), "application/xml");
    const pics = [...doc.getElementsByTagNameNS(NS.p, "pic")];
    const idOf = pic => { const c = pic.getElementsByTagNameNS(NS.p, "cNvPr")[0]; return c ? { id: c.getAttribute("id"), name: c.getAttribute("name") } : {}; };
    const digits = (String(match.id || "").match(/\d+/) || [""])[0];
    let pic = pics.find(p => idOf(p).id === digits) || null;
    if (!pic) { const byName = pics.filter(p => idOf(p).name === match.name); if (byName.length === 1) pic = byName[0]; }
    if (!pic && pics.length === 1) pic = pics[0];
    if (!pic) return null; // not a picture we can swap

    const blipFill = pic.getElementsByTagNameNS(NS.p, "blipFill")[0];
    const blip = blipFill.getElementsByTagNameNS(NS.a, "blip")[0];

    // new media file + relationship
    const n = Date.now().toString(36);
    const mediaName = `continu-swap-${n}.${img.ext}`;
    zip.file(`ppt/media/${mediaName}`, img.bytes);
    const rels = new DOMParser().parseFromString(await zip.file(relsPath).async("string"), "application/xml");
    const relRoot = rels.documentElement, RNS = relRoot.namespaceURI;
    const rid = "rIdSwap" + n;
    const rel = rels.createElementNS(RNS, "Relationship");
    rel.setAttribute("Id", rid);
    rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
    rel.setAttribute("Target", `../media/${mediaName}`);
    relRoot.appendChild(rel);
    zip.file(relsPath, new XMLSerializer().serializeToString(rels));
    blip.setAttributeNS(NS.r, "r:embed", rid);
    // drop any linked or svg versions of the old picture
    blip.removeAttributeNS(NS.r, "link");
    [...blip.getElementsByTagNameNS(NS.a, "extLst")].forEach(e => e.parentNode.removeChild(e));

    // content type for the new extension
    const ctPath = "[Content_Types].xml";
    const ct = new DOMParser().parseFromString(await zip.file(ctPath).async("string"), "application/xml");
    const types = ct.documentElement, TNS = types.namespaceURI;
    if (![...types.getElementsByTagNameNS(TNS, "Default")].some(d => d.getAttribute("Extension").toLowerCase() === img.ext)) {
      const d = ct.createElementNS(TNS, "Default");
      d.setAttribute("Extension", img.ext);
      d.setAttribute("ContentType", img.ext === "png" ? "image/png" : "image/jpeg");
      types.insertBefore(d, types.firstChild);
      zip.file(ctPath, new XMLSerializer().serializeToString(ct));
    }

    // crop the new image to fill the frame (like "Fill"), centred
    let fw = frame && frame.w, fh = frame && frame.h;
    const ext = pic.getElementsByTagNameNS(NS.p, "spPr")[0]?.getElementsByTagNameNS(NS.a, "ext")[0];
    if (ext) { fw = +ext.getAttribute("cx"); fh = +ext.getAttribute("cy"); }
    let l = 0, t = 0, r = 0, b = 0;
    if (fw && fh && img.w && img.h) {
      const A = fw / fh, B = img.w / img.h;
      if (B > A) { const c = (1 - A / B) / 2; l = r = Math.round(c * 100000); }
      else if (B < A) { const c = (1 - B / A) / 2; t = b = Math.round(c * 100000); }
    }
    [...blipFill.getElementsByTagNameNS(NS.a, "srcRect")].forEach(e => e.parentNode.removeChild(e));
    const src = doc.createElementNS(NS.a, "a:srcRect");
    if (l) { src.setAttribute("l", l); src.setAttribute("r", r); }
    if (t) { src.setAttribute("t", t); src.setAttribute("b", b); }
    blip.parentNode.insertBefore(src, blip.nextSibling);
    let stretch = blipFill.getElementsByTagNameNS(NS.a, "stretch")[0];
    [...blipFill.getElementsByTagNameNS(NS.a, "tile")].forEach(e => e.parentNode.removeChild(e));
    if (!stretch) { stretch = doc.createElementNS(NS.a, "a:stretch"); stretch.appendChild(doc.createElementNS(NS.a, "a:fillRect")); blipFill.appendChild(stretch); }

    zip.file(slidePath, new XMLSerializer().serializeToString(doc));
    return await zip.generateAsync({ type: "base64" });
  }
  window.swapPictureInSlide = swapPictureInSlide;
})();

/* ---------- Edit selected items by rebuilding the slide ---------- */
(function () {
  const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const kids = (el, ns, name) => [...el.getElementsByTagNameNS(ns, name)];
  const first = (el, ns, name) => el.getElementsByTagNameNS(ns, name)[0] || null;
  const mk = (doc, ns, q) => doc.createElementNS(ns, q);

  // zipB64: one-slide pptx; targets: [{id,name,w,h}] (w/h in EMU); mutate(el, ctx) -> true if changed
  async function editSlideXml(zipB64, targets, mutate) {
    const zip = await JSZip.loadAsync(zipB64, { base64: true });
    const path = Object.keys(zip.files).find(f => /^ppt\/slides\/slide\d+\.xml$/.test(f));
    const doc = new DOMParser().parseFromString(await zip.file(path).async("string"), "application/xml");
    const all = [...doc.getElementsByTagNameNS(P, "sp"), ...doc.getElementsByTagNameNS(P, "pic")];
    let changed = 0;
    for (const t of targets) {
      const digits = (String(t.id || "").match(/\d+/) || [""])[0];
      const el = all.find(e => { const c = first(e, P, "cNvPr"); return c && c.getAttribute("id") === digits; })
        || all.find(e => { const c = first(e, P, "cNvPr"); return c && c.getAttribute("name") === t.name; });
      if (el && await mutate(el, { doc, zip, path, frame: t })) changed++;
    }
    if (!changed) return null;
    zip.file(path, new XMLSerializer().serializeToString(doc));
    return await zip.generateAsync({ type: "base64" });
  }

  // ----- lists -----
  function setBullets(kind) {
    return el => {
      const body = first(el, P, "txBody"); if (!body) return false;
      kids(body, A, "p").forEach(p => {
        const doc = p.ownerDocument;
        let pPr = first(p, A, "pPr");
        if (!pPr) { pPr = mk(doc, A, "a:pPr"); p.insertBefore(pPr, p.firstChild); }
        ["buNone", "buChar", "buAutoNum", "buFont", "buBlip", "buSzPct", "buSzPts", "buClr"].forEach(n => kids(pPr, A, n).filter(x => x.parentNode === pPr).forEach(x => pPr.removeChild(x)));
        if (!pPr.getAttribute("marL")) { pPr.setAttribute("marL", "285750"); pPr.setAttribute("indent", "-285750"); }
        const before = [...pPr.children].find(c => ["tabLst", "defRPr", "extLst"].includes(c.localName)) || null;
        const add = e => pPr.insertBefore(e, before);
        if (kind === "ticks") {
          const clr = mk(doc, A, "a:buClr"); const s = mk(doc, A, "a:srgbClr"); s.setAttribute("val", "B5EB03"); clr.appendChild(s); add(clr);
          const f = mk(doc, A, "a:buFont"); f.setAttribute("typeface", "Wingdings"); f.setAttribute("pitchFamily", "2"); f.setAttribute("charset", "2"); add(f);
          const c = mk(doc, A, "a:buChar"); c.setAttribute("char", "ü"); add(c);
        } else {
          const n = mk(doc, A, "a:buAutoNum"); n.setAttribute("type", "arabicPeriod"); add(n);
        }
      });
      return true;
    };
  }

  // ----- pictures -----
  async function naturalSize(el, ctx) {
    const blip = first(el, A, "blip"); if (!blip) return null;
    const rid = blip.getAttributeNS(R, "embed");
    const relsPath = ctx.path.replace("slides/", "slides/_rels/") + ".rels";
    const rels = new DOMParser().parseFromString(await ctx.zip.file(relsPath).async("string"), "application/xml");
    const rel = [...rels.documentElement.children].find(r => r.getAttribute("Id") === rid); if (!rel) return null;
    const media = "ppt/" + rel.getAttribute("Target").replace(/^\.\.\//, "");
    const f = ctx.zip.file(media); if (!f) return null;
    const bmp = await createImageBitmap(new Blob([await f.async("uint8array")]));
    return { w: bmp.width, h: bmp.height };
  }
  function frameOf(el, ctx) {
    const sp = first(el, P, "spPr"); const xfrm = sp && first(sp, A, "xfrm");
    const off = xfrm && first(xfrm, A, "off"), ext = xfrm && first(xfrm, A, "ext");
    if (ext) return { sp, xfrm, off, ext, w: +ext.getAttribute("cx"), h: +ext.getAttribute("cy") };
    return { sp, w: ctx.frame.w, h: ctx.frame.h };
  }
  function coverCrop(el, nat, w, h) {
    const bf = first(el, P, "blipFill"); const blip = first(bf, A, "blip");
    kids(bf, A, "srcRect").forEach(x => x.parentNode.removeChild(x));
    const src = mk(el.ownerDocument, A, "a:srcRect");
    const Ar = w / h, B = nat.w / nat.h;
    if (B > Ar) { const c = Math.round((1 - Ar / B) / 2 * 100000); src.setAttribute("l", c); src.setAttribute("r", c); }
    else if (B < Ar) { const c = Math.round((1 - B / Ar) / 2 * 100000); src.setAttribute("t", c); src.setAttribute("b", c); }
    blip.parentNode.insertBefore(src, blip.nextSibling);
    kids(bf, A, "tile").forEach(x => x.parentNode.removeChild(x));
    if (!first(bf, A, "stretch")) { const st = mk(el.ownerDocument, A, "a:stretch"); st.appendChild(mk(el.ownerDocument, A, "a:fillRect")); bf.appendChild(st); }
  }
  async function fillFrame(el, ctx) {
    if (el.localName !== "pic") return false;
    const nat = await naturalSize(el, ctx); if (!nat) return false;
    const f = frameOf(el, ctx); coverCrop(el, nat, f.w, f.h); return true;
  }
  async function cropCircle(el, ctx) {
    if (el.localName !== "pic") return false;
    const nat = await naturalSize(el, ctx); if (!nat) return false;
    const doc = el.ownerDocument; let f = frameOf(el, ctx);
    if (!f.ext) { // placeholder picture without its own position: give it one from the frame PowerPoint reported
      const x = mk(doc, A, "a:xfrm"), o = mk(doc, A, "a:off"), e = mk(doc, A, "a:ext");
      o.setAttribute("x", Math.round(ctx.frame.x)); o.setAttribute("y", Math.round(ctx.frame.y));
      e.setAttribute("cx", Math.round(ctx.frame.w)); e.setAttribute("cy", Math.round(ctx.frame.h));
      x.appendChild(o); x.appendChild(e); f.sp.insertBefore(x, f.sp.firstChild); f = frameOf(el, ctx);
    }
    const s = Math.min(f.w, f.h);
    f.off.setAttribute("x", Math.round(+f.off.getAttribute("x") + (f.w - s) / 2));
    f.off.setAttribute("y", Math.round(+f.off.getAttribute("y") + (f.h - s) / 2));
    f.ext.setAttribute("cx", s); f.ext.setAttribute("cy", s);
    kids(f.sp, A, "prstGeom").concat(kids(f.sp, A, "custGeom")).forEach(g => g.parentNode.removeChild(g));
    const g = mk(doc, A, "a:prstGeom"); g.setAttribute("prst", "ellipse"); g.appendChild(mk(doc, A, "a:avLst"));
    f.sp.insertBefore(g, f.xfrm.nextSibling);
    coverCrop(el, nat, s, s);
    return true;
  }

  window.continuEdits = { editSlideXml, ticks: setBullets("ticks"), numbers: setBullets("numbers"), fillFrame, cropCircle };
})();
