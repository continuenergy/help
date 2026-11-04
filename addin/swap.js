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
    const digits = String(match.id || "").replace(/\D+/g, "");
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
