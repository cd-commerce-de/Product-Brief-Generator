/* xlsx-images.js
 * Extracts embedded images from a PD sheet (.xlsx). SheetJS's public API
 * (XLSX.read) does not expose embedded drawings/pictures - an .xlsx is a
 * zip container, so this reads it as a raw zip (via JSZip) and walks the
 * OOXML relationship chain by hand:
 *   xl/workbook.xml            -> sheet name <-> rId
 *   xl/_rels/workbook.xml.rels -> rId -> worksheets/sheetN.xml
 *   xl/worksheets/_rels/sheetN.xml.rels -> drawing relationship -> drawingN.xml
 *   xl/drawings/drawingN.xml   -> anchor cell (row/col) + blip r:embed id
 *   xl/drawings/_rels/drawingN.xml.rels -> embed id -> xl/media/imageN.ext
 *
 * PD sheet photos are often several MB each at full camera resolution - far
 * too large to embed directly into a .docx or keep in localStorage - so
 * every image is downscaled via <canvas> before being returned.
 */
(function (global) {
  const MAX_DIM = 700; // longest edge, px
  const JPEG_QUALITY = 0.72;
  const MAX_HERO_IMAGES = 4;
  const MAX_FEATURE_IMAGES = 10;

  function parseXml(text) {
    return new DOMParser().parseFromString(text, "application/xml");
  }

  function relsPathFor(partPath) {
    const idx = partPath.lastIndexOf("/");
    return partPath.slice(0, idx + 1) + "_rels/" + partPath.slice(idx + 1) + ".rels";
  }

  async function resizeImageBlob(blob) {
    const bitmap = await createImageBitmap(blob);
    let { width, height } = bitmap;
    const scale = Math.min(1, MAX_DIM / Math.max(width, height));
    width = Math.max(1, Math.round(width * scale));
    height = Math.max(1, Math.round(height * scale));
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d");
    ctx.drawImage(bitmap, 0, 0, width, height);
    if (bitmap.close) bitmap.close();
    const dataUrl = canvas.toDataURL("image/jpeg", JPEG_QUALITY);
    return { dataUrl, width, height };
  }

  // Returns { hero: [{dataUrl,width,height}], feature: [...] }
  async function extractImages(arrayBuffer) {
    const hero = [];
    const feature = [];

    let zip;
    try {
      zip = await JSZip.loadAsync(arrayBuffer);
    } catch (e) {
      return { hero, feature }; // not a valid zip / no embedded content - fine
    }

    const workbookFile = zip.file("xl/workbook.xml");
    const wbRelsFile = zip.file("xl/_rels/workbook.xml.rels");
    if (!workbookFile || !wbRelsFile) return { hero, feature };

    const wbDoc = parseXml(await workbookFile.async("text"));
    const sheets = Array.from(wbDoc.getElementsByTagName("sheet")).map((el) => ({
      name: el.getAttribute("name") || "",
      rId: el.getAttribute("r:id"),
    }));

    const wbRelsDoc = parseXml(await wbRelsFile.async("text"));
    const relMap = {};
    Array.from(wbRelsDoc.getElementsByTagName("Relationship")).forEach((el) => {
      relMap[el.getAttribute("Id")] = el.getAttribute("Target");
    });

    for (const sheet of sheets) {
      const isOverview = /overview/i.test(sheet.name);
      const isSpecSheet = /product spec/i.test(sheet.name);
      if (!isOverview && !isSpecSheet) continue; // only care about these two

      const target = relMap[sheet.rId];
      if (!target) continue;
      const sheetPath = "xl/" + target.replace(/^\.\.\//, "");
      const sheetRelsFile = zip.file(relsPathFor(sheetPath));
      if (!sheetRelsFile) continue;

      const sheetRelsDoc = parseXml(await sheetRelsFile.async("text"));
      let drawingTarget = null;
      Array.from(sheetRelsDoc.getElementsByTagName("Relationship")).forEach((el) => {
        if ((el.getAttribute("Type") || "").endsWith("/drawing")) drawingTarget = el.getAttribute("Target");
      });
      if (!drawingTarget) continue;

      const drawingPath = "xl/drawings/" + drawingTarget.split("/").pop();
      const drawingFile = zip.file(drawingPath);
      if (!drawingFile) continue;
      const drawingDoc = parseXml(await drawingFile.async("text"));

      const drawingRelsFile = zip.file(relsPathFor(drawingPath));
      const drawingRelMap = {};
      if (drawingRelsFile) {
        const drawingRelsDoc = parseXml(await drawingRelsFile.async("text"));
        Array.from(drawingRelsDoc.getElementsByTagName("Relationship")).forEach((el) => {
          drawingRelMap[el.getAttribute("Id")] = el.getAttribute("Target");
        });
      }

      const anchors = Array.from(drawingDoc.getElementsByTagNameNS("*", "twoCellAnchor"))
        .concat(Array.from(drawingDoc.getElementsByTagNameNS("*", "oneCellAnchor")));

      for (const anchor of anchors) {
        const bucket = isOverview ? hero : feature;
        const cap = isOverview ? MAX_HERO_IMAGES : MAX_FEATURE_IMAGES;
        if (bucket.length >= cap) continue;

        const blipEls = anchor.getElementsByTagNameNS("*", "blip");
        if (!blipEls.length) continue;
        const embedId = blipEls[0].getAttributeNS(
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed"
        );
        const mediaTarget = drawingRelMap[embedId];
        if (!mediaTarget) continue;
        const mediaFile = zip.file("xl/media/" + mediaTarget.split("/").pop());
        if (!mediaFile) continue;

        try {
          const blob = await mediaFile.async("blob");
          const resized = await resizeImageBlob(blob);
          bucket.push(resized);
        } catch (e) {
          // Unsupported/corrupt image - skip it rather than fail the whole import
        }
      }
    }

    return { hero, feature };
  }

  global.PBImages = { extractImages };
})(window);
