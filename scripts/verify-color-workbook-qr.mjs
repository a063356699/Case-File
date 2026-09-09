import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { PNG } from "pngjs";
import jsQR from "../app/jsQR.cjs";
import { colorWorkbookQrPayload, createColorWorkbookQr, placeColorWorkbookQr, pngDataUrlBytes } from "../app/colorWorkbookQr.ts";

const propertyNo = process.argv[2] || "EG0507820";
const outputPath = path.resolve(process.argv[3] || "outputs/qr-verification/彩色表_QR驗證_EG0507820.xlsx");
const payload = colorWorkbookQrPayload(propertyNo);
const dataUrl = await createColorWorkbookQr(payload);
const template = await fs.readFile(new URL("../app/assets/彩色表範本.xlsm", import.meta.url));
const zip = await JSZip.loadAsync(template);
await placeColorWorkbookQr(zip, pngDataUrlBytes(dataUrl), new DOMParser(), new XMLSerializer());
zip.remove("xl/vbaProject.bin");
zip.remove("xl/calcChain.xml");
const relsPath = "xl/_rels/workbook.xml.rels";
const relsXml = await zip.file(relsPath)?.async("string");
if (relsXml) {
  const document = new DOMParser().parseFromString(relsXml, "application/xml");
  for (const node of Array.from(document.getElementsByTagName("Relationship"))) {
    if (/(?:vbaProject|calcChain)/i.test(node.getAttribute("Type") || "") || /(?:vbaProject\.bin|calcChain\.xml)/i.test(node.getAttribute("Target") || "")) node.parentNode?.removeChild(node);
  }
  zip.file(relsPath, new XMLSerializer().serializeToString(document));
}
const typesPath = "[Content_Types].xml";
const typesXml = await zip.file(typesPath)?.async("string");
if (typesXml) {
  const document = new DOMParser().parseFromString(typesXml, "application/xml");
  for (const node of Array.from(document.getElementsByTagName("Override"))) {
    if (/(?:vbaProject\.bin|calcChain\.xml)/i.test(node.getAttribute("PartName") || "")) node.parentNode?.removeChild(node);
    else if ((node.getAttribute("PartName") || "") === "/xl/workbook.xml") node.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
  }
  zip.file(typesPath, new XMLSerializer().serializeToString(document));
}
await fs.mkdir(path.dirname(outputPath), { recursive: true });
await fs.writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer" }));

const saved = await JSZip.loadAsync(await fs.readFile(outputPath));
const qr = await saved.file("xl/media/color-sheet-case-qr.png")?.async("nodebuffer");
if (!qr) throw new Error("Saved workbook does not contain QR image");
await fs.writeFile(path.join(path.dirname(outputPath), "embedded-qr.png"), qr);
const decodedPng = PNG.sync.read(qr);
const decoded = jsQR(decodedPng.data, decodedPng.width, decodedPng.height, { inversionAttempts: "attemptBoth" })?.data || "";
if (decoded !== payload) throw new Error(`Decoded QR mismatch: expected ${payload}, got ${decoded || "NO_DECODE"}`);
for (const drawingNumber of [2, 4]) {
  const drawing = await saved.file(`xl/drawings/drawing${drawingNumber}.xml`)?.async("string") || "";
  const relations = await saved.file(`xl/drawings/_rels/drawing${drawingNumber}.xml.rels`)?.async("string") || "";
  if (!drawing.includes("rIdColorSheetCaseQr") || !relations.includes("../media/color-sheet-case-qr.png")) throw new Error(`Drawing ${drawingNumber} does not reference QR image`);
}
for (const invalid of ["eg0507820", "EG050782", "EG05078200", "EG05078A0", "0495445"]) {
  let rejected = false;
  try { colorWorkbookQrPayload(invalid); } catch { rejected = true; }
  if (!rejected) throw new Error(`Invalid contract number was accepted: ${invalid}`);
}
console.log(JSON.stringify({ outputPath, qrPath: path.join(path.dirname(outputPath), "embedded-qr.png"), payload, decoded }, null, 2));
