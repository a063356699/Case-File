import QRCode from "qrcode";
import type JSZip from "jszip";

const COLOR_WORKBOOK_COMPANY_CODE = "TC";
const COLOR_WORKBOOK_STORE_CODE = "PB033";
const PROPERTY_NO_PATTERN = /^[A-Z]{2}\d{7}$/;

export function colorWorkbookQrPayload(propertyNoValue: unknown) {
  const raw = String(propertyNoValue ?? "");
  const propertyNo = raw.trim();
  if (!propertyNo) throw new Error("無法產生 QR Code：契約編號未填寫，請先核對總表原始資料。");
  if (raw !== propertyNo) throw new Error(`無法產生 QR Code：契約編號「${raw}」前後含有空白，請先核對總表原始資料。`);
  if (/[a-z]/.test(propertyNo)) throw new Error(`無法產生 QR Code：契約編號「${propertyNo}」含小寫英文；正確規則為 2 碼大寫英文＋7 碼數字，共 9 碼。請先核對原始資料，不會自動改碼。`);
  if (!PROPERTY_NO_PATTERN.test(propertyNo)) {
    const details = [
      propertyNo.length !== 9 ? `目前共 ${propertyNo.length} 碼，正確應為 9 碼` : "",
      !/^[A-Z]{2}/.test(propertyNo) ? "前 2 碼應為大寫英文" : "",
      !/^.{2}\d{7}$/.test(propertyNo) ? "後 7 碼應為完整數字" : "",
    ].filter(Boolean).join("；");
    throw new Error(`無法產生 QR Code：契約編號「${propertyNo}」格式異常${details ? `（${details}）` : ""}。請先逐字核對總表原始資料，不會自行猜測或補碼。`);
  }
  return `${COLOR_WORKBOOK_COMPANY_CODE},${COLOR_WORKBOOK_STORE_CODE},${propertyNo}`;
}

export async function createColorWorkbookQr(payload: string) {
  return QRCode.toDataURL(payload, {
    type: "image/png",
    errorCorrectionLevel: "M",
    margin: 2,
    width: 320,
    color: { dark: "#000000", light: "#FFFFFF" },
  });
}

export function pngDataUrlBytes(dataUrl: string) {
  const encoded = dataUrl.match(/^data:image\/png;base64,(.+)$/)?.[1];
  if (!encoded) throw new Error("QR Code 圖片格式錯誤，未放入 Excel。");
  const binary = atob(encoded);
  return Uint8Array.from(binary, character => character.charCodeAt(0));
}

export async function placeColorWorkbookQr(
  zip: JSZip,
  qrPng: Uint8Array,
  parser: DOMParser,
  serializer: XMLSerializer,
) {
  const elementChildren = (node: Element) => Array.from(node.childNodes).filter((child): child is Element => child.nodeType === 1) as Element[];
  const drawingNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
  const drawingMainNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const officeRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const packageRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
  const relationId = "rIdColorSheetCaseQr";
  zip.file("xl/media/color-sheet-case-qr.png", qrPng);

  const targets = [
    { drawingPath: "xl/drawings/drawing2.xml", relationsPath: "xl/drawings/_rels/drawing2.xml.rels", shapeId: "12" },
    { drawingPath: "xl/drawings/drawing4.xml", relationsPath: "xl/drawings/_rels/drawing4.xml.rels", shapeId: "19" },
  ];

  for (const target of targets) {
    const drawingXml = await zip.file(target.drawingPath)?.async("string");
    const relationsXml = await zip.file(target.relationsPath)?.async("string");
    if (!drawingXml || !relationsXml) throw new Error("彩色表範本缺少右上角 QR Code 方框。");

    const drawingDocument = parser.parseFromString(drawingXml, "application/xml");
    const shape = Array.from(drawingDocument.getElementsByTagNameNS(drawingNs, "sp")).find(node =>
      Array.from(node.getElementsByTagNameNS(drawingNs, "cNvPr")).some(item => item.getAttribute("id") === target.shapeId),
    );
    const shapeProperties = shape && elementChildren(shape).find(node => node.localName === "spPr");
    if (!shape || !shapeProperties) throw new Error("找不到彩色表右上角原有小方框，QR Code 未放入 Excel。");

    elementChildren(shapeProperties).filter(node => ["solidFill", "gradFill", "pattFill", "noFill", "blipFill"].includes(node.localName)).forEach(node => node.parentNode?.removeChild(node));
    const blipFill = drawingDocument.createElementNS(drawingMainNs, "a:blipFill");
    const blip = drawingDocument.createElementNS(drawingMainNs, "a:blip");
    blip.setAttributeNS(officeRelNs, "r:embed", relationId);
    const stretch = drawingDocument.createElementNS(drawingMainNs, "a:stretch");
    stretch.appendChild(drawingDocument.createElementNS(drawingMainNs, "a:fillRect"));
    blipFill.appendChild(blip);
    blipFill.appendChild(stretch);
    const line = elementChildren(shapeProperties).find(node => node.localName === "ln");
    shapeProperties.insertBefore(blipFill, line || null);
    zip.file(target.drawingPath, serializer.serializeToString(drawingDocument));

    const relationsDocument = parser.parseFromString(relationsXml, "application/xml");
    const root = relationsDocument.documentElement;
    Array.from(root.getElementsByTagNameNS(packageRelNs, "Relationship")).filter(node => node.getAttribute("Id") === relationId).forEach(node => node.parentNode?.removeChild(node));
    const relationship = relationsDocument.createElementNS(packageRelNs, "Relationship");
    relationship.setAttribute("Id", relationId);
    relationship.setAttribute("Type", `${officeRelNs}/image`);
    relationship.setAttribute("Target", "../media/color-sheet-case-qr.png");
    root.appendChild(relationship);
    zip.file(target.relationsPath, serializer.serializeToString(relationsDocument));
  }
}
