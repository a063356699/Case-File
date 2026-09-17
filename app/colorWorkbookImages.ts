import JSZip from "jszip";

type WorkbookRecord = { caseName?: unknown; address?: unknown };

const compact = (value: unknown) => String(value || "").normalize("NFKC").toLowerCase().replace(/[\s　_－—–\-｜|·・,，、。.．\/\\()（）\[\]【】]/g, "");
const imageFile = (file: File) => /\.(?:jpe?g|png)$/i.test(file.name);
const relativeName = (file: File) => String((file as File & { webkitRelativePath?: string }).webkitRelativePath || file.name);

const matchesCaseOrAddress = (file: File, record: WorkbookRecord) => {
  const path = compact(relativeName(file));
  const caseName = compact(record.caseName);
  const address = compact(record.address);
  return Boolean((caseName && path.includes(caseName)) || (address && path.includes(address)));
};

const onlyUnique = (files: File[]) => files.length === 1 ? files[0] : undefined;

export const colorWorkbookPhotoMatches = (files: File[], record: WorkbookRecord) =>
  files.filter(file => imageFile(file) && matchesCaseOrAddress(file, record) && /^1\.(?:jpe?g|png)$/i.test(file.name));

export function matchColorWorkbookPhoto(files: File[], record: WorkbookRecord) {
  return onlyUnique(colorWorkbookPhotoMatches(files, record));
}

export const colorWorkbookLayoutMatches = (files: File[], record: WorkbookRecord) =>
  files.filter(file => imageFile(file) && matchesCaseOrAddress(file, record));

export function matchColorWorkbookLayout(files: File[], record: WorkbookRecord) {
  return onlyUnique(colorWorkbookLayoutMatches(files, record));
}

const ensureImageContentType = async (zip: JSZip, extension: string, contentType: string) => {
  const path = "[Content_Types].xml";
  const xml = await zip.file(path)?.async("string");
  if (!xml) return;
  const parser = new DOMParser(), serializer = new XMLSerializer();
  const document = parser.parseFromString(xml, "application/xml");
  const contentTypesNs = "http://schemas.openxmlformats.org/package/2006/content-types";
  const exists = Array.from(document.getElementsByTagNameNS(contentTypesNs, "Default")).some(node => node.getAttribute("Extension")?.toLowerCase() === extension);
  if (!exists) {
    const item = document.createElementNS(contentTypesNs, "Default");
    item.setAttribute("Extension", extension);
    item.setAttribute("ContentType", contentType);
    document.documentElement.appendChild(item);
    zip.file(path, serializer.serializeToString(document));
  }
};

export async function placeColorWorkbookImage(zip: JSZip, file: File, kind: "photo" | "layout", isLand: boolean, parser: DOMParser, serializer: XMLSerializer) {
  const xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
  const a = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const officeRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const packageRel = "http://schemas.openxmlformats.org/package/2006/relationships";
  const target = isLand
    ? { drawingPath: "xl/drawings/drawing4.xml", relationsPath: "xl/drawings/_rels/drawing4.xml.rels", shapeId: kind === "photo" ? "10" : "11" }
    : { drawingPath: "xl/drawings/drawing2.xml", relationsPath: "xl/drawings/_rels/drawing2.xml.rels", shapeId: kind === "photo" ? "35" : "33" };
  const drawingXml = await zip.file(target.drawingPath)?.async("string");
  const relationsXml = await zip.file(target.relationsPath)?.async("string");
  if (!drawingXml || !relationsXml) throw new Error(`彩色表範本缺少${kind === "photo" ? "照片" : "格局圖"}方框。`);

  const extension = /\.png$/i.test(file.name) ? "png" : "jpg";
  const contentType = extension === "png" ? "image/png" : "image/jpeg";
  const mediaPath = `xl/media/color-sheet-${kind}.${extension}`;
  const relationId = `rIdColorSheet${kind === "photo" ? "Photo" : "Layout"}`;
  zip.file(mediaPath, new Uint8Array(await file.arrayBuffer()));
  await ensureImageContentType(zip, extension, contentType);

  const drawingDocument = parser.parseFromString(drawingXml, "application/xml");
  const shape = Array.from(drawingDocument.getElementsByTagNameNS(xdr, "sp")).find(node =>
    Array.from(node.getElementsByTagNameNS(xdr, "cNvPr")).some(item => item.getAttribute("id") === target.shapeId),
  );
  const shapeProperties = shape && Array.from(shape.children).find(node => node.localName === "spPr");
  if (!shape || !shapeProperties) throw new Error(`找不到彩色表右側${kind === "photo" ? "第一個照片" : "第二個格局圖"}框。`);
  Array.from(shapeProperties.children).filter(node => ["solidFill", "gradFill", "pattFill", "noFill", "blipFill"].includes(node.localName)).forEach(node => node.remove());
  const blipFill = drawingDocument.createElementNS(a, "a:blipFill");
  const blip = drawingDocument.createElementNS(a, "a:blip");
  blip.setAttributeNS(officeRel, "r:embed", relationId);
  const stretch = drawingDocument.createElementNS(a, "a:stretch");
  stretch.appendChild(drawingDocument.createElementNS(a, "a:fillRect"));
  blipFill.append(blip, stretch);
  const line = Array.from(shapeProperties.children).find(node => node.localName === "ln");
  shapeProperties.insertBefore(blipFill, line || null);
  zip.file(target.drawingPath, serializer.serializeToString(drawingDocument));

  const relationsDocument = parser.parseFromString(relationsXml, "application/xml");
  Array.from(relationsDocument.getElementsByTagNameNS(packageRel, "Relationship")).filter(node => node.getAttribute("Id") === relationId).forEach(node => node.remove());
  const relationship = relationsDocument.createElementNS(packageRel, "Relationship");
  relationship.setAttribute("Id", relationId);
  relationship.setAttribute("Type", `${officeRel}/image`);
  relationship.setAttribute("Target", `../media/color-sheet-${kind}.${extension}`);
  relationsDocument.documentElement.appendChild(relationship);
  zip.file(target.relationsPath, serializer.serializeToString(relationsDocument));
}
