import type JSZip from "jszip";

type MapRecord = Record<string, unknown>;
type LocatedCase = { latitude: number; longitude: number; matchedAddress: string; score: number };
const normalizeTaiwanText = (value: unknown) => String(value ?? "").trim().replace(/臺/g, "台").replace(/[\s　,，。]/g, "").replace(/之/g, "-");
const addressParts = (value: unknown) => { const normalized = normalizeTaiwanText(value); return { city: normalized.match(/^(台[^縣市]{1,4}[市縣]|[^縣市]{1,4}[縣市])/)?.[1] || "", district: normalized.match(/(?:縣|市)([^區鄉鎮市]{1,5}[區鄉鎮市])/)?.[1] || "", road: normalized.match(/([^縣市區鄉鎮]{1,12}(?:路|街|大道))/)?.[1] || "", number: normalized.match(/(\d+(?:-\d+)?號)/)?.[1] || "", landSection: normalized.match(/([^區鄉鎮市]{1,12}(?:段|小段))/)?.[1] || "", landNumber: normalized.match(/(\d+(?:[-/]\d+)*(?:地號)?)/)?.[1] || "" }; };
const firstLandTarget = (value: unknown) => {
  const normalized = normalizeTaiwanText(value);
  const sectionMatch = normalized.match(/([^區鄉鎮市]{1,12}(?:段|小段))([^段]*?)(\d+(?:-\d+)?)(?:地號|[、/]|$)/);
  return sectionMatch ? { section: sectionMatch[1], number: sectionMatch[3] } : null;
};
const coordinatesFromInput = (value: unknown) => {
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  let decoded = raw;
  try { decoded = decodeURIComponent(raw); } catch {}
  const pairs = [
    ...decoded.matchAll(/@(-?\d{1,2}(?:\.\d+)?),\s*(-?\d{2,3}(?:\.\d+)?)/g),
    ...decoded.matchAll(/(?:^|[?&#=/\s])(-?\d{1,2}(?:\.\d+)?),\s*(-?\d{2,3}(?:\.\d+)?)(?:$|[?&#/\s])/g),
  ];
  for (const pair of pairs) {
    const latitude = Number(pair[1]), longitude = Number(pair[2]);
    if (latitude >= 20 && latitude <= 27 && longitude >= 118 && longitude <= 123) return { latitude, longitude };
  }
  return null;
};

const jsonp = <T,>(url: string, timeoutMs = 15000) => new Promise<T>((resolve, reject) => {
  const callbackName = `__caseFileMap_${Date.now()}_${Math.random().toString(36).slice(2)}`;
  const script = document.createElement("script");
  const finish = (error?: Error, value?: T) => { window.clearTimeout(timer); script.remove(); delete (window as unknown as Record<string, unknown>)[callbackName]; if (error) reject(error); else resolve(value as T); };
  const timer = window.setTimeout(() => finish(new Error("位置查詢逾時，請稍後再試。")), timeoutMs);
  (window as unknown as Record<string, unknown>)[callbackName] = (value: T) => finish(undefined, value);
  script.onerror = () => finish(new Error("位置查詢服務目前無法連線。"));
  script.src = `${url}${url.includes("?") ? "&" : "?"}callback=${encodeURIComponent(callbackName)}`;
  document.head.appendChild(script);
});

export async function locateColorWorkbookCase(record: MapRecord): Promise<LocatedCase> {
  const propertyNo = String(record.propertyNo || "").trim();
  const locationInput = String(record.locationMapInput || "").trim();
  const suppliedCoordinates = coordinatesFromInput(locationInput);
  if (suppliedCoordinates) return { ...suppliedCoordinates, matchedAddress: locationInput, score: 100 };
  if (/^https?:\/\//i.test(locationInput)) throw new Error("位置圖定位短網址無法直接取得座標，請貼上經緯度或完整地址。");
  const isLand = /^(?:LG|LA)/i.test(propertyNo) || /土地|建地|農地|地號/.test(String(record.type || ""));
  const suppliedLandAddress = [record.locationLandCity, record.locationLandDistrict, record.locationLandSection, record.locationLandNumber ? `${record.locationLandNumber}地號` : ""].map(value => String(value || "").trim()).filter(Boolean).join("");
  const rawAddress = locationInput || (isLand && suppliedLandAddress ? suppliedLandAddress : String(record.address || "").trim());
  if (!rawAddress) throw new Error("無法產生位置圖：案件未填完整地址或地號資料。請先核對案件資料，不會猜測位置。");
  const expected = addressParts(rawAddress);
  const landTarget = isLand ? firstLandTarget(rawAddress) : null;
  // 樓層、棟別及括號備註常使地理編碼找不到門牌；依序查完整地址與「縣市＋行政區＋路名＋門牌」。
  // 每個結果仍須重新核對路名與門牌，不能因縮短查詢字串而採用相似地址。
  const baseHouseAddress = rawAddress
    .replace(/[（(][^）)]*(?:未保存|地號|建號|增建)[^）)]*[）)]/g, "")
    .replace(/(?:地下)?\d+樓(?:之\d+)?(?:[、,，及與]\d+樓(?:之\d+)?)?.*$/g, "")
    .trim();
  const structuredHouseAddress = [expected.city, expected.district, expected.road, expected.number].filter(Boolean).join("");
  // 多地段或多筆地號只取第一個完整的「地段＋地號」定位，避免要求所有地號同時命中。
  const structuredLandAddress = landTarget ? [expected.city, expected.district, landTarget.section, `${landTarget.number}地號`].filter(Boolean).join("") : "";
  const querySources = isLand ? [structuredLandAddress, rawAddress] : [rawAddress, baseHouseAddress, structuredHouseAddress];
  const queries = [...new Set(querySources.map(value => value.replace(/(\d+)之(\d+)/g, "$1-$2").replace(/\s+/g, "")).filter(Boolean))];
  type ArcCandidate = { address?: string; score?: number; location?: { x?: number; y?: number }; attributes?: Record<string, unknown> };
  const candidates: Array<LocatedCase & { searchable: string }> = [];
  for (const query of queries) {
    const url = `https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/findAddressCandidates?f=json&maxLocations=8&countryCode=TWN&outFields=*&SingleLine=${encodeURIComponent(query)}`;
    let result: { candidates?: ArcCandidate[] };
    try { result = await jsonp<{ candidates?: ArcCandidate[] }>(url); }
    catch { continue; }
    for (const candidate of result.candidates || []) {
      const attributes = candidate.attributes || {};
      const searchable = normalizeTaiwanText([candidate.address, attributes.LongLabel, attributes.ShortLabel, attributes.Match_addr, attributes.City, attributes.District, attributes.Subregion, attributes.Region, attributes.Address].filter(Boolean).join(" "));
      const located = { latitude: Number(candidate.location?.y), longitude: Number(candidate.location?.x), matchedAddress: String(candidate.address || attributes.LongLabel || ""), score: Number(candidate.score || 0), searchable };
      if (Number.isFinite(located.latitude) && Number.isFinite(located.longitude)) candidates.push(located);
    }
  }
  const exact = candidates.sort((a, b) => b.score - a.score).find(candidate => {
    const matched = candidate.searchable;
    const localityMatches = (!expected.city || matched.includes(expected.city)) && (!expected.district || matched.includes(expected.district));
    const houseMatches = candidate.score >= 90 && localityMatches && !!expected.road && !!expected.number && matched.includes(expected.road) && matched.includes(expected.number);
    const targetSection = landTarget?.section || expected.landSection;
    const targetNumber = landTarget?.number || expected.landNumber.replace(/地號$/, "");
    const landMatches = candidate.score >= 95 && localityMatches && !!targetSection && !!targetNumber && matched.includes(targetSection) && matched.includes(targetNumber);
    return isLand ? landMatches : houseMatches;
  });
  if (!exact) { const reason = isLand ? "土地案件須能逐字核對縣市、行政區、地段／小段與完整地號；目前公開定位結果不足。" : "查詢結果無法同時核對縣市、行政區、路名及門牌。"; throw new Error(`無法產生位置圖：${reason}請補齊或修正資料後再下載，不會採用相似地址。`); }
  return { latitude: exact.latitude, longitude: exact.longitude, matchedAddress: exact.matchedAddress, score: exact.score };
}

const imageFromBytes = async (bytes: ArrayBuffer, contentType: string) => {
  // 先把外部底圖轉成本機 blob URL，再繪製到 canvas。這可避免不同瀏覽器
  // 對跨網站圖片畫布的限制，導致畫面看得到但輸出 PNG 時被擋下。
  const objectUrl = URL.createObjectURL(new Blob([bytes], { type: contentType || "image/png" }));
  try {
    return await new Promise<HTMLImageElement>((resolve, reject) => {
      const image = new Image();
      image.onload = () => resolve(image);
      image.onerror = () => reject(new Error("國土測繪中心道路底圖無法轉成 Excel 圖片。"));
      image.src = objectUrl;
    });
  } finally {
    // 圖片完成解碼後才釋放；已解碼內容仍可安全畫入 canvas。
    window.setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
  }
};

export async function createColorWorkbookMap(record: MapRecord) {
  // Excel 位置圖採較近的街區視野，讓周邊道路、學校與公園名稱清楚可讀，
  // 並保留足夠範圍辨識位置，不放大到只剩單一街廓。
  const located = await locateColorWorkbookCase(record), width = 900, height = 680, lonSpan = 0.0085, latSpan = lonSpan * height / width;
  const bbox = [located.longitude - lonSpan / 2, located.latitude - latSpan / 2, located.longitude + lonSpan / 2, located.latitude + latSpan / 2].join(",");
  const mapUrl = `https://wms.nlsc.gov.tw/wms?SERVICE=WMS&VERSION=1.1.1&REQUEST=GetMap&BBOX=${bbox}&SRS=EPSG:4326&WIDTH=${width}&HEIGHT=${height}&LAYERS=EMAP&STYLES=&FORMAT=image/png&DPI=96&MAP_RESOLUTION=96&FORMAT_OPTIONS=dpi:96&TRANSPARENT=FALSE`;
  const mapResponse = await fetch(mapUrl);
  if (!mapResponse.ok) throw new Error(`國土測繪中心道路底圖載入失敗（${mapResponse.status}）。`);
  const contentType = mapResponse.headers.get("content-type") || "";
  if (!contentType.toLowerCase().includes("image")) throw new Error("國土測繪中心未回傳有效的道路底圖。");
  const mapImage = await imageFromBytes(await mapResponse.arrayBuffer(), contentType), canvas = document.createElement("canvas"); canvas.width = width; canvas.height = height;
  const context = canvas.getContext("2d"); if (!context) throw new Error("無法建立位置圖畫布。"); context.drawImage(mapImage, 0, 0, width, height);
  const x = width / 2, y = height / 2; context.save(); context.shadowColor = "rgba(0,0,0,.35)"; context.shadowBlur = 7; context.beginPath(); context.arc(x, y - 13, 17, 0, Math.PI * 2); context.moveTo(x - 11, y - 1); context.lineTo(x, y + 24); context.lineTo(x + 11, y - 1); context.closePath(); context.fillStyle = "#e7272d"; context.fill(); context.restore(); context.beginPath(); context.arc(x, y - 13, 6, 0, Math.PI * 2); context.fillStyle = "#fff"; context.fill();
  context.fillStyle = "rgba(255,255,255,.9)"; context.fillRect(0, height - 28, width, 28); context.fillStyle = "#26343b"; context.font = '18px "Microsoft JhengHei", sans-serif'; context.textAlign = "right"; context.textBaseline = "middle"; context.fillText("底圖來源：內政部國土測繪中心 臺灣通用電子地圖", width - 10, height - 14);
  const blob = await new Promise<Blob>((resolve, reject) => canvas.toBlob(value => value ? resolve(value) : reject(new Error("位置圖轉檔失敗。")), "image/png")); return { png: new Uint8Array(await blob.arrayBuffer()), located };
}

export async function placeColorWorkbookMap(zip: JSZip, mapPng: Uint8Array, isLand: boolean, parser: DOMParser, serializer: XMLSerializer) {
  const xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", a = "http://schemas.openxmlformats.org/drawingml/2006/main", officeRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships", packageRel = "http://schemas.openxmlformats.org/package/2006/relationships";
  const target = isLand ? { drawingPath: "xl/drawings/drawing4.xml", relationsPath: "xl/drawings/_rels/drawing4.xml.rels", shapeId: "12" } : { drawingPath: "xl/drawings/drawing2.xml", relationsPath: "xl/drawings/_rels/drawing2.xml.rels", shapeId: "22" };
  const drawingXml = await zip.file(target.drawingPath)?.async("string"), relationsXml = await zip.file(target.relationsPath)?.async("string"); if (!drawingXml || !relationsXml) throw new Error("彩色表範本缺少第三個位置圖框。");
  const drawingDocument = parser.parseFromString(drawingXml, "application/xml"); const shape = Array.from(drawingDocument.getElementsByTagNameNS(xdr, "sp")).find(node => Array.from(node.getElementsByTagNameNS(xdr, "cNvPr")).some(item => item.getAttribute("id") === target.shapeId)); const shapeProperties = shape && Array.from(shape.children).find(node => node.localName === "spPr"); if (!shape || !shapeProperties) throw new Error("找不到彩色表右側第三個位置圖框。");
  const relationId = "rIdColorSheetLocationMap"; zip.file("xl/media/color-sheet-location-map.png", mapPng); Array.from(shapeProperties.children).filter(node => ["solidFill", "gradFill", "pattFill", "noFill", "blipFill"].includes(node.localName)).forEach(node => node.remove());
  const blipFill = drawingDocument.createElementNS(a, "a:blipFill"), blip = drawingDocument.createElementNS(a, "a:blip"), stretch = drawingDocument.createElementNS(a, "a:stretch"); blip.setAttributeNS(officeRel, "r:embed", relationId); stretch.appendChild(drawingDocument.createElementNS(a, "a:fillRect")); blipFill.append(blip, stretch); const line = Array.from(shapeProperties.children).find(node => node.localName === "ln"); shapeProperties.insertBefore(blipFill, line || null); zip.file(target.drawingPath, serializer.serializeToString(drawingDocument));
  const relationsDocument = parser.parseFromString(relationsXml, "application/xml"); Array.from(relationsDocument.getElementsByTagNameNS(packageRel, "Relationship")).filter(node => node.getAttribute("Id") === relationId).forEach(node => node.remove()); const relationship = relationsDocument.createElementNS(packageRel, "Relationship"); relationship.setAttribute("Id", relationId); relationship.setAttribute("Type", `${officeRel}/image`); relationship.setAttribute("Target", "../media/color-sheet-location-map.png"); relationsDocument.documentElement.appendChild(relationship); zip.file(target.relationsPath, serializer.serializeToString(relationsDocument));
}
