"use client";

import { ChangeEvent, Dispatch, Fragment, SetStateAction, useEffect, useMemo, useRef, useState } from "react";
import PptxGenJS from "pptxgenjs";
import JSZip from "jszip";
import jsQR from "./jsQR.cjs";
import colorWorkbookTemplateUrl from "./assets/彩色表範本.xlsm?url";
import { sourceBalconyFixes } from "./source-balcony-fixes";

type RecordItem = Record<string, string> & { id: string; photos: string[]; archived?: string };
type AdvancedFilter = Record<string, string>;
type Person = { id: string; sequence?: string; name: string; nationalId: string; phone?: string; role?: "業務" | "秘書"; status: "在職" | "離職" };
type Settings = { personnel: Person[]; staffName?: string; staffId?: string; supabaseUrl: string; supabaseKey: string; supabaseTable: string; supabaseRecord: string; expiry591?: string; expiry5168?: string; brokerExpiry?: string; bookReviewCurrentDate?: string; bookReviewNextDate?: string };
type CloudSession = { accessToken: string; refreshToken?: string; email?: string };
type IntakeData = { id: string; values: Record<string, string>; propertyKind: "房屋" | "純土地"; createdAt: string; modifiedAt?: string; raw?: string; linkedRecordId?: string; enteredAt?: string; groupViewDate?: string; printedForSalesAt?: string };
type TourItem = { id: string; recordId?: string; sequence: string; temporary?: boolean; data: RecordItem };
type TourHistory = { id: string; date: string; title: string; items: TourItem[]; completedAt: string };

const STORAGE_KEY = "property-desk-v1";
const SETTINGS_KEY = "property-desk-settings-v1";
const INTAKE_KEY = "property-desk-intake-draft-v1";
const PHOTO_INTAKE_CLEANUP_KEY = "property-desk-photo-intake-cleanup-v2";
const TOUR_KEY = "property-desk-tour-plan-v1";
const PPT_WEEK_SELECTIONS_KEY = "property-desk-ppt-week-selections-v1";
const DAILY_HIDDEN_KEY = "property-desk-daily-hidden-v1";
const MISSING_REMINDER_DATE_KEY = "property-desk-missing-reminder-date-v1";
const CLOUD_SESSION_KEY = "property-desk-supabase-session-v1";
const CLOUD_LAST_UPLOAD_KEY = "property-desk-supabase-last-upload-v1";
const CASE_FILE_SUPABASE_URL = "https://oiywtmjbasoonfuxemtr.supabase.co";
const CASE_FILE_SUPABASE_TABLE = "case_file_state";
const newCaseReminderPending = (record: RecordItem) => !["housingListingCompleted", "newBookCompleted", "wangReviewCompleted"].every(key => record[key] === "1");
// This is Supabase's browser-safe publishable key. Access is protected by the
// signed-in user and the table's row-level security policy, never by this key.
const CASE_FILE_SUPABASE_PUBLISHABLE_KEY = "sb_publishable_dhr81918k0zRtar14yepIA_ZWmTKT10";
const newId = () => globalThis.crypto?.randomUUID?.() || `id-${Date.now()}-${Math.random().toString(36).slice(2)}`;
const jsonExportFileName = (date = new Date()) => {
  const rocYear = date.getFullYear() - 1911;
  const period = date.getHours() < 12 ? "早上" : "下午";
  const hour = date.getHours() % 12 || 12;
  return `${rocYear}年${date.getMonth() + 1}月${date.getDate()}日${period}${hour}點${date.getMinutes()}分${String(date.getSeconds()).padStart(2, "0")}秒.json`;
};

const fields = [
  ["propertyNo", "物件編號"], ["type", "種類"], ["contractType", "契種"], ["status", "物件狀態"],
  ["area", "地區"], ["caseName", "案名"], ["caseNameNote", "案名後方備註"], ["showingFollowUpDueDate", "暫停帶看追蹤日期"], ["address", "地址"], ["price", "開價（萬）"], ["reducedPrice", "降價/調價(萬)"],
  ["direction", "朝向"], ["builtYear", "完工年份（民國／西元）"], ["floor", "樓層"], ["layout", "格局"],
  ["indoorPing", "室內坪"], ["buildingPing", "建坪"], ["landPing", "地坪"], ["parking", "車位"],
  ["managementFee", "管理費（／月）"], ["key", "鑰匙"], ["currentState", "現況"], ["road", "臨路（米）"],
  ["frontage", "面寬（米）"], ["depth", "深度（米）"], ["zoning", "使用分區"], ["coverage", "建蔽率（%）"],
  ["far", "容積率（%）"], ["developer", "開發業務"], ["entrustStart", "委託開始"], ["entrustEnd", "委託結束"],
  ["reportDate", "進案報件日期"], ["updateDate", "更新日期"], ["groupViewDate", "團看日期"], ["bookLocationType", "物件本"],
  ["bookLocationDate", "物件本日期"], ["bookLocationNo", "旁放編號"], ["salesBook", "銷售本"], ["salesBookDate", "銷售本日期"], ["notes", "備註欄"], ["photoInfo", "照片"],
  ["areaPaste", "面積資料貼串"], ["completionDate", "建築完成日期"], ["registryBuildingPing", "總建物坪數"], ["registryIndoorPing", "室內坪數"],
  ["mainBuildingPing", "主建物坪數"], ["auxiliaryBuildingPing", "附屬建物坪數"], ["commonAreaPing", "共同使用坪數"], ["buildingOtherPing", "建物其他坪數"], ["basementPing", "地下室坪數"], ["landSharePing", "土地坪數"],
  ["buildingName", "大樓名稱"], ["unitsPerFloor", "每層戶數"], ["elevatorCount", "電梯數"], ["managementMethod", "管理方式"], ["titleFloor", "權狀樓層"], ["currentFloor", "現況樓別"],
  ["parkingOwnership", "車位產權"], ["parkingType", "車位型態"], ["parkingMethod", "停車方式"], ["parkingNo", "車位編號"],
  ["market", "市場"], ["park", "公園"], ["elementarySchool", "國小"], ["juniorHighSchool", "國中"], ["seniorHighSchool", "高中"], ["collegeSchool", "大專"], ["school", "學校彙整"], ["generalLandValueTax", "一般增值稅"], ["selfUseLandValueTax", "自用增值稅"], ["deedTax", "契稅"],
  ["feature1", "特色說明1"], ["feature2", "特色說明2"], ["feature3", "特色說明3"], ["feature4", "特色說明4"],
  ["additionNotes", "增建說明"], ["attentionNotes", "注意事項"], ["middleman", "中人"],
  ["coverBottomPrice", "底價（萬）"], ["coverBottomSource", "底價依據"], ["coverBottomPercent", "底價％數"], ["coverPercentSource", "％數依據"],
  ["coverChangeNo", "進案契變編號"], ["coverChangeDate", "契變日期"], ["coverNoChange", "契變狀態"], ["coverChangePurpose", "用途"], ["coverOriginalType", "原稿／草稿"],
  ["landTitleCount", "土地權狀（張）"], ["buildingTitleCount", "建物權狀（張）"], ["titleUndertaking", "切結"],
  ["zoningDocumentStatus", "使用分區文件"], ["authorizationStatus", "授權書狀態"], ["authorizationCopyType", "授權書本別"],
  ["platform591", "591"], ["platform591Expiry", "591到期日期"], ["yes319", "YES319"], ["houseinfor", "HOUSEINFOR"], ["windowAd", "櫥窗廣告（只上專簽）"],
  ["led", "LED（只上專簽）"], ["homeWeb", "我家網"], ["price5168", "5168實價網"], ["price5168Expiry", "5168到期日期"], ["goldExposure", "黃金曝光"], ["goldExposureExpiry", "黃金曝光到期日期"],
] as const;

const recordEditOrder = [
  "propertyNo", "type", "contractType", "status", "currentState", "key",
  "entrustStart", "entrustEnd", "contractChangeNo", "reportDate", "updateDate", "groupViewDate",
  "area", "caseName", "caseNameNote", "showingFollowUpDueDate", "address",
  "developer", "price", "completionDate", "direction",
  "buildingPing", "indoorPing", "landPing", "road", "frontage", "depth", "zoning", "coverage",
  "titleFloor", "currentFloor", "layout", "managementMethod", "managementFee",
  "photoInfo", "bookLocationType", "bookLocationDate", "salesBook",
  "notes",
  "colorSheetHeader", "areaPaste",
  "landSharePing", "registryBuildingPing", "registryIndoorPing", "buildingOtherPing", "mainBuildingPing",
  "auxiliaryBuildingPing", "commonAreaPing", "basementPing",
  "buildingName", "unitsPerFloor", "elevatorCount", "generalLandValueTax", "selfUseLandValueTax",
  "parkingOwnership", "parkingMethod", "parkingNo",
  "market", "park", "elementarySchool", "juniorHighSchool", "seniorHighSchool", "collegeSchool", "feature1", "feature2", "feature3", "feature4", "attentionNotes",
  "coverHeader", "coverBottomPrice", "coverBottomSource", "coverBottomPercent", "coverPercentSource", "coverChangeNo", "coverChangeDate", "coverNoChange", "coverChangePurpose",
  "landTitleCount", "buildingTitleCount", "titleUndertaking", "zoningDocumentStatus", "authorizationStatus", "authorizationCopyType",
  "websiteHeader",
  "yes319", "houseinfor", "homeWeb", "platform591", "price5168", "windowAd", "led", "goldExposure",
] as const;
const recordEditLabels: Record<string, string> = { price: "開價（萬）", contractChangeNo: "契變編號", direction: "朝向", builtYear: "完工年份（民國／西元）", managementFee: "管理費（/月）", coverage: "建蔽率%/容積率%", parkingMethod: "車位型態", notes: "內部備註欄", areaPaste: "房管面積資料貼串", attentionNotes: "彩色表注意事項", windowAd: "櫥窗（專）", led: "LED（專）", price5168: "5168網（專）", colorSheetHeader: "彩色表補充資料", coverHeader: "新進資料封面", websiteHeader: "網站編號" };
const recordEditClass = (key: string) => `edit-cell edit-${key}`;

const dateKeys = new Set(["entrustStart", "entrustEnd", "reportDate", "updateDate", "groupViewDate", "showingFollowUpDueDate", "bookLocationDate", "salesBookDate", "coverChangeDate", "platform591Expiry", "price5168Expiry", "goldExposureExpiry"]);
const selectOptions: Record<string, string[]> = {
  status: ["委託中", "售出下架", "成交下架", "下架洽開發"],
  type: ["公寓：5樓含以下（無電梯）", "華廈：10樓含以下（有電梯）", "大樓：11樓含以上（有電梯）", "透天", "廠房", "土地"],
  currentState: ["空屋", "自住", "出租中"],
};

const activeColumns = ["displaySequence", "type", "propertyNo", "entrustStart", "entrustEnd", "area", "caseName", "address", "price", "direction", "age", "floor", "layout", "indoorPing", "buildingPing", "landPing", "parking", "managementFee", "key", "currentState", "road", "frontage", "depth", "zoning", "coverageFar", "developer", "notes", "reportDate", "updateDate", "groupViewDate", "bookLocation", "salesBook", "photoInfo", "platform591", "yes319", "houseinfor", "windowAd", "led", "homeWeb", "price5168", "goldExposure"];
const archiveColumns = ["propertyNo", "entrustPeriod", "caseName", "address", "price", "developer", "key", "housingRemoval", "bookLocation", "salesBook", "photoInfo", "platform591", "yes319", "houseinfor", "windowAd", "led", "homeWeb", "price5168", "goldExposure"];
const publicColumns = ["propertyNo", "area", "caseName", "address", "price", "direction", "age", "floor", "layout", "indoorPing", "buildingPing", "landPing", "parking", "managementFee", "key", "currentState", "road", "frontage", "depth", "zoning", "coverageFar", "developer", "notes", "reportDate", "updateDate", "groupViewDate", "photoInfo"];
const dailyColumns = publicColumns.filter(column => !["propertyNo", "photoInfo", "contractType", "status"].includes(column));
const labels: Record<string, string> = Object.fromEntries(fields);
Object.assign(labels, { age: "屋齡", bookLocation: "物件本位置", housingRemoval: "房管下架", entrustPeriod: "委託期間", archived: "封存日期", coverageFar: "建蔽率／容積率" });
labels.displaySequence = "序";

const archiveWebsiteTasks = [
  ["platform591", "591"], ["price5168", "5168"], ["goldExposure", "黃金曝光"], ["windowAd", "櫥窗（專）"],
  ["led", "LED（專）"], ["homeWeb", "我家網"], ["houseinfor", "HOUSE INFOR"], ["yes319", "YES319"],
] as const;
const archiveCleanupTasks = (record: RecordItem) => [
  { key: "housingDownDate", label: "房管下架" },
  { key: "bookDownDate", label: `${record.bookLocationType === "旁5" ? "旁5" : "物件本"}下架` },
  { key: "salesBookDownDate", label: "銷售本下架" },
  ...archiveWebsiteTasks.filter(([field]) => record.bookLocationType !== "旁5" && record[`${field}None`] !== "1" && String(record[field] || "").trim() && record[field] !== "旁5").map(([field, label]) => ({ key: `${field}DownDate`, label: `${label}下架` })),
];
const archiveCleanupTaskLabel = (record: RecordItem, key: string) => archiveCleanupTasks(record).find(task => task.key === key)?.label || "下架動作";
const dealCompletionTasks = [
  ["dealCardsPrintedDate", "列印紅卡、粉卡"],
  ["dealFacebookPostedDate", "臉書PO網"],
  ["dealPerformanceReportedDate", "房管報業績"],
  ["dealSellerPayoutAppliedDate", "申請賣方出款"],
  ["dealPropertyCompletedDate", "成交物件製作"],
] as const;

const blankRecord = (): RecordItem => ({ ...Object.fromEntries(fields.map(([k]) => [k, ""])), id: newId(), photos: [], status: "委託中", bookLocationType: "架上", salesBook: "製作" });
const blankAdvancedFilter = (): AdvancedFilter => ({ ...Object.fromEntries(["propertyNo", "entrustStartFrom", "entrustStartTo", "entrustEndFrom", "entrustEndTo", "caseName", "address", "type", "priceFrom", "priceTo", "direction", "ageFrom", "ageTo", "floor", "rooms", "halls", "baths", "indoorFrom", "indoorTo", "landFrom", "landTo", "buildingFrom", "buildingTo", "parking", "key", "currentState", "roadFrom", "roadTo", "frontageFrom", "frontageTo", "depthFrom", "depthTo", "zoning", "developer", "reportFrom", "reportTo", "updateFrom", "updateTo", "groupFrom", "groupTo", "bookFrom", "bookTo", "salesBook", "salesFrom", "salesTo", "photoInfo", "platform591", "yes319", "houseinfor", "windowAd", "led", "homeWeb", "price5168", "goldExposure", "sourceActive", "sourceArchive"].map(key => [key, ""])), sourceActive: "1" });
const filterNumber = (value: unknown, last = false) => { const values = String(value ?? "").replace(/,/g, "").match(/-?\d+(?:\.\d+)?/g) || []; const text = last ? values.at(-1) : values[0]; return text === undefined ? NaN : Number(text); };
const pingFieldKeys = ["indoorPing", "buildingPing", "landPing", "registryBuildingPing", "registryIndoorPing", "mainBuildingPing", "auxiliaryBuildingPing", "commonAreaPing", "buildingOtherPing", "basementPing", "landSharePing"] as const;
const formatPingValue = (value: unknown) => {
  const text = String(value ?? "").trim();
  if (!text) return "";
  const matched = text.match(/^(-?\d+(?:\.\d+)?)\s*(?:坪)?$/);
  if (!matched) return text;
  const number = Number(matched[1]);
  return Number.isFinite(number) ? String(Math.round((number + Number.EPSILON) * 1000) / 1000) : text;
};
const normalizeRecordPings = (record: RecordItem): RecordItem => ({
  ...record,
  ...Object.fromEntries(pingFieldKeys.map(key => [key, formatPingValue(record[key])])),
});
const filterInRange = (value: unknown, from = "", to = "", last = false) => { const number = filterNumber(value, last); return (!from || (Number.isFinite(number) && number >= Number(from))) && (!to || (Number.isFinite(number) && number <= Number(to))); };
const filterDateInRange = (value: unknown, from = "", to = "") => { const date = normalizeDateInput(String(value || "")); const start = normalizeDateInput(from) || from; const end = normalizeDateInput(to) || to; return (!start || (date && date >= start)) && (!end || (date && date <= end)); };
const today = () => { const value = new Date(); return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}-${String(value.getDate()).padStart(2, "0")}`; };
const BOOK_REVIEW_START = "2026-07-30";
const bookReviewCycleKey = () => { const elapsed = Math.floor((Date.parse(`${today()}T00:00:00`) - Date.parse(`${BOOK_REVIEW_START}T00:00:00`)) / 86400000); return elapsed < 0 ? "" : `global-${Math.floor(elapsed / 60)}`; };
const bookReviewCycleStart = () => { const elapsed = Math.floor((Date.parse(`${today()}T00:00:00`) - Date.parse(`${BOOK_REVIEW_START}T00:00:00`)) / 86400000); if (elapsed < 0) return ""; const date = new Date(`${BOOK_REVIEW_START}T00:00:00`); date.setDate(date.getDate() + Math.floor(elapsed / 60) * 60); return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`; };
const addDaysIso = (iso: string, days: number) => { if (!iso) return ""; const date = new Date(`${iso}T00:00:00`); date.setDate(date.getDate() + days); return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`; };
const sortPeopleBySequence = (people: Person[]) => people.slice().sort((a, b) => Number(a.sequence || 9999) - Number(b.sequence || 9999) || String(a.sequence || "").localeCompare(String(b.sequence || ""), "zh-TW", { numeric: true }) || a.name.localeCompare(b.name, "zh-TW"));
const displayPhone = (value = "") => { const digits = value.replace(/\D/g, ""); return digits.length === 10 ? `${digits.slice(0, 4)}-${digits.slice(4, 7)}-${digits.slice(7)}` : value; };
const parseAreaPaste = (raw: string, record: RecordItem): RecordItem => {
  const text = String(raw || "").replace(/\r/g, "");
  const numberAfter = (label: string) => text.match(new RegExp(`${label}\\s*([\\d,.]+)\\s*坪`))?.[1]?.replace(/,/g, "") || "";
  const nextLine = (label: string) => text.match(new RegExp(`${label}\\s*\\n\\s*([^\\n]+)`))?.[1]?.trim() || "";
  const main = numberAfter("主建物"), auxiliary = numberAfter("附屬建物"), common = numberAfter("共同使用"), basement = numberAfter("地下室");
  const completion = text.match(/建築完成日\s*(?:民國\s*)?(\d{2,3})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/) || [];
  const completionDate = completion[1] ? `${completion[1].padStart(3, "0")}.${completion[2].padStart(2, "0")}.${completion[3].padStart(2, "0")}` : record.completionDate || "";
  const landShare = numberAfter("土地持分面積");
  const indoor = [main, auxiliary].reduce((sum, value) => sum + Number(value || 0), 0);
  return {
    ...record, areaPaste: raw,
    registryBuildingPing: numberAfter("建物總面積") || record.registryBuildingPing || "",
    mainBuildingPing: main || record.mainBuildingPing || "", auxiliaryBuildingPing: auxiliary || record.auxiliaryBuildingPing || "",
    commonAreaPing: common || record.commonAreaPing || "", basementPing: basement || record.basementPing || "",
    buildingOtherPing: basement || record.buildingOtherPing || "", registryIndoorPing: indoor ? String(Math.round(indoor * 1000) / 1000) : record.registryIndoorPing || "",
    completionDate, builtYear: completion[1] ? String(Number(completion[1])) : record.builtYear || "",
    landSharePing: landShare || record.landSharePing || "", landPing: landShare || record.landPing || ""
  };
};
const suppliedPersonnel: Array<Pick<Person, "sequence" | "name" | "nationalId" | "phone">> = [];
const contactDirectoryOrder = [
  "王啟山", "王若芸", "王俞云", "王妤宸", "林玉環", "林顯昌", "林姿岑", "林俊嘉", "陳帝元", "陳珮菁", "陳信良",
  "郭建佑", "謝馨儀", "蔡宇育", "田庭宇", "張小曼", "李享嶧", "楊巧甄", "阮氏金水", "柯育婷", "余沛臻", "葉翊緁",
  "宋喜輝", "吳佩玲", "黃文成", "買淑玲", "劉勝仁", "李麗卉",
];
const contactDirectorySequence = new Map(contactDirectoryOrder.map((name, index) => [name, String(index + 1)]));
const contactDirectoryPhoneOverrides: Record<string, string> = { 郭建佑: "0938-839-308" };
const mergeSuppliedPersonnel = (people: Person[]) => {
  const remaining = people.slice();
  const supplied = suppliedPersonnel.map(entry => {
    const index = remaining.findIndex(person => person.nationalId.toUpperCase() === entry.nationalId || person.name.trim() === entry.name);
    const existing = index >= 0 ? remaining.splice(index, 1)[0] : undefined;
    return { ...(existing || {}), id: existing?.id || `staff-${entry.nationalId}`, ...entry, role: ["李麗卉", "施紹薇"].includes(entry.name) ? "秘書" as const : existing?.role || "業務" as const, status: "在職" as const };
  });
  return [...supplied, ...remaining].map(person => ({
    ...person,
    sequence: contactDirectorySequence.get(person.name.trim()) || person.sequence,
    phone: contactDirectoryPhoneOverrides[person.name.trim()] || person.phone,
  }));
};
const recordUpdateHistory = (record: RecordItem): Record<string, string[]> => { try { return JSON.parse(record._updateHistory || "{}"); } catch { return {}; } };
const editedFieldKeys = (record: RecordItem): string[] => { try { const value = JSON.parse(record._editedFields || "[]"); return Array.isArray(value) ? value.map(String) : []; } catch { return []; } };
const trackedValue = (value: unknown) => {
  const text = String(value ?? "").trim().replace(/\s+/g, "");
  return /^(?:nan|null|undefined)$/i.test(text) ? "" : text;
};
const websiteTrackingKeys = new Set(["platform591", "platform591Expiry", "platform591None", "price5168", "price5168Expiry", "price5168None", "goldExposure", "goldExposureExpiry", "goldExposureNone", "yes319", "yes319None", "houseinfor", "houseinforNone", "homeWeb", "homeWebNone", "windowAd", "windowAdNone", "led", "ledNone"]);
const dailyActivityUpdateKeys = new Set([
  "area", "caseName", "caseNameNote", "address", "price", "reducedPrice", "direction", "completionDate", "builtYear",
  "floor", "titleFloor", "currentFloor", "layout", "indoorPing", "registryIndoorPing", "buildingPing", "registryBuildingPing",
  "landPing", "landSharePing", "parking", "parkingOwnership", "parkingType", "parkingMethod", "parkingNo", "managementFee", "key",
  "currentState", "road", "frontage", "depth", "zoning", "coverage", "far", "developer", "notes"
]);
const correctedWebsiteOnlyUpdate = (record: RecordItem, date: string) => date === "2026-08-07" && [
  "東區德昌路優質建地", "興達港全新大地坪大面寬電梯雙車豪墅", "高雄九如商圈高投報整新四套房1樓公寓"
].some(name => String(record.caseName || "").includes(name)) || date === "2026-08-07" && ["自由路27", "大智路53", "大昌一路43"].some(address => String(record.address || "").includes(address));
const dailyUpdateFields = (record: RecordItem, date: string) => {
  if (correctedWebsiteOnlyUpdate(record, date)) return [];
  const restoredToday = String(record._restoredAt || "").slice(0, 10) === date && String(record.caseNameNote || "").includes(`${shortRocMonthDay(date)}重新上架`);
  const restoredAt = Date.parse(String(record.caseNameNoteModifiedAt || ""));
  const laterEdited = Number.isFinite(restoredAt) && Date.parse(String(record.lastModifiedAt || "")) > restoredAt;
  // 舊的重新上架資料可能仍保有舊欄位歷程；畫面不採用，但絕不改寫資料本身。
  if (restoredToday && !laterEdited) return [];
  const fields = (recordUpdateHistory(record)[date] || []).filter(key => dailyActivityUpdateKeys.has(key)).filter(key => {
    // 舊資料的土地完工年／無車位組合由系統轉換產生，並不是這次人工修改。
    if (key === "builtYear" && /^(土地|建地|nan)?$/i.test(trackedValue(record.completionDate || record.builtYear))) return false;
    if (key === "parking" && trackedValue(record.parking) === trackedValue(record.parkingOwnership) && !trackedValue(record.parkingType) && !trackedValue(record.parkingMethod) && !trackedValue(record.parkingNo)) return false;
    return true;
  });
  const reducedPriceEditedToday = String(record.reducedPriceModifiedAt || "").slice(0, 10) === date;
  const originalPriceEditedToday = String(record.priceModifiedAt || "").slice(0, 10) === date;
  const correctedFields = reducedPriceEditedToday && !originalPriceEditedToday && fields.includes("price") && !fields.includes("reducedPrice") ? [...fields.filter(key => key !== "price"), "reducedPrice"] : fields;
  const displayColumn: Record<string, string> = { caseNameNote: "caseName", completionDate: "age", builtYear: "age", titleFloor: "floor", currentFloor: "floor", registryIndoorPing: "indoorPing", registryBuildingPing: "buildingPing", landSharePing: "landPing", parkingOwnership: "parking", parkingType: "parking", parkingMethod: "parking", parkingNo: "parking", coverage: "coverageFar", far: "coverageFar" };
  return [...new Set(correctedFields.map(key => displayColumn[key] || key))];
};
const dailyChangedLabels = (keys: string[]) => [...new Set(keys.map(key => ({
  area: "地區", caseName: "案名", caseNameNote: "案名", address: "地址", price: "總價", reducedPrice: "總價", direction: "朝向",
  completionDate: "屋齡", builtYear: "屋齡", floor: "樓層", titleFloor: "樓層", currentFloor: "樓層", layout: "格局",
  indoorPing: "室內坪", registryIndoorPing: "室內坪", buildingPing: "建坪", registryBuildingPing: "建坪", landPing: "地坪", landSharePing: "地坪",
  parking: "車位", parkingOwnership: "車位", parkingType: "車位", parkingMethod: "車位", parkingNo: "車位", managementFee: "管理費", key: "鑰匙",
  currentState: "現況", road: "臨路", frontage: "面寬", depth: "深度", zoning: "使用分區", coverage: "建蔽率", far: "容積率", developer: "開發業務", notes: "備註欄"
} as Record<string, string>)[key] || labels[key] || key))];
const displayModifiedAt = (value = "") => { if (!value) return ""; const date = new Date(value); if (Number.isNaN(date.getTime())) return value; return `${date.getFullYear() - 1911}/${date.getMonth() + 1}/${date.getDate()} ${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}`; };
const displayHomeModifiedAt = (value = "") => { if (!value) return ""; const date = new Date(value); if (Number.isNaN(date.getTime())) return value; const weekdays = ["日", "一", "二", "三", "四", "五", "六"]; return `${date.getFullYear() - 1911}年${date.getMonth() + 1}月${date.getDate()}日 星期${weekdays[date.getDay()]} ${String(date.getHours()).padStart(2, "0")}點${String(date.getMinutes()).padStart(2, "0")}分${String(date.getSeconds()).padStart(2, "0")}秒`; };
const withTrackedUpdate = (previous: RecordItem, next: RecordItem, date = today(), onlyFields?: string[]) => {
  const ignored = new Set(["id", "updateDate", "lastModifiedAt", "priceModifiedAt", "reducedPriceModifiedAt", "groupViewDate", "_updateHistory", "_dailyAnnotation", "_dailyHighlight"]);
  const candidates = onlyFields?.length ? onlyFields : Object.keys(next);
  const changed = candidates.filter(key => !ignored.has(key) && !websiteTrackingKeys.has(key) && trackedValue(previous[key]) !== trackedValue(next[key]));
  if (!changed.length) return next;
  const history = recordUpdateHistory(previous);
  // 每日物件動態只顯示本次儲存實際修改的欄位，不累加同一天較早的修改。
  history[date] = [...new Set(changed)];
  return { ...next, updateDate: date, lastModifiedAt: new Date().toISOString(), _updateHistory: JSON.stringify(history) };
};
const isoLocalDate = (date: Date) => `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
const pptWeekOf = (anchor = today()) => { const parsed = /^\d{4}-\d{2}-\d{2}$/.test(anchor) ? new Date(`${anchor}T00:00:00`) : new Date(); parsed.setHours(0, 0, 0, 0); const start = new Date(parsed); start.setDate(parsed.getDate() - ((parsed.getDay() + 1) % 7)); const end = new Date(start); end.setDate(start.getDate() + 6); const meeting = new Date(end); meeting.setDate(end.getDate() + 3); return { start: isoLocalDate(start), end: isoLocalDate(end), meeting: isoLocalDate(meeting) }; };
const currentPptWeek = () => pptWeekOf(today());
const displayPptWeekDate = (iso = "") => { const date = new Date(`${iso}T00:00:00`); if (!iso || Number.isNaN(date.getTime())) return "—"; return `${date.getFullYear() - 1911}/${date.getMonth() + 1}/${date.getDate()}(${["日", "一", "二", "三", "四", "五", "六"][date.getDay()]})`; };
const displayPptWeekLabel = (week: ReturnType<typeof pptWeekOf>) => `${displayPptWeekDate(week.start)} ~ ${displayPptWeekDate(week.end)}於 ${displayPptWeekDate(week.meeting)} 報告`;
const pptMeetingDateLabel = (week = currentPptWeek()) => { const [year, month, day] = week.meeting.split("-").map(Number); return `${year - 1911}年${month}月${day}日`; };
const isPptWeek = (value = "", week = currentPptWeek()) => {
  const date = normalizeDateInput(String(value || ""));
  const start = normalizeDateInput(String(week.start || ""));
  const end = normalizeDateInput(String(week.end || ""));
  return !!date && !!start && !!end && date >= start && date <= end;
};
const pptStoredList = (value = "") => { try { const parsed = JSON.parse(value || "[]"); return Array.isArray(parsed) ? parsed : []; } catch { return []; } };
const pptStoredChoices = (value = "") => { try { const parsed = JSON.parse(value || "{}"); return parsed && typeof parsed === "object" ? parsed : {}; } catch { return {}; } };
const nextPptWeekStart = (start: string) => addDaysIso(start, 7);
const belongsToPptWeek = (record: RecordItem, week: ReturnType<typeof pptWeekOf>) => isPptWeek(record.reportDate, week) || pptStoredList(record._pptExtraWeeks).includes(week.start);
const excludedFromPptWeek = (record: RecordItem, weekStart: string) => pptStoredList(record._pptExcludedWeeks).includes(weekStart);
const isCurrentPptWeek = (value = "") => isPptWeek(value, currentPptWeek());
const pptIncluded = (record: RecordItem) => record.pptSelected === "1" || (record.pptSelected !== "0" && isCurrentPptWeek(record.reportDate));
const pptCategory = (record: RecordItem) => record.pptCategory || (isCurrentPptWeek(record.reportDate) ? "本週進案" : "臨時新增");
const validDate = (v = "") => !v || /^\d{4}-\d{2}-\d{2}$/.test(v) && !Number.isNaN(Date.parse(v));
const displayRocDate = (v = "") => {
  const text = String(v || "").trim();
  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return `${Number(iso[1]) - 1911}/${Number(iso[2])}/${Number(iso[3])}`;
  const roc = text.match(/^(\d{2,3})[.\/-](\d{1,2})[.\/-](\d{1,2})$/);
  return roc ? `${Number(roc[1])}/${Number(roc[2])}/${Number(roc[3])}` : text;
};
const displayRocDotDate = (v = "") => {
  const text = String(v || "").trim();
  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) return `${String(Number(iso[1]) - 1911).padStart(3, "0")}.${String(Number(iso[2])).padStart(2, "0")}.${String(Number(iso[3])).padStart(2, "0")}`;
  const roc = text.match(/^(\d{2,3})[.\/-](\d{1,2})[.\/-](\d{1,2})$/);
  return roc ? `${roc[1].padStart(3, "0")}.${roc[2].padStart(2, "0")}.${roc[3].padStart(2, "0")}` : text;
};
const shortRocMonthDay = (v = "") => {
  const parts = displayRocDate(v).split("/");
  return parts.length === 3 ? `${parts[1]}/${parts[2]}` : displayRocDate(v);
};
const compactRestoredNote = (value = "") => String(value || "").replace(/(\d{1,3})\/(\d{1,2})\/(\d{1,2})重新上架/g, "$2/$3").replace(/(\d{1,2}\/\d{1,2})重新上架/g, "$1");
const normalizeDateInput = (v = "") => {
  let text = v.trim(); if (!text) return "";
  const embeddedDate = text.match(/\d{2,4}[.\/-]\d{1,2}[.\/-]\d{1,2}/)?.[0];
  if (embeddedDate) text = embeddedDate;
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return validDate(text) ? text : text;
  const parts = text.replace(/[.\-]/g, "/").split("/").map(Number);
  let year: number, month: number, day: number;
  if (parts.length === 2) [year, month, day] = [new Date().getFullYear(), parts[0], parts[1]];
  else if (parts.length === 3) [year, month, day] = [parts[0] < 1911 ? parts[0] + 1911 : parts[0], parts[1], parts[2]];
  else return text;
  const iso = `${year.toString().padStart(4, "0")}-${month.toString().padStart(2, "0")}-${day.toString().padStart(2, "0")}`;
  const d = new Date(`${iso}T00:00:00`); return d.getFullYear() === year && d.getMonth() + 1 === month && d.getDate() === day ? iso : text;
};
const websiteEffectiveDate = (value = "") => {
  const text = String(value || "");
  const matched = text.match(/有效期\s*[:：]?\s*(\d{3,4}[.\/-]\d{1,2}[.\/-]\d{1,2})/) || text.match(/(\d{3,4}[.\/-]\d{1,2}[.\/-]\d{1,2})\s*刊登到期/);
  return matched ? normalizeDateInput(matched[1]) : "";
};
const isExpired = (r: RecordItem) => validDate(r.entrustEnd) && !!r.entrustEnd && r.entrustEnd < today();
const displayStatus = (r: RecordItem) => isExpired(r) && r.status === "委託中" ? "到期下架" : (r.status || "委託中");
const ageOf = (r: RecordItem) => {
  // 建築完成日期是編輯畫面的來源；舊 builtYear 僅作為沒有日期時的備用。
  const completionYear = String(r.completionDate || "").match(/\d{2,4}/)?.[0] || "";
  const raw = Number(completionYear || r.builtYear); if (!raw) return "";
  const year = raw < 1911 ? raw + 1911 : raw;
  const age = new Date().getFullYear() - year;
  const rocYear = year - 1911;
  return age >= 0 ? `${rocYear}年建（${age}年屋）` : "年份有誤";
};
const locationOf = (r: RecordItem) => `${r.bookLocationDate ? r.bookLocationDate.slice(5).replace("-", "/") : ""}${r.bookLocationType || "架上"}`;
const contractFromNo = (value: string) => ({ EG: "房屋一般約", EA: "房屋專約", LG: "土地一般約", LA: "土地專約", EB: "租賃一般約", EC: "租賃專約", RG: "預售一般約", RA: "預售專約" }[value.trim().toUpperCase().slice(0, 2)] || "");
const typeShort = (value = "") => value.includes("廠房") || value.includes("廠辦") ? "廠房" : value.includes("透天") ? "透天" : value.includes("公寓") ? "公寓" : value.includes("華廈") || value.includes("華厦") || value.includes("華夏") ? "華廈" : value.includes("大樓") ? "大樓" : value.includes("土地") ? "土地" : value;
const parkingShort = (value = "") => { if (!value) return ""; if (value.includes("無車位") || value.includes("無產權")) return "無"; if (value.includes("獨立車庫")) return "車庫"; const mechanical = value.match(/昇降[／/]機械\s*(.*)$/); if (mechanical) { const suffix = mechanical[1].replace(/[（(]\s*上層\s*[）)]/g, "上").replace(/[（(]\s*下層\s*[）)]/g, "下").replace(/號$/g, "").replace(/^[、，,：:\s]+/, ""); return `昇機${suffix}`; } const horizontalOrRampMechanical = value.match(/(平移|坡道)[／/]機械\s*(.*)$/); if (horizontalOrRampMechanical) { const suffix = horizontalOrRampMechanical[2].replace(/號$/g, "").replace(/^[、，,：:\s]+/, ""); return `${horizontalOrRampMechanical[1] === "平移" ? "平機" : "坡機"}${suffix}`; } const parts = value.split(/[／/]/).map(part => part.trim()).filter(Boolean); const number = (parts[parts.length - 1] || "").replace(/號$/, ""); if (value.includes("昇降") && value.includes("平面")) return `昇平${number}`; if (value.includes("坡道") && value.includes("平面")) return `坡平${number}`; return value; };
const floorShort = (value = "") => {
  if (!value) return "";
  const normalized = value
    .replace(/[０-９]/g, digit => String(digit.charCodeAt(0) - 0xfee0))
    .replace(/[一壹]/g, "1").replace(/[二貳]/g, "2").replace(/[三參]/g, "3")
    .replace(/[四肆]/g, "4").replace(/[五伍]/g, "5").replace(/[六陸]/g, "6")
    .replace(/[七柒]/g, "7").replace(/[八捌]/g, "8").replace(/[九玖]/g, "9")
    .replace(/[～〜－–—]/g, "~");
  const parts = normalized.split(/[／/]/).map(part => part.trim()).filter(Boolean);
  const totalPart = parts.find(part => /^共/.test(part)) || parts[0] || "";
  const currentPart = parts.find(part => /(?:現況)?在/.test(part)) || parts[1] || "";
  const rangeOf = (text: string) => text.match(/B?\d+(?:\s*~\s*\d+)?/i)?.[0]?.replace(/\s/g, "") || "";
  const total = rangeOf(totalPart);
  const current = rangeOf(currentPart);
  if (!current) return total || normalized;
  if (/現況在/.test(currentPart) && total && current.toUpperCase() === total.toUpperCase()) {
    return current.toUpperCase().startsWith("B") ? `${current}T` : `${current.split("~").pop()}T`;
  }
  const totalTop = total.split("~").pop() || total;
  return totalTop ? `${current}/${totalTop}` : current;
};
const floorShortFixed = (value = "") => {
  if (!value) return "";
  const digits: Record<string, string> = { "一": "1", "壹": "1", "二": "2", "貳": "2", "兩": "2", "三": "3", "參": "3", "四": "4", "肆": "4", "五": "5", "伍": "5", "六": "6", "陸": "6", "七": "7", "柒": "7", "八": "8", "捌": "8", "九": "9", "玖": "9" };
  const normalized = Array.from(value).map(char => { const code = char.charCodeAt(0); return code >= 0xff10 && code <= 0xff19 ? String(code - 0xff10) : (digits[char] || char); }).join("").replace(/[～〜－—–]/g, "~");
  const pieces = normalized.split(/[／/\n]/).map(part => part.trim()).filter(Boolean);
  const totalText = pieces.find(part => part.includes("共")) || pieces[0] || "";
  const currentText = pieces.find(part => part.includes("現況") || part.includes("在")) || pieces[1] || "";
  const range = (text: string) => text.match(/B?\d+(?:\s*~\s*\d+)?/i)?.[0]?.replace(/\s/g, "") || "";
  const total = range(totalText), current = range(currentText);
  if (!current) return total || normalized;
  if (total && current.toUpperCase() === total.toUpperCase()) return current.toUpperCase().startsWith("B") ? `${current}T` : `${current.split("~").pop()}T`;
  const totalRange = total.match(/^(\d+)~(\d+)$/), currentRange = current.match(/^(\d+)~(\d+)$/);
  const legalBottom = totalRange ? Number(totalRange[1]) : (/^\d+$/.test(total) ? Number(total) : Number.NaN);
  const legalTop = totalRange ? Number(totalRange[2]) : legalBottom;
  if (currentRange && currentRange[1] === String(legalBottom) && Number(currentRange[2]) > legalTop) {
    const currentTop = Number(currentRange[2]);
    const additions = Array.from({ length: currentTop - legalTop }, (_, index) => legalTop + index + 1).join(".");
    return `${legalTop}T/${additions}樓增建`;
  }
  const totalTop = total.split("~").pop() || total;
  return totalTop ? `${current}/${totalTop}` : current;
};
const floorPptDisplay = (value = "") => {
  if (!value) return "";
  const digits: Record<string, string> = { "一": "1", "壹": "1", "二": "2", "貳": "2", "兩": "2", "三": "3", "參": "3", "四": "4", "肆": "4", "五": "5", "伍": "5", "六": "6", "陸": "6", "七": "7", "柒": "7", "八": "8", "捌": "8", "九": "9", "玖": "9" };
  const normalized = Array.from(value).map(char => { const code = char.charCodeAt(0); return code >= 0xff10 && code <= 0xff19 ? String(code - 0xff10) : (digits[char] || char); }).join("").replace(/[～〜－—–]/g, "~");
  const pieces = normalized.split(/[／/\n]/).map(part => part.trim()).filter(Boolean);
  const totalText = pieces.find(part => part.includes("共")) || pieces[0] || "";
  const currentText = pieces.find(part => part.includes("現況") || part.includes("在")) || pieces[1] || "";
  const range = (text: string) => text.match(/B?\d+(?:\s*~\s*\d+)?/i)?.[0]?.replace(/\s/g, "") || "";
  const total = range(totalText), current = range(currentText);
  if (!current) return total ? `共${total.split("~").pop()}樓` : value;
  if (total && current.toUpperCase() === total.toUpperCase() && current.toUpperCase().startsWith("B")) return value;
  const totalTop = total.split("~").pop() || total;
  return totalTop ? `在${current}樓/共${totalTop}樓` : `在${current}樓`;
};
const contractShort = (value = "") => value.includes("一般") ? "一般約" : value.includes("專") ? "專約" : value;
const layoutFull = (value = "", type = "") => {
  const raw = String(value || "").trim();
  if (!raw) return "";
  if (typeShort(type) === "土地" || /^(?:土地\s*)+$/.test(raw)) return "土地";
  const numberOf = (pattern: RegExp) => raw.match(pattern)?.[1] || "";
  let room = numberOf(/(\d+)\s*房/);
  let hall = numberOf(/(\d+)\s*廳/);
  let bath = numberOf(/(\d+)\s*衛(?:浴)?/);
  let balcony = numberOf(/(\d+)\s*陽台/);
  if (!room && !hall && !bath && !balcony) {
    const parts = raw.match(/\d+/g) || [];
    [room, hall, bath, balcony] = [parts[0] || "", parts[1] || "", parts[2] || "", parts[3] || ""];
  }
  if (!room && !hall && !bath && !balcony) return raw;
  return `${room || "0"}房${hall || "0"}廳${bath || "0"}衛${balcony || "0"}陽台`;
};
const layoutForHouseWorkbook = (value = "", type = "") => {
  const full = layoutFull(value, type);
  return full.replace(/0陽台$/, "");
};
const layoutShort = (value = "", type = "") => {
  const full = layoutFull(value, type);
  if (full === "土地") return "土地";
  const room = full.match(/(\d+)\s*房/)?.[1];
  const hall = full.match(/(\d+)\s*廳/)?.[1];
  const bath = full.match(/(\d+)\s*衛/)?.[1];
  return [room, hall, bath].every(Boolean) ? [room, hall, bath].join(".") : full;
};
const nameMatches = (developer = "", person = "") => { const a = developer.replace(/\s/g, ""); const b = person.replace(/\s/g, ""); return !!a && !!b && (a.includes(b) || b.includes(a)); };
const isImportedDashPlaceholder = (value: unknown) => /^[\s－—-]*$/.test(String(value ?? ""));
const clearImportedContractChangePlaceholders = (record: RecordItem): RecordItem => {
  const contractChangeKeys = ["contractChangeNo", "coverChangeNo", "coverChangeDate", "coverChangePurpose", "coverNoChange"];
  const cleared = Object.fromEntries(contractChangeKeys.filter(key => isImportedDashPlaceholder(record[key])).map(key => [key, ""]));
  return Object.keys(cleared).length ? { ...record, ...cleared } : record;
};
const normalizeCompleteDateInput = (value = "") => /^\d{2,4}[.\/-]\d{1,2}[.\/-]\d{2}$/.test(value.trim()) ? normalizeDateInput(value) : value;
const clearImportedLandLabels = (record: RecordItem): RecordItem => {
  const landOnly = (value: unknown) => /^(?:土地|建地)+$/.test(String(value ?? "").replace(/[\s／/]/g, ""));
  const cleared = Object.fromEntries(["titleFloor", "currentFloor", "floor", "layout"].filter(key => landOnly(record[key])).map(key => [key, ""]));
  return Object.keys(cleared).length ? { ...record, ...cleared } : record;
};
const moveRestoredTextToCaseNote = (record: RecordItem): RecordItem => {
  const caseName = String(record.caseName || "");
  const restoredTexts = [...caseName.matchAll(/(?:\d{2,4}[.\/-]\d{1,2}[.\/-]\d{1,2}\s*)?重新上架/g)].map(match => match[0].trim());
  if (!restoredTexts.length) return record;
  const cleanCaseName = caseName.replace(/(?:\d{2,4}[.\/-]\d{1,2}[.\/-]\d{1,2}\s*)?重新上架/g, "").replace(/[　\s｜|·・—-]+$/g, "").trim();
  const existingNote = String(record.caseNameNote || "").trim();
  const additions = restoredTexts.filter(text => !existingNote.includes(text));
  const caseNameNote = [existingNote, ...additions].filter(Boolean).join("　");
  return { ...record, caseName: cleanCaseName, caseNameNote, caseNameNoteModifiedAt: record.caseNameNoteModifiedAt || new Date().toISOString() };
};
const districtFromAddress = (value = "") => {
  const address = String(value || "").replace(/臺/g, "台").trim();
  const tainan = address.match(/^台南市(.{1,4}?區)/)?.[1];
  if (tainan) return tainan;
  const outside = address.match(/^(.{2,3}?)(?:市|縣)(.{1,4}?)(?:區|鄉|鎮|市)/);
  return outside ? `${outside[1]}${outside[2]}` : "";
};
const salesBookDateCorrections = new Set(["EG0522899", "EG0522916", "LG0132934", "EG0522910", "EG0522911", "EG0522912", "EG0522915", "EG0522908"]);
const applySourceLayoutFixes = (records: RecordItem[]) => records.map(record => {
  record = moveRestoredTextToCaseNote(clearImportedLandLabels(clearImportedContractChangePlaceholders(record)));
  // 已到期下架的案件若委託結束已被改到今天或未來，自動恢復委託中；重新上架不是每日「更新物件」。
  if (record.archived && record.status === "到期下架" && validDate(record.entrustEnd) && !isExpired(record)) {
    const history = recordUpdateHistory(record);
    delete history[today()];
    const reopenNote = `${shortRocMonthDay(today())}重新上架`;
    const existingNote = String(record.caseNameNote || "").trim();
    const normalizedNote = existingNote === shortRocMonthDay(today()) ? reopenNote : existingNote;
    record = {
      ...record,
      status: "委託中",
      archived: "",
      _restoredAt: today(),
      caseNameNote: normalizedNote.includes(reopenNote) ? normalizedNote : [normalizedNote, reopenNote].filter(Boolean).join("　"),
      caseNameNoteModifiedAt: new Date().toISOString(),
      _updateHistory: JSON.stringify(history),
    };
  }
  // V73 前的新市真祥家當日歷程被累加；依實際操作保留本次「總價」更新。
  if (String(record.caseName || "").includes("新市真祥家3輕齡四套房車墅")) {
    const history = recordUpdateHistory(record);
    if (history[today()] && record._dailyUpdateNormalizedV75 !== "1") {
      record = { ...record, _updateHistory: JSON.stringify({ ...history, [today()]: ["price"] }), _dailyUpdateNormalizedV75: "1" };
    }
  }
  const normalizedReportDate = normalizeDateInput(String(record.reportDate || ""));
  const normalizedUpdateDate = normalizeDateInput(String(record.updateDate || ""));
  record = { ...record, ...(!String(record.area || "").trim() && districtFromAddress(record.address) ? { area: districtFromAddress(record.address) } : {}), ...(salesBookDateCorrections.has(String(record.propertyNo || "").trim()) && !record.salesBookDate ? { salesBookDate: "2026-07-31", salesBook: "製作" } : {}), ...(validDate(normalizedReportDate) ? { reportDate: normalizedReportDate } : {}), ...(validDate(normalizedUpdateDate) ? { updateDate: normalizedUpdateDate } : {}) };
  if (String(record.propertyNo || "").trim().toUpperCase() === "LG0107941" && String(record.depth || "").replace(/\s/g, "") !== "45.8") record = { ...record, depth: "45.8" };
  if (String(record.propertyNo || "").trim().toUpperCase() === "LG0113338" && String(record.depth || "").replace(/\s/g, "") !== "33.5") record = { ...record, depth: "33.5" };
  if (String(record.propertyNo || "").trim().toUpperCase() === "LG0118365" && String(record.depth || "").trim()) record = { ...record, depth: "" };
  if (String(record.propertyNo || "").trim().toUpperCase() === "EG0464136") record = { ...record, road: "6", frontage: "32", depth: "80" };
  if (String(record.propertyNo || "").trim().toUpperCase() === "LA0063144") record = { ...record, zoning: "商業區", coverage: "60/80", far: "360/320", coverageCombined: "60/360、80/320" };
  if (String(record.propertyNo || "").trim().toUpperCase() === "LA0064339") record = { ...record, depth: "61.5/66.8" };
  if (String(record.propertyNo || "").trim().toUpperCase() === "EG0390636") record = { ...record, parkingOwnership: "固定車位", parkingMethod: "坡道/平面", parkingNo: "B1 7號", parking: "坡道/平面B1 7號" };
  if (String(record.propertyNo || "").trim().toUpperCase() === "EG0438672") record = { ...record, parkingOwnership: "無車位", parkingMethod: "", parkingType: "", parkingNo: "", parking: "無車位" };
  const sourceBalcony = sourceBalconyFixes[String(record.propertyNo || "").trim().toUpperCase()];
  if (!sourceBalcony || typeShort(record.type) === "土地" || /^(?:LG|LA)/i.test(record.propertyNo || "")) return record;
  const balcony = sourceBalcony.match(/(\d+)\s*陽台/)?.[1];
  const current = String(record.layout || "").trim();
  const normalized = layoutFull(current, record.type);
  if (!balcony || !normalized || normalized === "土地") return record;
  const corrected = normalized.replace(/\d+陽台$/, `${balcony}陽台`);
  return corrected === normalized ? record : { ...record, layout: corrected };
});
const daysUntil = (date = "") => validDate(date) && date ? Math.ceil((Date.parse(`${date}T00:00:00`) - Date.parse(`${today()}T00:00:00`)) / 86400000) : 99999;
const nextDate = (date = "") => { const normalized = normalizeDateInput(date); if (!validDate(normalized)) return today(); const value = new Date(`${normalized}T00:00:00`); value.setDate(value.getDate() + 1); return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}-${String(value.getDate()).padStart(2, "0")}`; };
const activeGroupKey = (record: RecordItem) => {
  const address = String(record.address || "").replace(/臺/g, "台").trim();
  const area = String(record.area || "").replace(/臺/g, "台").trim();
  const addressParts = address.match(/^(.{2,3}?[市縣])(.{1,4}?[區鄉鎮市])/);
  const city = addressParts?.[1] || (/^(台南市|高雄市|嘉義市|嘉義縣|屏東縣|台北市|新北市|桃園市|台中市|彰化縣|雲林縣)/.exec(area)?.[1] || "");
  const district = addressParts?.[2] || area.replace(city, "") || "未填地區";
  const outside = city && city !== "台南市";
  const typeOrder: Record<string, string> = { "透天": "1", "華廈": "2", "大樓": "3", "公寓": "4", "廠房": "5", "土地": "6" };
  const kind = typeShort(record.type);
  return `${outside ? "1外縣市" : "0台南市"}|${outside ? city : ""}|${district}|${typeOrder[kind] || "9"}${kind}`;
};
const areaCategory = (record: RecordItem) => {
  const address = String(record.address || "").trim();
  if (!address.includes("台南市")) return "外縣市";
  const district = districtFromAddress(address);
  return ["北區", "東區", "中西區", "南區", "永康區", "安平區", "仁德區", "安南區"].includes(district) ? district : "其他區";
};
const activeGroupKeyFixed = (record: RecordItem) => {
  const areaOrder: Record<string, string> = { "北區": "01", "東區": "02", "中西區": "03", "南區": "04", "永康區": "05", "安平區": "06", "仁德區": "07", "安南區": "08", "其他區": "09", "外縣市": "10" };
  const typeOrder: Record<string, string> = { "透天": "1", "華廈": "2", "大樓": "3", "公寓": "4", "廠房": "5", "土地": "6" };
  const area = areaCategory(record), kind = typeShort(record.type);
  return `${areaOrder[area] || "99"}|${typeOrder[kind] || "9"}${kind}`;
};
const isRentalRecord = (record: RecordItem) => /^(EB|EC)/i.test(String(record.propertyNo || "").trim()) || String(record.contractType || "").includes("租賃");
const sortActiveRecords = (items: RecordItem[]) => items.map((record, index) => ({ record, index })).sort((a, b) => {
  const rentalOrderA = isRentalRecord(a.record) ? "0" : "1";
  const rentalOrderB = isRentalRecord(b.record) ? "0" : "1";
  const group = `${rentalOrderA}|${activeGroupKeyFixed(a.record)}`.localeCompare(`${rentalOrderB}|${activeGroupKeyFixed(b.record)}`, "zh-TW", { numeric: true });
  if (group) return group;
  const date = String(a.record.reportDate || "9999-12-31").localeCompare(String(b.record.reportDate || "9999-12-31"), "zh-TW", { numeric: true });
  return date || a.index - b.index;
}).map(item => item.record);
// 封存依實際封存日（到期自動下架則以委託結束日）由新到舊排列。
const sortArchivedRecords = (items: RecordItem[]) => items.map((record, index) => ({ record, index })).sort((a, b) => {
  const dateOf = (item: RecordItem) => String(item.archived || item._archiveActionDate || item.entrustEnd || "");
  const date = dateOf(b.record).localeCompare(dateOf(a.record), "zh-TW", { numeric: true });
  return date || b.index - a.index;
}).map(item => item.record);
const chunkText = (value = "", size = 16) => { const actualSize = size === 20 ? 25 : size === 15 || size === 10 ? 12 : size; const chars = Array.from(value); return Array.from({ length: Math.ceil(chars.length / actualSize) }, (_, index) => chars.slice(index * actualSize, index * actualSize + actualSize).join("")); };
let developerPersonnelForDisplay: Person[] = [];
const knownDeveloperFullNames = ["王啟山", "王若芸", "王俞云", "王妤宸", "林玉環", "林顯昌", "林姿岑", "林俊嘉", "陳帝元", "陳珮菁", "陳信良", "郭建佑", "謝馨儀", "蔡宇育", "田庭宇", "吳佩玲", "黃文成", "買淑玲", "劉勝仁", "張小曼", "李享嶧", "阮氏金水", "宋喜輝", "柯育婷", "李麗卉", "施紹薇", "余沛臻", "葉翊緁", "楊巧甄"];
const developerNameLines = (value = "", personnel: Person[] = []) => {
  const rawConfiguredPeople = personnel.length ? personnel : developerPersonnelForDisplay;
  const knownAliases = new Set(knownDeveloperFullNames.map(name => Array.from(name).slice(-2).join("")));
  const configuredPeople = rawConfiguredPeople.filter(person => { const name = String(person.name || "").trim(); return Array.from(name).length >= 3 || !knownAliases.has(name); });
  const configuredNames = new Set(configuredPeople.map(person => String(person.name || "").trim()).filter(Boolean));
  const fallbackPeople = knownDeveloperFullNames.filter(name => !configuredNames.has(name)).map((name, index) => ({ id: `known-developer-${index}`, name, status: "在職" } as Person));
  const people = [...configuredPeople, ...fallbackPeople];
  const aliases = people.flatMap(person => { const full = String(person.name || "").trim(); const chars = Array.from(full); return full ? [{ token: full, full }, ...(chars.length >= 3 ? [{ token: chars.slice(-2).join(""), full }] : [])] : []; });
  const expand = (segment: string) => {
    if (segment === "王總") return ["王啟山"];
    const chars = Array.from(segment); const names: string[] = []; let index = 0;
    while (index < chars.length) {
      const remaining = chars.slice(index).join("");
      const candidates = aliases.filter(item => remaining.startsWith(item.token)).sort((a, b) => Array.from(b.token).length - Array.from(a.token).length);
      if (candidates.length) {
        const longest = Array.from(candidates[0].token).length;
        const sameToken = candidates.filter(item => Array.from(item.token).length === longest && item.token === candidates[0].token);
        const fullNames = [...new Set(sameToken.map(item => item.full))];
        names.push(fullNames.length === 1 ? fullNames[0] : candidates[0].token);
        index += longest;
      } else {
        names.push(chars.slice(index, index + 2).join(""));
        index += Math.min(2, chars.length - index);
      }
    }
    return names;
  };
  return String(value || "").split(/[\/／,，、。]+/).map(name => name.trim()).filter(Boolean).flatMap(expand);
};
const developerFullNameText = (value = "", personnel: Person[] = []) => developerNameLines(value, personnel).join("、");
const stripRestoredDisplay = (value = "") => String(value || "").replace(/(?:\d{2,4}[.\/-]\d{1,2}[.\/-]\d{1,2}\s*)?重新上架/g, "").replace(/[　\s｜|·・—-]+$/g, "").trim();
const archiveDisplayRecord = (record: RecordItem): RecordItem => ({ ...record, caseName: stripRestoredDisplay(record.caseName), caseNameNote: stripRestoredDisplay(record.caseNameNote) });
const showingFollowUpDisplayRecord = (record: RecordItem): RecordItem => record;
const sortPptRecords = (items: RecordItem[]) => items.map((record, index) => ({ record, index })).sort((a, b) => {
  const priority = (record: RecordItem) => /王啟山|蔡宇育/.test(record.developer || "") ? 0 : 1;
  const priorityDiff = priority(a.record) - priority(b.record); if (priorityDiff) return priorityDiff;
  const dateDiff = String(a.record.reportDate || "9999-12-31").localeCompare(String(b.record.reportDate || "9999-12-31"), "zh-TW", { numeric: true });
  return dateDiff || a.index - b.index;
}).map(item => item.record);

function parseTsv(text: string) {
  const rows: string[][] = []; let row: string[] = []; let cell = ""; let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') { if (quoted && text[i + 1] === '"') { cell += '"'; i++; } else quoted = !quoted; }
    else if (ch === "\t" && !quoted) { row.push(cell.trim()); cell = ""; }
    else if ((ch === "\n" || ch === "\r") && !quoted) { if (ch === "\r" && text[i + 1] === "\n") i++; row.push(cell.trim()); if (row.some(Boolean)) rows.push(row); row = []; cell = ""; }
    else cell += ch;
  }
  row.push(cell.trim()); if (row.some(Boolean)) rows.push(row);
  return rows;
}

const defaultIntakeHeaders = [
  "時間戳記", "表單填寫人", "開發１/開發２", "委託主約編號:", "委託開始 日期", "委託結束 日期", "案名", "物件(完整)地址", "(物件)現況", "鑰匙位置", "物件型態", "契約開價 (萬)", "總建坪", "室內坪=(主建物+附屬建物)", "地坪", "增建說明/坪數說明", "臨路 (米)", "面寬 (透天*、土地*)必填", "深度 (土地*)必填", "使用分區 (土地*)必填", "建蔽率/容積率 (土地*必填)(透天選填)", "建築完成日期 (房屋*必填) 格式: 082.7.17", "格局 (房)", "格局 (廳)", "格局 (衛浴)", "格局 (陽台)", "有無電梯", "權狀層數＿例如：共１～３樓", "透天請寫『現況在１～３樓』 大樓請寫『在６樓』", "大樓名稱:", "電梯數", "每層戶數", "朝向 [土地朝]", "朝向 [房屋朝]", "朝向 [大門朝]", "警衛管理", "管理費 (月/元)", "車位", "車位型態", "車位型態", "車位型態", "車位型態", "車位型態", "車位編號 (如有上下層/地下B請標記)", "特色說明1. (簡單明瞭，請勿冗長)", "特色說明2. (簡單明瞭，請勿冗長)", "特色說明3. (簡單明瞭，請勿冗長)", "特色說明4. (簡單明瞭，請勿冗長)", "鄰近國小:", "鄰近國中:", "鄰近高中:", "鄰近大專:", "市場/購物:", "公園綠地:", "注意事項", "當下進案文件 [物件相片]", "進案文件 [格局圖]", "進案文件 [現詢調]", "進案文件 [主約 契約拍照 上傳系統]", "進案文件 [契變 照片上傳系統]", "進案文件 [物件照片 上傳系統]", "中人(介紹費)"
];

function parseIntakes(text: string): IntakeData[] {
  const rows = parseTsv(text).filter(row => row.some(cell => cell.trim()));
  if (!rows.length) return [];
  const hasHeaders = rows[0].some(cell => cell.includes("時間戳記") || cell.includes("表單填寫人") || cell.includes("委託主約編號"));
  const headers = hasHeaders ? rows[0] : defaultIntakeHeaders;
  const dataRows = hasHeaders ? rows.slice(1) : rows;
  return dataRows.filter(data => data.some(cell => cell.trim())).map(data => {
    const values: Record<string, string> = {};
    headers.forEach((header, i) => { if (!header) return; const key = values[header] === undefined ? header : `${header}#${i}`; values[key] = data[i] || ""; });
    const type = intakeValue(values, "物件型態");
    const createdAt = new Date().toISOString();
    return { id: newId(), values, propertyKind: type.includes("土地") ? "純土地" : "房屋", createdAt, modifiedAt: createdAt };
  });
}

function intakeValue(values: Record<string, string>, ...needles: string[]) {
  for (const needle of needles) { const found = Object.entries(values).find(([key, value]) => key.includes(needle) && value); if (found) return found[1]; }
  return "";
}

function intakeAll(values: Record<string, string>, needle: string) { return Object.entries(values).filter(([key, value]) => key.includes(needle) && value).map(([, value]) => value); }

const directionFacing = (value = "") => {
  const text = String(value || "").trim();
  const match = text.match(/朝\s*([東南西北]{1,2})/) || text.match(/([東南西北]{1,2})$/);
  return match?.[1] || text;
};

const directionShort = (value = "") => {
  const parts = String(value || "").split(/[／/、,，]+/).map(part => part.trim()).filter(Boolean);
  if (!parts.length) return "";
  const parsed = parts.map(part => ({
    label: /房屋/.test(part) ? "房屋" : /大門/.test(part) ? "大門" : /土地/.test(part) ? "土地" : "",
    facing: directionFacing(part),
  }));
  if (parsed.length === 1) return parsed[0].facing;
  if (new Set(parsed.map(part => part.facing)).size === 1) return parsed[0].facing;
  return parsed.map(part => `${part.label || "朝向"}朝${part.facing}`).join("／");
};

const displayNoteSegments = (value = "") => String(value || "").split(/[；;]/).map(part => part.trim()).filter(part => part && !/^(?:0|無)$/.test(part)).map(part => part.includes("開發%") && !/^中人[:：]/.test(part) ? `中人:${part}` : part);
const websiteCellDisplay = (record: RecordItem, key: string) => {
  const raw = String(record[key] || "").trim();
  const down = record[`${key}DownDate`] ? `${displayRocDate(record[`${key}DownDate`])}下架` : "";
  const price5168Listed = key === "price5168" ? raw.match(/^(\d{3,4}[.\/-]\d{1,2}[.\/-]\d{1,2})\s*(上)$/) : null;
  if (price5168Listed) return { value: price5168Listed[1], notes: [price5168Listed[2], down].filter(Boolean) };
  const expiryKey = key === "platform591" ? "platform591Expiry" : key === "price5168" ? "price5168Expiry" : key === "goldExposure" ? "goldExposureExpiry" : "";
  const detectedExpiry = expiryKey ? websiteEffectiveDate(raw) : "";
  const expiryDate = expiryKey ? (record[expiryKey] || detectedExpiry) : "";
  const expiry = expiryDate ? `${displayRocDate(expiryDate)}到期` : "";
  const visibleRaw = raw.replace(/有效期\s*[:：]?\s*\d{3,4}[.\/-]\d{1,2}[.\/-]\d{1,2}/g, "").replace(/\d{3,4}[.\/-]\d{1,2}[.\/-]\d{1,2}\s*刊登到期/g, "").replace(/[\s,，;；]+$/g, "").trim();
  if (!visibleRaw || /^(?:無|旁5)$/i.test(visibleRaw)) return expiryDate ? { value: displayRocDate(expiryDate), notes: ["到期", down].filter(Boolean) } : { value: visibleRaw || "—", notes: [down].filter(Boolean) };
  const lines = visibleRaw.split(/\r?\n/).map(value => value.trim()).filter(Boolean);
  const compact = lines.join(" ");
  const leadingCode = compact.match(/^([0-9A-Za-zＡ-Ｚａ-ｚ][0-9A-Za-zＡ-Ｚａ-ｚ_－—\-]*)/);
  if (leadingCode) {
    const value = leadingCode[1];
    const note = compact.slice(value.length).replace(/^[\s／/、,，.:：#＃()（）\[\]【】－—\-]+/, "").trim();
    return { value, notes: [note, expiry, down].filter(Boolean) };
  }
  const parts = compact.split(/[－—\-\n]/).map(value => value.trim()).filter(Boolean);
  return { value: parts.shift() || "—", notes: [...parts, expiry, down].filter(Boolean) };
};
const colorSheetAttention = (value = "", additionNotes = "") => {
  // Excel 注意事項只使用委託中編輯的「彩色表注意事項」。
  const notes = displayNoteSegments(value).filter(part => !/(?:中人|介紹費|開發\s*%)/.test(part));
  return notes.join("；");
};
const schoolSummary = (record: RecordItem) => [record.elementarySchool, record.juniorHighSchool, record.seniorHighSchool, record.collegeSchool].map(value => String(value || "").trim()).filter(Boolean).join("／") || String(record.school || "").trim();

function intakeToRecord(intake: IntakeData, existing?: RecordItem): RecordItem {
  const v = intake.values; const no = intakeValue(v, "委託主約編號");
  const layout = [intakeValue(v, "格局 (房)"), intakeValue(v, "格局 (廳)"), intakeValue(v, "格局 (衛浴)"), intakeValue(v, "格局 (陽台)")].filter(Boolean).join("");
  const coverageFar = intakeValue(v, "建蔽率/容積率").split(/[／/]/);
  const completion = intakeValue(v, "建築完成日期");
  const middleman = intakeValue(v, "中人");
  const notes = [intakeValue(v, "增建說明"), middleman ? `中人:${middleman}` : "", intakeValue(v, "注意事項")].filter(Boolean).join("；");
  const parking = [intakeValue(v, "車位"), ...intakeAll(v, "車位型態"), intakeValue(v, "車位編號")].filter(Boolean).join("／");
  const parkingOwnership = /無車位/.test(parking) ? "無車位" : /停自有地|自有地/.test(parking) ? "停自有地" : /車位另租/.test(parking) ? "車位另租" : /抽籤/.test(parking) ? "抽籤決定" : /固定/.test(parking) ? "固定車位" : "";
  const parkingType = /先到先停/.test(parking) ? "先到先停" : /排隊/.test(parking) ? "排隊等候" : /停自有地|自有地/.test(parking) ? "停自有地" : /車位另租/.test(parking) ? "車位另租" : /抽籤/.test(parking) ? "抽籤決定" : /固定/.test(parking) ? "固定車位" : "";
  const parkingMethod = parking.match(/坡道[／/]平面|坡道[／/]機械|昇降[／/]平面|昇降[／/]機械|庭院|平移[／/]機械/)?.[0] || "";
  const elementarySchool = intakeValue(v, "鄰近國小"); const juniorHighSchool = intakeValue(v, "鄰近國中"); const seniorHighSchool = intakeValue(v, "鄰近高中"); const collegeSchool = intakeValue(v, "鄰近大專");
  return { ...blankRecord(), ...(existing || {}), propertyNo: no, contractType: contractFromNo(no), type: intakeValue(v, "物件型態"), status: existing?.status || "委託中", area: intakeValue(v, "物件(完整)地址").replace(/^.*?[市縣]/, "").slice(0, 3), caseName: intakeValue(v, "案名"), address: intakeValue(v, "物件(完整)地址"), price: intakeValue(v, "契約開價"), direction: intakeValue(v, "朝向 [房屋朝]", "朝向 [大門朝]", "朝向 [土地朝]"), completionDate: completion, builtYear: completion ? String(Number(completion.split(/[./]/)[0])) : "", titleFloor: intakeValue(v, "權狀層數"), currentFloor: intakeValue(v, "透天請寫"), floor: [intakeValue(v, "權狀層數"), intakeValue(v, "透天請寫")].filter(Boolean).join("／"), layout, indoorPing: intakeValue(v, "室內坪"), buildingPing: intakeValue(v, "總建坪"), landPing: intakeValue(v, "地坪"), parking, parkingOwnership, parkingType, parkingMethod, parkingNo: intakeValue(v, "車位編號"), buildingName: intakeValue(v, "大樓名稱"), elevatorCount: intakeValue(v, "電梯數"), unitsPerFloor: intakeValue(v, "每層戶數"), managementMethod: intakeValue(v, "警衛管理"), market: intakeValue(v, "市場/購物"), park: intakeValue(v, "公園綠地"), elementarySchool, juniorHighSchool, seniorHighSchool, collegeSchool, school: [elementarySchool, juniorHighSchool, seniorHighSchool, collegeSchool].filter(Boolean).join("／") || existing?.school || "", feature1: intakeValue(v, "特色說明1"), feature2: intakeValue(v, "特色說明2"), feature3: intakeValue(v, "特色說明3"), feature4: intakeValue(v, "特色說明4"), attentionNotes: [intakeValue(v, "增建說明"), intakeValue(v, "注意事項")].filter(Boolean).join("；"), managementFee: intakeValue(v, "管理費"), key: intakeValue(v, "鑰匙位置"), currentState: intakeValue(v, "(物件)現況"), road: intakeValue(v, "臨路"), frontage: intakeValue(v, "面寬"), depth: intakeValue(v, "深度"), zoning: intakeValue(v, "使用分區"), coverage: coverageFar[0] || "", far: coverageFar[1] || "", developer: intakeValue(v, "開發１/開發２"), entrustStart: normalizeDateInput(intakeValue(v, "委託開始")), entrustEnd: normalizeDateInput(intakeValue(v, "委託結束")), reportDate: existing?.reportDate || today(), updateDate: today(), groupViewDate: existing?.groupViewDate || intake.groupViewDate || "", notes, photoInfo: existing?.photoInfo || "" };
}

function recordToIntake(record: RecordItem): IntakeData {
  const values: Record<string, string> = {
    "時間戳記": new Date().toLocaleString("zh-TW"), "表單填寫人": "", "開發１/開發２": developerFullNameText(record.developer),
    "委託主約編號:": record.propertyNo, "委託開始 日期": displayRocDate(record.entrustStart), "委託結束 日期": displayRocDate(record.entrustEnd),
    "案名": record.caseName, "物件(完整)地址": record.address, "(物件)現況": record.currentState, "鑰匙位置": record.key,
    "物件型態": record.type, "契約開價 (萬)": record.price, "總建坪": record.buildingPing, "室內坪=(主建物+附屬建物)": record.indoorPing,
    "地坪": record.landPing, "增建說明/坪數說明": "", "臨路 (米)": record.road, "面寬 (透天*、土地*)必填": record.frontage,
    "深度 (土地*)必填": record.depth, "使用分區 (土地*)必填": record.zoning, "建蔽率/容積率 (土地*必填)(透天選填)": [record.coverage, record.far].filter(Boolean).join("/"),
    "建築完成日期 (房屋*必填) 格式: 082.7.17": record.completionDate || record.builtYear, "權狀層數＿例如：共１～３樓": record.titleFloor || record.floor, "透天請寫『現況在１～３樓』 大樓請寫『在６樓』": record.currentFloor,
    "大樓名稱:": record.buildingName, "電梯數": record.elevatorCount, "每層戶數": record.unitsPerFloor, "警衛管理": record.managementMethod,
    "管理費 (月/元)": record.managementFee, "車位": [record.parkingOwnership, record.parkingType, record.parkingMethod].filter(Boolean).join("／") || record.parking, "車位編號 (如有上下層/地下B請標記)": record.parkingNo,
    "特色說明1. (簡單明瞭，請勿冗長)": record.feature1, "特色說明2. (簡單明瞭，請勿冗長)": record.feature2, "特色說明3. (簡單明瞭，請勿冗長)": record.feature3, "特色說明4. (簡單明瞭，請勿冗長)": record.feature4, "注意事項": record.attentionNotes || record.notes,
  };
  return { id: newId(), values, propertyKind: record.type.includes("土地") ? "純土地" : "房屋", createdAt: new Date().toISOString(), groupViewDate: record.groupViewDate || "" };
}

function syncRecordToDraftValues(draft: IntakeData, record: RecordItem): Record<string, string> {
  const values = { ...draft.values };
  const setValue = (needle: string, value: string, fallback = needle) => { const key = Object.keys(values).find(name => name.includes(needle)) || fallback; values[key] = value || ""; };
  setValue("開發１/開發２", developerFullNameText(record.developer), "開發１/開發２"); setValue("委託主約編號", record.propertyNo, "委託主約編號:");
  setValue("委託開始", displayRocDate(record.entrustStart), "委託開始 日期"); setValue("委託結束", displayRocDate(record.entrustEnd), "委託結束 日期");
  setValue("案名", record.caseName); setValue("物件(完整)地址", record.address); setValue("(物件)現況", record.currentState); setValue("鑰匙位置", record.key);
  setValue("物件型態", record.type); setValue("契約開價", record.price, "契約開價 (萬)"); setValue("朝向", record.direction, "朝向 [房屋朝]"); setValue("總建坪", record.buildingPing); setValue("室內坪", record.indoorPing);
  setValue("地坪", record.landPing); setValue("臨路", record.road, "臨路 (米)"); setValue("面寬", record.frontage); setValue("深度", record.depth); setValue("使用分區", record.zoning);
  setValue("建蔽率/容積率", [record.coverage, record.far].filter(Boolean).join("/")); setValue("建築完成日期", record.completionDate || record.builtYear); setValue("權狀層數", record.titleFloor || record.floor); setValue("透天請寫", record.currentFloor);
  setValue("格局 (房)", record.layout.match(/(\d+)\s*房/)?.[0] || ""); setValue("格局 (廳)", record.layout.match(/(\d+)\s*廳/)?.[0] || ""); setValue("格局 (衛浴)", record.layout.match(/(\d+)\s*衛(?:浴)?/)?.[0] || ""); setValue("格局 (陽台)", record.layout.match(/(\d+)\s*陽台/)?.[0] || "");
  setValue("大樓名稱", record.buildingName); setValue("電梯數", record.elevatorCount); setValue("每層戶數", record.unitsPerFloor); setValue("警衛管理", record.managementMethod);
  setValue("特色說明1", record.feature1); setValue("特色說明2", record.feature2); setValue("特色說明3", record.feature3); setValue("特色說明4", record.feature4);
  setValue("管理費", record.managementFee); setValue("車位", [record.parkingOwnership, record.parkingType, record.parkingMethod].filter(Boolean).join("／") || record.parking); setValue("車位編號", record.parkingNo); setValue("市場/購物", record.market); setValue("公園綠地", record.park); setValue("鄰近國小", record.elementarySchool); setValue("鄰近國中", record.juniorHighSchool); setValue("鄰近高中", record.seniorHighSchool); setValue("鄰近大專", record.collegeSchool); setValue("注意事項", record.attentionNotes || record.notes);
  return values;
}

const sample: RecordItem = {
  ...blankRecord(), propertyNo: "EA024", type: "華廈：10樓含以下（有電梯）", contractType: "房屋專約", status: "委託中", area: "北區",
  caseName: "晴光花園美宅", address: "中清路一段 168 號", price: "1,688", direction: "坐北朝南", builtYear: "103",
  floor: "1-3樓", layout: "4房2廳3衛", indoorPing: "32.6", buildingPing: "48.2", landPing: "26.4", parking: "門前停車",
  managementFee: "0", key: "公司鑰匙櫃 A3", currentState: "自住", road: "8米", frontage: "5.2米", depth: "16米",
  zoning: "住二", coverage: "60", far: "180", developer: "王小明", entrustStart: today(), entrustEnd: "2026-12-31",
  reportDate: today(), updateDate: today(), groupViewDate: "", bookLocationType: "架上", bookLocationDate: today(),
  salesBook: "已製作", notes: "屋況佳，帶看前請先聯絡。", photoInfo: "雲端相簿編號 P024", platform591: "591-246810", yes319: "", houseinfor: "H-13579",
  windowAd: "櫥窗 A2", led: "LED-07", homeWeb: "W-024", price5168: "5168-024", goldExposure: "", photos: [],
};

export default function Home() {
  const [internalView] = useState(() => typeof window !== "undefined" && new URLSearchParams(window.location.search).get("view") === "internal");
  const [records, setRecords] = useState<RecordItem[]>([]);
  const [settings, setSettings] = useState<Settings>({ personnel: [], supabaseUrl: CASE_FILE_SUPABASE_URL, supabaseKey: CASE_FILE_SUPABASE_PUBLISHABLE_KEY, supabaseTable: CASE_FILE_SUPABASE_TABLE, supabaseRecord: "main", bookReviewCurrentDate: "2026-07-30", bookReviewNextDate: "2026-09-30" });
  const [storageReady, setStorageReady] = useState(false);
  developerPersonnelForDisplay = settings.personnel;
  const [cloudSession, setCloudSession] = useState<CloudSession | null>(null);
  const [cloudLastUploadAt, setCloudLastUploadAt] = useState(() => typeof window !== "undefined" ? localStorage.getItem(CLOUD_LAST_UPLOAD_KEY) || "" : "");
  const [tab, setTab] = useState<"active" | "archive" | "activity" | "inventory" | "tour" | "keys" | "public" | "settings" | "intake">("active");
  const [query, setQuery] = useState("");
  const [archiveQuery, setArchiveQuery] = useState("");
  const [missingDataReminderOpen, setMissingDataReminderOpen] = useState(false);
  const [websiteFilter, setWebsiteFilter] = useState("");
  const [showingFollowUpOnly, setShowingFollowUpOnly] = useState(false);
  const [showingFollowUpOpen, setShowingFollowUpOpen] = useState(false);
  const [showingFollowUpRecordId, setShowingFollowUpRecordId] = useState("");
  const [showingFollowUpStart, setShowingFollowUpStart] = useState(today());
  const [showingFollowUpDue, setShowingFollowUpDue] = useState("");
  const [advancedFilterOpen, setAdvancedFilterOpen] = useState(false);
  const [advancedFilter, setAdvancedFilter] = useState<AdvancedFilter>(() => blankAdvancedFilter());
  const [bookReviewOpenRequest, setBookReviewOpenRequest] = useState(0);
  const [monthlyProgressOpen, setMonthlyProgressOpen] = useState(false);
  const [frontLastLogins, setFrontLastLogins] = useState<Record<string, string>>({});
  const [editing, setEditing] = useState<RecordItem | null>(null);
  useEffect(() => {
    setEditing(current => current ? records.find(record => record.id === current.id) || current : current);
  }, [records]);
  const [newCaseReminder, setNewCaseReminder] = useState<RecordItem | null>(null);
  const [newCaseReminderBatchIds, setNewCaseReminderBatchIds] = useState<string[]>([]);
  const deferredNewCaseReminderIds = useRef<Set<string>>(new Set());
  const [expiryReminderOpen, setExpiryReminderOpen] = useState(false);
  const [websitePoReminderOpen, setWebsitePoReminderOpen] = useState(false);
  const [archiveCleanupReminderOpen, setArchiveCleanupReminderOpen] = useState(false);
  const [archiveCleanupSort, setArchiveCleanupSort] = useState<{ key: string; direction: "asc" | "desc" }>({ key: "", direction: "asc" });
  const suppressArchiveCleanupReminderRef = useRef(false);
  const [dealCompletionReminderOpen, setDealCompletionReminderOpen] = useState(false);
  const [printEditor, setPrintEditor] = useState<{ kind: "color"; data: RecordItem } | null>(null);
  const [colorWorkbookDownloading, setColorWorkbookDownloading] = useState(false);
  const [restoreChoiceRecord, setRestoreChoiceRecord] = useState<RecordItem | null>(null);
  const [archiveChoice, setArchiveChoice] = useState<{ record: RecordItem; status: string; date: string; salesPerson: string; reason: string } | null>(null);
  const [publicUnlocked, setPublicUnlocked] = useState(false);
  const [publicPersonId, setPublicPersonId] = useState("");
  const [publicScope, setPublicScope] = useState<"activity" | "mine" | "all" | "contacts">("mine");
  const [publicExpiryFilter, setPublicExpiryFilter] = useState<"all" | "15" | "30">("all");
  const [publicQuery, setPublicQuery] = useState("");
  const [publicZoom, setPublicZoom] = useState(() => typeof window !== "undefined" && window.innerWidth <= 1100 ? 50 : 100);
  const [password, setPassword] = useState("");
  const [notice, setNotice] = useState("");
  const [intakeRaw, setIntakeRaw] = useState("");
  const [intakeDrafts, setIntakeDrafts] = useState<IntakeData[]>([]);
  const [selectedIntakeId, setSelectedIntakeId] = useState("");
  const [pptPickerOpen, setPptPickerOpen] = useState(false);
  const [pptExtraSearch, setPptExtraSearch] = useState("");
  const [pptAdHocOpen, setPptAdHocOpen] = useState(false);
  const [pptAdHocDraft, setPptAdHocDraft] = useState<RecordItem>(() => blankRecord());
  const [pptAdHocRecords, setPptAdHocRecords] = useState<RecordItem[]>([]);
  const [pptConfirmedSnapshots, setPptConfirmedSnapshots] = useState<Record<string, RecordItem>>({});
  const [pptPreviewRecord, setPptPreviewRecord] = useState<RecordItem | null>(null);
  const [pptShowExtras, setPptShowExtras] = useState(false);
  const [pptExtraIds, setPptExtraIds] = useState<string[]>([]);
  const [pptOrderIds, setPptOrderIds] = useState<string[]>([]);
  const [pptWeekStart, setPptWeekStart] = useState(currentPptWeek().start);
  const [pptWeekMenuOpen, setPptWeekMenuOpen] = useState(false);
  const [pptCustomStart, setPptCustomStart] = useState("");
  const [pptCustomEnd, setPptCustomEnd] = useState("");
  const [pptCustomMeeting, setPptCustomMeeting] = useState("");
  const pptWeekLoadedRef = useRef("");
  const reminderAnchorRecord = records.find(record => String(record.address || "").includes("富農街一段188巷40號"));
  const reminderAnchorDraft = reminderAnchorRecord ? intakeDrafts.find(draft => draft.linkedRecordId === reminderAnchorRecord.id) : undefined;
  const reminderAnchorTime = reminderAnchorDraft?.enteredAt || reminderAnchorDraft?.createdAt || "";
  const intakeReminderDraft = (record: RecordItem) => intakeDrafts.find(draft => draft.linkedRecordId === record.id && !!draft.enteredAt);
  const isIntakeReminderRecord = (record: RecordItem) => {
    if (record._newCaseReminderSource === "intake") return true;
    const draft = intakeReminderDraft(record);
    if (!draft || !reminderAnchorTime) return false;
    return (draft.enteredAt || draft.createdAt || "") >= reminderAnchorTime;
  };
  const pendingIntakeReminderRecords = records.filter(record => isIntakeReminderRecord(record) && newCaseReminderPending(record) && !deferredNewCaseReminderIds.current.has(record.id));
  const reminderCandidate = pendingIntakeReminderRecords[0];
  const newCaseReminderBatchRecords = newCaseReminderBatchIds.map(id => records.find(record => record.id === id) || (newCaseReminder?.id === id ? newCaseReminder : null)).filter((record): record is RecordItem => !!record);
  useEffect(() => {
    // 業務內部總表僅供查看；新案件完成提醒只在管理模式處理。
    if (internalView || !storageReady || editing || newCaseReminder || !reminderCandidate) return;
    setNewCaseReminder({ ...reminderCandidate });
    setNewCaseReminderBatchIds(pendingIntakeReminderRecords.map(record => record.id));
  }, [internalView, storageReady, editing, newCaseReminder, reminderCandidate]);
  useEffect(() => { setPptCustomStart(""); setPptCustomEnd(""); setPptCustomMeeting(""); }, [pptWeekStart]);
  useEffect(() => { if (!pptPickerOpen) { setPptCustomStart(""); setPptCustomEnd(""); setPptCustomMeeting(""); setPptExtraSearch(""); setPptAdHocOpen(false); } }, [pptPickerOpen]);
  useEffect(() => {
    try {
      const saved = JSON.parse(localStorage.getItem(PPT_WEEK_SELECTIONS_KEY) || "{}")[pptWeekStart] || {};
      pptWeekLoadedRef.current = pptWeekStart;
      setPptExtraIds(Array.isArray(saved.extraIds) ? saved.extraIds : []);
      setPptOrderIds(Array.isArray(saved.orderIds) ? saved.orderIds : []);
      setPptAdHocRecords(Array.isArray(saved.adHocRecords) ? saved.adHocRecords : []);
      setPptConfirmedSnapshots(saved.confirmedSnapshots && typeof saved.confirmedSnapshots === "object" ? saved.confirmedSnapshots : {});
    } catch {
      pptWeekLoadedRef.current = pptWeekStart;
      setPptExtraIds([]); setPptOrderIds([]); setPptAdHocRecords([]); setPptConfirmedSnapshots({});
    }
  }, [pptWeekStart]);
  useEffect(() => {
    if (!storageReady || pptWeekLoadedRef.current !== pptWeekStart) return;
    try {
      const all = JSON.parse(localStorage.getItem(PPT_WEEK_SELECTIONS_KEY) || "{}");
      all[pptWeekStart] = { extraIds: pptExtraIds, orderIds: pptOrderIds, adHocRecords: pptAdHocRecords, confirmedSnapshots: pptConfirmedSnapshots };
      localStorage.setItem(PPT_WEEK_SELECTIONS_KEY, JSON.stringify(all));
    } catch {}
  }, [storageReady, pptWeekStart, pptExtraIds, pptOrderIds, pptAdHocRecords, pptConfirmedSnapshots]);
  const [tourItems, setTourItems] = useState<TourItem[]>([]);
  const [tourDate, setTourDate] = useState(today());
  const [tourTitle, setTourTitle] = useState(`${displayRocDate(today()).replace(/\//g, ".")}團看`);
  const [tourModifiedAt, setTourModifiedAt] = useState("");
  const [tourHistory, setTourHistory] = useState<TourHistory[]>([]);
  const selectedIntakeRef = useRef("");
  const cloudSyncBaselineRef = useRef("");
  const cloudSyncTimerRef = useRef<number | null>(null);
  const cloudAutoPullRef = useRef("");
  const editingInitialRef = useRef("");
  const editingInitialIdRef = useRef("");
  const personnelNameSignature = settings.personnel.map(person => `${person.id}:${person.name}:${person.status}`).join("|");
  const developerNormalizationSignature = records.map(record => `${record.id}:${record.developer || ""}`).join("|");
  const draftDeveloperSignature = intakeDrafts.map(draft => `${draft.id}:${intakeValue(draft.values, "開發１/開發２")}`).join("|");

  useEffect(() => {
    const usablePeople = settings.personnel.filter(person => String(person.name || "").trim());
    setRecords(previous => { let changed = false; const next = previous.map(record => {
      const developer = developerFullNameText(record.developer || "", usablePeople);
      if (developer && developer !== record.developer) { changed = true; return { ...record, developer }; }
      return record;
    }); return changed ? next : previous; });
    setIntakeDrafts(previous => { let anyChanged = false; const next = previous.map(draft => {
      let changed = false;
      const values = Object.fromEntries(Object.entries(draft.values).map(([key, value]) => {
        if (!/開發/.test(key) || !String(value || "").trim()) return [key, value];
        const fullName = developerFullNameText(String(value), usablePeople);
        if (fullName && fullName !== value) changed = true;
        return [key, fullName || value];
      }));
      if (changed) anyChanged = true;
      return changed ? { ...draft, values } : draft;
    }); return anyChanged ? next : previous; });
    setEditing(previous => {
      if (!previous) return previous;
      const developer = developerFullNameText(previous.developer || "", usablePeople);
      return developer && developer !== previous.developer ? { ...previous, developer } : previous;
    });
  }, [personnelNameSignature, developerNormalizationSignature, draftDeveloperSignature]);

  useEffect(() => {
    if (!editing) {
      editingInitialRef.current = "";
      editingInitialIdRef.current = "";
      return;
    }
    const editingKey = `${editing.id}|${editing._intakeDraftId || ""}`;
    if (editingInitialIdRef.current !== editingKey) {
      editingInitialIdRef.current = editingKey;
      editingInitialRef.current = JSON.stringify(editing);
    }
  }, [editing?.id, editing?._intakeDraftId]);

  const requestCloseEditing = () => {
    if (!editing) return;
    if (editingInitialRef.current && JSON.stringify(editing) !== editingInitialRef.current) {
      if (!confirm("資料尚未儲存，確定要直接關閉視窗嗎？")) return;
    }
    setEditing(null);
  };

  const copyEditingFeatures = async () => {
    if (!editing) return;
    const text = [1, 2, 3, 4].map(number => `${number}.${String(editing[`feature${number}`] || "").trim()}`).join("\n");
    try { await navigator.clipboard.writeText(text); }
    catch { const textarea = document.createElement("textarea"); textarea.value = text; textarea.style.position = "fixed"; textarea.style.opacity = "0"; document.body.appendChild(textarea); textarea.select(); document.execCommand("copy"); textarea.remove(); }
    flash("已複製特色說明1～4");
  };

  useEffect(() => {
    if (tab !== "active") {
      document.body.classList.remove("active-list-stuck");
      return;
    }
    let toolbarAnchor = 0;
    const updateStickyListOffsets = () => {
      const topbar = document.querySelector<HTMLElement>(".topbar");
      const nav = document.querySelector<HTMLElement>(".topbar > .nav, body > .nav");
      const toolbar = document.querySelector<HTMLElement>(".active-list-toolbar");
      const content = document.querySelector<HTMLElement>(".active-list-page");
      if (!toolbar) return;
      const headerBottom = Math.max(topbar?.getBoundingClientRect().bottom || 0, nav?.getBoundingClientRect().bottom || 0);
      const toolbarHeight = toolbar.getBoundingClientRect().height;
      if (!document.body.classList.contains("active-list-stuck")) toolbarAnchor = toolbar.getBoundingClientRect().top + window.scrollY;
      const alreadyStuck = document.body.classList.contains("active-list-stuck");
      const stickPoint = toolbarAnchor - headerBottom;
      const shouldStick = alreadyStuck ? window.scrollY > Math.max(0, stickPoint - 40) : window.scrollY >= stickPoint;
      document.documentElement.style.setProperty("--active-toolbar-top", `${Math.round(headerBottom)}px`);
      document.documentElement.style.setProperty("--active-table-top", `${Math.round(headerBottom + toolbarHeight)}px`);
      if (content) {
        const bounds = content.getBoundingClientRect();
        document.documentElement.style.setProperty("--active-list-left", `${Math.round(bounds.left)}px`);
        document.documentElement.style.setProperty("--active-list-width", `${Math.round(bounds.width)}px`);
      }
      document.body.classList.toggle("active-list-stuck", shouldStick);
      if (shouldStick) {
        const actualToolbarBottom = toolbar.getBoundingClientRect().bottom;
        document.documentElement.style.setProperty("--active-table-top", `${Math.ceil(actualToolbarBottom)}px`);
      }
    };
    const frame = requestAnimationFrame(updateStickyListOffsets);
    const observer = new ResizeObserver(updateStickyListOffsets);
    [document.querySelector(".topbar"), document.querySelector(".topbar > .nav, body > .nav"), document.querySelector(".active-list-toolbar")].forEach(element => element && observer.observe(element));
    window.addEventListener("resize", updateStickyListOffsets);
    window.addEventListener("scroll", updateStickyListOffsets, { passive: true });
    return () => { cancelAnimationFrame(frame); observer.disconnect(); window.removeEventListener("resize", updateStickyListOffsets); window.removeEventListener("scroll", updateStickyListOffsets); document.body.classList.remove("active-list-stuck"); };
  }, [tab]);

  useEffect(() => {
    const closeTopModalWithEscape = (event: KeyboardEvent) => {
      if (event.key !== "Escape") return;
      const pptPreview = document.querySelector<HTMLElement>(".ppt-slide-preview-backdrop");
      if (pptPreview?.getClientRects().length) {
        event.preventDefault();
        event.stopImmediatePropagation();
        pptPreview.querySelector<HTMLButtonElement>(".ppt-slide-preview-close")?.click();
        return;
      }
      const backdrops = [...document.querySelectorAll<HTMLElement>(".modal-backdrop, .qr-scanner-backdrop")].filter(element => {
        const style = window.getComputedStyle(element);
        return style.display !== "none" && style.visibility !== "hidden" && element.getClientRects().length > 0;
      });
      const topBackdrop = backdrops[backdrops.length - 1];
      if (!topBackdrop) return;
      event.preventDefault();
      event.stopPropagation();
      const closeButton = topBackdrop.querySelector<HTMLButtonElement>("button.close, .modal-head button[aria-label='關閉']");
      if (closeButton) closeButton.click();
      else topBackdrop.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    };
    window.addEventListener("keydown", closeTopModalWithEscape, true);
    return () => window.removeEventListener("keydown", closeTopModalWithEscape, true);
  }, []);

  useEffect(() => {
    if (!editing || editing._generalWebsiteDefaultsApplied === "1") return;
    const contract = contractFromNo(editing.propertyNo) || editing.contractType || "";
    if (!/一般/.test(contract)) return;
    const defaultNoneKeys = ["platform591", "price5168", "led", "goldExposure", "windowAd"];
    const defaults = Object.fromEntries(defaultNoneKeys.map(key => [`${key}None`, String(editing[key] || "").trim() ? editing[`${key}None`] || "" : "1"]));
    const nextEditing = { ...editing, ...defaults, _generalWebsiteDefaultsApplied: "1" };
    editingInitialRef.current = JSON.stringify(nextEditing);
    setEditing(nextEditing);
  }, [editing?.id, editing?.propertyNo, editing?._generalWebsiteDefaultsApplied]);

  useEffect(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      let loadedRecords: RecordItem[] = saved ? JSON.parse(saved) : [sample];
      const officialRecords = ((window as any).__PROPERTY_OFFICIAL_RECORDS__ || []) as RecordItem[]; const appRestoreMarker = "property-desk-app-restore-216-v3"; if (officialRecords.length && localStorage.getItem(appRestoreMarker) !== "1") { const keyOf = (record: RecordItem) => String(record.propertyNo || record.id || "").trim(); const merged = new Map(officialRecords.map(record => [keyOf(record), record])); loadedRecords.forEach(record => { const key = keyOf(record); if (!key) return; const base = merged.get(key) || {} as RecordItem; const next = { ...base, ...record }; if (!String(next.bookLocationDate || "").trim() && String(base.bookLocationDate || "").trim()) next.bookLocationDate = base.bookLocationDate; if (!String(next.bookLocationType || "").trim() && String(base.bookLocationType || "").trim()) next.bookLocationType = base.bookLocationType; merged.set(key, next); }); loadedRecords = [...merged.values()]; localStorage.setItem(STORAGE_KEY, JSON.stringify(loadedRecords)); localStorage.setItem(appRestoreMarker, "1"); }
      loadedRecords = applySourceLayoutFixes(loadedRecords).map(record => normalizeRecordPings({ ...record, salesBook: record.salesBook || "製作" }));
      const savedSettings = localStorage.getItem(SETTINGS_KEY); setSettings(s => { const old = savedSettings ? JSON.parse(savedSettings) : {}; const personnel = old.personnel || (old.staffName || old.staffId ? [{ id: newId(), name: old.staffName || "", nationalId: old.staffId || "", status: "在職" }] : []); return { ...s, ...old, supabaseUrl: old.supabaseUrl || CASE_FILE_SUPABASE_URL, supabaseKey: old.supabaseKey || CASE_FILE_SUPABASE_PUBLISHABLE_KEY, supabaseTable: old.supabaseTable === "property_app_state" || !old.supabaseTable ? CASE_FILE_SUPABASE_TABLE : old.supabaseTable, supabaseRecord: old.supabaseRecord || "main", personnel: mergeSuppliedPersonnel(personnel) }; });
      const savedCloudSession = localStorage.getItem(CLOUD_SESSION_KEY); if (savedCloudSession) setCloudSession(JSON.parse(savedCloudSession));
      const savedIntake = localStorage.getItem(INTAKE_KEY); if (savedIntake) { const saved = JSON.parse(savedIntake); const savedDrafts: IntakeData[] = saved.drafts || (saved.parsed ? [{ ...saved.parsed, raw: saved.raw || "" }] : []); const drafts = reconcileIntakeDraftLinks(savedDrafts, loadedRecords); if (!localStorage.getItem(PHOTO_INTAKE_CLEANUP_KEY)) { const legacyPhotoValues = new Map(drafts.filter(draft => draft.linkedRecordId).map(draft => [draft.linkedRecordId!, new Set(Object.entries(draft.values).filter(([key, value]) => value && (key.includes("進案文件") || key.includes("當下進案文件"))).map(([, value]) => value.trim()))])); loadedRecords = loadedRecords.map(record => { const values = legacyPhotoValues.get(record.id); const current = String(record.photoInfo || "").split(/[／/]/).map(value => value.trim()).filter(Boolean); return values && current.length && current.every(value => values.has(value)) ? { ...record, photoInfo: "" } : record; }); localStorage.setItem(PHOTO_INTAKE_CLEANUP_KEY, "1"); } const linkedDrafts = new Map(drafts.filter(draft => draft.linkedRecordId).map(draft => [draft.linkedRecordId!, draft])); loadedRecords = loadedRecords.map(record => { const draft = linkedDrafts.get(record.id); return draft ? intakeToRecord(draft, record) : record; }); setIntakeDrafts(drafts); setSelectedIntakeId(saved.selectedId || drafts.find(draft => !draft.linkedRecordId)?.id || ""); setIntakeRaw(saved.raw || ""); }
      const savedTour = localStorage.getItem(TOUR_KEY); if (savedTour) { const tour = JSON.parse(savedTour); setTourItems(Array.isArray(tour.items) ? tour.items : []); setTourDate(tour.date || today()); setTourTitle(tour.title || `${displayRocDate(tour.date || today()).replace(/\//g, ".")}團看`); setTourModifiedAt(tour.modifiedAt || new Date().toISOString()); setTourHistory(Array.isArray(tour.history) ? tour.history : []); }
      setRecords(loadedRecords);
      setStorageReady(true);
    } catch { setRecords([sample]); }
  }, []);
  useEffect(() => { if (internalView) setTab("public"); }, [internalView]);
  useEffect(() => {
    if (!internalView) return;
    try {
      const saved = JSON.parse(localStorage.getItem("case-file-public-daily-login") || "{}");
      if (saved.date === today() && saved.personId) {
        setPublicPersonId(saved.personId);
        setPublicScope("mine");
        setPublicUnlocked(true);
      } else {
        localStorage.removeItem("case-file-public-daily-login");
      }
    } catch { localStorage.removeItem("case-file-public-daily-login"); }
  }, [internalView]);
  useEffect(() => {
    if (!internalView || !publicUnlocked) return;
    const timer = window.setInterval(() => {
      try {
        const saved = JSON.parse(localStorage.getItem("case-file-public-daily-login") || "{}");
        if (saved.date !== today()) {
          localStorage.removeItem("case-file-public-daily-login");
          setPublicUnlocked(false);
          setPublicPersonId("");
        }
      } catch {
        setPublicUnlocked(false);
        setPublicPersonId("");
      }
    }, 60000);
    return () => window.clearInterval(timer);
  }, [internalView, publicUnlocked]);
  useEffect(() => {
    if (internalView) document.title = "連城內部總表 請勿外流";
  }, [internalView]);
  useEffect(() => {
    if (!records.length) return;
    const cleanupKey = "property-desk-cleanup-website-only-2026-08-07-v1";
    if (localStorage.getItem(cleanupKey) === "1") return;
    const websiteOnlyCases = ["東區德昌路優質建地", "興達港全新大地坪大面寬電梯雙車豪墅", "高雄九如商圈高投報整新四套房1樓公寓"];
    setRecords(previous => previous.map(record => {
      if (!websiteOnlyCases.some(name => String(record.caseName || "").includes(name))) return record;
      const history = recordUpdateHistory(record);
      if (!history["2026-08-07"]) return record;
      delete history["2026-08-07"];
      return { ...record, _updateHistory: JSON.stringify(history) };
    }));
    localStorage.setItem(cleanupKey, "1");
  }, [records.length]);
  useEffect(() => { if (records.length) localStorage.setItem(STORAGE_KEY, JSON.stringify(records)); }, [records]);
  // 正式進案日與業務交件日分開：進案統計沿用 reportDate，
  // 每日動態則顯示在助理實際按下正式進案的日期。
  useEffect(() => {
    if (!storageReady || !intakeDrafts.length) return;
    const enteredAtByRecord = new Map(intakeDrafts.filter(draft => draft.linkedRecordId && draft.enteredAt).map(draft => [draft.linkedRecordId!, draft.enteredAt!]));
    setRecords(previous => {
      let changed = false;
      const next = previous.map(record => {
        if (record._dailyAddedAt) return record;
        const enteredAt = enteredAtByRecord.get(record.id);
        if (!enteredAt) return record;
        changed = true;
        return { ...record, _dailyAddedAt: enteredAt };
      });
      return changed ? next : previous;
    });
  }, [storageReady, intakeDrafts]);
  useEffect(() => {
    if (!storageReady) return;
    setRecords(previous => previous.map(record => String(record.address || "").includes("富農街一段188巷40號")
      ? { ...record, _newCaseReminderEnabled: "1", _newCaseReminderSource: "intake", housingListingCompleted: "1", newBookCompleted: "1", wangReviewCompleted: "1" }
      : record));
  }, [storageReady]);
  useEffect(() => {
    const syncAcrossTabs = (event: StorageEvent) => {
      if (event.key === STORAGE_KEY && event.newValue) {
        try { setRecords(applySourceLayoutFixes(JSON.parse(event.newValue)).map(normalizeRecordPings)); } catch {}
      }
      if (event.key === SETTINGS_KEY && event.newValue) {
        try { const next = JSON.parse(event.newValue); setSettings(previous => ({ ...previous, ...next, personnel: mergeSuppliedPersonnel(next.personnel || previous.personnel) })); } catch {}
      }
    };
    window.addEventListener("storage", syncAcrossTabs);
    return () => window.removeEventListener("storage", syncAcrossTabs);
  }, []);
  // 先完成本機設定讀取，才允許寫回；避免匯入後被初始空白人員覆蓋。
  useEffect(() => { if (storageReady) localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings)); }, [settings, storageReady]);
  // 等草稿讀取完成後才寫回，避免剛開新版本時用初始空白內容覆寫舊草稿。
  useEffect(() => { if (storageReady) localStorage.setItem(INTAKE_KEY, JSON.stringify({ raw: intakeRaw, drafts: intakeDrafts, selectedId: selectedIntakeId })); }, [intakeRaw, intakeDrafts, selectedIntakeId, storageReady]);
  useEffect(() => { localStorage.setItem(TOUR_KEY, JSON.stringify({ date: tourDate, title: tourTitle, items: tourItems, modifiedAt: tourModifiedAt, history: tourHistory })); }, [tourDate, tourTitle, tourItems, tourModifiedAt, tourHistory]);
  // V181 起保留完成團看歷史；先把舊資料中已有團看日期、但當時尚未建立歷史的案件補成可查看列表。
  useEffect(() => {
    if (!storageReady) return;
    const byDate = new Map<string, TourItem[]>();
    records.filter(record => String(record.groupViewDate || "").trim()).forEach(record => {
      const date = normalizeDateInput(record.groupViewDate || ""); if (!date) return;
      const list = byDate.get(date) || []; list.push({ id: `history-${date}-${record.id}`, recordId: record.id, sequence: String(list.length + 1), data: { ...record } }); byDate.set(date, list);
    });
    if (!byDate.size) return;
    setTourHistory(previous => {
      const known = new Set(previous.map(item => item.date));
      const recovered = [...byDate.entries()].filter(([date]) => !known.has(date)).map(([date, items]) => ({ id: `recovered-${date}`, date, title: `${displayRocDate(date).replace(/\//g, ".")}團看`, items, completedAt: date }));
      return recovered.length ? [...recovered, ...previous].sort((a, b) => String(b.date).localeCompare(String(a.date))) : previous;
    });
  }, [storageReady, records]);
  // 115/8/12 已完成團看的五筆，補回當時路線快照。前兩筆雖已不在現有委託中，歷史仍須可重新列印。
  useEffect(() => {
    if (!storageReady) return;
    const date = "2026-08-12";
    const seeds = [
      { propertyNo: "", caseName: "【永康區】中正二街 地坪近百坪甲工廠房", address: "台南市永康區中正二街172號", developer: "陳帝元", price: "3188" },
      { propertyNo: "", caseName: "永康四維商圈｜低總價樓中樓四房", address: "台南市永康區四維街37號五樓之2", developer: "謝馨儀、張小曼", price: "588" },
      { propertyNo: "EG0507746", caseName: "東區復興學區整新四套房車墅", address: "台南市東區富農街一段188巷40號", developer: "陳珮菁", price: "1788" },
      { propertyNo: "EG0540823", caseName: "東區仁和新世界整新明亮通風三房", address: "台南市東區仁和路79號6樓之3", developer: "柯育婷、蔡宇育", price: "786" },
      { propertyNo: "EG0522899", caseName: "東門圓環一樓套房＋平車", address: "台南市東區東門路一段84號之13", developer: "張小曼", price: "298" },
    ];
    setTourHistory(previous => {
      const existing = previous.find(item => normalizeDateInput(item.date) === date);
      const items = seeds.map((seed, index) => {
        const current = records.find(record => (seed.propertyNo && String(record.propertyNo || "") === seed.propertyNo) || String(record.address || "") === seed.address || String(record.caseName || "") === seed.caseName);
        const prior = existing?.items.find(item => String(item.data.address || "") === seed.address || String(item.data.caseName || "") === seed.caseName);
        const activeSaved = tourItems.find(item => String(item.data.address || "") === seed.address || String(item.data.caseName || "") === seed.caseName);
        const intakeDraft = intakeDrafts.find(draft => String(intakeValue(draft.values, "物件(完整)地址") || "") === seed.address || String(intakeValue(draft.values, "案名") || "") === seed.caseName);
        const draftData = intakeDraft ? intakeToRecord(intakeDraft) : undefined;
        // 進案草稿是歷史團看的完整資料來源，必須覆蓋先前已存下來的空白歷史版本。
        const data = draftData || activeSaved?.data || current || prior?.data || ({ id: `history-1150812-${index + 1}`, status: "委託中", ...seed } as RecordItem);
        return { id: `history-${date}-${index + 1}`, recordId: current?.id || prior?.recordId || activeSaved?.recordId || intakeDraft?.linkedRecordId, sequence: String(index + 1), data: { ...seed, ...data, _tourHistorySnapshot: "1" } } as TourItem;
      });
      const entry: TourHistory = { id: existing?.id || "recovered-2026-08-12", date, title: "115.08.12團看", items, completedAt: existing?.completedAt || date };
      const unchanged = existing && existing.items.length === 5 && existing.items.every((item, index) => item.data.caseName === seeds[index]?.caseName && Object.keys(item.data || {}).length > 45);
      return unchanged ? previous : [entry, ...previous.filter(item => normalizeDateInput(item.date) !== date)].sort((a, b) => String(b.date).localeCompare(String(a.date)));
    });
  }, [storageReady, records, intakeDrafts, tourItems]);

  const archived = useMemo(() => sortArchivedRecords(records.filter(r => r.archived || isExpired(r) || r.status !== "委託中")), [records]);
  const active = useMemo(() => records.filter(r => !r.archived && !isExpired(r) && (r.status || "委託中") === "委託中"), [records]);
  const bookReviewCycleStart = normalizeDateInput(settings.bookReviewCurrentDate || "") || "2026-07-30";
  const bookReviewDueCount = today() >= bookReviewCycleStart ? active.filter(record => {
    const confirmedDate = normalizeDateInput(record._bookReviewAt || record.bookLocationDate || "");
    return !confirmedDate || confirmedDate < bookReviewCycleStart;
  }).length : 0;
  const latestRecordModifiedAt = records.reduce((latest, record) => {
    const candidate = String(record.lastModifiedAt || record.caseNameNoteModifiedAt || record.reducedPriceModifiedAt || "");
    if (!candidate || Number.isNaN(new Date(candidate).getTime())) return latest;
    return !latest || new Date(candidate).getTime() > new Date(latest).getTime() ? candidate : latest;
  }, "");
  const latestDraftModifiedAt = intakeDrafts.reduce((latest, draft) => {
    const candidate = String(draft.modifiedAt || draft.createdAt || "");
    if (!candidate || Number.isNaN(new Date(candidate).getTime())) return latest;
    return !latest || new Date(candidate).getTime() > new Date(latest).getTime() ? candidate : latest;
  }, "");
  const latestModifiedAt = [latestRecordModifiedAt, latestDraftModifiedAt, tourModifiedAt].filter(Boolean).reduce((latest, candidate) => !latest || new Date(candidate).getTime() > new Date(latest).getTime() ? candidate : latest, "");
  const cloudUploadStatus = !cloudSession?.accessToken ? "未登入" : cloudLastUploadAt && (!latestModifiedAt || Date.parse(cloudLastUploadAt) >= Date.parse(latestModifiedAt)) ? "上傳完成" : "ING";
  const showingFollowUpRecords = active.filter(record => record.showingFollowUp === "暫停帶看／等待業務回覆");
  useEffect(() => {
    const dueToday = showingFollowUpRecords.filter(record => normalizeDateInput(record.showingFollowUpDueDate || "") === today());
    const reminderKey = `property-desk-showing-follow-up-${today()}`;
    if (!dueToday.length || localStorage.getItem(reminderKey) === "1") return;
    localStorage.setItem(reminderKey, "1");
    setTimeout(() => alert(`今日有 ${dueToday.length} 筆帶看追蹤到期：\n${dueToday.map(record => record.caseName || record.propertyNo).join("\n")}`), 100);
  }, [showingFollowUpRecords.map(record => `${record.id}:${record.showingFollowUpDueDate}`).join("|")]);
  const controlledKeyCount = new Set(records.map(record => String(record.key || "").match(/公司\s*[#＃]?\s*(\d+)/)?.[1]).filter(Boolean)).size;
  const pendingArchiveCleanup = archived.map(record => ({ record, tasks: archiveCleanupTasks(record).filter(task => !record[task.key]) })).filter(item => item.record.archived && item.tasks.length > 0);
  const toggleArchiveCleanupSort = (key: string) => setArchiveCleanupSort(previous => previous.key === key ? { key, direction: previous.direction === "asc" ? "desc" : "asc" } : { key, direction: "asc" });
  const sortedPendingArchiveCleanup = [...pendingArchiveCleanup].sort((a, b) => {
    if (!archiveCleanupSort.key) return 0;
    const key = archiveCleanupSort.key;
    const aValue = key === "case" ? String(a.record.caseName || a.record.propertyNo || "") : a.record[key] ? "1" : "0";
    const bValue = key === "case" ? String(b.record.caseName || b.record.propertyNo || "") : b.record[key] ? "1" : "0";
    const compared = aValue.localeCompare(bValue, "zh-Hant");
    return archiveCleanupSort.direction === "asc" ? compared : -compared;
  });
  const archiveCleanupHeader = (key: string, label: string) => <th className="archive-cleanup-sort-header" onClick={() => toggleArchiveCleanupSort(key)}>{label}<b>{archiveCleanupSort.key === key ? archiveCleanupSort.direction === "asc" ? " ↑" : " ↓" : " ↕"}</b></th>;
  useEffect(() => { if (suppressArchiveCleanupReminderRef.current) { suppressArchiveCleanupReminderRef.current = false; setArchiveCleanupReminderOpen(false); return; } if (pendingArchiveCleanup.length) setArchiveCleanupReminderOpen(true); else setArchiveCleanupReminderOpen(false); }, [pendingArchiveCleanup.map(item => `${item.record.id}:${item.tasks.map(task => task.key).join(",")}`).join("|")]);
  const pendingDealCompletion = archived.map(record => ({ record, tasks: dealCompletionTasks.filter(([key]) => !record[key]) })).filter(item => item.record.archived && item.record.status === "成交下架" && item.tasks.length > 0);
  useEffect(() => { if (pendingDealCompletion.length) setDealCompletionReminderOpen(true); else setDealCompletionReminderOpen(false); }, [pendingDealCompletion.map(item => `${item.record.id}:${item.tasks.map(([key]) => key).join(",")}`).join("|")]);
  const missingDataOf = (record: RecordItem) => [!record.bookLocationDate ? "物件本日期" : "", (!record.salesBookDate || !record.salesBook) ? "銷售本" : "", (!record.photoInfo && !record.photos?.length) ? "照片" : ""].filter(Boolean);
  const missingDataRecords = active.filter(record => missingDataOf(record).length > 0);
  useEffect(() => { if (!missingDataRecords.length) return; const date = today(); if (localStorage.getItem(MISSING_REMINDER_DATE_KEY) !== date) { localStorage.setItem(MISSING_REMINDER_DATE_KEY, date); setMissingDataReminderOpen(true); } }, [missingDataRecords.map(record => `${record.id}:${missingDataOf(record).join(",")}`).join("|")]);
  const websiteFieldMap: Record<string, string> = { platform591: "591", price5168: "5168", led: "LED", goldExposure: "黃金曝光", windowAd: "櫥窗（專）", homeWeb: "我家網", houseinfor: "HOUSE INFOR", yes319: "YES319" };
  const isExclusiveRecord = (record: RecordItem) => /專/.test(contractFromNo(record.propertyNo) || record.contractType || "");
  const exclusiveWebsiteKeys = ["platform591", "price5168", "led", "goldExposure", "windowAd"];
  const generalWebsiteKeys = ["homeWeb", "houseinfor", "yes319"];
  const allWebsiteKeys = [...exclusiveWebsiteKeys, ...generalWebsiteKeys];
  const requiredWebsiteKeys = (record: RecordItem) => isExclusiveRecord(record) ? allWebsiteKeys : generalWebsiteKeys;
  const isWebsiteMissing = (record: RecordItem, key: string) => record.bookLocationType !== "旁5" && record[`${key}None`] !== "1" && (isExclusiveRecord(record) ? allWebsiteKeys.includes(key) : generalWebsiteKeys.includes(key)) && !String(record[key] || "").trim();
  const daysSinceReport = (record: RecordItem) => { const date = normalizeDateInput(record.reportDate || ""); if (!date) return 0; return Math.floor((new Date(`${today()}T00:00:00`).getTime() - new Date(`${date}T00:00:00`).getTime()) / 86400000); };
  const overduePoRecords = active.map(record => ({ record, missing: requiredWebsiteKeys(record).filter(key => isWebsiteMissing(record, key)) })).filter(item => daysSinceReport(item.record) > 14 && item.missing.length > 0);
  useEffect(() => { if (overduePoRecords.length) setWebsitePoReminderOpen(true); else setWebsitePoReminderOpen(false); }, [overduePoRecords.map(item => `${item.record.id}:${item.missing.join(",")}`).join("|")]);
  const websiteExpiryRecords = active.flatMap(record => record.bookLocationType === "旁5" ? [] : [
    ["591", "platform591", record.platform591Expiry], ["5168", "price5168", record.price5168Expiry], ["黃金曝光", "goldExposure", record.goldExposureExpiry],
  ].filter((entry): entry is [string, string, string] => record[`${entry[1]}None`] !== "1" && Boolean(entry[2]) && normalizeDateInput(entry[2]) < today()).map(([site, , date]) => ({ record, site, date })));
  const systemExpiryReminders = [
    { label: "591", date: normalizeDateInput(settings.expiry591 || ""), leadDays: 1 },
    { label: "5168", date: normalizeDateInput(settings.expiry5168 || ""), leadDays: 1 },
    { label: "經紀人", date: normalizeDateInput(settings.brokerExpiry || ""), leadDays: 30 },
  ].filter(item => validDate(item.date) && item.date <= addDaysIso(today(), item.leadDays));
  const expiredUnarchived = useMemo(() => records.filter(r => !r.archived && (r.status || "委託中") === "委託中" && isExpired(r)), [records]);
  useEffect(() => { if (expiredUnarchived.length) setExpiryReminderOpen(true); else setExpiryReminderOpen(false); }, [expiredUnarchived.map(record => record.id).join("|")]);
  const matchesAdvancedFilter = (r: RecordItem) => {
    const f = advancedFilter;
    const includes = (key: string, value = r[key] || "") => !f[key] || String(value).toLowerCase().includes(f[key].toLowerCase());
    const layout = String(r.layout || "");
    const floor = String(r.currentFloor || r.floor || "").replace(/[～〜]/g, "~");
    const floorMatch = !f.floor || new RegExp(`(^|\\D)${f.floor.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}(?=\\D|$)`).test(floor);
    const parkingMatch = !f.parking || String(r.parking || [r.parkingType, r.parkingMethod, r.parkingOwnership].join("／")).includes(f.parking);
    const websiteMatch = Object.keys(websiteFieldMap).every(key => !f[key] || (f[key] === "filled" ? !!String(r[key] || "").trim() && !/^(無|旁5)$/i.test(String(r[key] || "").trim()) : !String(r[key] || "").trim() || /^(無|旁5)$/i.test(String(r[key] || "").trim())));
    const photoMatch = !f.photoInfo || (f.photoInfo === "filled" ? !!(String(r.photoInfo || "").trim() || r.photos?.length) : !(String(r.photoInfo || "").trim() || r.photos?.length));
    return includes("propertyNo") && includes("caseName") && includes("address") && (!f.type || typeShort(r.type) === f.type) && includes("direction") && includes("zoning") && includes("developer") && includes("key") && includes("currentState") && includes("salesBook") && (!f.rooms || new RegExp(`${f.rooms}\\s*房`).test(layout)) && (!f.halls || new RegExp(`${f.halls}\\s*廳`).test(layout)) && (!f.baths || new RegExp(`${f.baths}\\s*衛`).test(layout)) && floorMatch && parkingMatch && photoMatch && websiteMatch && filterDateInRange(r.entrustStart, f.entrustStartFrom, f.entrustStartTo) && filterDateInRange(r.entrustEnd, f.entrustEndFrom, f.entrustEndTo) && filterDateInRange(r.reportDate, f.reportFrom, f.reportTo) && filterDateInRange(r.updateDate, f.updateFrom, f.updateTo) && filterDateInRange(r.groupViewDate, f.groupFrom, f.groupTo) && filterDateInRange(r.bookLocationDate, f.bookFrom, f.bookTo) && filterDateInRange(r.salesBookDate, f.salesFrom, f.salesTo) && filterInRange(r.price, f.priceFrom, f.priceTo, true) && filterInRange(ageOf(r), f.ageFrom, f.ageTo) && filterInRange(r.indoorPing, f.indoorFrom, f.indoorTo) && filterInRange(r.landPing, f.landFrom, f.landTo) && filterInRange(r.buildingPing, f.buildingFrom, f.buildingTo) && filterInRange(r.road, f.roadFrom, f.roadTo) && filterInRange(r.frontage, f.frontageFrom, f.frontageTo) && filterInRange(r.depth, f.depthFrom, f.depthTo);
  };
  const currentListQuery = tab === "archive" ? archiveQuery : query;
  const advancedSourceRecords = advancedFilter.sourceArchive === "1"
    ? advancedFilter.sourceActive === "1" ? [...sortActiveRecords(active), ...archived] : archived
    : sortActiveRecords(active);
  const shown = (tab === "archive" ? archived : advancedSourceRecords).filter(r => (!showingFollowUpOnly || tab !== "active" || r.showingFollowUp === "暫停帶看／等待業務回覆") && (!websiteFilter || tab === "archive" || (websiteFilter === "missing_sales_book_date" ? !r.salesBookDate : websiteFilter === "missing_book_location_date" ? !r.bookLocationDate : websiteFilter === "missing_photo_info" ? (!r.photoInfo && !r.photos?.length) : websiteFilter === "all" ? requiredWebsiteKeys(r).some(key => isWebsiteMissing(r, key)) : isWebsiteMissing(r, websiteFilter))) && (tab !== "active" || matchesAdvancedFilter(r)) && [r.propertyNo, r.area, r.caseName, r.address, r.developer].join(" ").toLowerCase().includes(currentListQuery.toLowerCase()));
  const defaultSelectedPptWeek = pptWeekOf(pptWeekStart);
  const selectedPptWeek = {
    start: normalizeDateInput(pptCustomStart) || defaultSelectedPptWeek.start,
    end: normalizeDateInput(pptCustomEnd) || defaultSelectedPptWeek.end,
    meeting: normalizeDateInput(pptCustomMeeting) || defaultSelectedPptWeek.meeting,
  };
  const resetPptCustomWeek = () => { setPptCustomStart(""); setPptCustomEnd(""); setPptCustomMeeting(""); };
  const pptCurrentStart = currentPptWeek().start;
  const standardPptWeeks = Array.from({ length: 24 }, (_, index) => addDaysIso(pptCurrentStart, index * -7));
  const pptWeekOptions = [...new Set([...standardPptWeeks, ...records.flatMap(record => [normalizeDateInput(record.reportDate || ""), ...pptStoredList(record._pptExtraWeeks)]).filter(date => /^\d{4}-\d{2}-\d{2}$/.test(date)).map(date => pptWeekOf(date).start).filter(start => start <= pptCurrentStart)])].sort((a, b) => b.localeCompare(a));
  const weeklyPptRecords = sortPptRecords(records.filter(record => belongsToPptWeek(record, selectedPptWeek) && !excludedFromPptWeek(record, selectedPptWeek.start)));
  const deferredPptRecords = sortPptRecords(records.filter(record => belongsToPptWeek(record, selectedPptWeek) && excludedFromPptWeek(record, selectedPptWeek.start)));
  const pptDraftRecords = intakeDrafts.filter(draft => !draft.linkedRecordId).map(draft => ({ ...intakeToRecord(draft), id: `ppt-draft-${draft.id}`, reportDate: "", status: "尚未進案", _intakeDraftId: draft.id, _notEntered: "1" }));
  const pptExtraCandidates = [...records.filter(record => !belongsToPptWeek(record, selectedPptWeek)), ...pptDraftRecords].filter(record => [record.caseName, record.address].join(" ").toLowerCase().includes(pptExtraSearch.trim().toLowerCase()));
  const selectedPptBaseRecords = (() => {
    const standard = sortPptRecords([...records.filter(record => weeklyPptRecords.some(item => item.id === record.id) || pptExtraIds.includes(record.id)), ...pptDraftRecords.filter(record => pptExtraIds.includes(record.id)), ...pptAdHocRecords]);
    const position = new Map(pptOrderIds.map((id, index) => [id, index]));
    return standard.slice().sort((a, b) => (position.get(a.id) ?? Number.MAX_SAFE_INTEGER) - (position.get(b.id) ?? Number.MAX_SAFE_INTEGER));
  })();
  const selectedPptRecords = selectedPptBaseRecords.map(record => pptConfirmedSnapshots[record.id] || record);
  const pendingPptConfirmCount = selectedPptBaseRecords.filter(record => !pptConfirmedSnapshots[record.id]).length;

  const flash = (text: string) => { setNotice(text); setTimeout(() => setNotice(""), 2600); };
  const movePptOrder = (id: string, direction: -1 | 1) => {
    const ids = selectedPptBaseRecords.map(record => record.id);
    const index = ids.indexOf(id), target = index + direction;
    if (index < 0 || target < 0 || target >= ids.length) return;
    [ids[index], ids[target]] = [ids[target], ids[index]];
    setPptOrderIds(ids);
  };
  const setPptSequence = (id: string, value: string) => {
    const ids = selectedPptBaseRecords.map(record => record.id);
    const from = ids.indexOf(id);
    const requested = Math.round(Number(value));
    if (from < 0 || !Number.isFinite(requested)) return;
    ids.splice(from, 1);
    ids.splice(Math.max(0, Math.min(ids.length, requested - 1)), 0, id);
    setPptOrderIds(ids);
  };
  useEffect(() => {
    if (!pptPickerOpen) return;
    const body = document.querySelector<HTMLElement>(".ppt-picker-body");
    const list = body?.querySelector<HTMLElement>(".ppt-case-list");
    if (!body || !list) return;
    const rows = Array.from(list.querySelectorAll<HTMLElement>(".ppt-case-row.auto"));
    const rowById = new Map<string, HTMLElement>();
    rows.forEach((row, index) => {
      const record = weeklyPptRecords.find(item => item.id === row.dataset.pptRecordId) || weeklyPptRecords[index];
      if (!record) return;
      row.dataset.pptRecordId = record.id;
      rowById.set(record.id, row);
      const order = Math.max(0, selectedPptBaseRecords.findIndex(item => item.id === record.id)) + 1;
      const marker = row.querySelector<HTMLElement>("strong");
      if (marker) {
        marker.textContent = "";
        const input = document.createElement("input");
        input.className = "ppt-sequence-input";
        input.type = "number";
        input.min = "1";
        input.max = String(selectedPptBaseRecords.length);
        input.value = String(order);
        input.title = "修改順序";
        input.addEventListener("change", () => setPptSequence(record.id, input.value));
        marker.append(input);
      }
      const actions = row.querySelector<HTMLElement>(".ppt-inline-confirm");
      actions?.querySelectorAll(".ppt-preview-button").forEach(button => button.remove());
      if (actions) {
        const preview = document.createElement("button");
        preview.type = "button";
        preview.className = "ppt-preview-button";
        preview.textContent = "圖";
        preview.title = "查看 PPT 預覽";
        preview.addEventListener("click", () => setPptPreviewRecord(pptConfirmedSnapshots[record.id] || record));
        actions.append(preview);
      }
    });
    selectedPptBaseRecords.forEach(record => {
      const row = rowById.get(record.id);
      if (row) list.append(row);
    });
  }, [pptPickerOpen, pptOrderIds, pptExtraIds, pptAdHocRecords, pptConfirmedSnapshots, weeklyPptRecords.map(record => record.id).join("|"), selectedPptBaseRecords.map(record => record.id).join("|")]);
  useEffect(() => {
    if (!pptPreviewRecord) return;
    const record = pptPreviewRecord;
    const land = typeShort(record.type) === "土地" || /^(LG|LA)/i.test(record.propertyNo || "");
    const clean = (value = "") => /^(?:0|無|0；無|0;無)$/.test(String(value).trim()) ? "" : String(value || "").trim();
    const num = (value = "") => String(value || "").replace(/,/g, "").match(/-?\d+(?:\.\d+)?/)?.[0] || "";
    const price = clean(record.reducedPrice || record.price).replace(/萬/g, "");
    const landPing = num(record.landPing);
    const unitPrice = price && landPing ? `$${Math.round(Number(price) * 10000 / Number(landPing)).toLocaleString("en-US")}` : "";
    const linkedDraft = intakeDrafts.find(draft => draft.id === record._intakeDraftId || draft.linkedRecordId === record.id);
    const draftValue = (...keys: string[]) => linkedDraft ? intakeValue(linkedDraft.values, ...keys) : "";
    const previewCompletionDate = (value = "") => {
      const text = String(value || "").trim();
      const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
      if (iso) return `${String(Number(iso[1]) - 1911).padStart(3, "0")}.${String(Number(iso[2])).padStart(2, "0")}.${String(Number(iso[3])).padStart(2, "0")}`;
      const roc = text.match(/^(\d{2,3})[.\/-](\d{1,2})[.\/-](\d{1,2})$/);
      return roc ? `${roc[1].padStart(3, "0")}.${roc[2].padStart(2, "0")}.${roc[3].padStart(2, "0")}` : text;
    };
    const overlay = document.createElement("div"); overlay.className = "ppt-slide-preview-backdrop";
    const close = document.createElement("button"); close.type = "button"; close.className = "ppt-slide-preview-close"; close.textContent = "×"; close.title = "關閉預覽"; close.addEventListener("click", () => setPptPreviewRecord(null));
    const slide = document.createElement("canvas"); slide.className = "ppt-slide-preview ppt-slide-preview-canvas"; slide.width = 1000; slide.height = 750;
    const context = slide.getContext("2d"); if (!context) return;
    const px = (inch: number) => inch * 100;
    const fontPx = (pt: number) => pt * 100 / 72;
    const font = '"DFKai-SB","標楷體",serif';
    const line = (x: number, y: number, w: number) => { context.beginPath(); context.moveTo(px(x), px(y)); context.lineTo(px(x + w), px(y)); context.strokeStyle = "#555"; context.lineWidth = 1; context.stroke(); };
    const draw = (value: string, x: number, y: number, w: number, h: number, pt: number, options: { color?: string; bold?: boolean; align?: CanvasTextAlign; wrap?: boolean } = {}) => {
      const text = String(value || ""); let size = pt; const maxWidth = px(w) - 4;
      const setFont = () => { context.font = `${options.bold ? "bold " : ""}${fontPx(size)}px ${font}`; };
      setFont(); while (size > 8 && !options.wrap && context.measureText(text).width > maxWidth) { size -= .5; setFont(); }
      context.fillStyle = options.color || "#000"; context.textAlign = options.align || "center"; context.textBaseline = "middle";
      const anchorX = options.align === "left" ? px(x) + 2 : options.align === "right" ? px(x + w) - 2 : px(x + w / 2);
      if (!options.wrap) { context.fillText(text, anchorX, px(y + h / 2), maxWidth); return; }
      const rows: string[] = []; let row = "";
      Array.from(text).forEach(char => { const next = row + char; if (row && context.measureText(next).width > maxWidth) { rows.push(row); row = char; } else row = next; }); if (row) rows.push(row);
      while (size > 8 && rows.length * fontPx(size) * 1.08 > px(h)) { size -= .5; setFont(); }
      rows.slice(0, 3).forEach((item, index) => context.fillText(item, anchorX, px(y + h / 2) + (index - (rows.length - 1) / 2) * fontPx(size) * 1.08, maxWidth));
    };
    const alignedLabel = (label: string, x: number, y: number, pt = 15) => { const chars = Array.from(label); chars.forEach((char, index) => draw(char, x + (chars.length === 1 ? .34 : index * (.68 / (chars.length - 1))), y, .22, .4, pt)); };
    const field = (label: string, value: string, x: number, y: number, w: number, valuePt = 19, labelPt = 15, lineOffset = .39) => { alignedLabel(label, x, y, labelPt); draw("：", x + .88, y, .2, .36, 15); draw(value, x + 1.08, y - .04, w - 1.08, .44, valuePt); line(x + 1.08, y + lineOffset, w - 1.13); };
    context.fillStyle = "#fff"; context.fillRect(0, 0, 1000, 750);
    draw("案名", .05, .16, .72, .55, 14); draw(record.caseName || "", .78, .08, 5.23, .72, (land ? 28 : 25) - 2);
    draw("地址", 6.08, .16, .68, .55, 14); draw(record.address || "", 6.78, .08, 3.12, .72, 17, { wrap: true }); line(0, .84, 10);
    context.fillStyle = "#fce9d9"; context.fillRect(0, px(.87), px(2.756), px(land ? .94 : .787));
    draw("開發", .06, land ? 1.14 : 1.07, .58, .35, 13, { bold: true }); draw(developerFullNameText(record.developer) || "", .66, land ? 1.08 : 1, 2.09, .52, 17, { bold: true, color: "#0000ff" });
    draw(land ? "開價" : "委託總價", 2.76, land ? 1.15 : 1.08, 1.12, .38, 16, { align: "right" }); draw("：", 3.9, land ? 1.15 : 1.08, .2, .38, 16); draw(price, 4.1, .89, 1.48, .72, 35, { bold: true, color: "#e00000" }); draw("萬", 5.55, 1.24, .3, .3, 14); line(0, land ? 1.82 : 1.68, 6.05);
    if (land) {
      field("總地坪", landPing ? `${landPing}坪` : "", .08, 2.12, 2.85, 23, 18, .58); field("每坪單價", unitPrice, 3.02, 2.12, 2.82, 23, 18, .58);
      field("臨路", num(record.road) ? `${num(record.road)}米` : clean(record.road), .08, 3.19, 2.85, 23, 18, .58); field("座向", record.direction, 3.02, 3.19, 2.82, 23, 18, .58);
      field("面寬", num(record.frontage) ? `${num(record.frontage)}米` : clean(record.frontage), .08, 4.26, 2.85, 23, 18, .58); field("深度", num(record.depth) ? `${num(record.depth)}米` : clean(record.depth), 3.02, 4.26, 2.82, 23, 18, .58);
      field("建蔽容積", [record.coverage, record.far].filter(Boolean).join("／"), .08, 5.33, 2.85, 18, 12, .58); field("使用分區", record.zoning, 3.02, 5.33, 2.82, 20, 18, .58);
    } else {
      const entries: [string,string,string,string,number][] = [["總地坪",landPing ? `${landPing}坪` : "","總建坪",num(record.buildingPing) ? `${num(record.buildingPing)}坪` : "",1.8],["室內坪",num(record.indoorPing) ? `${num(record.indoorPing)}坪` : "","格局",record.layout || "",2.36],["面寬",num(record.frontage) ? `${num(record.frontage)}米` : clean(record.frontage),"深度",num(record.depth) ? `${num(record.depth)}米` : clean(record.depth),2.92],["臨路",num(record.road) ? `${num(record.road)}米` : clean(record.road),"樓層",floorPptDisplay(record.floor),3.48],["朝向",record.direction || "","車位",parkingShort(record.parking),4.04],["管理費",/^\$?0(?:\.0+)?(?:元)?$/.test(record.managementFee || "") ? "無0" : record.managementFee || "","社區名稱",draftValue("大樓名稱") || record.communityName || "",4.6],["建築完成日期",previewCompletionDate(draftValue("建築完成日期") || record.completionDate || record.builtYear || ""),"屋齡",ageOf(record),5.16],["現況",record.currentState || "","鑰匙",record.key || "",5.72]];
      entries.forEach(([leftLabel,leftValue,rightLabel,rightValue,y]) => { field(leftLabel,leftValue,.08,y,2.85); field(rightLabel,rightValue,3.02,y,2.82,rightLabel === "車位" ? 15.5 : rightLabel === "社區名稱" || rightLabel === "建築完成日期" ? 16 : 19); });
    }
    const notesY = land ? 6.4 : 6.28; alignedLabel("備註", .08, notesY); draw("：", .96, notesY, .2, .36, 15); draw(displayNoteSegments(record.notes).join("；"), 1.16, notesY, 4.61, .72, 14, { align: "left", wrap: true });
    line(0, 7.18, 6.05); draw(`進案報件日期：${displayRocDate(record.reportDate) || ""}`, .05, 7.22, 2.8, .24, 9, { align: "left" }); draw(`物件編號：${record.propertyNo || ""}`, 3.1, 7.22, 2.9, .24, 9);
    overlay.append(close, slide); overlay.addEventListener("mousedown", event => { if (event.target === overlay) setPptPreviewRecord(null); }); document.body.append(overlay);
    return () => overlay.remove();
  }, [pptPreviewRecord]);
  const removeFromPptWeek = (record: RecordItem) => setRecords(previous => previous.map(item => {
    if (item.id !== record.id) return item;
    const excluded = [...new Set([...pptStoredList(item._pptExcludedWeeks), selectedPptWeek.start])];
    const choices = { ...pptStoredChoices(item._pptWeekChoices) }; delete choices[selectedPptWeek.start];
    return { ...item, _pptExcludedWeeks: JSON.stringify(excluded), _pptWeekChoices: JSON.stringify(choices), lastModifiedAt: new Date().toISOString() };
  }));
  const decideRemovedPpt = (record: RecordItem, choice: "next" | "skip") => setRecords(previous => previous.map(item => {
    if (item.id !== record.id) return item;
    const nextWeek = nextPptWeekStart(selectedPptWeek.start);
    const extras = pptStoredList(item._pptExtraWeeks).filter((week: string) => choice === "next" || week !== nextWeek);
    if (choice === "next" && !extras.includes(nextWeek)) extras.push(nextWeek);
    const choices = { ...pptStoredChoices(item._pptWeekChoices), [selectedPptWeek.start]: choice };
    return { ...item, _pptExtraWeeks: JSON.stringify(extras), _pptWeekChoices: JSON.stringify(choices), lastModifiedAt: new Date().toISOString() };
  }));
  const addShowingFollowUp = () => {
    if (!showingFollowUpRecordId) return flash("請先選擇委託中案件");
    if (!normalizeDateInput(showingFollowUpStart)) return flash("請填寫暫停日期");
    if (!normalizeDateInput(showingFollowUpDue)) return flash("請填寫追蹤日期");
    setRecords(previous => previous.map(record => record.id === showingFollowUpRecordId ? { ...record, showingFollowUp: "暫停帶看／等待業務回覆", showingFollowUpDate: normalizeDateInput(showingFollowUpStart), showingFollowUpDueDate: normalizeDateInput(showingFollowUpDue), lastModifiedAt: new Date().toISOString() } : record));
    setShowingFollowUpRecordId(""); setShowingFollowUpStart(today()); setShowingFollowUpDue(""); flash("已加入帶看追蹤");
  };
  const clearShowingFollowUp = (record: RecordItem) => {
    setRecords(previous => previous.map(item => item.id === record.id ? { ...item, showingFollowUp: "", showingFollowUpDate: "", showingFollowUpDueDate: "", lastModifiedAt: new Date().toISOString() } : item));
    flash("已解除帶看追蹤，案件仍保留委託中");
  };
  const updateEditingRecord = (next: RecordItem) => setEditing(previous => {
    if (!previous) return next;
    const ignored = new Set(["id", "lastModifiedAt", "priceModifiedAt", "reducedPriceModifiedAt", "groupViewDate", "_updateHistory", "_dailyAnnotation", "_dailyHighlight"]);
    const changed = Object.keys(next).filter(key => !ignored.has(key) && !websiteTrackingKeys.has(key) && trackedValue(previous[key]) !== trackedValue(next[key]));
    if (!changed.length) return next;
    const edited = new Set([...editedFieldKeys(previous), ...changed]);
    return { ...next, _materialEditAt: "1", _editedFields: JSON.stringify([...edited]), lastModifiedAt: new Date().toISOString() };
  });
  const saveRecord = (publishDaily = false) => {
    if (!editing) return;
    const normalizedEditing = normalizeRecordPings(clearImportedLandLabels({
      ...editing,
      developer: developerFullNameText(editing.developer || "", settings.personnel) || editing.developer,
      ...Object.fromEntries([...dateKeys].map(key => [key, normalizeDateInput(editing[key] || "")])),
      completionDate: normalizeDateInput(editing.completionDate || ""),
    }));
    const keyNumber = String(editing.key || "").match(/公司\s*[#＃]?\s*(\d+)/)?.[1];
    const allowedKeyNumbers = new Set([1, 2, 3, 5, 6, 7, 8, 17, 18, 19, 20, 21, 22, 23, 24, 33, 34, 35, 36, 37, 38, 39, 49, 50, 51, 52, 53, 55, 56, 65, 66, 67, 68, 69, 70, 71, 72, 81, 82, 83, 85, 86, 87, 88]);
    if (keyNumber && !allowedKeyNumbers.has(Number(keyNumber))) { alert(`鑰匙編號公司#${keyNumber}不在鑰匙總表的有效標號內，請重新輸入。`); return; }
    const duplicatedKey = keyNumber ? records.find(record => record.id !== editing.id && String(record.key || "").match(/公司\s*[#＃]?\s*(\d+)/)?.[1] === keyNumber) : undefined;
    if (duplicatedKey) { alert(`鑰匙編號公司#${keyNumber}已由「${duplicatedKey.caseName || duplicatedKey.propertyNo}」使用，請確認後再儲存。`); return; }
    if (normalizedEditing._intakeDraftId) {
      setIntakeDrafts(previous => previous.map(draft => draft.id === normalizedEditing._intakeDraftId ? { ...draft, values: syncRecordToDraftValues(draft, normalizedEditing), propertyKind: normalizedEditing.type.includes("土地") ? "純土地" : "房屋" } : draft));
      setTourItems(previous => previous.map(item => item.data._intakeDraftId === normalizedEditing._intakeDraftId ? { ...item, data: { ...item.data, ...normalizedEditing, reportDate: "", status: "尚未進案", _notEntered: "1" } } : item));
      setEditing(null); flash("草稿與團看資料已同步更新"); return;
    }
    if (!normalizedEditing.propertyNo || !normalizedEditing.caseName) return flash("請填寫物件編號與案名");
    const invalid = [...dateKeys].find(k => k !== "groupViewDate" && !validDate(normalizedEditing[k])); if (invalid) return flash(`${labels[invalid]}日期格式錯誤`);
    const existing = records.find(record => record.id === normalizedEditing.id);
    const combinedFloor = [normalizedEditing.titleFloor, normalizedEditing.currentFloor].filter(Boolean).join("／");
    const combinedParking = [normalizedEditing.parkingOwnership, normalizedEditing.parkingType, normalizedEditing.parkingMethod, normalizedEditing.parkingNo].filter(Boolean).join("／");
    const preparedWithMarker = { ...normalizedEditing, school: schoolSummary(normalizedEditing), builtYear: normalizedEditing.completionDate ? String(Number(normalizedEditing.completionDate.split(/[./]/)[0])) : normalizedEditing.builtYear, floor: combinedFloor || normalizedEditing.floor, parking: combinedParking || normalizedEditing.parking, contractType: contractFromNo(normalizedEditing.propertyNo) || normalizedEditing.contractType, status: normalizedEditing.status || "委託中", updateDate: normalizedEditing.updateDate || existing?.updateDate || today(), lastModifiedAt: normalizedEditing.lastModifiedAt || existing?.lastModifiedAt || "" };
    const { _materialEditAt, _editedFields, ...prepared } = preparedWithMarker;
    // 到期下架案件只要把委託結束改為今天或之後，即視為重新上架。
    const autoReopen = Boolean(existing?.archived) && existing?.status === "到期下架" && validDate(prepared.entrustEnd) && !isExpired(prepared);
    const reopenNote = `${shortRocMonthDay(today())}重新上架`;
    const preparedForSave = autoReopen ? {
      ...prepared,
      status: "委託中",
      archived: "",
      _restoredAt: today(),
      caseNameNote: String(prepared.caseNameNote || "").includes(reopenNote) ? prepared.caseNameNote : [prepared.caseNameNote, reopenNote].filter(Boolean).join("　"),
      caseNameNoteModifiedAt: new Date().toISOString(),
    } : prepared;
    const changedFields = existing ? editedFieldKeys(preparedWithMarker).filter(key => trackedValue(existing[key]) !== trackedValue(preparedForSave[key])) : [];
    if (existing && changedFields.length) {
      const changedLabels = changedFields.map(key => labels[key] || recordEditLabels[key] || key).join("、");
      const action = publishDaily ? "儲存並發布到每日物件動態" : "僅儲存，不發布到每日物件動態";
      if (!confirm(`本次修改欄位：\n${changedLabels}\n\n確定要${action}嗎？`)) return;
    }
    let next = existing
      ? (autoReopen
        ? withTrackedUpdate(existing, preparedForSave, today())
        : _materialEditAt
          ? (publishDaily
            ? withTrackedUpdate(existing, preparedForSave, today(), changedFields)
            : { ...preparedForSave, updateDate: existing.updateDate, lastModifiedAt: new Date().toISOString(), _updateHistory: existing._updateHistory })
          : { ...preparedForSave, updateDate: existing.updateDate, lastModifiedAt: existing.lastModifiedAt, _updateHistory: existing._updateHistory })
      : { ...preparedForSave };
    // 重新上架本身在每日動態只顯示「月/日重新上架」，不另列為「更新：案名」。
    if (autoReopen) {
      const history = recordUpdateHistory(next);
      if (history[today()]) history[today()] = history[today()].filter(key => key !== "caseNameNote" && key !== "caseNameNoteModifiedAt");
      next = { ...next, _updateHistory: JSON.stringify(history) };
    }
    setRecords(prev => prev.some(r => r.id === next.id) ? prev.map(r => r.id === next.id ? next : r) : [next, ...prev]);
    setIntakeDrafts(previous => previous.map(draft => draft.linkedRecordId === next.id ? { ...draft, values: syncRecordToDraftValues(draft, next) } : draft));
    setEditing(null);
    flash(autoReopen ? "已重新上架，並記錄在案名後方備註" : "物件已儲存");
  };
  const saveNewCaseReminder = () => {
    if (!newCaseReminderBatchIds.length) return;
    newCaseReminderBatchIds.forEach(id => deferredNewCaseReminderIds.current.add(id));
    setNewCaseReminder(null);
    setNewCaseReminderBatchIds([]);
    flash("新案件完成提醒已儲存");
  };
  const deferNewCaseReminder = () => {
    newCaseReminderBatchIds.forEach(id => deferredNewCaseReminderIds.current.add(id));
    setNewCaseReminder(null);
    setNewCaseReminderBatchIds([]);
  };
  const updateNewCaseReminderItem = (id: string, key: string, checked: boolean) => {
    const value = checked ? "1" : "";
    const savedAt = new Date().toISOString();
    setRecords(previous => previous.map(record => record.id === id ? { ...record, [key]: value, lastModifiedAt: savedAt } : record));
    setNewCaseReminder(previous => previous?.id === id ? { ...previous, [key]: value, lastModifiedAt: savedAt } : previous);
  };
  const temporarilySaveRecord = () => {
    if (!editing) return;
    const savedAt = new Date().toISOString();
    const next = normalizeRecordPings(clearImportedLandLabels({
      ...editing,
      developer: developerFullNameText(editing.developer || "", settings.personnel) || editing.developer,
      ...Object.fromEntries([...dateKeys].map(key => [key, normalizeDateInput(editing[key] || "")])),
      completionDate: normalizeDateInput(editing.completionDate || ""),
      lastModifiedAt: savedAt,
    }));
    if (next._intakeDraftId) {
      setIntakeDrafts(previous => previous.map(draft => draft.id === next._intakeDraftId ? { ...draft, values: syncRecordToDraftValues(draft, next), propertyKind: next.type.includes("土地") ? "土地" : "房屋" } : draft));
      setTourItems(previous => previous.map(item => item.data._intakeDraftId === next._intakeDraftId ? { ...item, data: { ...item.data, ...next } } : item));
    } else {
      setRecords(previous => previous.some(record => record.id === next.id) ? previous.map(record => record.id === next.id ? { ...record, ...next } : record) : [next, ...previous]);
      setIntakeDrafts(previous => previous.map(draft => draft.linkedRecordId === next.id ? { ...draft, values: syncRecordToDraftValues(draft, next) } : draft));
    }
    editingInitialRef.current = JSON.stringify(next);
    setEditing(next);
    flash("已暫存，可繼續填寫");
  };

  const saveColorSheetIssue = (data: RecordItem) => {
    const existing = records.find(record => record.id === data.id);
    if (!existing) return flash("此案件尚未正式進案，無法儲存缺件紀錄");
    const colorSheetIssue = String(data.colorSheetIssue || "").trim();
    const next = withTrackedUpdate(existing, { ...existing, colorSheetIssue });
    setRecords(previous => previous.map(record => record.id === next.id ? next : record));
    setEditing(previous => previous?.id === next.id ? { ...previous, colorSheetIssue: next.colorSheetIssue } : previous);
    flash(colorSheetIssue ? "已儲存彩色表缺件紀錄" : "已清除彩色表缺件紀錄");
  };
  const archiveRecord = (r: RecordItem, status = "下架洽開發", archiveDate = today(), salesPerson = r.salesPerson || "", reason = r.archiveReason || "") => {
    const cleanupKeys = ["housingDownDate", "bookDownDate", "salesBookDownDate", ...archiveWebsiteTasks.map(([field]) => `${field}DownDate`), ...dealCompletionTasks.map(([key]) => key)];
    setRecords(prev => prev.map(x => x.id === r.id ? { ...x, status, archived: normalizeDateInput(archiveDate) || today(), salesPerson: salesPerson.trim(), archiveReason: status === "下架洽開發" ? reason.trim() : "", _archiveActionDate: today(), ...Object.fromEntries(cleanupKeys.map(key => [key, ""])) } : x)); setArchiveChoice(null); flash("已移至封存，並建立下架待辦");
  };
  const completeArchiveCleanup = (record: RecordItem, key: string) => setRecords(previous => previous.map(item => item.id === record.id ? { ...item, [key]: item[key] ? "" : today() } : item));
  const completeArchiveCleanupAndRefresh = (record: RecordItem, key: string) => {
    const value = record[key] ? "" : today();
    suppressArchiveCleanupReminderRef.current = true;
    setArchiveCleanupReminderOpen(false);
    completeArchiveCleanup(record, key);
    setEditing(current => current?.id === record.id ? { ...current, [key]: value } : current);
  };
  const completeDealTask = (record: RecordItem, key: string) => setRecords(previous => previous.map(item => item.id === record.id ? { ...item, [key]: today() } : item));
  const requestArchive = (record: RecordItem, status: string) => setArchiveChoice({ record, status, date: status === "到期下架" ? nextDate(record.entrustEnd) : today(), salesPerson: record.salesPerson || "", reason: record.archiveReason || "" });
  const restoreRecord = (r: RecordItem, reopened = true) => {
    const cleanupKeys = ["housingDownDate", "bookDownDate", "salesBookDownDate", ...archiveWebsiteTasks.map(([field]) => `${field}DownDate`)];
    setRecords(prev => prev.map(x => x.id === r.id ? { ...x, status: "委託中", archived: "", entrustEnd: x.entrustEnd < today() ? "" : x.entrustEnd, ...(reopened ? { _restoredAt: today() } : Object.fromEntries(cleanupKeys.map(key => [key, ""]))) } : x)); setRestoreChoiceRecord(null); flash(reopened ? "已重新上架" : "已恢復委託中物件");
  };
  const removeRecord = (r: RecordItem) => { if (confirm(`確定永久刪除「${r.caseName}」？`)) { setRecords(prev => prev.filter(x => x.id !== r.id)); setIntakeDrafts(prev => prev.filter(draft => draft.linkedRecordId !== r.id)); setEditing(null); } };
  const returnToIntake = (r: RecordItem) => {
    if (!confirm(`確定將「${r.caseName}」退出委託中，退回進案草稿等待？`)) return;
    const linkedDraft = intakeDrafts.find(draft => draft.linkedRecordId === r.id);
    setIntakeDrafts(prev => linkedDraft ? prev.map(draft => draft.id === linkedDraft.id ? { ...draft, linkedRecordId: undefined, enteredAt: undefined } : draft) : [recordToIntake(r), ...prev]);
    setRecords(prev => prev.filter(item => item.id !== r.id)); setEditing(null); setTab("intake"); flash("已退回進案草稿等待");
  };

  const download = (name: string, blob: Blob) => { const a = document.createElement("a"); const url = URL.createObjectURL(blob); a.href = url; a.download = name; a.style.display = "none"; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 3000); };
  const exportJson = () => download(jsonExportFileName(), new Blob([JSON.stringify({ version: 2, exportedAt: new Date().toISOString(), settings: { personnel: settings.personnel, bookReviewCurrentDate: settings.bookReviewCurrentDate, bookReviewNextDate: settings.bookReviewNextDate, expiry591: settings.expiry591, expiry5168: settings.expiry5168, brokerExpiry: settings.brokerExpiry }, records }, null, 2)], { type: "application/json" }));
  const importJson = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]; if (!file) return;
    const reader = new FileReader(); reader.onload = () => { try {
      // Windows exports can contain a UTF-8 BOM; remove it before JSON parsing.
      const data = JSON.parse(String(reader.result).replace(/^\uFEFF/, "")); const list = Array.isArray(data) ? data : data.records; if (!Array.isArray(list)) throw new Error();
      const repairDuplicates = !Array.isArray(data) && data.importMode === "repair-duplicate-property-no";
      if (repairDuplicates) {
        if (!confirm("將依物件編號合併重複案件，保留原有資料並補上空白欄位，確定嗎？")) return;
        const grouped = new Map<string, RecordItem[]>();
        records.forEach(record => { const key = String(record.propertyNo || record.id || "").trim(); grouped.set(key, [...(grouped.get(key) || []), record]); });
        let removed = 0;
        const nextRecords = [...grouped.values()].map(group => {
          if (group.length === 1) return group[0]; removed += group.length - 1;
          const ordered = group.slice().sort((a, b) => Number(String(b.id || "").startsWith("official-")) - Number(String(a.id || "").startsWith("official-")) || Object.values(b).filter(Boolean).length - Object.values(a).filter(Boolean).length);
          return ordered.slice(1).reduce((merged, candidate) => {
            const additions = Object.fromEntries(Object.entries(candidate).filter(([key, value]) => key !== "id" && key !== "photos" && String(value ?? "").trim() && !String(merged[key] || "").trim()));
            return { ...merged, ...additions, photos: merged.photos?.length ? merged.photos : (candidate.photos || []) };
          }, ordered[0]);
        });
        localStorage.setItem(STORAGE_KEY, JSON.stringify(nextRecords)); alert(`已修復 ${removed} 筆重複案件，目前共 ${nextRecords.length} 件。`); location.reload(); return;
      }
      const supplement = !Array.isArray(data) && data.importMode === "active-sales-form-supplement";
      if (supplement) {
        if (!confirm(`將依物件編號補入 ${list.length} 件售屋資料表欄位；已填寫的欄位不會被覆蓋，確定嗎？`)) return;
        const patches = new Map(list.map((item: RecordItem) => [String(item.propertyNo || "").trim(), item]));
        let changed = 0;
        const nextRecords = records.map(record => {
          const patch = patches.get(String(record.propertyNo || "").trim()); if (!patch) return record;
          const additions = Object.fromEntries(Object.entries(patch).filter(([key, value]) => key !== "propertyNo" && key !== "id" && key !== "photos" && String(value ?? "").trim() && !String(record[key] || "").trim()));
          if (Object.keys(additions).length) changed += 1;
          return { ...record, ...additions };
        });
        localStorage.setItem(STORAGE_KEY, JSON.stringify(nextRecords)); alert(`售屋資料表補齊完成：${changed} 件有新增欄位，原本已填資料均已保留。`); location.reload(); return;
      }
      const fullBackup = !Array.isArray(data) && Number(data.version || 0) >= 2 && Array.isArray(data.records); const replaceExisting = fullBackup || (!Array.isArray(data) && data.replaceExisting === true); const message = replaceExisting ? `將以備份中的 ${list.length} 筆資料取代 Edge 目前物件，確定嗎？` : `將匯入 ${list.length} 筆資料，並與現有資料合併，確定嗎？`; if (confirm(message)) { const normalized = applySourceLayoutFixes(list.map((r: RecordItem) => ({ ...blankRecord(), ...r, id: r.id || newId(), photos: Array.isArray(r.photos) ? r.photos : [] }))); let nextRecords = normalized; if (!replaceExisting) { const map = new Map(records.map(r => [r.id, r])); normalized.forEach((r: RecordItem) => map.set(r.id, r)); nextRecords = [...map.values()]; } localStorage.setItem(STORAGE_KEY, JSON.stringify(nextRecords)); if (!Array.isArray(data) && data.settings) { let savedSettings: any = {}; try { savedSettings = JSON.parse(localStorage.getItem(SETTINGS_KEY) || "{}"); } catch {} localStorage.setItem(SETTINGS_KEY, JSON.stringify({ ...savedSettings, ...(Array.isArray(data.settings.personnel) ? { personnel: mergeSuppliedPersonnel(data.settings.personnel) } : {}), ...(data.settings.bookReviewCurrentDate ? { bookReviewCurrentDate: data.settings.bookReviewCurrentDate } : {}), ...(data.settings.bookReviewNextDate ? { bookReviewNextDate: data.settings.bookReviewNextDate } : {}), ...(data.settings.expiry591 ? { expiry591: data.settings.expiry591 } : {}), ...(data.settings.expiry5168 ? { expiry5168: data.settings.expiry5168 } : {}), ...(data.settings.brokerExpiry ? { brokerExpiry: data.settings.brokerExpiry } : {}) })); } localStorage.setItem("property-desk-import-prefer-local-once", "1"); localStorage.setItem("property-desk-official-merge-2026-07-30-v6", "1"); localStorage.setItem("property-desk-app-restore-216-v3", "1"); alert(`JSON 匯入完成，共 ${nextRecords.length} 筆，現在重新載入。`); location.reload(); }
    } catch (error) {
      console.error("JSON import failed", error);
      const detail = error instanceof Error ? error.message : String(error || "");
      flash(detail ? `JSON 匯入失敗：${detail}` : "JSON 檔案格式錯誤");
    } e.target.value = ""; }; reader.readAsText(file);
  };
  const exportExcel = () => {
    const cols = activeColumns.filter(k => k !== "photos");
    const esc = (v: string) => String(v || "").replace(/&/g, "&amp;").replace(/</g, "&lt;");
    const rows = records.map(r => `<tr>${cols.map(k => `<td>${esc(cellValue(r, k))}</td>`).join("")}</tr>`).join("");
    const html = `<html><head><meta charset="UTF-8"></head><body><table border="1"><tr>${cols.map(k => `<th>${labels[k]}</th>`).join("")}</tr>${rows}</table></body></html>`;
    download(`物件總表_${today()}.xls`, new Blob(["\ufeff", html], { type: "application/vnd.ms-excel" }));
  };
  const exportPptOrderImage = () => {
    if (!selectedPptRecords.length) return flash("目前沒有選擇要產生圖片的物件");
    const chunkText = (value = "", size = 20) => { const chars = Array.from(value); return Array.from({ length: Math.ceil(chars.length / size) }, (_, index) => chars.slice(index * size, index * size + size).join("")); };
    const width = 1600; const margin = 0; const tableWidth = width; const headerHeight = 58;
    const rowHeights = selectedPptRecords.map(record => Math.max(64, chunkText(record.address || "—", 20).length * 34 + 18));
    const standards = [
      "(必)*三個優點三個缺點",
      "本件成交行情", "周邊成交行情", "周遭生活機能/設施", "本棟/本區在售物件簡述", "(專簽除外)如一般約幾間在賣",
      "開價/底價/趴數/注意事項(介紹人、產權、坪數增建說明)/鑰匙", "帶看時間/帶看須知", "本件適合客戶族群",
      "土地", "必要報告:是否臨路+幾米/面寬/深度/建蔽/容積", "建地:適合蓋甚麼/建築線", "農地:農用證明/容許"
    ];
    const tableTop = 160; const standardsTop = tableTop + headerHeight + rowHeights.reduce((sum, value) => sum + value, 0) + 58; const height = standardsTop + 70 + standards.length * 39;
    const canvas = document.createElement("canvas"); canvas.width = width; canvas.height = height; const context = canvas.getContext("2d"); if (!context) return flash("圖片產生失敗");
    const setFittedFont = (text: string, maxWidth: number, startSize: number, bold = false) => { let size = startSize; do { context.font = `${bold ? "bold " : ""}${size}px "Microsoft JhengHei", sans-serif`; size -= 1; } while (size >= 17 && context.measureText(text).width > maxWidth); };
    context.fillStyle = "#ffffff"; context.fillRect(0, 0, width, height); context.textBaseline = "middle"; context.font = 'bold 48px "Microsoft JhengHei", sans-serif'; context.fillStyle = "#c62828"; context.textAlign = "center"; context.fillText("下周一 開會進案報告順序 請同仁準備", width / 2, 76);
    const columns = [margin, margin + 110, margin + 650, margin + 1320, width - margin];
    context.fillStyle = "#dcece6"; context.fillRect(margin, tableTop, tableWidth, headerHeight); context.strokeStyle = "#78988d"; context.lineWidth = 2; context.strokeRect(margin, tableTop, tableWidth, headerHeight);
    context.font = 'bold 30px "Microsoft JhengHei", sans-serif'; context.fillStyle = "#173f35"; ["序", "案名", "地址", "開發"].forEach((label, index) => context.fillText(label, (columns[index] + columns[index + 1]) / 2, tableTop + headerHeight / 2));
    let y = tableTop + headerHeight;
    selectedPptRecords.forEach((record, index) => { const rowHeight = rowHeights[index]; context.fillStyle = index % 2 ? "#f7faf8" : "#ffffff"; context.fillRect(margin, y, tableWidth, rowHeight); context.strokeStyle = "#a7bbb4"; context.strokeRect(margin, y, tableWidth, rowHeight); columns.slice(1, -1).forEach(x => { context.beginPath(); context.moveTo(x, y); context.lineTo(x, y + rowHeight); context.stroke(); }); context.fillStyle = "#1f2926"; setFittedFont(String(index + 1), columns[1] - columns[0] - 8, 30); context.fillText(String(index + 1), (columns[0] + columns[1]) / 2, y + rowHeight / 2); const caseName = record.caseName || "未命名案件"; setFittedFont(caseName, columns[2] - columns[1] - 8, 32, true); context.fillText(caseName, (columns[1] + columns[2]) / 2, y + rowHeight / 2); const addressLines = chunkText(record.address || "—", 20); addressLines.forEach((line, lineIndex) => { setFittedFont(line, columns[3] - columns[2] - 8, 30); context.fillText(line, (columns[2] + columns[3]) / 2, y + rowHeight / 2 + (lineIndex - (addressLines.length - 1) / 2) * 34); }); const developer = developerFullNameText(record.developer) || "—"; setFittedFont(developer, columns[4] - columns[3] - 8, 30); context.fillText(developer, (columns[3] + columns[4]) / 2, y + rowHeight / 2); y += rowHeight; });
    context.textAlign = "left"; let standardY = standardsTop;
    standards.forEach(line => { if (line === "土地") { standardY += 18; context.fillStyle = "#173f35"; context.font = 'bold 30px "Microsoft JhengHei", sans-serif'; } else if (line.startsWith("(必)")) { context.fillStyle = "#b32929"; context.font = 'bold 28px "Microsoft JhengHei", sans-serif'; } else { context.fillStyle = "#222222"; context.font = '25px "Microsoft JhengHei", sans-serif'; } context.fillText(line, 8, standardY); standardY += 39; });
    canvas.toBlob(blob => { if (!blob) return flash("圖片產生失敗"); download(`公告${pptMeetingDateLabel(selectedPptWeek)}開會星期一PPT.png`, blob); flash(`已產生圖片，共 ${selectedPptRecords.length} 筆物件`); }, "image/png");
  };
  const exportPptLegacy = async () => {
    const selected = selectedPptRecords;
    if (!selected.length) return flash("目前沒有勾選要加入 PPT 的物件");
    flash(`正在產生 PPT，共 ${selected.length} 筆物件，請稍候…`);
    const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_4x3"; pptx.author = "台慶不動產"; pptx.subject = "進案物件"; pptx.title = "進案物件"; pptx.lang = "zh-TW";
    const addLine = (slide: any, label: string, value: string, x: number, y: number, w = 4.15) => { slide.addText(label, { x, y, w: .98, h: .36, fontFace: "Microsoft JhengHei", fontSize: 14, bold: true, color: "245A4A", align: "right", margin: 0 }); slide.addText("：", { x: x + .98, y, w: .22, h: .36, fontFace: "Microsoft JhengHei", fontSize: 14, bold: true, color: "245A4A", margin: 0 }); slide.addText(value || "—", { x: x + 1.2, y, w: w - 1.2, h: .36, fontFace: "Microsoft JhengHei", fontSize: 15, color: "222222", margin: 0, fit: "shrink", breakLine: false }); };
    selected.forEach(record => {
      const slide = pptx.addSlide(); slide.background = { color: "FFFFFF" }; const noteText = displayNoteSegments(record.notes).join("；"); const land = typeShort(record.type) === "土地" || /^(LG|LA)/i.test(record.propertyNo || ""); const linkedDraft = intakeDrafts.find(draft => draft.id === record._intakeDraftId || draft.linkedRecordId === record.id); const draftValue = (...keys: string[]) => linkedDraft ? intakeValue(linkedDraft.values, ...keys) : ""; const numberOf = (value = "") => Number(String(value).replace(/,/g, "").match(/-?\d+(?:\.\d+)?/)?.[0] || 0); const unitPrice = numberOf(record.price) && numberOf(record.landPing) ? `$${Math.round(numberOf(record.price) * 10000 / numberOf(record.landPing)).toLocaleString("en-US")}` : "—";
      slide.addText("案名：", { x: .32, y: .16, w: 1.05, h: .45, fontFace: "Microsoft JhengHei", fontSize: 18, bold: true, color: "111111", margin: 0 });
      slide.addText(record.caseName || "未命名物件", { x: 1.32, y: .11, w: 8.35, h: .55, fontFace: "Microsoft JhengHei", fontSize: 25, bold: true, color: "111111", margin: 0, fit: "shrink" });
      slide.addShape(pptx.ShapeType.line, { x: 0, y: .7, w: 10, h: 0, line: { color: "555555", width: 1 } });
      slide.addText("地址：", { x: .32, y: .82, w: 1.05, h: .38, fontFace: "Microsoft JhengHei", fontSize: 17, bold: true, color: "111111", margin: 0 });
      slide.addText(record.address || "—", { x: 1.32, y: .75, w: 8.35, h: .52, fontFace: "Microsoft JhengHei", fontSize: 20, bold: true, color: "111111", margin: 0, fit: "shrink" });
      slide.addShape(pptx.ShapeType.line, { x: 0, y: 1.35, w: 10, h: 0, line: { color: "555555", width: 1 } });
      slide.addShape(pptx.ShapeType.roundRect, { x: .25, y: 1.5, w: 5.45, h: 3.85, rectRadius: .08, line: { color: "B7CCC4", width: 1.2 }, fill: { color: "FFFFFF" } });
      slide.addShape(pptx.ShapeType.rect, { x: 5.95, y: 1.55, w: 3.8, h: 2.55, line: { color: "4E5D58", width: 1.3 }, fill: { color: "FFFFFF" } });
      slide.addShape(pptx.ShapeType.rect, { x: 5.95, y: 4.35, w: 3.8, h: 2.55, line: { color: "4E5D58", width: 1.3 }, fill: { color: "FFFFFF" } });
      slide.addText("開發", { x: .42, y: 1.68, w: .75, h: .42, fontFace: "Microsoft JhengHei", fontSize: 16, bold: true, color: "111111", align: "right", margin: 0 });
      slide.addText("：", { x: 1.17, y: 1.68, w: .18, h: .42, fontFace: "Microsoft JhengHei", fontSize: 16, bold: true, color: "111111", margin: 0 });
      slide.addText(developerFullNameText(record.developer) || "—", { x: 1.35, y: 1.6, w: 1.55, h: .55, fontFace: "Microsoft JhengHei", fontSize: 21, bold: true, color: "1648D8", align: "center", margin: 0, fit: "shrink", fill: { color: "FCE9D9" }, line: { color: "D6A77B", width: .8 } });
      slide.addText(land ? "開價" : "委託總價", { x: 3.05, y: 1.68, w: 1, h: .42, fontFace: "Microsoft JhengHei", fontSize: 15, bold: true, color: "111111", align: "right", margin: 0, fit: "shrink" });
      slide.addText("：", { x: 4.05, y: 1.68, w: .2, h: .42, fontFace: "Microsoft JhengHei", fontSize: 15, bold: true, color: "111111", margin: 0 });
      slide.addText(`${record.reducedPrice || record.price || "—"}${record.reducedPrice && record.price ? `（原${record.price}）` : ""}`, { x: 4.25, y: 1.56, w: 1.12, h: .62, fontFace: "Microsoft JhengHei", fontSize: 27, bold: true, color: "E02020", align: "center", margin: 0, fit: "shrink" });
      if (land) {
        addLine(slide, "總地坪", record.landPing ? `${record.landPing.replace(/\s*坪\s*$/, "")} 坪` : "—", .42, 2.25, 2.55); addLine(slide, "每坪單價", unitPrice, 3.05, 2.25, 2.45);
        addLine(slide, "臨路", record.road, .42, 2.78, 2.55); addLine(slide, "座向", record.direction, 3.05, 2.78, 2.45);
        addLine(slide, "面寬", record.frontage, .42, 3.31, 2.55); addLine(slide, "深度", record.depth, 3.05, 3.31, 2.45);
        addLine(slide, "建蔽/容積率", [record.coverage, record.far].filter(Boolean).join("／"), .42, 3.84, 2.55); addLine(slide, "使用分區", record.zoning, 3.05, 3.84, 2.45);
        slide.addText("備註：", { x: .42, y: 4.48, w: 1.05, h: .38, fontFace: "Microsoft JhengHei", fontSize: 14, bold: true, color: "245A4A", margin: 0 });
        slide.addText(noteText, { x: 1.45, y: 4.43, w: 4.2, h: .72, fontFace: "Microsoft JhengHei", fontSize: 14, color: "222222", margin: 0, valign: "top", fit: "shrink" });
      } else {
        addLine(slide, "總地坪", record.landPing ? `${record.landPing.replace(/\s*坪\s*$/, "")} 坪` : "—", .42, 2.2, 2.55); addLine(slide, "總建坪", record.buildingPing ? `${record.buildingPing.replace(/\s*坪\s*$/, "")} 坪` : "—", 3.05, 2.2, 2.45);
        addLine(slide, "室內坪", record.indoorPing ? `${record.indoorPing.replace(/\s*坪\s*$/, "")} 坪` : "—", .42, 2.63, 2.55); addLine(slide, "格局", record.layout, 3.05, 2.63, 2.45);
        addLine(slide, "面寬", record.frontage, .42, 3.06, 2.55); addLine(slide, "深度", record.depth, 3.05, 3.06, 2.45);
        addLine(slide, "臨路", record.road, .42, 3.49, 2.55); addLine(slide, "樓層", floorPptDisplay(record.floor), 3.05, 3.49, 2.45);
        addLine(slide, "朝向", record.direction, .42, 3.92, 2.55); addLine(slide, "車位", parkingShort(record.parking), 3.05, 3.92, 2.45);
        addLine(slide, "管理費", /^\$?0(?:\.0+)?(?:元)?$/.test(record.managementFee || "") ? "無" : record.managementFee ? `${record.managementFee.replace(/\s*\/月\s*$/, "")}/月` : "—", .42, 4.35, 2.55); addLine(slide, "完工日&屋齡", `${draftValue("建築完成日期") || record.builtYear || "—"} ${ageOf(record)}`, 3.05, 4.35, 2.45);
        addLine(slide, "社區名稱", draftValue("大樓名稱") || record.communityName, .42, 4.78, 2.55); addLine(slide, "現況", record.currentState, 3.05, 4.78, 2.45);
        addLine(slide, "鑰匙", record.key, .42, 5.21, 2.55);
        slide.addText("備註：", { x: 3.05, y: 5.21, w: 1.05, h: .38, fontFace: "Microsoft JhengHei", fontSize: 14, bold: true, color: "245A4A", margin: 0 });
        slide.addText(noteText, { x: 4.08, y: 5.18, w: 1.55, h: .55, fontFace: "Microsoft JhengHei", fontSize: 13, color: "222222", margin: 0, valign: "top", fit: "shrink" });
      }
      slide.addText(`進案報件日期：${displayRocDate(record.reportDate) || "—"}`, { x: .32, y: 6.88, w: 3.4, h: .34, fontFace: "Microsoft JhengHei", fontSize: 12, bold: true, color: "333333", margin: 0 });
      slide.addText(`物件編號：${record.propertyNo || "—"}`, { x: 3.6, y: 6.88, w: 2.8, h: .34, fontFace: "Microsoft JhengHei", fontSize: 12, bold: true, color: "333333", align: "center", margin: 0 });
      if (typeShort(record.type) === "土地" || /^(LG|LA)/i.test(record.propertyNo || "")) { const blankOne = pptx.addSlide(); blankOne.background = { color: "FFFFFF" }; const blankTwo = pptx.addSlide(); blankTwo.background = { color: "FFFFFF" }; }
    });
    try { const pptBlob = await pptx.write({ outputType: "blob" }) as Blob; download(`${pptMeetingDateLabel(selectedPptWeek)}開會PPT.pptx`, pptBlob); setPptPickerOpen(false); flash(`已產生 PPT，共 ${selected.length} 筆物件`); } catch (error) { console.error(error); flash("PPT 產生失敗，請重新整理後再試一次"); }
  };
  const exportPpt = async () => {
    const selected = selectedPptRecords;
    if (!selected.length) return flash("請先選擇要放入 PPT 的物件");
    flash(`正在產生 PPT，共 ${selected.length} 筆物件…`);
    const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_4x3"; pptx.author = "台慶不動產"; pptx.lang = "zh-TW";
    const font = "標楷體";
    const clean = (value = "") => /^(?:0|無|0；無|0;無)$/.test(String(value).trim()) ? "" : String(value || "").trim();
    const num = (value = "") => String(value || "").replace(/,/g, "").match(/-?\d+(?:\.\d+)?/)?.[0] || "";
    const pptCompletionDate = (value = "") => {
      const text = String(value || "").trim();
      const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
      if (iso) return `${String(Number(iso[1]) - 1911).padStart(3, "0")}.${String(Number(iso[2])).padStart(2, "0")}.${String(Number(iso[3])).padStart(2, "0")}`;
      const roc = text.match(/^(\d{2,3})[.\/-](\d{1,2})[.\/-](\d{1,2})$/);
      return roc ? `${roc[1].padStart(3, "0")}.${roc[2].padStart(2, "0")}.${roc[3].padStart(2, "0")}` : text;
    };
    const pptBuildingAge = (value = "") => {
      const text = String(value || "").trim();
      const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
      const roc = text.match(/^(\d{2,3})[.\/-](\d{1,2})[.\/-](\d{1,2})$/);
      const year = iso ? Number(iso[1]) : roc ? Number(roc[1]) + 1911 : 0;
      const age = year ? new Date().getFullYear() - year : NaN;
      return Number.isFinite(age) && age >= 0 ? `約${age}` : "";
    };
    const pptManagementFeeRuns = (value = "") => {
      const raw = String(value || "").trim();
      const noMonthlySuffix = raw.replace(/\s*[\/／]\s*月\s*$/, "").trim();
      if (!noMonthlySuffix || /^\$?0(?:\.0+)?(?:\s*[\/／]\s*月)?$/.test(raw) || /^無(?:\s*[\/／]\s*月)?$/.test(raw)) return [{ text: "無$0", options: { fontFace: font, fontSize: 21 } }];
      if (/[\/／]\s*年\s*$/.test(raw)) return [{ text: raw, options: { fontFace: font, fontSize: 21 } }];
      return [{ text: noMonthlySuffix, options: { fontFace: font, fontSize: 21 } }, { text: "/月", options: { fontFace: font, fontSize: 12 } }];
    };
    const addAlignedLabel = (slide: any, label: string, x: number, y: number, labelSize = 17) => {
      const chars = Array.from(label);
      if (chars.length >= 2 && chars.length <= 4) chars.forEach((char, index) => slide.addText(char, { x: x + index * (.68 / (chars.length - 1)), y, w: .22, h: .4, fontFace: font, fontSize: labelSize, color: "000000", align: "center", margin: 0, breakLine: false }));
      else slide.addText(label, { x, y, w: .9, h: .4, fontFace: font, fontSize: labelSize > 17 ? 16 : 14, color: "000000", align: "center", margin: 0, breakLine: false, fit: "shrink" });
    };
    const addLabel = (slide: any, label: string, value: string, x: number, y: number, w: number, valueColor = "000000", valueSize = 21, labelSize = 17, lineOffset = .39) => {
      addAlignedLabel(slide, label, x, y, labelSize);
      slide.addText("：", { x: x + .88, y, w: .2, h: .36, fontFace: font, fontSize: 17, color: "000000", margin: 0 });
      slide.addText(value || "", { x: x + 1.08, y: y - .04, w: w - 1.08, h: .44, fontFace: font, fontSize: valueSize, color: valueColor, margin: 0, align: "center", fit: "shrink", breakLine: false });
      slide.addShape(pptx.ShapeType.line, { x: x + 1.08, y: y + lineOffset, w: w - 1.13, h: 0, line: { color: "555555", width: .7 } });
    };
    const addRichLabel = (slide: any, label: string, runs: any[], x: number, y: number, w: number) => {
      addAlignedLabel(slide, label, x, y);
      slide.addText("：", { x: x + .88, y, w: .2, h: .36, fontFace: font, fontSize: 17, color: "000000", margin: 0 });
      const safeRuns = runs.filter(run => String(run?.text || "").length > 0);
      slide.addText(safeRuns.length ? safeRuns : [{ text: " ", options: { fontFace: font, fontSize: 12 } }], { x: x + 1.08, y: y - .04, w: w - 1.08, h: .44, fontFace: font, fontSize: 21, color: "000000", margin: 0, align: "center", fit: "shrink", breakLine: false });
      slide.addShape(pptx.ShapeType.line, { x: x + 1.08, y: y + .39, w: w - 1.13, h: 0, line: { color: "555555", width: .7 } });
    };
    selected.forEach(record => {
      const slide = pptx.addSlide(); slide.background = { color: "FFFFFF" };
      const land = typeShort(record.type) === "土地" || /^(LG|LA)/i.test(record.propertyNo || "");
      const frameColor = "222222";
      const linkedDraft = intakeDrafts.find(draft => draft.id === record._intakeDraftId || draft.linkedRecordId === record.id);
      const draftValue = (...keys: string[]) => linkedDraft ? intakeValue(linkedDraft.values, ...keys) : "";
      const noteText = displayNoteSegments(record.notes).join("；");
      const price = num(record.reducedPrice || record.price);
      const landPing = num(record.landPing);
      const unitPrice = price && landPing ? `$${Math.round(Number(price) * 10000 / Number(landPing)).toLocaleString("en-US")}` : "";
      const pptAge = ageOf(record).match(/\d+年屋/)?.[0] || ageOf(record);
      const leftW = 6.05, rightX = 6.05;
      slide.addText("案名", { x: .05, y: .16, w: .72, h: .55, fontFace: font, fontSize: 16, color: "000000", align: "center", valign: "mid", margin: 0, breakLine: false });
      slide.addText(record.caseName || "", { x: .78, y: .08, w: 5.23, h: .72, fontFace: font, fontSize: land ? 28 : 25, color: "000000", align: "center", valign: "mid", margin: 0, fit: "shrink", breakLine: false });
      slide.addText("地址", { x: 6.08, y: .16, w: .68, h: .55, fontFace: font, fontSize: 16, color: "000000", align: "center", valign: "mid", margin: 0, breakLine: false });
      slide.addText(record.address || "", { x: 6.78, y: .08, w: 3.12, h: .72, fontFace: font, fontSize: 19, color: "000000", align: "center", valign: "mid", margin: 0, fit: "shrink", breakLine: false });
      slide.addShape(pptx.ShapeType.line, { x: 0, y: .84, w: 10, h: 0, line: { color: frameColor, width: 1 } });
      slide.addShape(pptx.ShapeType.rect, { x: 0, y: .87, w: 2.756, h: land ? .94 : .787, line: { color: "FCE9D9", transparency: 100 }, fill: { color: "FCE9D9" } });
      slide.addText("開發", { x: .06, y: land ? 1.14 : 1.07, w: .58, h: .35, fontFace: font, fontSize: 15, bold: true, color: "000000", align: "center", margin: 0, breakLine: false });
      slide.addText(developerFullNameText(record.developer) || "", { x: .66, y: land ? 1.08 : 1.0, w: 2.09, h: .52, fontFace: font, fontSize: 19, bold: true, color: "0000FF", align: "center", valign: "mid", margin: 0, fit: "shrink", breakLine: false });
      slide.addText(land ? "開價" : "委託總價", { x: 2.76, y: land ? 1.15 : 1.08, w: 1.12, h: .38, fontFace: font, fontSize: 18, color: "000000", align: "right", margin: 0, breakLine: false });
      slide.addText("：", { x: 3.9, y: land ? 1.15 : 1.08, w: .2, h: .38, fontFace: font, fontSize: 18, margin: 0 });
      slide.addText([{ text: price || "", options: { fontFace: font, fontSize: 37, bold: true, color: "E00000" } }, { text: "萬", options: { fontFace: font, fontSize: 16, color: "000000", breakLine: false } }], { x: 4.1, y: .89, w: 1.8, h: .72, fontFace: font, align: "center", valign: "mid", margin: 0, fit: "shrink", breakLine: false });
      slide.addShape(pptx.ShapeType.line, { x: 0, y: land ? 1.82 : 1.68, w: leftW, h: 0, line: { color: frameColor, width: 1 } });
      if (land) {
        addLabel(slide, "總地坪", landPing ? `${landPing}坪` : "", .08, 2.12, 2.85, "0000FF", 25, 20, .58); addLabel(slide, "每坪單價", unitPrice, 3.02, 2.12, 2.82, "000000", 25, 20, .58);
        addLabel(slide, "臨路", num(record.road) ? `${num(record.road)}米` : clean(record.road), .08, 3.19, 2.85, "000000", 25, 20, .58); addLabel(slide, "座向", record.direction, 3.02, 3.19, 2.82, "000000", 25, 20, .58);
        addLabel(slide, "面寬", num(record.frontage) ? `${num(record.frontage)}米` : clean(record.frontage), .08, 4.26, 2.85, "000000", 25, 20, .58); addLabel(slide, "深度", num(record.depth) ? `${num(record.depth)}米` : clean(record.depth), 3.02, 4.26, 2.82, "000000", 25, 20, .58);
        addAlignedLabel(slide, "建蔽率", .08, 5.08, 14);
        addAlignedLabel(slide, "容積率", .08, 5.48, 14);
        slide.addText("：", { x: .96, y: 5.28, w: .2, h: .36, fontFace: font, fontSize: 17, margin: 0 });
        slide.addText([record.coverage, record.far].filter(Boolean).join("／"), { x: 1.16, y: 5.27, w: 1.67, h: .44, fontFace: font, fontSize: 20, align: "center", margin: 0, fit: "shrink" });
        slide.addShape(pptx.ShapeType.line, { x: 1.16, y: 5.91, w: 1.67, h: 0, line: { color: "555555", width: .7 } });
        addLabel(slide, "使用分區", record.zoning, 3.02, 5.33, 2.82, "000000", 22, 20, .58);
      } else {
        addLabel(slide, "總地坪", landPing ? `${landPing}坪` : "", .08, 1.8, 2.85, "0000FF"); addLabel(slide, "總建坪", num(record.buildingPing) ? `${num(record.buildingPing)}坪` : "", 3.02, 1.8, 2.82, "0000FF");
        addLabel(slide, "室內坪", num(record.indoorPing) ? `${num(record.indoorPing)}坪` : "", .08, 2.36, 2.85, "0000FF"); addRichLabel(slide, "格局", Array.from(record.layout || "").map(char => ({ text: char, options: { fontFace: font, fontSize: /[房廳衛浴陽台]/.test(char) ? 13 : 21 } })), 3.02, 2.36, 2.82);
        addLabel(slide, "面寬", num(record.frontage) ? `${num(record.frontage)}米` : clean(record.frontage), .08, 2.92, 2.85); addLabel(slide, "深度", num(record.depth) ? `${num(record.depth)}米` : clean(record.depth), 3.02, 2.92, 2.82);
        addLabel(slide, "臨路", num(record.road) ? `${num(record.road)}米` : clean(record.road), .08, 3.48, 2.85); addRichLabel(slide, "樓層", Array.from(floorPptDisplay(record.floor)).map(char => ({ text: char, options: { fontFace: font, fontSize: /[\u4e00-\u9fff]/.test(char) ? 12 : 21 } })), 3.02, 3.48, 2.82);
        addLabel(slide, "朝向", record.direction, .08, 4.04, 2.85); addLabel(slide, "車位", parkingShort(record.parking), 3.02, 4.04, 2.82, "000000", 17.5);
        addRichLabel(slide, "管理費", pptManagementFeeRuns(record.managementFee), .08, 4.6, 2.85); addLabel(slide, "社區名稱", draftValue("大樓名稱") || record.communityName, 3.02, 4.6, 2.82, "000000", 18);
        const completionForPpt = draftValue("建築完成日期") || record.completionDate || record.builtYear;
        addLabel(slide, "建築完成日期", pptCompletionDate(completionForPpt), .08, 5.16, 2.85, "000000", 18); addRichLabel(slide, "屋齡", [{ text: pptBuildingAge(completionForPpt), options: { fontFace: font, fontSize: 21 } }, { text: pptBuildingAge(completionForPpt) ? "年屋" : "", options: { fontFace: font, fontSize: 12 } }], 3.02, 5.16, 2.82);
        addLabel(slide, "現況", record.currentState, .08, 5.72, 2.85); addLabel(slide, "鑰匙", record.key, 3.02, 5.72, 2.82);
      }
      const notesY = land ? 6.40 : 6.28;
      addAlignedLabel(slide, "備註", .08, notesY);
      slide.addText("：", { x: .96, y: notesY, w: .2, h: .36, fontFace: font, fontSize: 17, color: "000000", margin: 0 });
      slide.addText(noteText, { x: 1.16, y: notesY, w: 4.61, h: .72, fontFace: font, fontSize: 16, color: "000000", margin: 0, align: "left", valign: "top", fit: "shrink" });
      slide.addShape(pptx.ShapeType.line, { x: 0, y: 7.18, w: leftW, h: 0, line: { color: frameColor, width: 1 } });
      slide.addText(`進案報件日期：${displayRocDate(record.reportDate) || ""}`, { x: .05, y: 7.22, w: 2.8, h: .24, fontFace: font, fontSize: 11, color: "000000", margin: 0 });
      slide.addText(`物件編號：${record.propertyNo || ""}`, { x: 3.1, y: 7.22, w: 2.9, h: .24, fontFace: font, fontSize: 11, color: "000000", align: "center", margin: 0 });
      slide.addShape(pptx.ShapeType.rect, { x: rightX, y: 1.0, w: 3.94, h: 3.15, line: { color: "FFFFFF", transparency: 100 }, fill: { color: "FFFFFF", transparency: 100 } });
      slide.addShape(pptx.ShapeType.rect, { x: rightX, y: 4.15, w: 3.94, h: 3.15, line: { color: "FFFFFF", transparency: 100 }, fill: { color: "FFFFFF", transparency: 100 } });
      if (land) { const blankOne = pptx.addSlide(); blankOne.background = { color: "FFFFFF" }; const blankTwo = pptx.addSlide(); blankTwo.background = { color: "FFFFFF" }; }
    });
    try { const pptBlob = await pptx.write({ outputType: "blob" }) as Blob; download(`${pptMeetingDateLabel(selectedPptWeek)}開會PPT.pptx`, pptBlob); setPptPickerOpen(false); flash(`已產生 PPT，共 ${selected.length} 筆物件`); } catch (error) { console.error(error); flash("PPT 產生失敗，請再試一次"); }
  };
  // Cloud backup intentionally excludes locally stored photo payloads.
  // Textual photo notes stay with the record, while real images remain on this computer.
  const cloudData = () => ({
    records: records.map(({ photos, ...record }) => record),
    settings: { personnel: settings.personnel, bookReviewCurrentDate: settings.bookReviewCurrentDate, bookReviewNextDate: settings.bookReviewNextDate, expiry591: settings.expiry591, expiry5168: settings.expiry5168, brokerExpiry: settings.brokerExpiry },
    intake: { raw: intakeRaw, drafts: intakeDrafts, selectedId: selectedIntakeId },
    tour: { date: tourDate, title: tourTitle, items: tourItems, modifiedAt: tourModifiedAt, history: tourHistory },
    pptWeeks: (() => {
      try {
        const weeks = JSON.parse(localStorage.getItem(PPT_WEEK_SELECTIONS_KEY) || "{}");
        return Object.fromEntries(Object.entries(weeks).map(([week, value]) => {
          const saved = value && typeof value === "object" ? value as { extraIds?: string[]; orderIds?: string[]; adHocRecords?: RecordItem[]; confirmedSnapshots?: Record<string, RecordItem> } : {};
          return [week, {
            extraIds: Array.isArray(saved.extraIds) ? saved.extraIds : [],
            orderIds: Array.isArray(saved.orderIds) ? saved.orderIds : [],
            adHocRecords: Array.isArray(saved.adHocRecords) ? saved.adHocRecords.map(({ photos, ...record }) => record) : [],
            confirmedSnapshots: Object.fromEntries(Object.entries(saved.confirmedSnapshots || {}).map(([id, snapshot]) => {
              const { photos, ...record } = snapshot;
              return [id, record];
            })),
          }];
        }));
      } catch { return {}; }
    })(),
  });
  const cloudTokenExpiresSoon = (token = "") => {
    try {
      const encoded = token.split(".")[1] || "";
      const normalized = encoded.replace(/-/g, "+").replace(/_/g, "/").padEnd(Math.ceil(encoded.length / 4) * 4, "=");
      const payload = JSON.parse(atob(normalized));
      return !payload.exp || Number(payload.exp) * 1000 <= Date.now() + 60000;
    } catch { return true; }
  };
  const refreshCloudSession = async (force = false) => {
    const current = cloudSession;
    if (!current?.accessToken) return null;
    if (!force && !cloudTokenExpiresSoon(current.accessToken)) return current;
    if (!current.refreshToken || !settings.supabaseUrl || !settings.supabaseKey) return null;
    try {
      const res = await fetch(`${settings.supabaseUrl.replace(/\/$/, "")}/auth/v1/token?grant_type=refresh_token`, { method: "POST", headers: { apikey: settings.supabaseKey, "Content-Type": "application/json" }, body: JSON.stringify({ refresh_token: current.refreshToken }) });
      const data = await res.json();
      if (!res.ok || !data.access_token) return null;
      const next = { accessToken: data.access_token, refreshToken: data.refresh_token || current.refreshToken, email: data.user?.email || current.email };
      setCloudSession(next);
      localStorage.setItem(CLOUD_SESSION_KEY, JSON.stringify(next));
      return next;
    } catch { return null; }
  };
  const cloudHeaders = (session = cloudSession) => ({ apikey: settings.supabaseKey, Authorization: `Bearer ${session?.accessToken || ""}`, "Content-Type": "application/json" });
  const supabasePush = async (quiet = false) => {
    if (!cloudSession?.accessToken) { if (!quiet) flash("請先登入雲端帳號"); return false; }
    if (!settings.supabaseUrl || !settings.supabaseKey) { if (!quiet) flash("請先填入 Supabase Publishable key"); return false; }
    try {
      const session = await refreshCloudSession();
      if (!session) throw new Error("cloud session expired");
      const url = `${settings.supabaseUrl.replace(/\/$/, "")}/rest/v1/${settings.supabaseTable}`;
      const uploadedAt = new Date().toISOString();
      const res = await fetch(url, { method: "POST", headers: { ...cloudHeaders(session), Prefer: "resolution=merge-duplicates,return=minimal" }, body: JSON.stringify({ id: settings.supabaseRecord, data: cloudData(), updated_at: uploadedAt }) });
      if (!res.ok) throw new Error(await res.text());
      setCloudLastUploadAt(uploadedAt);
      localStorage.setItem(CLOUD_LAST_UPLOAD_KEY, uploadedAt);
      if (!quiet) flash("雲端同步完成");
      return true;
    } catch { if (!quiet) flash("雲端同步失敗，請檢查登入與設定"); return false; }
  };
  const supabasePull = async (automatic = false) => {
    if (!cloudSession?.accessToken) return flash("請先登入雲端帳號");
    try {
      const session = await refreshCloudSession();
      if (!session) throw new Error("cloud session expired");
      const url = `${settings.supabaseUrl.replace(/\/$/, "")}/rest/v1/${settings.supabaseTable}?id=eq.${encodeURIComponent(settings.supabaseRecord)}&select=data,updated_at`;
      const res = await fetch(url, { headers: cloudHeaders(session) }); const rows = await res.json(); const data = rows[0]?.data;
      if (!res.ok || !data?.records) throw new Error();
      if (rows[0]?.updated_at) { setCloudLastUploadAt(rows[0].updated_at); localStorage.setItem(CLOUD_LAST_UPLOAD_KEY, rows[0].updated_at); }
      if (automatic || confirm("雲端資料將與本機資料合併，本機已修改但尚未同步的同一筆資料將以雲端版本為準。確定嗎？")) {
        setRecords(prev => {
          const map = new Map(prev.map(r => [r.id, r]));
          data.records.forEach((r: RecordItem) => {
            const normalized = normalizeRecordPings(applySourceLayoutFixes([r])[0]);
            const local = map.get(normalized.id);
            map.set(normalized.id, { ...local, ...normalized, photos: normalized.photos || local?.photos || [] });
          });
          return [...map.values()];
        });
        if (data.settings) setSettings(previous => ({ ...previous, ...(data.settings.bookReviewCurrentDate ? { bookReviewCurrentDate: data.settings.bookReviewCurrentDate } : {}), ...(data.settings.bookReviewNextDate ? { bookReviewNextDate: data.settings.bookReviewNextDate } : {}), ...(data.settings.expiry591 ? { expiry591: data.settings.expiry591 } : {}), ...(data.settings.expiry5168 ? { expiry5168: data.settings.expiry5168 } : {}), ...(data.settings.brokerExpiry ? { brokerExpiry: data.settings.brokerExpiry } : {}), ...(Array.isArray(data.settings.personnel) && data.settings.personnel.length > 0 ? { personnel: mergeSuppliedPersonnel(data.settings.personnel) } : {}) }));
        if (data.intake) {
          setIntakeRaw(previous => previous || data.intake.raw || "");
          setIntakeDrafts(previous => reconcileIntakeDraftLinks(mergeIntakeDrafts(previous, Array.isArray(data.intake.drafts) ? data.intake.drafts : []), [...records, ...data.records]));
          setSelectedIntakeId(previous => previous || data.intake.selectedId || "");
        }
        if (data.tour) {
          const remoteModifiedAt = String(data.tour.modifiedAt || "");
          const remoteStamp = Date.parse(remoteModifiedAt);
          const localStamp = Date.parse(tourModifiedAt || "");
          if (!tourModifiedAt || (Number.isFinite(remoteStamp) && remoteStamp > localStamp)) {
            setTourDate(data.tour.date || today());
            setTourTitle(data.tour.title || "");
            setTourItems(data.tour.items || []);
            setTourHistory(Array.isArray(data.tour.history) ? data.tour.history : []);
            setTourModifiedAt(remoteModifiedAt || new Date().toISOString());
          }
        }
        if (data.pptWeeks && typeof data.pptWeeks === "object") {
          try {
            const localWeeks = JSON.parse(localStorage.getItem(PPT_WEEK_SELECTIONS_KEY) || "{}");
            const mergedWeeks = { ...localWeeks, ...data.pptWeeks };
            Object.entries(data.pptWeeks).forEach(([week, value]) => {
              const cloudWeek = value && typeof value === "object" ? value as { adHocRecords?: RecordItem[]; confirmedSnapshots?: Record<string, RecordItem> } : {};
              const localWeek = localWeeks[week] || {};
              const localSnapshots = localWeek.confirmedSnapshots || {};
              const localAdHoc = new Map<string, RecordItem>((localWeek.adHocRecords || []).map((record: RecordItem) => [record.id, record]));
              mergedWeeks[week] = {
                ...localWeek,
                ...cloudWeek,
                adHocRecords: (cloudWeek.adHocRecords || []).map(record => ({ ...localAdHoc.get(record.id), ...record, photos: localAdHoc.get(record.id)?.photos || [] })),
                confirmedSnapshots: Object.fromEntries(Object.entries(cloudWeek.confirmedSnapshots || {}).map(([id, record]) => {
                  const local = localSnapshots[id] || records.find(item => item.id === id);
                  return [id, { ...local, ...record, photos: local?.photos || [] }];
                })),
              };
            });
            localStorage.setItem(PPT_WEEK_SELECTIONS_KEY, JSON.stringify(mergedWeeks));
            const current = mergedWeeks[pptWeekStart] || {};
            setPptExtraIds(Array.isArray(current.extraIds) ? current.extraIds : []);
            setPptOrderIds(Array.isArray(current.orderIds) ? current.orderIds : []);
            setPptAdHocRecords(Array.isArray(current.adHocRecords) ? current.adHocRecords : []);
            setPptConfirmedSnapshots(current.confirmedSnapshots && typeof current.confirmedSnapshots === "object" ? current.confirmedSnapshots : {});
          } catch {}
        } else {
          window.setTimeout(() => { void supabasePush(true); }, 0);
        }
        flash(automatic ? "已自動同步最新雲端資料" : "雲端資料已合併到本機");
      }
    } catch { flash(automatic ? "自動同步失敗，請檢查雲端登入" : "雲端讀取失敗，請檢查登入與設定"); }
  };
  const supabaseSignIn = async (email: string, password: string, signUp = false) => {
    if (!settings.supabaseKey) return flash("請先貼上 Supabase Publishable key");
    try {
      const endpoint = signUp ? "/auth/v1/signup" : "/auth/v1/token?grant_type=password";
      const res = await fetch(`${settings.supabaseUrl.replace(/\/$/, "")}${endpoint}`, { method: "POST", headers: { apikey: settings.supabaseKey, "Content-Type": "application/json" }, body: JSON.stringify({ email, password }) });
      const data = await res.json(); if (!res.ok) throw new Error(data?.message || data?.error_description || "");
      if (!data.access_token) { flash("註冊完成，請先到 Email 收信完成驗證後再登入"); return; }
      const session = { accessToken: data.access_token, refreshToken: data.refresh_token, email: data.user?.email || email };
      setCloudSession(session); localStorage.setItem(CLOUD_SESSION_KEY, JSON.stringify(session)); flash("雲端帳號登入完成");
    } catch { flash(signUp ? "註冊失敗，請檢查 Email 與密碼" : "登入失敗，請檢查 Email 與密碼"); }
  };
  const supabaseSignOut = () => { setCloudSession(null); localStorage.removeItem(CLOUD_SESSION_KEY); flash("已登出雲端帳號"); };
  const openMonthlyProgress = async () => {
    setMonthlyProgressOpen(true);
    if (!cloudSession?.accessToken) return;
    try {
      const session = await refreshCloudSession();
      if (!session) return;
      const response = await fetch(`${CASE_FILE_SUPABASE_URL}/rest/v1/case_file_front_access?select=person_id,last_entered_at`, { headers: cloudHeaders(session) });
      const rows = await response.json();
      if (response.ok && Array.isArray(rows)) setFrontLastLogins(Object.fromEntries(rows.map((row: { person_id: string; last_entered_at: string }) => [row.person_id, row.last_entered_at])));
    } catch {}
  };
  const cloudSnapshot = JSON.stringify({ records, personnel: settings.personnel, expiry591: settings.expiry591, expiry5168: settings.expiry5168, brokerExpiry: settings.brokerExpiry, intakeRaw, intakeDrafts, selectedIntakeId, tourDate, tourTitle, tourItems, tourModifiedAt, tourHistory, pptWeekStart, pptExtraIds, pptOrderIds, pptAdHocRecords, pptConfirmedSnapshots });
  useEffect(() => {
    if (!cloudSyncBaselineRef.current) { cloudSyncBaselineRef.current = cloudSnapshot; return; }
    if (cloudSyncBaselineRef.current === cloudSnapshot) return;
    cloudSyncBaselineRef.current = cloudSnapshot;
    if (!cloudSession?.accessToken || !settings.supabaseKey) return;
    if (cloudSyncTimerRef.current) window.clearTimeout(cloudSyncTimerRef.current);
    cloudSyncTimerRef.current = window.setTimeout(() => { void supabasePush(true); }, 6000);
    return () => { if (cloudSyncTimerRef.current) window.clearTimeout(cloudSyncTimerRef.current); };
  }, [cloudSnapshot, cloudSession?.accessToken, settings.supabaseKey]);
  useEffect(() => {
    if (!cloudSession?.accessToken || !settings.supabaseUrl || !settings.supabaseKey) return;
    const pullKey = `${settings.supabaseUrl}|${settings.supabaseTable}|${settings.supabaseRecord}|${cloudSession.email || "signed-in"}`;
    if (cloudAutoPullRef.current === pullKey) return;
    cloudAutoPullRef.current = pullKey;
    if (localStorage.getItem("property-desk-import-prefer-local-once") === "1") { localStorage.removeItem("property-desk-import-prefer-local-once"); return; }
    void supabasePull(true);
  }, [cloudSession?.accessToken, cloudSession?.email, settings.supabaseUrl, settings.supabaseKey, settings.supabaseTable, settings.supabaseRecord]);

  const openPublic = () => { setTab("public"); setPublicUnlocked(false); setPublicPersonId(""); setPublicScope("mine"); setPassword(""); };
  const logoutPublic = () => { localStorage.removeItem("case-file-public-daily-login"); setPublicUnlocked(false); setPublicPersonId(""); setPublicScope("mine"); setPublicExpiryFilter("all"); setPublicQuery(""); setPassword(""); flash("已登出業務帳號"); };
  const rememberPublicLogin = (personId: string) => localStorage.setItem("case-file-public-daily-login", JSON.stringify({ date: today(), personId }));
  const recordPublicEntry = async (nationalId: string) => {
    try { await fetch(`${CASE_FILE_SUPABASE_URL}/rest/v1/rpc/case_file_front_touch`, { method: "POST", headers: { apikey: CASE_FILE_SUPABASE_PUBLISHABLE_KEY, "Content-Type": "application/json" }, body: JSON.stringify({ p_national_id: nationalId }) }); } catch {}
  };
  useEffect(() => {
    if (!(internalView || tab === "public") || !publicUnlocked) return;
    const header = document.querySelector<HTMLElement>(".public-head");
    if (!header || header.querySelector(".public-logout")) return;
    const button = document.createElement("button");
    button.type = "button";
    button.className = "public-logout";
    button.textContent = "登出";
    button.addEventListener("click", logoutPublic);
    header.firstElementChild?.append(button);
    return () => { button.removeEventListener("click", logoutPublic); button.remove(); };
  }, [internalView, tab, publicUnlocked]);
  const unlock = async () => {
    const normalizeLoginId = (value = "") => value.trim().replace(/\s+/g, "").toUpperCase();
    const activePeople = settings.personnel.filter(p => (p.status || "在職") === "在職" && normalizeLoginId(p.nationalId));
    const loginId = normalizeLoginId(password);
    const person = activePeople.find(p => loginId === normalizeLoginId(p.nationalId));
    if (person) { setPublicPersonId(person.id); setPublicScope("mine"); setPublicUnlocked(true); rememberPublicLogin(person.id); setPassword(""); void recordPublicEntry(loginId); return; }
    if (!loginId) return flash("請輸入身分證字號");
    try {
      const response = await fetch(`${CASE_FILE_SUPABASE_URL}/rest/v1/rpc/case_file_front_login`, { method: "POST", headers: { apikey: CASE_FILE_SUPABASE_PUBLISHABLE_KEY, "Content-Type": "application/json" }, body: JSON.stringify({ p_national_id: loginId }) });
      const data = await response.json();
      if (!response.ok || !data?.personId || !Array.isArray(data.records)) return flash("身分證字號錯誤，或人員目前不是在職");
      const nextRecords = applySourceLayoutFixes(data.records).map((record: RecordItem) => normalizeRecordPings({ ...record, photos: [] }));
      const nextPeople = (Array.isArray(data.personnel) ? data.personnel : []).map((entry: Partial<Person>) => ({ ...entry, id: entry.id || newId(), name: entry.name || "", nationalId: "", status: entry.status || "在職" })) as Person[];
      setRecords(nextRecords);
      setSettings(previous => ({ ...previous, personnel: nextPeople }));
      setPublicPersonId(data.personId);
      setPublicScope("mine");
      setPublicUnlocked(true);
      rememberPublicLogin(data.personId);
      setPassword("");
    } catch { flash("目前無法連接雲端，請確認網路後再試一次"); }
  };
  const publicPerson = settings.personnel.find(p => p.id === publicPersonId);
  const contactPeople = sortPeopleBySequence(settings.personnel.filter(person => person.status === "在職" && person.name.trim() && String(person.phone || "").trim()));
  const myProperties = active.filter(r => nameMatches(r.developer, publicPerson?.name || ""));
  const expiryAlerts = myProperties.filter(r => daysUntil(r.entrustEnd) >= 0 && daysUntil(r.entrustEnd) <= 30);
  const publicBaseRecords = publicScope === "mine" ? myProperties.map(record => ({ ...record, _showEntrustEndInPropertyNo: "1" })) : active;
  const publicExpiryRecords = publicScope !== "mine" || publicExpiryFilter === "all"
    ? publicBaseRecords
    : publicBaseRecords.filter(record => {
        const remaining = daysUntil(record.entrustEnd);
        return remaining >= 0 && remaining <= Number(publicExpiryFilter);
      });
  const normalizedPublicQuery = publicQuery.trim().toLowerCase();
  const publicTableRecords = !normalizedPublicQuery
    ? publicExpiryRecords
    : publicExpiryRecords.filter(record => [record.propertyNo, record.area, record.caseName, record.address, developerFullNameText(record.developer)]
        .some(value => String(value || "").toLowerCase().includes(normalizedPublicQuery)));
  const publicExpiry15Count = publicBaseRecords.filter(record => {
    const remaining = daysUntil(record.entrustEnd);
    return remaining >= 0 && remaining <= 15;
  }).length;
  const publicExpiry30Count = publicBaseRecords.filter(record => {
    const remaining = daysUntil(record.entrustEnd);
    return remaining >= 0 && remaining <= 30;
  }).length;
  const intakeDraft = intakeDrafts.find(d => d.id === selectedIntakeId) || null;
  const analyzeIntake = () => {
    const parsed = parseIntakes(intakeRaw); if (!parsed.length) return flash("無法辨識資料，請貼上業務回應的一列資料");
    const duplicateNames = parsed.flatMap(item => {
      const no = intakeValue(item.values, "委託主約編號").trim(); const caseName = intakeValue(item.values, "案名").trim(); const address = intakeValue(item.values, "物件(完整)地址").trim();
      const sameDraft = intakeDrafts.some(draft => no ? intakeValue(draft.values, "委託主約編號").trim() === no : !!caseName && !!address && intakeValue(draft.values, "案名").trim() === caseName && intakeValue(draft.values, "物件(完整)地址").trim() === address);
      const sameRecord = records.some(record => no ? record.propertyNo.trim() === no : !!caseName && !!address && record.caseName.trim() === caseName && record.address.trim() === address);
      return sameDraft || sameRecord ? [no || caseName || "未命名案件"] : [];
    });
    if (duplicateNames.length && !confirm(`發現相同案件：${duplicateNames.join("、")}\n資料已存在草稿或總表，仍要再次加入嗎？`)) return;
    const stamp = new Date().toISOString();
    const saved = parsed.map(item => ({ ...item, raw: intakeRaw, modifiedAt: stamp })); setIntakeDrafts(prev => [...saved.reverse(), ...prev]); setSelectedIntakeId(""); selectedIntakeRef.current = ""; setIntakeRaw(""); flash(`已新增 ${saved.length} 筆進案草稿`);
  };
  const addManualIntakeDraft = () => {
    const createdAt = new Date().toISOString();
    const draft: IntakeData = { id: newId(), values: {}, propertyKind: "房屋", createdAt, modifiedAt: createdAt, raw: "手動新增例外案件" };
    setIntakeDrafts(previous => [draft, ...previous]);
    selectedIntakeRef.current = draft.id;
    setSelectedIntakeId(draft.id);
    flash("已建立例外案件草稿，確認後再按正式進案");
  };
  const updateIntakeValue = (needle: string, value: string) => {
    const target = intakeDrafts.find(draft => draft.id === selectedIntakeId); if (!target) return;
    const values = { ...target.values }; const key = Object.keys(values).find(k => k.includes(needle)) || needle; values[key] = value;
    const updated: IntakeData = { ...target, values, modifiedAt: new Date().toISOString(), propertyKind: needle === "物件型態" ? (value.includes("土地") ? "純土地" : "房屋") : target.propertyKind };
    setIntakeDrafts(prev => prev.map(draft => draft.id === updated.id ? updated : draft));
    if (updated.linkedRecordId) setRecords(prev => prev.map(record => record.id === updated.linkedRecordId ? normalizeRecordPings(intakeToRecord(updated, record)) : record));
  };
  const selectIntakeDraft = (id: string) => { selectedIntakeRef.current = id; setSelectedIntakeId(id); };
  const openRecordIntakeDraft = (record: RecordItem) => {
    const linked = intakeDrafts.find(draft => draft.linkedRecordId === record.id);
    const baseDraft = recordToIntake(record);
    const draft = linked || { ...baseDraft, values: syncRecordToDraftValues(baseDraft, record), linkedRecordId: record.id, enteredAt: new Date().toISOString() };
    if (!linked) setIntakeDrafts(previous => [draft, ...previous]);
    setEditing(null);
    setTab("intake");
    selectIntakeDraft(draft.id);
    flash(linked ? "已開啟原本的進案草稿；修改會同步總表" : "已建立連結進案草稿；修改會同步總表");
  };
  const markIntakeDraftPrintedForSales = (id: string) => {
    const target = intakeDrafts.find(draft => draft.id === id);
    if (!target) return;
    const printedForSalesAt = target.printedForSalesAt ? undefined : new Date().toISOString();
    setIntakeDrafts(previous => previous.map(draft => draft.id === id ? { ...draft, printedForSalesAt, modifiedAt: new Date().toISOString() } : draft));
    flash(printedForSalesAt ? `已記錄「${intakeValue(target.values, "案名") || "未命名案件"}」列印草稿` : "已取消列印草稿紀錄");
  };
  const confirmIntake = (draftId?: string) => { const targetId = draftId || selectedIntakeRef.current; const target = targetId ? intakeDrafts.find(d => d.id === targetId) : intakeDraft; if (!target) return; const targetNo = intakeValue(target.values, "委託主約編號"); const existing = target.linkedRecordId ? records.find(record => record.id === target.linkedRecordId) : records.find(record => !!targetNo && record.propertyNo === targetNo); const record = intakeToRecord(target, existing); if (!record.propertyNo || !record.caseName) return flash("缺少物件編號或案名，請先確認表單內容"); if (existing) { const firstFormalEntry = !target.enteredAt; const tracked = { ...withTrackedUpdate(existing, record), ...(firstFormalEntry ? { _newCaseReminderEnabled: "1", _newCaseReminderSource: "intake" } : {}) }; const enteredAt = target.enteredAt || new Date().toISOString(); setRecords(prev => prev.map(item => item.id === existing.id ? tracked : item)); setIntakeDrafts(prev => prev.map(draft => draft.id === target.id ? { ...draft, linkedRecordId: existing.id, enteredAt, modifiedAt: enteredAt } : draft)); if (firstFormalEntry) { setNewCaseReminder({ ...tracked }); setNewCaseReminderBatchIds(previous => [...new Set([...previous, ...pendingIntakeReminderRecords.map(item => item.id), tracked.id])]); } flash(firstFormalEntry ? "已正式進案；請完成新案件提醒" : "已連結並同步更新原總表資料"); return; } if (!confirm(`確定文件已收到，將「${record.caseName}」正式加入總表？`)) return; const enteredAt = new Date().toISOString(); const entered = { ...record, _newCaseReminderEnabled: "1", _newCaseReminderSource: "intake" }; setRecords(prev => [entered, ...prev]); setIntakeDrafts(prev => prev.map(draft => draft.id === target.id ? { ...draft, linkedRecordId: entered.id, enteredAt, modifiedAt: enteredAt } : draft)); setNewCaseReminder({ ...entered }); setNewCaseReminderBatchIds(previous => [...new Set([...previous, ...pendingIntakeReminderRecords.map(item => item.id), entered.id])]); flash("已正式進案；請完成新案件提醒"); };
  const removeIntakeDraft = (id: string) => { const target = intakeDrafts.find(d => d.id === id); if (!target || !confirm(`確定刪除「${intakeValue(target.values, "案名") || "未命名草稿"}」？`)) return; setIntakeDrafts(prev => prev.filter(d => d.id !== id)); if (selectedIntakeId === id) setSelectedIntakeId(""); };
  const updateIntakeDraftCaseName = (id: string, caseName: string) => setIntakeDrafts(previous => previous.map(draft => { if (draft.id !== id) return draft; const key = Object.keys(draft.values).find(name => name.includes("案名")) || "案名"; return { ...draft, values: { ...draft.values, [key]: caseName } }; }));
  const submitMonthlyPropertyReport = (record: RecordItem, status: string, reason: string) => {
    if (!publicPerson) return;
    if (status === "下架洽開發" && !reason.trim()) return flash("下架洽開發必須填寫原因");
    const reportKey = `${today().slice(0, 7)}:${publicPerson.id}`; let reports: Record<string, any> = {}; try { reports = JSON.parse(record._monthlyReports || "{}"); } catch {}
    const due = new Date(`${today()}T00:00:00`); due.setDate(due.getDate() + 7); const dueDate = `${due.getFullYear()}-${String(due.getMonth() + 1).padStart(2, "0")}-${String(due.getDate()).padStart(2, "0")}`;
    const reportedAt = new Date().toISOString();
    reports[reportKey] = { personId: publicPerson.id, personName: publicPerson.name, status, reason: reason.trim(), reportedAt, dueDate: status === "待確認" ? dueDate : "" };
    // A second developer's active confirmation completes the first developer's
    // request automatically, so management does not need to clear it manually.
    if (status === "委託中") Object.entries(reports).forEach(([key, report]: [string, any]) => {
      if (key !== reportKey && report?.status === "請跟開發業務2確認" && !report.adminHandledAt && report.personId !== publicPerson.id) {
        reports[key] = { ...report, adminHandledAt: reportedAt, autoResolvedAt: reportedAt, autoResolvedBy: publicPerson.name };
      }
    });
    const updatesDate = status === "委託中";
    // 業務前台不使用管理者帳號；回報時只寫入這一筆確認紀錄，讓手機與電腦讀到同一份資料。
    void fetch(`${CASE_FILE_SUPABASE_URL}/rest/v1/rpc/case_file_front_report`, {
      method: "POST",
      headers: { apikey: CASE_FILE_SUPABASE_PUBLISHABLE_KEY, "Content-Type": "application/json" },
      body: JSON.stringify({ p_person_id: publicPerson.id, p_record_id: record.id, p_status: status, p_reason: reason.trim() })
    }).then(response => {
      if (!response.ok) throw new Error("front report sync failed");
    }).catch(() => flash("回報暫存於這台裝置；雲端同步失敗，請確認網路後再送出"));
    setRecords(previous => previous.map(item => item.id === record.id ? { ...item, _monthlyReports: JSON.stringify(reports), ...(updatesDate ? { updateDate: today(), lastModifiedAt: new Date().toISOString() } : {}) } : item));
    flash(status === "待確認" ? `已設定 ${displayRocDate(dueDate)} 再次確認` : "物件回報完成");
  };
  const resolveMonthlyPropertyReport = (record: RecordItem, reportKey: string, keepActive = false) => {
    const handledAt = new Date().toISOString();
    setRecords(previous => previous.map(item => {
      if (item.id !== record.id) return item;
      let reports: Record<string, any> = {}; try { reports = JSON.parse(item._monthlyReports || "{}"); } catch {}
      if (!reports[reportKey]) return item;
      const reportedAtValue = String(reports[reportKey].reportedAt || "");
      const reportedAtDate = reportedAtValue ? new Date(reportedAtValue) : new Date();
      const reportedDate = Number.isNaN(reportedAtDate.getTime()) ? today() : isoLocalDate(reportedAtDate);
      reports[reportKey] = { ...reports[reportKey], adminHandledAt: handledAt };
      return { ...item, _monthlyReports: JSON.stringify(reports), ...(keepActive ? { updateDate: reportedDate, lastModifiedAt: handledAt } : {}) };
    }));
    flash(keepActive ? "已確認維持委託中，更新日期已更新" : "已標示處理完成");
  };
  const submitBookReviews = (reviews: Array<{ record: RecordItem; status: string }>) => { const stamp = new Date().toISOString(), cycle = bookReviewCycleKey(); setRecords(previous => previous.map(item => reviews.some(value => value.record.id === item.id) ? { ...item, bookLocationDate: today(), _bookReviewAt: stamp, _bookReviewCycle: cycle, lastModifiedAt: stamp } : item)); flash(`已確認 ${reviews.length} 筆物件本，日期更新為今天`); };
  const updateTourDate = (date: string) => { setTourDate(date); setTourModifiedAt(new Date().toISOString()); };
  const updateTourTitle = (title: string) => { setTourTitle(title); setTourModifiedAt(new Date().toISOString()); };
  const updateTourItems = (items: TourItem[]) => { setTourItems(items); setTourModifiedAt(new Date().toISOString()); };

  const showingPublic = internalView || tab === "public";

  return <main lang="en-GB" className={internalView ? "internal-public-app" : ""}>
    {!internalView && <header className="topbar">
      <div className="topbar-row"><div className="brand"><h1>總表　管理模式 <small className="app-version">V204</small></h1></div>
      <div className="header-actions"><button className="action-monthly-progress" onClick={() => void openMonthlyProgress()}>45天確認進度</button>{pendingDealCompletion.length > 0 && <button className="deal-reminder-header-button" onClick={() => setDealCompletionReminderOpen(true)}>成交後續提醒 {pendingDealCompletion.length}</button>}{pendingArchiveCleanup.length > 0 && <button className="archive-reminder-header-button" onClick={() => setArchiveCleanupReminderOpen(true)}>下架提醒 {pendingArchiveCleanup.length}</button>}{bookReviewDueCount > 0 && <button className="book-review-header-button action-book-review" onClick={() => { setTab("active"); setBookReviewOpenRequest(value => value + 1); }}>物件本確認 {bookReviewDueCount}</button>}<button className="ppt-export-button action-ppt" onClick={() => { setPptShowExtras(false); setPptPickerOpen(true); }}>產生 PPT</button><button className="action-excel" onClick={exportExcel}>匯出 Excel</button><label className="file-button action-import-json">匯入 JSON<input type="file" accept=".json,application/json" onChange={importJson}/></label><button className="action-export-json" onClick={exportJson}>匯出 JSON</button><button className="key-tag action-keys" onClick={() => setTab("keys")}>🔑 鑰匙總表 <b>{controlledKeyCount}</b></button></div></div>
      <nav className="nav">
      <button className={tab === "active" ? "active" : ""} onClick={() => setTab("active")}>委託中 <span>{active.length}</span></button>
      <button className={tab === "archive" ? "active" : ""} onClick={() => setTab("archive")}>封存{pendingArchiveCleanup.length > 0 && <small className="archive-pending-count">待下架 {pendingArchiveCleanup.length}</small>}</button>
      <button className={tab === "activity" ? "active" : ""} onClick={() => setTab("activity")}>每日物件動態</button>
      <button className={tab === "inventory" ? "active" : ""} onClick={() => setTab("inventory")}>物件庫存</button>
      <button className={tab === "tour" ? "active" : ""} onClick={() => setTab("tour")}>團看安排</button>
      <button className={tab === "intake" ? "active" : ""} onClick={() => { setTab("intake"); selectIntakeDraft(""); }}>進案草稿{intakeDrafts.filter(draft => !draft.enteredAt).length > 0 && <b className="intake-draft-nav-count">{intakeDrafts.filter(draft => !draft.enteredAt).length}</b>}</button>
      <button className={tab === "public" ? "active" : ""} onClick={openPublic}>前台總表</button>
      <button className={tab === "settings" ? "active" : ""} onClick={() => setTab("settings")}>設定</button>
      <span className="home-last-modified home-sync-times"><span className="local-modified-line"><b>最後修改:</b><em>{latestModifiedAt ? displayHomeModifiedAt(latestModifiedAt) : "尚無紀錄"}</em></span><span className="cloud-upload-line"><b>Supabase上傳:</b><em>{cloudLastUploadAt ? displayHomeModifiedAt(cloudLastUploadAt) : "尚無紀錄"}<i className={`cloud-upload-state ${cloudUploadStatus === "上傳完成" ? "complete" : cloudUploadStatus === "ING" ? "uploading" : "signed-out"}`}>{cloudUploadStatus}</i></em></span></span>
      </nav>
    </header>}

    {!internalView && tab === "active" && systemExpiryReminders.length > 0 && <div className="system-expiry-reminder"><b>系統到期提醒</b>{systemExpiryReminders.map(item => <span key={item.label}>{item.label}：{displayRocDate(item.date)} 到期{item.date < today() ? "（已到期）" : item.leadDays === 30 ? "（30天內）" : "（明天到期）"}</span>)}</div>}
    {!internalView && tab === "active" && <PropertyBookReview records={records} settings={settings} openRequest={bookReviewOpenRequest} submit={submitBookReviews} openRecord={setEditing}/>}
    {monthlyProgressOpen && <MonthlyConfirmationProgress records={active} people={settings.personnel} lastLogins={frontLastLogins} close={() => setMonthlyProgressOpen(false)}/>}
    {!internalView && tab === "active" && <BusinessReportInbox records={active} resolve={resolveMonthlyPropertyReport} archive={(record, status) => setArchiveChoice({ record, status, date: today(), salesPerson: record.salesPerson || "", reason: record.archiveReason || "" })}/>}
    {showingPublic && publicUnlocked && publicScope === "mine" && publicPerson && <MonthlyPropertyReport records={myProperties} person={publicPerson} submit={submitMonthlyPropertyReport}/>}
    {pptPickerOpen && selectedPptBaseRecords.length > 0 && <aside className="ppt-order-float"><b>調整本次排序</b><small>此順序會保留在本週</small><ol>{selectedPptBaseRecords.map((record, index) => <li key={record.id}><span>{index + 1}. {record.caseName || "未命名案件"}</span><div><button type="button" disabled={index === 0} onClick={() => movePptOrder(record.id, -1)}>↑</button><button type="button" disabled={index === selectedPptBaseRecords.length - 1} onClick={() => movePptOrder(record.id, 1)}>↓</button></div></li>)}</ol></aside>}

    {tab === "settings" ? <SettingsPanel settings={settings} setSettings={setSettings} supabasePush={supabasePush} supabasePull={supabasePull} cloudSession={cloudSession} supabaseSignIn={supabaseSignIn} supabaseSignOut={supabaseSignOut} /> :
    tab === "intake" ? <IntakePanel raw={intakeRaw} setRaw={setIntakeRaw} drafts={intakeDrafts} draft={intakeDraft} selectDraft={selectIntakeDraft} deleteDraft={removeIntakeDraft} analyze={analyzeIntake} addManualDraft={addManualIntakeDraft} updateValue={updateIntakeValue} clear={() => setIntakeRaw("")} confirmIntake={confirmIntake} markPrintedForSales={markIntakeDraftPrintedForSales} /> :
    tab === "tour" ? <TourPlanner records={records} drafts={intakeDrafts} items={tourItems} setItems={updateTourItems} history={tourHistory} setHistory={setTourHistory} editRecord={setEditing} updateDraftCaseName={updateIntakeDraftCaseName} tourDate={tourDate} setTourDate={updateTourDate} tourTitle={tourTitle} setTourTitle={updateTourTitle} notify={flash} complete={(date, recordIds, draftIds) => { setRecords(previous => previous.map(record => recordIds.includes(record.id) ? withTrackedUpdate(record, { ...record, groupViewDate: date, updateDate: today() }) : record)); setIntakeDrafts(previous => previous.map(draft => draftIds.includes(draft.id) ? { ...draft, groupViewDate: date } : draft)); updateTourItems([]); flash("團看日期已同步到物件與草稿"); }} /> :
    tab === "activity" ? <DailyActivity records={records} onEdit={setEditing} /> :
    tab === "inventory" ? <BusinessInventory records={records} people={settings.personnel} /> :
    tab === "keys" ? <KeySummary records={records}/> :
    showingPublic ? <section className="public-shell">{!publicUnlocked ? <div className="login-card"><h2>內部總表</h2><p>請輸入業務人員身分證字號進入</p><input type="password" value={password} onChange={e => setPassword(e.target.value)} onKeyDown={e => e.key === "Enter" && unlock()} placeholder="請輸入密碼"/><button className="primary wide" onClick={unlock}>進入內部總表</button><small>連城不動產開發有限公司物件資料屬公司重要資訊，僅限內部使用；未經授權，禁止外傳、外流、轉載或提供他人使用。</small></div> : <><div className="public-head"><div><p className="logged-person">登錄人員：{publicPerson?.name || "業務人員"}</p><h2>{publicScope === "activity" ? "每日物件動態" : publicScope === "mine" ? "我的物件" : publicScope === "contacts" ? "通訊錄" : "物件總表"}</h2><div className="scope-tabs"><button className={publicScope === "activity" ? "selected" : ""} onClick={() => setPublicScope("activity")}>每日物件動態</button><button className={publicScope === "mine" ? "selected" : ""} onClick={() => setPublicScope("mine")}>我的物件 {myProperties.length}</button><button className={publicScope === "all" ? "selected" : ""} onClick={() => setPublicScope("all")}>物件總表 {active.length}</button><button className={publicScope === "contacts" ? "selected" : ""} onClick={() => setPublicScope("contacts")}>通訊錄</button></div></div></div>{publicScope === "mine" && <div className="public-expiry-filters"><button className={publicExpiryFilter === "all" ? "selected" : ""} onClick={() => setPublicExpiryFilter("all")}>全部</button><button className={publicExpiryFilter === "15" ? "selected" : ""} onClick={() => setPublicExpiryFilter("15")}>15天內到期 {publicExpiry15Count}</button><button className={publicExpiryFilter === "30" ? "selected" : ""} onClick={() => setPublicExpiryFilter("30")}>30天內到期 {publicExpiry30Count}</button></div>}{(publicScope === "mine" || publicScope === "all") && <div className="public-search-tools"><label className="search">⌕<input value={publicQuery} onChange={event => setPublicQuery(event.target.value)} placeholder="搜尋編號、地區、案名、地址、開發…"/></label><button type="button" onClick={() => { setPublicQuery(""); setPublicExpiryFilter("all"); }}>清除篩選</button><div className="public-zoom-tools"><button type="button" onClick={() => setPublicZoom(value => Math.max(25, value - 10))}>－縮小</button><b>{publicZoom}%</b><button type="button" onClick={() => setPublicZoom(value => Math.min(140, value + 10))}>＋放大</button></div></div>}{publicScope === "activity" ? <DailyActivity records={records} compact/> : publicScope === "contacts" ? <ContactDirectory people={contactPeople}/> : <>{publicScope === "mine" && expiryAlerts.length > 0 && <div className="expiry-alert"><b>委託到期提醒</b>{expiryAlerts.map(r => <span key={r.id}>{r.caseName}：還有 {daysUntil(r.entrustEnd)} 天到期{daysUntil(r.entrustEnd) <= 15 ? "（15天內）" : "（30天內）"}</span>)}</div>}<PropertyTable records={publicTableRecords} columns={publicColumns} publicMode zoom={publicZoom} expiryAnnotation={publicScope === "mine"} onEdit={() => {}} onArchive={() => {}} onRestore={() => {}} onRemove={() => {}}/></>}</>}</section> :
    <section className={`content ${tab === "archive" ? "archive-list-page" : "active-list-page"}`}>
      <div className="list-head active-list-toolbar"><div className="list-title-search"><SectionTitle title={tab === "archive" ? "封存物件" : "委託中物件"} subtitle={tab === "archive" ? "到期與已下架物件紀錄" : ""}/><label className="search">⌕<input value={currentListQuery} onChange={e => tab === "archive" ? setArchiveQuery(e.target.value) : setQuery(e.target.value)} placeholder="搜尋編號、地區、案名、地址…"/></label>{tab === "archive" && <button className="clear-filter-button" onClick={() => setArchiveQuery("")}>清除篩選</button>}</div><div className="tools">{tab === "active" && <><button className="showing-follow-up-filter" onClick={() => setShowingFollowUpOpen(true)}>暫停帶看 {showingFollowUpRecords.length}</button><button className={missingDataRecords.length ? "missing-data-button has-items" : "missing-data-button"} onClick={() => setMissingDataReminderOpen(true)}>待補資料 {missingDataRecords.length}</button><select className="website-filter" value={websiteFilter} onChange={event => setWebsiteFilter(event.target.value)}><option value="">全部物件</option><option value="missing_sales_book_date">銷售本缺日期</option><option value="missing_book_location_date">物件本位置缺日期</option><option value="missing_photo_info">照片缺填寫</option><option value="all">全部尚未 PO</option>{Object.entries(websiteFieldMap).map(([key, name]) => <option key={key} value={key}>{name} 尚未 PO</option>)}</select><button className="advanced-filter-button" onClick={() => setAdvancedFilterOpen(true)}>進階篩選</button><button className="clear-filter-button" onClick={() => { setShowingFollowUpOnly(false); setWebsiteFilter(""); setAdvancedFilter(blankAdvancedFilter()); setQuery(""); }}>清除篩選</button></>}</div></div>
      {tab === "active" && websiteExpiryRecords.length > 0 && <div className="website-expiry-reminder"><b>網站廣告到期提醒</b>{websiteExpiryRecords.map(({ record, site, date }) => <button key={`${record.id}-${site}`} onClick={() => setEditing(record)}>{record.caseName || record.propertyNo}：{site} 已於 {displayRocDate(date)} 到期</button>)}</div>}
      <PropertyTable records={tab === "archive" ? shown.map(archiveDisplayRecord) : shown.map(showingFollowUpDisplayRecord).map((record, index) => ({ ...record, displaySequence: String(index + 1) }))} columns={tab === "archive" ? archiveColumns : activeColumns} archiveMode={tab === "archive"} activeLead={tab === "active"} showAreaGroups={tab === "active"} onEdit={setEditing} onArchive={archiveRecord} onRestore={setRestoreChoiceRecord} onRemove={removeRecord}/>
      {advancedFilterOpen && <AdvancedFilterModal value={advancedFilter} setValue={setAdvancedFilter} onClose={() => setAdvancedFilterOpen(false)} onReset={() => setAdvancedFilter(blankAdvancedFilter())}/>}
      {!shown.length && <Empty text={query || tab === "active" && Object.values(advancedFilter).some(Boolean) ? "找不到符合的物件" : tab === "archive" ? "目前沒有封存物件" : "目前沒有委託中物件"}/>}
    </section>}

    {showingFollowUpOpen && tab === "active" && <div className="modal-backdrop"><div className="modal showing-follow-up-modal"><div className="modal-head"><div><span>委託中案件</span><h2>帶看追蹤 {showingFollowUpRecords.length}</h2></div><button className="close" onClick={() => setShowingFollowUpOpen(false)}>×</button></div><div className="showing-follow-up-table"><div className="showing-follow-up-heading"><b>案件</b><b>開發</b><b>追蹤日期</b><b>紀錄</b><b>操作</b></div>{sortActiveRecords(showingFollowUpRecords).map(record => <div className="showing-follow-up-row" key={record.id}><section><b>{record.caseName || record.propertyNo}</b><small>{record.propertyNo}　{record.address}</small></section><span className="showing-follow-up-developer">{developerFullNameText(record.developer) || "—"}</span><strong>{displayRocDate(record.showingFollowUpDueDate) || "—"}</strong><input className="showing-follow-up-note" type="text" value={record.showingFollowUpNote || ""} onChange={event => setRecords(previous => previous.map(item => item.id === record.id ? { ...item, showingFollowUpNote: event.target.value, lastModifiedAt: new Date().toISOString() } : item))} placeholder="輸入追蹤紀錄"/><div><button onClick={() => { setShowingFollowUpOpen(false); setEditing(record); }}>編輯</button><button onClick={() => clearShowingFollowUp(record)}>解除</button></div></div>)}{!showingFollowUpRecords.length && <p className="showing-follow-up-empty">目前沒有帶看追蹤案件</p>}</div><div className="modal-foot"><span>紀錄為帶看追蹤專用，不會帶入物件備註欄</span><button onClick={() => setShowingFollowUpOpen(false)}>完成</button></div></div></div>}

    {dealCompletionReminderOpen && tab !== "public" && pendingDealCompletion.length > 0 && !expiryReminderOpen && !websitePoReminderOpen && !missingDataReminderOpen && !archiveCleanupReminderOpen && <div className="modal-backdrop archive-cleanup-backdrop"><div className="modal archive-cleanup-modal"><div className="modal-head"><div><span>成交後續提醒</span><h2>有 {pendingDealCompletion.length} 筆成交案件尚未完成後續作業</h2></div><button className="close" onClick={() => setDealCompletionReminderOpen(false)}>×</button></div><div className="archive-cleanup-table-wrap"><table className="archive-cleanup-table deal-completion-table"><thead><tr><th>案件</th><th>開發／銷售</th>{dealCompletionTasks.map(([, label]) => <th key={label}>{label}</th>)}</tr></thead><tbody>{pendingDealCompletion.map(({ record, tasks }) => <tr key={record.id}><td className="archive-cleanup-case-cell"><b>案名：{stripRestoredDisplay(record.caseName) || "未命名案件"}</b><small>{record.propertyNo || "—"}　{displayRocDate(record.archived || "")}成交下架</small></td><td className="deal-completion-people"><span>開發：{developerFullNameText(record.developer, settings.personnel) || "—"}</span><span>銷售：{record.salesPerson || "—"}</span></td>{dealCompletionTasks.map(([key]) => <td key={key}>{tasks.some(([pendingKey]) => pendingKey === key) ? <button onClick={() => completeDealTask(record, key)}>完成</button> : <span className="archive-cleanup-done">已完成</span>}</td>)}</tr>)}</tbody></table></div><div className="modal-foot"><span>五項全部完成前，下次開啟系統仍會提醒</span><button onClick={() => setDealCompletionReminderOpen(false)}>稍後處理</button></div></div></div>}

    {missingDataReminderOpen && tab !== "public" && !expiryReminderOpen && <div className="modal-backdrop missing-data-reminder-backdrop"><div className="modal expiry-reminder-modal missing-data-reminder-modal"><div className="modal-head"><div><span>每日待補資料提醒</span><h2>{missingDataRecords.length ? `有 ${missingDataRecords.length} 筆物件尚待補資料` : "目前沒有待補資料"}</h2></div><button className="close" onClick={() => setMissingDataReminderOpen(false)}>×</button></div><div className="missing-data-table-wrap">{missingDataRecords.length ? <table className="missing-data-table"><thead><tr><th>物件編號</th><th>案名</th><th>地址</th><th>缺</th></tr></thead><tbody>{missingDataRecords.map(record => <tr key={record.id}><td>{record.propertyNo || "—"}</td><td><button className="missing-data-case-link" onClick={() => { setMissingDataReminderOpen(false); setEditing(record); }}>{record.caseName || "未命名案件"}</button></td><td>{record.address || "未填寫地址"}</td><td>{missingDataOf(record).join("　")}</td></tr>)}</tbody></table> : <div className="contact-empty">所有委託中物件資料均已補齊。</div>}</div><div className="modal-foot"><span>點選案名可直接前往補資料；仍有待補資料時，每天第一次開啟會提醒一次</span><button onClick={() => setMissingDataReminderOpen(false)}>稍後處理</button></div></div></div>}
    {expiryReminderOpen && tab !== "public" && expiredUnarchived.length > 0 && <div className="modal-backdrop expiry-reminder-backdrop"><div className="modal expiry-reminder-modal"><div className="modal-head"><div><span>委託期限提醒</span><h2>有 {expiredUnarchived.length} 筆到期物件尚未下架</h2></div><button className="close" onClick={() => setExpiryReminderOpen(false)}>×</button></div><div className="expiry-reminder-list">{expiredUnarchived.map(record => <div className="expiry-reminder-row" key={record.id}><div><b><span className="expiry-property-no">{record.propertyNo || "—"}</span>　<button type="button" className="expiry-case-name" onClick={() => { setExpiryReminderOpen(false); setEditing(record); }}>{record.caseName || "未命名案件"}</button></b><span className="expiry-reminder-detail"><span className="expiry-address">{record.address || "—"}</span><span>開發：{record.developer || "—"}</span><span>委託結束：<strong>{displayRocDate(record.entrustEnd) || "—"}</strong></span></span></div><select defaultValue="" onChange={event => { if (event.target.value) requestArchive(record, event.target.value); }}><option value="">選擇封存原因…</option><option>售出下架</option><option>成交下架</option><option>下架洽開發</option><option>到期下架</option></select></div>)}</div>
<div className="modal-foot"><span>未封存前，下次開啟系統仍會提醒</span><button onClick={() => setExpiryReminderOpen(false)}>稍後處理</button></div></div></div>}
    {websitePoReminderOpen && tab !== "public" && overduePoRecords.length > 0 && !expiryReminderOpen && !missingDataReminderOpen && <div className="modal-backdrop website-po-reminder-backdrop"><div className="modal archive-cleanup-modal website-po-reminder-modal"><div className="modal-head"><div><span>PO 網提醒</span><h2>進案超過 14 天，尚有 {overduePoRecords.length} 筆未完成</h2></div><button className="close" onClick={() => setWebsitePoReminderOpen(false)}>×</button></div><div className="archive-cleanup-table-wrap"><table className="archive-cleanup-table website-po-reminder-table"><thead><tr><th>案件</th><th>地址</th><th>開發</th><th>進案日期</th><th>尚未 PO</th><th>操作</th></tr></thead><tbody>{overduePoRecords.map(({ record, missing }) => <tr key={record.id}><td className="archive-cleanup-case-cell"><b>{record.caseName || "未命名案件"}</b><small>{record.propertyNo || "—"}</small></td><td className="website-po-address">{record.address || "—"}</td><td>{record.developer || "—"}</td><td>{displayRocDate(record.reportDate) || "—"}</td><td className="website-po-missing">{missing.map(key => websiteFieldMap[key]).join("、")}</td><td><button onClick={() => { setWebsitePoReminderOpen(false); setEditing(record); }}>前往填寫</button></td></tr>)}</tbody></table></div>
<div className="modal-foot"><span>填入網站編號、勾選「無」或設為旁5後，該項即停止提醒</span><button onClick={() => setWebsitePoReminderOpen(false)}>稍後處理</button></div></div></div>}

    {archiveCleanupReminderOpen && tab !== "public" && pendingArchiveCleanup.length > 0 && !expiryReminderOpen && !websitePoReminderOpen && !missingDataReminderOpen && <div className="modal-backdrop archive-cleanup-backdrop"><div className="modal archive-cleanup-modal"><div className="modal-head"><div><span>封存下架提醒</span><h2>有 {pendingArchiveCleanup.length} 筆封存案件尚未完成下架</h2></div><button className="close" onClick={() => setArchiveCleanupReminderOpen(false)}>×</button></div><div className="archive-cleanup-table-wrap"><table className="archive-cleanup-table"><thead><tr>{archiveCleanupHeader("case", "案件")}{archiveCleanupHeader("housingDownDate", "房管下架")}{archiveCleanupHeader("bookDownDate", "物件本／旁5下架")}{archiveCleanupHeader("salesBookDownDate", "銷售本下架")}{archiveWebsiteTasks.map(([field, label]) => <React.Fragment key={field}>{archiveCleanupHeader(`${field}DownDate`, label)}</React.Fragment>)}</tr></thead><tbody>{sortedPendingArchiveCleanup.map(({ record, tasks }) => { const allTasks = archiveCleanupTasks(record); const buttonFor = (key: string) => { const applicable = allTasks.find(item => item.key === key); if (!applicable) return <span className="archive-cleanup-na">—</span>; const pending = tasks.find(item => item.key === key); return pending ? <button onClick={() => completeArchiveCleanup(record, pending.key)}>完成下架</button> : <span className="archive-cleanup-done">已完成</span>; }; return <tr key={record.id}><td className="archive-cleanup-case-cell"><b>{stripRestoredDisplay(record.caseName) || "未命名案件"}</b><small className="archive-cleanup-address">{record.address || "未填寫地址"}</small><small className="archive-cleanup-status">{record.propertyNo || "—"}　{displayRocDate(record.archived || "")}{record.status || "下架"}</small></td><td>{buttonFor("housingDownDate")}</td><td>{buttonFor("bookDownDate")}</td><td>{buttonFor("salesBookDownDate")}</td>{archiveWebsiteTasks.map(([field]) => <td key={field}>{buttonFor(`${field}DownDate`)}</td>)}</tr>; })}</tbody></table></div><div className="modal-foot"><span>未完成前，下次開啟系統仍會提醒</span><button onClick={() => setArchiveCleanupReminderOpen(false)}>稍後處理</button></div></div></div>}
    {archiveChoice && <div className="modal-backdrop"><div className="modal archive-choice-modal"><div className="modal-head"><div><span>確認封存資料</span><h2>{archiveChoice.record.caseName || "未命名案件"}</h2></div><button className="close" onClick={() => setArchiveChoice(null)}>×</button></div><div className="archive-choice-body"><label className="field"><span>封存原因</span><select value={archiveChoice.status} onChange={event => { const status = event.target.value; setArchiveChoice({ ...archiveChoice, status, date: status === "到期下架" ? nextDate(archiveChoice.record.entrustEnd) : today() }); }}><option>售出下架</option><option>成交下架</option><option>下架洽開發</option><option>到期下架</option></select></label><label className="field"><span>案名後方顯示日期</span><input type="text" value={displayRocDate(archiveChoice.date)} onChange={event => setArchiveChoice({ ...archiveChoice, date: event.target.value })} onBlur={event => setArchiveChoice({ ...archiveChoice, date: normalizeDateInput(event.target.value) || today() })}/></label><div className="archive-party-fields"><span>開發：{developerFullNameText(archiveChoice.record.developer, settings.personnel) || "—"}</span>{archiveChoice.status === "成交下架" && <label className="field"><span>銷售</span><input type="text" value={archiveChoice.salesPerson} placeholder="填寫銷售人員" onChange={event => setArchiveChoice({ ...archiveChoice, salesPerson: event.target.value })}/></label>}</div>{archiveChoice.status === "下架洽開發" && <label className="field archive-reason-field"><span>下架原因（非必填）</span><input type="text" value={archiveChoice.reason} placeholder="例如：屋主已出租" onChange={event => setArchiveChoice({ ...archiveChoice, reason: event.target.value })}/></label>}<small>一般原因預設今天；到期下架預設委託結束隔天，可自行修改。每日物件動態仍依今天的實際操作日顯示。</small></div>
<div className="modal-foot"><button onClick={() => setArchiveChoice(null)}>取消</button><button className="primary" onClick={() => archiveRecord(archiveChoice.record, archiveChoice.status, archiveChoice.date, archiveChoice.salesPerson, archiveChoice.reason)}>確認封存</button></div></div></div>}
    {restoreChoiceRecord && <div className="modal-backdrop"><div className="modal restore-choice-modal"><div className="modal-head"><div><span>恢復封存物件</span><h2>{restoreChoiceRecord.caseName || "未命名案件"}</h2></div><button className="close" onClick={() => setRestoreChoiceRecord(null)}>×</button></div><div className="restore-choice-body"><p>請選擇這次恢復的原因：</p><button className="primary" onClick={() => restoreRecord(restoreChoiceRecord, true)}><b>重新上架</b><span>記錄今天日期，顯示重新上架紅字並列入每日動態</span></button><button onClick={() => restoreRecord(restoreChoiceRecord, false)}><b>恢復委託中物件</b><span>只恢復到委託中，不標示重新上架</span></button></div>
<div className="modal-foot"><button onClick={() => setRestoreChoiceRecord(null)}>取消</button></div></div></div>}
    {editing && <div className="modal-backdrop" onMouseDown={e => e.target === e.currentTarget && requestCloseEditing()}><div className="modal record-edit-modal"><div className="modal-head"><div className="record-modal-title"><span>{editing._intakeDraftId ? "編輯進案草稿" : records.some(r => r.id === editing.id) ? "編輯案件" : "建立新物件"}</span><h2><b>{editing.propertyNo || "尚無編號"}</b><em>{editing.caseName || "尚未命名"}</em></h2><small>{editing.address || "尚未填寫地址"}</small><p>開發業務：{developerFullNameText(editing.developer) || "尚未填寫"}</p>{editing.archived && <i className="record-archive-title-note">{displayRocDate(editing.archived)} {editing.status || "下架"}</i>}</div>
<div className="modal-head-actions">{records.some(record => record.id === editing.id) && !editing.archived && (editing.status || "委託中") === "委託中" && <><button className="record-print-button" type="button" onClick={() => openRecordIntakeDraft(editing)}>進案草稿</button><button className={`record-print-button color${editing.colorSheetIssue ? " has-issue" : ""}`} type="button" onClick={() => { const attention = colorSheetAttention(editing.attentionNotes || "", editing.additionNotes || ""); setPrintEditor({ kind: "color", data: { ...editing, notes: attention, attentionNotes: attention, photos: [...(editing.photos || [])] } }); }}>{editing.colorSheetIssue ? "彩色表 Excel ●" : "彩色表 Excel"}</button><button className="record-print-button cover" type="button" onClick={() => printRecordDocument(editing, "cover")}>列印新進封面</button></>}<button className="close" type="button" onClick={requestCloseEditing}>×</button></div></div>
<div className="form-grid record-edit-grid">{recordEditOrder.filter(key => key !== "area").map(key => {
  if (["feature2", "feature3", "feature4"].includes(key)) return null;
  if (key === "feature1") return <div className="edit-cell edit-features" key="features">{["feature1", "feature2", "feature3", "feature4"].map((featureKey, index) => <div className="editing-feature-row" key={featureKey}><Field fieldKey={featureKey} label={labels[featureKey]} record={editing} records={records} setRecord={updateEditingRecord}/>{index === 0 && <button type="button" className="copy-features-button" onClick={copyEditingFeatures}>複製特色1～4</button>}</div>)}</div>;
  const label = recordEditLabels[key] || labels[key] || key;
  return <div className={recordEditClass(key)} key={key}><Field fieldKey={key} label={label} record={editing} records={records} setRecord={updateEditingRecord}/></div>;
})}</div>
<div className="modal-foot"><div className="archive-in-editor">{records.some(r => r.id === editing.id) && (editing.archived || isExpired(editing) || editing.status !== "委託中") ? <><button type="button" className="restore-in-editor" onClick={() => { setRestoreChoiceRecord(editing); setEditing(null); }}>恢復</button>{editing.archived && <div className="archive-task-pills">{archiveCleanupTasks(editing).map(task => <button type="button" key={task.key} className={editing[task.key] ? "done" : ""} onClick={() => completeArchiveCleanup(editing, task.key)}>{editing[task.key] ? `${task.label}完成` : task.label}</button>)}</div>}</> : records.some(r => r.id === editing.id) && <select defaultValue="" onChange={e => { if (e.target.value) { requestArchive(editing, e.target.value); setEditing(null); } }}><option value="">封存案件…</option><option>售出下架</option><option>成交下架</option><option>下架洽開發</option><option>到期下架</option></select>}</div><div className="record-bottom-actions">{records.some(r => r.id === editing.id) && !editing.archived && editing.status === "委託中" && <><button type="button" onClick={() => returnToIntake(editing)}>退回進案草稿</button><button type="button" className="danger" onClick={() => removeRecord(editing)}>刪除</button></>}</div><div className="record-save-actions">{editing._intakeDraftId && <span className="draft-save-note">尚未進案～進案草稿中</span>}<button type="button" onClick={temporarilySaveRecord}>暫存</button>{editing._intakeDraftId ? <button className="primary" onClick={() => saveRecord(false)}>同步更新草稿</button> : <><button type="button" className="daily-publish-secondary" onClick={() => saveRecord(true)}>儲存並發布每日動態</button><button className="primary save-only-primary" onClick={() => saveRecord(false)}>僅儲存</button></>}</div></div></div></div>}
    {printEditor && <div className="modal-backdrop print-editor-backdrop"><div className="modal print-editor-modal compact"><div className="modal-head print-editor-head"><div><span>彩色銷售資料表</span><h2>下載彩色表 Excel</h2></div><button className="close" type="button" onClick={() => setPrintEditor(null)}>×</button></div><div className="print-editor-body compact-body"><section className="print-editor-section basic"><div className="print-editor-compact-grid">{[["developer", "開發人員"], ["price", "總價（萬）"], ["key", "KEY編號"], ["currentState", "現況"], ["caseName", "案名"], ["address", "地址"]].map(([key, label]) => <label className={`field print-editor-${key}`} key={key}><span>{label}</span><input type="text" value={printEditor.data[key] || ""} onChange={event => setPrintEditor({ ...printEditor, data: { ...printEditor.data, [key]: event.target.value } })}/></label>)}<label className="field print-editor-notes"><span>重點說明與備註</span><textarea rows={5} value={printEditor.data.attentionNotes || ""} onChange={event => { const attention = colorSheetAttention(event.target.value, printEditor.data.additionNotes || ""); setPrintEditor({ ...printEditor, data: { ...printEditor.data, notes: attention, attentionNotes: attention } }); }}/></label><label className="field print-editor-issue"><span>彩色表缺件／原因紀錄</span><textarea rows={4} value={printEditor.data.colorSheetIssue || ""} onChange={event => setPrintEditor({ ...printEditor, data: { ...printEditor.data, colorSheetIssue: event.target.value } })} placeholder="例如：缺房管坪數、照片尚未收到，暫時無法製作。"/></label></div></section></div><div className="modal-foot print-editor-foot"><span className="print-editor-warning">重點說明與備註只套用本次下載；缺件紀錄可另外儲存到案件。</span><button type="button" onClick={() => setPrintEditor(null)}>取消</button><button type="button" onClick={() => saveColorSheetIssue(printEditor.data)}>儲存缺件紀錄</button><button className="primary" type="button" disabled={colorWorkbookDownloading} onClick={async () => { const generalTax = String(printEditor.data.generalLandValueTax || "").trim(); const selfUseTax = String(printEditor.data.selfUseLandValueTax || "").trim(); const taxMessage = `請確認土增稅資料：\n一般增值稅：${generalTax || "未填寫"}\n自用增值稅：${selfUseTax || "未填寫"}\n\n是否確認並下載彩色表 Excel？`; if (!window.confirm(taxMessage)) return; setColorWorkbookDownloading(true); flash("正在建立彩色表 Excel…"); try { await downloadColorWorkbook(printEditor.data, settings.personnel); setPrintEditor(null); flash("彩色表 Excel 已下載"); } catch (error) { console.error(error); const reason = error instanceof Error ? error.message : String(error || "未知錯誤"); alert(`彩色表 Excel 下載失敗：${reason}`); flash("彩色表 Excel 下載失敗"); } finally { setColorWorkbookDownloading(false); } }}>{colorWorkbookDownloading ? "建立中…" : "下載 Excel"}</button></div></div></div>}
    {pptPickerOpen && <div className="modal-backdrop" onMouseDown={event => event.target === event.currentTarget && setPptPickerOpen(false)}><div className="modal ppt-picker-modal"><div className="modal-head"><div><span>產生 POWERPOINT</span><h2>選擇週報期別與進案物件</h2></div><button className="close" onClick={() => setPptPickerOpen(false)}>×</button></div><div className="ppt-picker-body"><div className="ppt-week-selector"><b>報告週期</b><div className="ppt-week-dropdown"><button type="button" className="ppt-week-current" onClick={() => setPptWeekMenuOpen(value => !value)}><span>{displayPptWeekLabel(selectedPptWeek)}</span><i>{pptWeekMenuOpen ? "▲" : "▼"}</i></button>{pptWeekMenuOpen && <div className="ppt-week-menu">{pptWeekOptions.map(start => { const week = pptWeekOf(start); return <button type="button" className={selectedPptWeek.start === week.start ? "selected" : ""} key={week.start} onClick={() => { setPptWeekStart(week.start); setPptExtraIds([]); setPptWeekMenuOpen(false); }}>{displayPptWeekLabel(week)}</button>; })}</div>}</div></div><div className="ppt-week-range"><b>本週進案</b><span>{displayPptWeekLabel(selectedPptWeek)}</span></div><div className="ppt-meeting-note">進案週期固定為星期六至星期五，下一個星期一報告；歷史週期會持續保留。</div><div className="ppt-case-list"><div className="ppt-case-heading"><span>序</span><span>案名</span><span>地址</span><span>開發</span><span>進案日期</span><span>確認／操作</span></div>{weeklyPptRecords.map((record, index) => <div className="ppt-case-row auto" key={record.id}><strong>{index + 1}</strong><button type="button" className="ppt-case-name-edit" onClick={() => { setPptConfirmedSnapshots(previous => { const next = { ...previous }; delete next[record.id]; return next; }); setEditing(record); }}>{record.caseName || "未命名案件"}</button><span className="ppt-case-address">{chunkText(record.address || "—", 20).map((line, lineIndex) => <span key={`${line}-${lineIndex}`}>{line}</span>)}</span><span>{developerFullNameText(record.developer) || "—"}</span><small>{displayRocDate(record.reportDate) || "—"}</small><div className="ppt-inline-confirm">{pptConfirmedSnapshots[record.id] ? <button type="button" className="confirmed" onClick={() => setPptConfirmedSnapshots(previous => { const next = { ...previous }; delete next[record.id]; return next; })}>已確認／解除</button> : <button type="button" className="primary" onClick={() => setPptConfirmedSnapshots(previous => ({ ...previous, [record.id]: { ...record, photos: [...(record.photos || [])] } }))}>確認並鎖定</button>}<button className="ppt-remove-pill" onClick={() => removeFromPptWeek(record)}>刪</button></div></div>)}{!weeklyPptRecords.length && <div className="ppt-empty">這個日期區間目前沒有進案物件</div>}</div>{deferredPptRecords.length > 0 && <div className="ppt-deferred-list"><h3>本週已移除／待決定</h3>{deferredPptRecords.map(record => { const choice = pptStoredChoices(record._pptWeekChoices)[selectedPptWeek.start] || ""; return <div className="ppt-deferred-row" key={record.id}><div><button type="button" className="ppt-case-name-edit" onClick={() => { setPptConfirmedSnapshots(previous => { const next = { ...previous }; delete next[record.id]; return next; }); setEditing(record); }}>{record.caseName || "未命名案件"}</button><small>{record.propertyNo}　原週期保留</small></div><button className={choice === "next" ? "selected" : ""} onClick={() => decideRemovedPpt(record, "next")}>下次報告</button><button className={choice === "skip" ? "selected skip" : ""} onClick={() => decideRemovedPpt(record, "skip")}>確定不上PPT</button></div>; })}</div>}{pptAdHocRecords.length > 0 && <div className="ppt-adhoc-added"><h3>本次新增案件</h3>{pptAdHocRecords.map(record => <div key={record.id}><button type="button" className="ppt-case-name-edit" onClick={() => { setPptAdHocDraft({ ...record }); setPptAdHocOpen(true); }}>{record.caseName || "未命名案件"}</button><span>{record.address || "—"}</span><div className="ppt-inline-confirm">{pptConfirmedSnapshots[record.id] ? <button type="button" className="confirmed" onClick={() => setPptConfirmedSnapshots(previous => { const next = { ...previous }; delete next[record.id]; return next; })}>已確認／解除</button> : <button type="button" className="primary" onClick={() => setPptConfirmedSnapshots(previous => ({ ...previous, [record.id]: { ...record, photos: [...(record.photos || [])] } }))}>確認並鎖定</button>}<button type="button" onClick={() => setPptAdHocRecords(previous => previous.filter(item => item.id !== record.id))}>刪</button></div></div>)}</div>}<div className="ppt-add-actions"><button type="button" className="ppt-extra-toggle" onClick={() => setPptShowExtras(value => !value)}>＋加入物件</button><button type="button" className="ppt-adhoc-toggle" onClick={() => { setPptAdHocDraft({ ...blankRecord(), id: newId(), status: "尚未進案" }); setPptAdHocOpen(true); }}>＋尚未填寫表單案件</button></div>{pptShowExtras && <><label className="ppt-extra-search">搜尋案名或地址<input type="search" value={pptExtraSearch} onChange={event => setPptExtraSearch(event.target.value)} placeholder="輸入案名或地址"/></label><div className="ppt-extra-list">{pptExtraCandidates.map(record => <label className="ppt-case-row extra" key={record.id}><input type="checkbox" checked={pptExtraIds.includes(record.id)} onChange={event => setPptExtraIds(previous => event.target.checked ? [...previous, record.id] : previous.filter(id => id !== record.id))}/><button type="button" className="ppt-case-name-edit" onClick={() => { setPptConfirmedSnapshots(previous => { const next = { ...previous }; delete next[record.id]; return next; }); setEditing(record); }}>{record.caseName || "未命名案件"}</button><span className="ppt-case-address">{chunkText(record.address || "—", 20).map((line, lineIndex) => <span key={`${line}-${lineIndex}`}>{line}</span>)}</span><span>{developerFullNameText(record.developer) || "—"}</span>{record._intakeDraftId ? <small className="ppt-source-status draft"><b>進案草稿中</b><em>尚未進案</em></small> : record.archived || isExpired(record) || record.status !== "委託中" ? <small className="ppt-source-status archived"><b>封存下架</b><em>{displayRocDate(record.archived || record.entrustEnd || "") || "—"}{record.archived ? (record.status || "下架") : isExpired(record) ? "到期下架" : (record.status || "下架")}</em></small> : <small className="ppt-source-status active">委託中</small>}<span className="ppt-inline-confirm">{pptExtraIds.includes(record.id) ? (pptConfirmedSnapshots[record.id] ? <button type="button" className="confirmed" onClick={event => { event.preventDefault(); setPptConfirmedSnapshots(previous => { const next = { ...previous }; delete next[record.id]; return next; }); }}>已確認／解除</button> : <button type="button" className="primary" onClick={event => { event.preventDefault(); setPptConfirmedSnapshots(previous => ({ ...previous, [record.id]: { ...record, photos: [...(record.photos || [])] } })); }}>確認並鎖定</button>) : null}</span></label>)}{!pptExtraCandidates.length && <div className="ppt-empty">沒有符合的委託中、封存或尚未進案草稿案件</div>}</div></>}</div><div className="modal-foot ppt-picker-foot"><span>本期 {weeklyPptRecords.length} 筆＋加入 {pptExtraIds.length + pptAdHocRecords.length} 筆　{pendingPptConfirmCount ? `／待確認 ${pendingPptConfirmCount} 筆` : "／全部已確認"}</span><button onClick={() => setPptPickerOpen(false)}>取消</button><button className="ppt-image-button" onClick={exportPptOrderImage}>產生圖片</button><button className="primary" onClick={exportPpt}>下載 PPT</button></div></div></div>}
    {pptAdHocOpen && <div className="modal-backdrop ppt-adhoc-backdrop"><div className="modal ppt-adhoc-modal"><div className="modal-head"><div><span>僅供本次 PPT</span><h2>{pptAdHocDraft.caseName || "尚未輸入案名"}</h2><p className="ppt-adhoc-live-title"><b>{pptAdHocDraft.address || "尚未輸入地址"}</b><em>開發：{developerFullNameText(pptAdHocDraft.developer || "") || "尚未輸入"}</em></p></div><button className="close" onClick={() => setPptAdHocOpen(false)}>×</button></div><div className="ppt-adhoc-grid">{([["type","種類"],["caseName","案名"],["address","地址"],["developer","開發"],["price","總價（萬）"],["landPing","地坪"],["buildingPing","建坪"],["indoorPing","室內坪"],["layout","格局"],["floor","樓層"],["direction","朝向"],["parking","車位"],["managementFee","管理費"],["completionDate","完工日期"],["communityName","社區名稱"],["currentState","現況"],["key","鑰匙"],["road","臨路"],["frontage","面寬"],["depth","深度"],["zoning","使用分區"],["coverage","建蔽率"],["far","容積率"],["notes","備註"]] as [string,string][]).map(([key,label]) => <label className={`field ppt-adhoc-${key}`} key={key}><span>{label}</span>{key === "type" ? <select value={pptAdHocDraft[key] || ""} onChange={event => setPptAdHocDraft({ ...pptAdHocDraft, [key]: event.target.value })}><option value="">請選擇</option><option>土地</option><option>透天</option><option>公寓</option><option>華廈</option><option>大樓</option><option>其他</option></select> : key === "notes" ? <textarea rows={3} value={pptAdHocDraft[key] || ""} onChange={event => setPptAdHocDraft({ ...pptAdHocDraft, [key]: event.target.value })}/> : <input type="text" value={pptAdHocDraft[key] || ""} onChange={event => setPptAdHocDraft({ ...pptAdHocDraft, [key]: event.target.value })}/>}</label>)}</div><div className="modal-foot"><span>資料只保留到這次關閉 PPT 視窗為止</span><button onClick={() => setPptAdHocOpen(false)}>取消</button><button className="primary" onClick={() => { const item = { ...pptAdHocDraft, id: pptAdHocDraft.id || newId(), status: "尚未進案", reportDate: "", _notEntered: "1" }; setPptAdHocRecords(previous => previous.some(record => record.id === item.id) ? previous.map(record => record.id === item.id ? item : record) : [...previous, item]); setPptAdHocOpen(false); flash("已加入本次 PPT"); }}>加入本次 PPT</button></div></div></div>}
    {pptPickerOpen && <div className="ppt-custom-period"><label><span>週期開始日期</span><input type="text" value={displayRocDate(pptCustomStart || defaultSelectedPptWeek.start)} onChange={event => setPptCustomStart(event.target.value)} onBlur={event => setPptCustomStart(normalizeDateInput(event.target.value) || defaultSelectedPptWeek.start)}/></label><label><span>週期結束日期</span><input type="text" value={displayRocDate(pptCustomEnd || defaultSelectedPptWeek.end)} onChange={event => setPptCustomEnd(event.target.value)} onBlur={event => setPptCustomEnd(normalizeDateInput(event.target.value) || defaultSelectedPptWeek.end)}/></label><label><span>報告日期</span><input type="text" value={displayRocDate(pptCustomMeeting || defaultSelectedPptWeek.meeting)} onChange={event => setPptCustomMeeting(event.target.value)} onBlur={event => setPptCustomMeeting(normalizeDateInput(event.target.value) || defaultSelectedPptWeek.meeting)}/></label><button type="button" onClick={resetPptCustomWeek}>恢復預設週期</button></div>}
    {!internalView && newCaseReminder && newCaseReminderBatchRecords.length > 0 && <div className="modal-backdrop new-case-reminder-backdrop" onMouseDown={event => event.target === event.currentTarget && deferNewCaseReminder()}><div className="modal new-case-reminder-modal new-case-reminder-list-modal"><div className="modal-head"><div><span>新增案件</span><h2>新案件完成提醒　{newCaseReminderBatchRecords.length} 筆</h2></div><button className="close" type="button" onClick={deferNewCaseReminder}>×</button></div><div className="new-case-reminder-table-wrap"><table className="new-case-reminder-table"><thead><tr><th>案名</th><th>地址</th><th>開發</th><th>進案日期</th><th>房管上架完成</th><th>NEW 本製作完成</th><th>王總審核完成</th></tr></thead><tbody>{newCaseReminderBatchRecords.map(record => <tr key={record.id}><td><b>{record.caseName || "未命名案件"}</b><small>{record.propertyNo || "—"}</small></td><td>{record.address || "未填寫地址"}</td><td>{developerFullNameText(record.developer || "", settings.personnel) || "未填寫"}</td><td>{displayRocDate(record.reportDate) || "—"}</td>{(["housingListingCompleted", "newBookCompleted", "wangReviewCompleted"] as string[]).map(key => <td key={key}><input type="checkbox" checked={record[key] === "1"} onChange={event => updateNewCaseReminderItem(record.id, key, event.target.checked)}/></td>)}</tr>)}</tbody></table></div><div className="modal-foot"><span>來源：進案草稿點選「進案」</span><button type="button" onClick={deferNewCaseReminder}>稍後處理</button><button type="button" className="primary" onClick={saveNewCaseReminder}>儲存提醒</button></div></div></div>}
    {notice && <div className="toast">{notice}</div>}
  </main>;
}

const tourColumns = ["propertyNo", "areaType", "caseName", "address", "price", "direction", "age", "floor", "layout", "indoorPing", "buildingPing", "landPing", "parking", "managementFee", "key", "currentState", "road", "frontageDepth", "zoning", "coverageFar", "developer", "notes", "reportDate"];
const tourDisplayColumns = tourColumns.filter(column => column !== "propertyNo");
const tourColumnLabel: Record<string, string> = { propertyNo: "物件編號", areaType: "地區\n種類", caseName: "案名", address: "地址", price: "總價", direction: "朝向", age: "屋齡", floor: "樓層", layout: "格局", indoorPing: "室內坪", buildingPing: "建坪", landPing: "地坪", parking: "車位", managementFee: "管理費", key: "鑰匙", currentState: "現況", road: "臨路", frontageDepth: "面寬\n深度", zoning: "使用分區", coverageFar: "建蔽\n容積", developer: "開發", notes: "備註欄", reportDate: "進案報件日期" };
const temporaryTourFields = [
  ["area", "地區"], ["type", "種類"], ["caseName", "案名"], ["developer", "開發"], ["address", "地址"], ["price", "總價（萬）"], ["direction", "朝向"], ["_tourAge", "屋齡"], ["floor", "樓層"], ["layout", "格局"], ["indoorPing", "室內坪"], ["buildingPing", "建坪"], ["landPing", "地坪"], ["parking", "車位"], ["managementFee", "管理費"], ["key", "鑰匙"], ["currentState", "現況"], ["road", "臨路"], ["frontage", "面寬"], ["depth", "深度"], ["zoning", "使用分區"], ["coverage", "建蔽"], ["far", "容積"], ["notes", "備註欄"],
] as const;

function TourPlanner({ records, drafts, items, setItems, history, setHistory, editRecord, updateDraftCaseName, tourDate, setTourDate, tourTitle, setTourTitle, notify, complete }: { records: RecordItem[]; drafts: IntakeData[]; items: TourItem[]; setItems: (items: TourItem[]) => void; history: TourHistory[]; setHistory: Dispatch<SetStateAction<TourHistory[]>>; editRecord: (record: RecordItem) => void; updateDraftCaseName: (id: string, caseName: string) => void; tourDate: string; setTourDate: (date: string) => void; tourTitle: string; setTourTitle: (title: string) => void; notify: (text: string) => void; complete: (date: string, recordIds: string[], draftIds: string[]) => void }) {
  const [search, setSearch] = useState("");
  const [pickerOpen, setPickerOpen] = useState(false);
  const [developerFilter, setDeveloperFilter] = useState("");
  const [reportDateFrom, setReportDateFrom] = useState("");
  const [reportDateTo, setReportDateTo] = useState("");
  const [propertyKindFilter, setPropertyKindFilter] = useState("");
  const [temporaryOpen, setTemporaryOpen] = useState(false);
  const [temporaryDraft, setTemporaryDraft] = useState<RecordItem>(() => blankRecord());
  const [temporaryEditId, setTemporaryEditId] = useState("");
  const unviewed = sortActiveRecords(records.filter(record => !record.archived && !isExpired(record) && (record.status || "委託中") === "委託中" && !record.groupViewDate && !items.some(item => item.recordId === record.id)));
  const draftCandidates = drafts.filter(draft => !draft.linkedRecordId && !draft.groupViewDate && !items.some(item => item.data._intakeDraftId === draft.id)).map(draft => ({ ...intakeToRecord(draft), id: `draft-${draft.id}`, reportDate: "", status: "尚未進案", _intakeDraftId: draft.id, _notEntered: "1" }));
  const candidates = [...draftCandidates, ...unviewed];
  const developers = [...new Set(candidates.flatMap(record => developerNameLines(record.developer || "")))].sort((a, b) => a.localeCompare(b, "zh-TW"));
  const available = candidates.filter(record => {
    const reportDate = String(record.reportDate || "").slice(0, 10);
    const kind = typeShort(record.type) === "土地" ? "土地" : "房屋";
    return (!developerFilter || developerNameLines(record.developer || "").includes(developerFilter))
      && (!reportDateFrom || (!!reportDate && reportDate >= reportDateFrom))
      && (!reportDateTo || (!!reportDate && reportDate <= reportDateTo))
      && (!propertyKindFilter || kind === propertyKindFilter)
      && [record.propertyNo, record.caseName, record.address, developerFullNameText(record.developer)].join(" ").includes(search);
  });
  const ordered = items.slice().sort((a, b) => Number(a.sequence || 9999) - Number(b.sequence || 9999));
  // 新增時只補新案件的下一個序號，保留原本路線已手動填寫的順序。
  const nextTourSequence = () => {
    const used = items.map(item => Number(String(item.sequence || "").trim())).filter(value => Number.isFinite(value) && value > 0);
    return String(used.length ? Math.max(...used) + 1 : 1);
  };
  const addRecord = (record: RecordItem) => setItems([...items, { id: newId(), recordId: record._intakeDraftId ? undefined : record.id, sequence: nextTourSequence(), temporary: !!record._intakeDraftId, data: { ...record } }]);
  const openTemporary = () => { const data = blankRecord(); data.propertyNo = "臨時"; data.status = "臨時團看"; data.reportDate = ""; data._notEntered = "1"; setTemporaryEditId(""); setTemporaryDraft(data); setTemporaryOpen(true); };
  const editTemporary = (item: TourItem) => { setTemporaryEditId(item.id); setTemporaryDraft({ ...item.data }); setTemporaryOpen(true); };
  const addTemporary = () => { const normalizedDraft = normalizeRecordPings({ ...temporaryDraft, developer: developerFullNameText(temporaryDraft.developer || "") || temporaryDraft.developer }); if (temporaryEditId) setItems(items.map(item => item.id === temporaryEditId ? { ...item, data: { ...normalizedDraft } } : item)); else setItems([...items, { id: newId(), sequence: nextTourSequence(), temporary: true, data: { ...normalizedDraft } }]); setTemporaryOpen(false); setTemporaryEditId(""); };
  const updateItem = (id: string, patch: Partial<TourItem>) => setItems(items.map(item => item.id === id ? { ...item, ...patch } : item));
  const updateTemp = (item: TourItem, key: string, value: string) => updateItem(item.id, { data: { ...item.data, [key]: value } });
  // 團看項目可能在 JSON 匯入或雲端合併後換了內部 ID；正常委託中案件優先重新對應目前總表，避免顯示加入當時的舊快照。
  const sourceOf = (item: TourItem) => {
    // 歷史路線叫回後，必須以完成當下的整筆快照呈現，不能被目前委託中資料覆蓋。
    if (item.data?._tourHistorySnapshot === "1") return item.data;
    if (!item.recordId) return item.data;
    const no = String(item.data?.propertyNo || "").trim();
    const name = String(item.data?.caseName || "").trim();
    const address = String(item.data?.address || "").trim();
    return records.find(record => record.id === item.recordId)
      || (no ? records.find(record => String(record.propertyNo || "").trim() === no) : undefined)
      || (name && address ? records.find(record => String(record.caseName || "").trim() === name && String(record.address || "").trim() === address) : undefined)
      || item.data;
  };
  const isNotEnteredTourRecord = (record: RecordItem) => record._notEntered === "1" || record.status === "臨時團看" || !!record._intakeDraftId;
  const renderValue = (record: RecordItem, column: string) => {
    if (column === "areaType") return <span className="tour-pair"><span>{areaCategory(record)}</span><span>{typeShort(record.type)}</span></span>;
    if (column === "age" && isNotEnteredTourRecord(record) && record._tourAge) return <>{record._tourAge}</>;
    if (column === "reportDate" && isNotEnteredTourRecord(record)) return <span className="not-entered">尚未進案</span>;
    if (column === "frontageDepth") return <span className="tour-pair"><span>{record.frontage || "—"}</span><span>{record.depth || "—"}</span></span>;
    if (column === "coverageFar") return <span className="tour-pair"><span>{record.coverage || "—"}</span><span>{record.far || "—"}</span></span>;
    return <CellContent record={record} column={column}/>;
  };
  const tempInput = (item: TourItem, column: string) => {
    const record = item.data;
    if (column === "age") return <input value={record._tourAge || ""} onChange={event => updateTemp(item, "_tourAge", event.target.value)} placeholder="屋齡"/>;
    if (column === "areaType") return <span className="tour-pair inputs"><input value={record.area || ""} onChange={event => updateTemp(item, "area", event.target.value)} placeholder="地區"/><input value={record.type || ""} onChange={event => updateTemp(item, "type", event.target.value)} placeholder="種類"/></span>;
    if (column === "frontageDepth") return <span className="tour-pair inputs"><input value={record.frontage || ""} onChange={event => updateTemp(item, "frontage", event.target.value)} placeholder="面寬"/><input value={record.depth || ""} onChange={event => updateTemp(item, "depth", event.target.value)} placeholder="深度"/></span>;
    if (column === "coverageFar") return <span className="tour-pair inputs"><input value={record.coverage || ""} onChange={event => updateTemp(item, "coverage", event.target.value)} placeholder="建蔽"/><input value={record.far || ""} onChange={event => updateTemp(item, "far", event.target.value)} placeholder="容積"/></span>;
    return <input value={record[column] || ""} onChange={event => updateTemp(item, column, event.target.value)} />;
  };
  const imageDeveloper = (value = "") => developerNameLines(value).join("\n");
  const supervisorRows = ordered.map(item => { const record = sourceOf(item); const price = String(record.price || "").replace(/\s*萬(?:元)?\s*/g, ""); return [item.sequence, areaCategory(record) || record.area || "", record.caseName || "", record.address || "", price ? `${price}萬` : "", imageDeveloper(record.developer || "")]; });
  const copyText = async (text: string, message: string) => {
    if (!text.trim()) return notify("目前沒有可複製的團看物件");
    try { await navigator.clipboard.writeText(text); notify(message); }
    catch { const box = document.createElement("textarea"); box.value = text; box.style.position = "fixed"; box.style.opacity = "0"; document.body.appendChild(box); box.select(); document.execCommand("copy"); box.remove(); notify(message); }
  };
  const copySupervisor = () => copyText([tourTitle, "序\t地區\t案名\t地址\t開價\t開發", ...supervisorRows.map(row => row.join("\t"))].join("\n"), "已複製主管用團看清單");
  const copyAvailable = () => copyText(["目前尚未安排團看", "案名\t地址\t開發\t開價\t現況\t進案日期", ...available.map(record => [record.caseName, record.address, developerFullNameText(record.developer) || "未填開發", record.price ? `${String(record.price).replace(/\s*萬(?:元)?\s*/g, "")}萬` : "", record.currentState || "", record._intakeDraftId ? "尚未進案" : displayRocDate(record.reportDate) || ""].join("\t"))].join("\n"), "已複製目前尚未安排團看列表");
  const copyRoute = () => copyText([tourTitle, ["案名", "地址", "開發", "開價", "現況", "進案日期"].join("\t"), ...ordered.map(item => { const record = sourceOf(item); const price = String(record.price || "").replace(/\s*萬(?:元)?\s*/g, ""); return [record.caseName || "", record.address || "", developerFullNameText(record.developer) || "", price ? `${price}萬` : "", record.currentState || "", isNotEnteredTourRecord(record) ? "尚未進案" : displayRocDate(record.reportDate) || ""].join("\t"); })].join("\n"), "已複製本次團看路線");
  const saveTourHistory = () => {
    if (!items.length) return;
    const entry: TourHistory = { id: newId(), date: tourDate, title: tourTitle || `${displayRocDate(tourDate).replace(/\//g, ".")}團看`, items: items.map(item => ({ ...item, data: { ...sourceOf(item), _tourHistorySnapshot: "1" } })), completedAt: new Date().toISOString() };
    setHistory(previous => [entry, ...previous.filter(item => !(item.date === entry.date && item.title === entry.title))].slice(0, 100));
  };
  const restoreTourHistory = (entry: TourHistory) => {
    if (!entry.items.length) return;
    if (!confirm(`叫回「${entry.title}」的 ${entry.items.length} 筆團看路線？目前正在編排的路線不會刪除，請您確認後再自行調整。`)) return;
    setItems(entry.items.map(item => ({ ...item, id: newId(), data: { ...item.data, _tourHistorySnapshot: "1" } })));
    setTourDate(entry.date);
    setTourTitle(entry.title);
    notify(`已叫回 ${entry.title}，可再次產生圖片或列印`);
  };
  const createNewTour = () => {
    if (items.length && !confirm("要建立新的團看路線嗎？目前尚未完成的本次清單會先保留在畫面外，請先確認不需要再使用。")) return;
    const nextDate = today();
    setItems([]);
    setTourDate(nextDate);
    setTourTitle(`${displayRocDate(nextDate).replace(/\//g, ".")}團看`);
    notify("已建立新的本次團看路線，請加入案件");
  };
  const exportTourImage = () => {
    if (!ordered.length) return notify("請先加入團看物件");
    try {
    const headers = ["序", "地區\n種類", "案名", "地址", "總價", "朝向", "屋齡", "樓層", "格局", "室內坪", "建坪", "地坪", "車位", "管理費", "鑰匙", "現況", "臨路", "面寬\n深度", "使用分區", "建蔽\n容積", "開發", "備註欄"];
      const widths = [52, 78, 170, 220, 90, 68, 120, 110, 82, 108, 105, 105, 120, 134, 112, 100, 88, 138, 105, 119, 105, 340];
    const textOf = (record: RecordItem, column: string) => String(cellValue(record, column) || record[column] || "");
    const imageLayout = (record: RecordItem) => { if (typeShort(record.type) === "土地" || /^(?:土地\s*)+$/.test(record.layout || "")) return "土地"; const room = (record.layout || "").match(/(\d+)\s*房/)?.[1], hall = (record.layout || "").match(/(\d+)\s*廳/)?.[1], bath = (record.layout || "").match(/(\d+)\s*衛(?:浴)?/)?.[1]; const lines = [[room, "房"], [hall, "廳"], [bath, "衛"]].filter(([value]) => value !== undefined).map(([value, label]) => `${value}${label}`); return lines.length ? lines.join("\n") : record.layout || ""; };
    const imageAge = (record: RecordItem) => isNotEnteredTourRecord(record) && record._tourAge ? record._tourAge : ageOf(record);
    const dimensionValue = (value: unknown, kind: "frontage" | "depth") => { const prefix = kind === "frontage" ? /^(?:面寬|面)\s*/ : /^(?:深度|深)\s*/; const text = String(value || "").trim().replace(prefix, "").trim(); return /^(?:[-—–－_／/、.．]|無|未填|沒有|0)?$/.test(text) ? "" : text; };
      const rows = ordered.map(item => { const record = sourceOf(item); const price = String(record.price || "").replace(/\s*萬(?:元)?\s*/g, ""), frontage = dimensionValue(record.frontage, "frontage"), depth = dimensionValue(record.depth, "depth"), vertical = (value: unknown) => Array.from(String(value || "").trim()).join("\n"); return [item.sequence, `${areaCategory(record) || record.area || ""}\n${typeShort(record.type) || ""}`, record.caseName || "", chunkText(record.address || "", 12).join("\n"), price ? `${price}\n萬` : "", vertical(record.direction), imageAge(record), floorShortFixed(record.floor || ""), imageLayout(record), record.indoorPing || "", record.buildingPing || "", record.landPing || "", parkingShort(record.parking || ""), record.managementFee || "", record.key || "", vertical(record.currentState), record.road || "", [frontage ? `面${frontage}` : "", depth ? `深${depth}` : ""].filter(Boolean).join("\n"), record.zoning || "", `${record.coverage || ""}\n${record.far || ""}`, imageDeveloper(record.developer || ""), displayNoteSegments(record.notes || "").join("； ")]; });
    const margin = 32, headerHeight = 112, titleHeight = 132, fontSize = 30, lineHeight = 38;
    const wrapLines = (text: string, width: number) => { const perLine = Math.max(1, Math.floor((width - 12) / fontSize)); return String(text || "—").split("\n").flatMap(part => Array.from({ length: Math.max(1, Math.ceil(Array.from(part || "—").length / perLine)) }, (_, index) => Array.from(part || "—").slice(index * perLine, (index + 1) * perLine).join(""))); };
    const rowHeights = rows.map(row => Math.max(58, ...row.map((text, index) => wrapLines(text, widths[index]).length * lineHeight + 18)));
    const tableWidth = widths.reduce((sum, width) => sum + width, 0);
    const canvas = document.createElement("canvas"); canvas.width = tableWidth + margin * 2; canvas.height = titleHeight + headerHeight + rowHeights.reduce((sum, height) => sum + height, 0) + margin * 2;
    const context = canvas.getContext("2d"); if (!context) return notify("圖片產生失敗"); context.fillStyle = "#ffffff"; context.fillRect(0, 0, canvas.width, canvas.height); context.textBaseline = "middle"; context.strokeStyle = "#8fa69e"; context.lineWidth = 1.5;
    context.fillStyle = "#000000"; context.font = "bold 50px Microsoft JhengHei"; context.textAlign = "center"; context.fillText(tourTitle || "團看", canvas.width / 2, margin + titleHeight / 2);
    const drawCell = (text: string, x: number, y: number, width: number, height: number, bold = false) => { context.strokeRect(x, y, width, height); context.fillStyle = "#000000"; context.font = `${bold ? "bold " : ""}${fontSize}px Microsoft JhengHei`; context.textAlign = "center"; const lines = wrapLines(text, width); lines.forEach((line, index) => context.fillText(line, x + width / 2, y + height / 2 + (index - (lines.length - 1) / 2) * lineHeight)); };
    const drawSizedCell = (text: string, x: number, y: number, width: number, height: number, size: number, bold = false, singleLine = false) => { context.strokeRect(x, y, width, height); context.fillStyle = "#000000"; context.font = `${bold ? "bold " : ""}${size}px Microsoft JhengHei`; context.textAlign = "center"; const lines = singleLine ? [String(text || "—").replace(/\s*\n\s*/g, "")] : String(text || "—").split("\n").filter(Boolean); const sizedLineHeight = Math.max(lineHeight, size + 4); lines.forEach((line, index) => context.fillText(line, x + width / 2, y + height / 2 + (index - (lines.length - 1) / 2) * sizedLineHeight)); };
    const drawSingleLineCell = (text: string, x: number, y: number, width: number, height: number, bold = false, size = 14) => { context.strokeRect(x, y, width, height); context.fillStyle = "#000000"; context.font = `${bold ? "bold " : ""}${size}px Microsoft JhengHei`; context.textAlign = "center"; context.fillText(String(text || "—").replace(/\s*\n\s*/g, ""), x + width / 2, y + height / 2); };
    const drawAgeCell = (text: string, x: number, y: number, width: number, height: number) => { context.strokeRect(x, y, width, height); const match = String(text).match(/^(.*?年建)（(.*?年屋)）$/); if (!match) return drawCell(text, x, y, width, height); context.fillStyle = "#000000"; context.textAlign = "center"; context.font = `${fontSize}px Microsoft JhengHei`; context.fillText(match[1], x + width / 2, y + height / 2 - 12); context.font = "15px Microsoft JhengHei"; context.fillText(match[2], x + width / 2, y + height / 2 + 14); };
    const drawFrontageDepthCell = (text: string, x: number, y: number, width: number, height: number) => { if (!String(text || "").trim() || text === "—") return drawCell("—", x, y, width, height); context.strokeRect(x, y, width, height); const lines = String(text).split("\n").filter(Boolean); context.fillStyle = "#000000"; context.textAlign = "left"; lines.forEach((line, index) => { const prefix = line.slice(0, 1), value = line.slice(1) || "—"; context.font = "14px Microsoft JhengHei"; const prefixWidth = context.measureText(prefix).width; context.font = "34px Microsoft JhengHei"; const valueWidth = context.measureText(value).width; const startX = x + (width - prefixWidth - valueWidth) / 2; const lineY = y + height / 2 + (index - (lines.length - 1) / 2) * lineHeight; context.font = "14px Microsoft JhengHei"; context.fillText(prefix, startX, lineY + 1); context.font = "34px Microsoft JhengHei"; context.fillText(value, startX + prefixWidth, lineY); }); };
    const drawNotesCell = (text: string, x: number, y: number, width: number, height: number) => { context.strokeRect(x, y, width, height); context.fillStyle = "#000000"; context.font = `${fontSize}px Microsoft JhengHei`; context.textAlign = "left"; const lines = wrapLines(text, width); lines.forEach((line, index) => context.fillText(line, x + 8, y + 12 + fontSize / 2 + index * lineHeight)); };
    let y = margin + titleHeight, x = margin; context.fillStyle = "#dcebe5"; context.fillRect(margin, y, tableWidth, headerHeight); headers.forEach((text, index) => { drawCell(text, x, y, widths[index], headerHeight, true); x += widths[index]; }); y += headerHeight;
    rows.forEach((row, rowIndex) => { x = margin; row.forEach((text, index) => { if (index === 4) drawSizedCell(text || "—", x, y, widths[index], rowHeights[rowIndex], 25, true); else if (index === 6) drawAgeCell(text || "—", x, y, widths[index], rowHeights[rowIndex]); else if (index === 8) drawSizedCell(text || "—", x, y, widths[index], rowHeights[rowIndex], 28, false); else if ([9, 10, 11].includes(index)) drawSizedCell(text || "—", x, y, widths[index], rowHeights[rowIndex], 24, false, true); else if (index === 17) drawFrontageDepthCell(text, x, y, widths[index], rowHeights[rowIndex]); else if (index === 20) drawSizedCell(text || "—", x, y, widths[index], rowHeights[rowIndex], 24, true); else if (index === 21) drawNotesCell(text || "—", x, y, widths[index], rowHeights[rowIndex]); else drawCell(text || "—", x, y, widths[index], rowHeights[rowIndex], index === 2); x += widths[index]; }); y += rowHeights[rowIndex]; });
    try {
      const link = document.createElement("a");
      link.href = canvas.toDataURL("image/png");
      link.download = `${(tourTitle || "團看路線").replace(/[\\/:*?"<>|]/g, "-")}.png`;
      document.body.appendChild(link); link.click(); link.remove();
      notify("已產生團看路線圖片");
    } catch { notify("圖片下載失敗，請再試一次"); }
    } catch (error) { console.error(error); notify("圖片產生失敗，請檢查團看欄位內容後再試一次"); }
  };
  return <>
    <section className="content tour-page">
      <div className="list-head"><SectionTitle title="團看安排" subtitle="挑選尚未團看的委託中物件，再依本次路線自行編排序號"/><div className="tour-head-actions"><button onClick={createNewTour}>＋ 新增本次團看</button><label className="tour-title-field">圖片標題<input type="text" value={tourTitle} onChange={event => setTourTitle(event.target.value)} placeholder="例如：115.08.13團看"/></label><button onClick={() => setPickerOpen(true)}>＋ 尚未團看物件</button><button onClick={openTemporary}>＋ 臨時案件</button></div></div>
      <div className="tour-list-head"><h3>本次團看路線</h3><div className="tour-copy-actions"><span>{items.length} 筆</span><button onClick={copyRoute}>複製文字</button><button className="primary" onClick={exportTourImage}>產生圖片（本次團看路線）</button></div></div>
      <div className="table-wrap tour-table"><table><thead><tr><th className="tour-remove"></th><th className="tour-sequence">序</th>{tourDisplayColumns.map(column => <th key={column} className={`tour-${column}`}>{tourColumnLabel[column].split("\n").map((line, index) => <Fragment key={line}>{index > 0 && <br/>}{line}</Fragment>)}</th>)}</tr></thead><tbody>{ordered.map(item => { const record = sourceOf(item); return <tr key={item.id}><td className="tour-remove"><button className="danger" title="移除" onClick={() => setItems(items.filter(current => current.id !== item.id))}>刪</button></td><td className="tour-sequence"><input type="number" min="1" value={item.sequence} onChange={event => updateItem(item.id, { sequence: event.target.value })}/></td>{tourDisplayColumns.map(column => <td key={column} className={`tour-${column}`}>{column === "caseName" ? <button className="case-link" onClick={() => record._intakeDraftId ? editRecord({ ...record, _editSource: "draft" }) : item.temporary ? editTemporary(item) : editRecord(record)}>{record.caseName || "未命名案件"}</button> : renderValue(record, column)}</td>)}</tr>;})}</tbody></table>{!items.length && <div className="contact-empty">目前尚未安排團看</div>}</div>
      <div className="tour-page-bottom"><div className="tour-complete-actions"><label>團看日期<input type="date" value={tourDate} onChange={event => setTourDate(event.target.value)}/></label><button className="primary" disabled={!items.length || !tourDate} onClick={() => { if (confirm(`確定完成 ${items.length} 筆團看，並將 ${displayRocDate(tourDate)} 寫入物件與草稿的團看日期？`)) { saveTourHistory(); complete(tourDate, items.map(item => item.recordId).filter(Boolean) as string[], items.map(item => item.data._intakeDraftId).filter(Boolean)); } }}>完成團看</button></div></div>
      {history.length > 0 && <section className="tour-history"><h3>歷史團看路線</h3>{history.map(entry => <details key={entry.id}><summary>{displayRocDate(entry.date)}　{entry.title}　{entry.items.length} 筆</summary><div className="tour-history-actions"><button className="primary" onClick={() => restoreTourHistory(entry)}>叫回此路線</button><span>叫回後可直接再次產生圖片或列印</span></div><ol>{entry.items.slice().sort((a, b) => Number(a.sequence || 9999) - Number(b.sequence || 9999)).map(item => <li key={item.id}><b>{item.data.caseName || "未命名案件"}</b><span>{item.data.address || ""}</span><em>{developerFullNameText(item.data.developer) || ""}</em></li>)}</ol></details>)}</section>}
    </section>
    {pickerOpen && <div className="modal-backdrop" onMouseDown={event => event.target === event.currentTarget && setPickerOpen(false)}><div className="modal tour-picker-modal"><div className="modal-head"><div><span>目前尚未安排團看</span><h2>選擇本次團看物件</h2></div><button className="close" onClick={() => setPickerOpen(false)}>×</button></div><div className="tour-picker-filters"><input value={search} onChange={event => setSearch(event.target.value)} placeholder="搜尋案名、地址、開發"/><label><span>進案日期起</span><input type="date" value={reportDateFrom} onChange={event => setReportDateFrom(event.target.value)}/></label><label><span>進案日期迄</span><input type="date" value={reportDateTo} onChange={event => setReportDateTo(event.target.value)}/></label><select value={developerFilter} onChange={event => setDeveloperFilter(event.target.value)}><option value="">全部開發人員</option>{developers.map(name => <option key={name}>{name}</option>)}</select><select value={propertyKindFilter} onChange={event => setPropertyKindFilter(event.target.value)}><option value="">土地＋房屋</option><option value="土地">土地</option><option value="房屋">房屋</option></select><button type="button" onClick={() => { setSearch(""); setReportDateFrom(""); setReportDateTo(""); setDeveloperFilter(""); setPropertyKindFilter(""); }}>清除篩選</button></div><div className="tour-candidate-list"><div className="tour-candidate-head"><b>案名</b><b>地址</b><b>開發</b><b>開價</b><b>現況</b><b>進案日期</b><b></b></div>{available.map(record => <div className="tour-candidate-row" key={record.id}><button className="case-link candidate-case-link" onClick={() => editRecord(record._intakeDraftId ? { ...record, _editSource: "draft" } : record)}>{record.caseName || "未命名案件"}</button><span>{record.address}</span><em>{developerFullNameText(record.developer) || "未填開發"}</em><span>{record.price ? `${String(record.price).replace(/\s*萬(?:元)?\s*/g, "")}萬` : "—"}</span><span>{record.currentState || "—"}</span><span className={record._intakeDraftId ? "not-entered" : ""}>{record._intakeDraftId ? "尚未進案" : displayRocDate(record.reportDate) || "—"}</span><button className="candidate-add" onClick={() => addRecord(record)}>＋ 加入</button></div>)}{!available.length && <div className="contact-empty">沒有符合篩選條件的尚未團看物件</div>}</div>
<div className="modal-foot"><span>共 {available.length} 筆</span><button onClick={copyAvailable}>複製目前列表</button><button className="primary" onClick={() => setPickerOpen(false)}>完成挑選</button></div></div></div>}
    {temporaryOpen && <div className="modal-backdrop" onMouseDown={event => event.target === event.currentTarget && setTemporaryOpen(false)}><div className="modal tour-temporary-modal"><div className="modal-head"><div><span>團看安排 · 所有欄位皆可留空</span><h2>{temporaryEditId ? "編輯團看案件" : "新增臨時案件"}</h2><p>先填目前知道的資料，之後可再次點案名修改。</p></div><button className="close" onClick={() => setTemporaryOpen(false)}>×</button></div><div className="temporary-form-grid">{temporaryTourFields.map(([key, label]) => <label className={`field temporary-field temporary-${key}${key === "address" || key === "notes" ? " temporary-wide" : ""}`} key={key}><span>{label}</span>{key === "notes" ? <textarea rows={4} value={temporaryDraft[key] || ""} onChange={event => setTemporaryDraft({ ...temporaryDraft, [key]: event.target.value })}/> : key === "type" ? <select value={temporaryDraft[key] || ""} onChange={event => setTemporaryDraft({ ...temporaryDraft, [key]: event.target.value })}><option value="">請選擇</option><option>土地</option><option>透天</option><option>公寓</option><option>華廈</option><option>大樓</option><option>其他</option></select> : <input type="text" value={temporaryDraft[key] || ""} onChange={event => setTemporaryDraft({ ...temporaryDraft, [key]: event.target.value })}/>}</label>)}</div><div className="modal-foot temporary-modal-foot"><span>所有欄位都不是必填</span><div className="temporary-save-actions">{temporaryEditId && !temporaryDraft._intakeDraftId && <span className="temporary-save-note">尚未進案～臨時加入團看</span>}<button onClick={() => setTemporaryOpen(false)}>取消</button><button className="primary" onClick={addTemporary}>{temporaryEditId ? "儲存修改" : "加入本次團看"}</button></div></div></div></div>}
  </>;
}

function ContactDirectory({ people }: { people: Person[] }) {
  const bySequence = new Map<number, Person>();
  contactDirectoryOrder.forEach((name, index) => {
    const person = people.find(item => item.name.trim() === name);
    if (person) bySequence.set(index + 1, { ...person, sequence: String(index + 1), phone: contactDirectoryPhoneOverrides[name] || person.phone });
  });
  const groupStarts = [1, 12, 23];
  return <section className="contact-directory"><h3>台慶文化崇明店～同仁通訊錄～</h3><div className="contact-table-scroll"><table><thead><tr>{groupStarts.map(start => <Fragment key={start}><th>序</th><th>姓名</th><th>手機</th></Fragment>)}</tr></thead><tbody>{Array.from({ length: 11 }, (_, row) => <tr key={row}>{groupStarts.map(start => { const sequence = start + row; const person = bySequence.get(sequence); return <Fragment key={sequence}><td>{sequence}</td><td>{person?.name || ""}</td><td>{person?.phone ? <a href={`tel:${String(person.phone).replace(/\D/g, "")}`}>{displayPhone(person.phone)}</a> : ""}</td></Fragment>; })}</tr>)}</tbody></table></div></section>;
}

const intakeDraftStamp = (draft: IntakeData) => Date.parse(draft.modifiedAt || draft.createdAt || "") || 0;
const intakeDraftIdentity = (draft: IntakeData) => intakeValue(draft.values, "委託主約編號").trim() || draft.linkedRecordId || `${intakeValue(draft.values, "案名").trim()}|${intakeValue(draft.values, "物件(完整)地址").trim()}` || draft.id;
const reconcileIntakeDraftLinks = (drafts: IntakeData[], records: RecordItem[]) => {
  const byId = new Map(records.map(record => [String(record.id || ""), record]));
  const byNo = new Map(records.filter(record => String(record.propertyNo || "").trim()).map(record => [String(record.propertyNo).trim(), record]));
  const byCaseAddress = new Map(records.filter(record => String(record.caseName || "").trim() && String(record.address || "").trim()).map(record => [`${String(record.caseName).trim()}|${String(record.address).trim()}`, record]));
  return drafts.map(draft => {
    if (draft.linkedRecordId && byId.has(draft.linkedRecordId)) return draft;
    const propertyNo = intakeValue(draft.values, "委託主約編號").trim();
    const caseAddress = `${intakeValue(draft.values, "案名").trim()}|${intakeValue(draft.values, "物件(完整)地址").trim()}`;
    const record = (propertyNo ? byNo.get(propertyNo) : undefined) || (caseAddress !== "|" ? byCaseAddress.get(caseAddress) : undefined);
    if (!record) return draft;
    const enteredAt = draft.enteredAt || record.reportDate || record.lastModifiedAt || new Date().toISOString();
    return { ...draft, linkedRecordId: record.id, enteredAt, modifiedAt: draft.modifiedAt || enteredAt };
  });
};
const mergeIntakeDrafts = (localDrafts: IntakeData[], cloudDrafts: IntakeData[]) => {
  const merged = new Map<string, IntakeData>();
  [...cloudDrafts, ...localDrafts].forEach(draft => {
    if (!draft?.id) return;
    const key = intakeDraftIdentity(draft);
    const existing = merged.get(key);
    if (!existing || intakeDraftStamp(draft) >= intakeDraftStamp(existing)) merged.set(key, draft);
  });
  return [...merged.values()].sort((a, b) => intakeDraftStamp(b) - intakeDraftStamp(a));
};

const keySummaryLeft = [1, 2, 3, 5, 6, 7, 8, 17, 18, 19, 20, 21, 22, 23, 24, 33, 34, 35, 36, 37, 38, 39];
const keySummaryRight = [49, 50, 51, 52, 53, 55, 56, 65, 66, 67, 68, 69, 70, 71, 72, 81, 82, 83, 85, 86, 87, 88];

async function downloadColorWorkbook(record: RecordItem, personnel: Person[] = []) {
  const spreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const response = await fetch(colorWorkbookTemplateUrl);
  if (!response.ok) throw new Error("無法讀取彩色表 Excel 範本");
  const zip = await JSZip.loadAsync(await response.arrayBuffer());
  const parser = new DOMParser();
  const serializer = new XMLSerializer();
  const sheetPath = "xl/worksheets/sheet1.xml"; // 回應：範本內所有彩色表的資料來源
  const sheetXml = await zip.file(sheetPath)?.async("string");
  if (!sheetXml) throw new Error("Excel 範本缺少資料工作表");
  const sheetDocument = parser.parseFromString(sheetXml, "application/xml");
  const row = Array.from(sheetDocument.getElementsByTagNameNS(spreadsheetNs, "row")).find(item => item.getAttribute("r") === "2");
  if (!row) throw new Error("Excel 範本缺少資料列");
  const setCell = (address: string, value: unknown) => {
    let cell = Array.from(row.getElementsByTagNameNS(spreadsheetNs, "c")).find(item => item.getAttribute("r") === address);
    if (!cell) {
      cell = sheetDocument.createElementNS(spreadsheetNs, "c");
      cell.setAttribute("r", address);
      row.appendChild(cell);
    }
    while (cell.firstChild) cell.removeChild(cell.firstChild);
    cell.setAttribute("t", "inlineStr");
    const inline = sheetDocument.createElementNS(spreadsheetNs, "is");
    const textNode = sheetDocument.createElementNS(spreadsheetNs, "t");
    textNode.setAttribute("xml:space", "preserve");
    textNode.textContent = String(value ?? "");
    inline.appendChild(textNode);
    cell.appendChild(inline);
  };
  const setWorksheetCell = (document: Document, address: string, value: unknown, fontSize = 0) => {
    const rowNumber = address.match(/\d+/)?.[0] || "";
    let targetRow = Array.from(document.getElementsByTagNameNS(spreadsheetNs, "row")).find(item => item.getAttribute("r") === rowNumber);
    if (!targetRow) return;
    let cell = Array.from(targetRow.getElementsByTagNameNS(spreadsheetNs, "c")).find(item => item.getAttribute("r") === address);
    if (!cell) { cell = document.createElementNS(spreadsheetNs, "c"); cell.setAttribute("r", address); targetRow.appendChild(cell); }
    while (cell.firstChild) cell.removeChild(cell.firstChild);
    cell.setAttribute("t", "inlineStr");
    const inline = document.createElementNS(spreadsheetNs, "is");
    if (fontSize) {
      const run = document.createElementNS(spreadsheetNs, "r");
      const properties = document.createElementNS(spreadsheetNs, "rPr");
      const size = document.createElementNS(spreadsheetNs, "sz"); size.setAttribute("val", String(fontSize)); properties.appendChild(size);
      const textNode = document.createElementNS(spreadsheetNs, "t"); textNode.setAttribute("xml:space", "preserve"); textNode.textContent = String(value ?? "");
      run.appendChild(properties); run.appendChild(textNode); inline.appendChild(run);
    } else {
      const textNode = document.createElementNS(spreadsheetNs, "t"); textNode.setAttribute("xml:space", "preserve"); textNode.textContent = String(value ?? ""); inline.appendChild(textNode);
    }
    cell.appendChild(inline);
  };
  const setWorksheetNumberCell = (document: Document, address: string, value: number) => {
    const rowNumber = address.match(/\d+/)?.[0] || "";
    const targetRow = Array.from(document.getElementsByTagNameNS(spreadsheetNs, "row")).find(item => item.getAttribute("r") === rowNumber);
    if (!targetRow) return;
    let cell = Array.from(targetRow.getElementsByTagNameNS(spreadsheetNs, "c")).find(item => item.getAttribute("r") === address);
    if (!cell) { cell = document.createElementNS(spreadsheetNs, "c"); cell.setAttribute("r", address); targetRow.appendChild(cell); }
    while (cell.firstChild) cell.removeChild(cell.firstChild);
    cell.removeAttribute("t");
    const number = document.createElementNS(spreadsheetNs, "v"); number.textContent = String(value); cell.appendChild(number);
  };
  const setWorksheetRichCell = (document: Document, address: string, value: unknown, options: { size?: number; font?: string; color?: string; bold?: boolean } = {}) => {
    const rowNumber = address.match(/\d+/)?.[0] || "";
    const targetRow = Array.from(document.getElementsByTagNameNS(spreadsheetNs, "row")).find(item => item.getAttribute("r") === rowNumber);
    if (!targetRow) return;
    let cell = Array.from(targetRow.getElementsByTagNameNS(spreadsheetNs, "c")).find(item => item.getAttribute("r") === address);
    if (!cell) { cell = document.createElementNS(spreadsheetNs, "c"); cell.setAttribute("r", address); targetRow.appendChild(cell); }
    while (cell.firstChild) cell.removeChild(cell.firstChild);
    cell.setAttribute("t", "inlineStr");
    const inline = document.createElementNS(spreadsheetNs, "is"); const run = document.createElementNS(spreadsheetNs, "r"); const properties = document.createElementNS(spreadsheetNs, "rPr");
    if (options.bold) properties.appendChild(document.createElementNS(spreadsheetNs, "b"));
    if (options.size) { const size = document.createElementNS(spreadsheetNs, "sz"); size.setAttribute("val", String(options.size)); properties.appendChild(size); }
    if (options.font) { const font = document.createElementNS(spreadsheetNs, "rFont"); font.setAttribute("val", options.font); properties.appendChild(font); }
    if (options.color) { const color = document.createElementNS(spreadsheetNs, "color"); color.setAttribute("rgb", options.color); properties.appendChild(color); }
    const textNode = document.createElementNS(spreadsheetNs, "t"); textNode.setAttribute("xml:space", "preserve"); textNode.textContent = String(value ?? "");
    run.appendChild(properties); run.appendChild(textNode); inline.appendChild(run); cell.appendChild(inline);
  };
  const replaceMergedField = (document: Document, original: string, labelRange: string, valueRange: string, oldValueCell: string, newValueCell: string) => {
    const mergeCells = document.getElementsByTagNameNS(spreadsheetNs, "mergeCells")[0];
    if (!mergeCells) return;
    const oldMerge = Array.from(mergeCells.getElementsByTagNameNS(spreadsheetNs, "mergeCell")).find(node => node.getAttribute("ref") === original);
    if (!oldMerge) return;
    const oldCell = Array.from(document.getElementsByTagNameNS(spreadsheetNs, "c")).find(node => node.getAttribute("r") === oldValueCell);
    const rowNumber = newValueCell.match(/\d+/)?.[0] || "";
    const targetRow = Array.from(document.getElementsByTagNameNS(spreadsheetNs, "row")).find(node => node.getAttribute("r") === rowNumber);
    if (targetRow) {
      let newCell = Array.from(targetRow.getElementsByTagNameNS(spreadsheetNs, "c")).find(node => node.getAttribute("r") === newValueCell);
      if (!newCell) { newCell = document.createElementNS(spreadsheetNs, "c"); newCell.setAttribute("r", newValueCell); targetRow.appendChild(newCell); }
      const style = oldCell?.getAttribute("s");
      if (style) newCell.setAttribute("s", style);
    }
    oldMerge.parentNode?.removeChild(oldMerge);
    [labelRange, valueRange].filter(ref => { const [start, end] = ref.split(":"); return !end || start !== end; }).forEach(ref => {
      const node = document.createElementNS(spreadsheetNs, "mergeCell");
      node.setAttribute("ref", ref); mergeCells.appendChild(node);
    });
    mergeCells.setAttribute("count", String(mergeCells.getElementsByTagNameNS(spreadsheetNs, "mergeCell").length));
  };
  const removeMergedRange = (document: Document, ref: string) => {
    const mergeCells = document.getElementsByTagNameNS(spreadsheetNs, "mergeCells")[0];
    const node = mergeCells && Array.from(mergeCells.getElementsByTagNameNS(spreadsheetNs, "mergeCell")).find(item => item.getAttribute("ref") === ref);
    node?.parentNode?.removeChild(node);
    if (mergeCells) mergeCells.setAttribute("count", String(mergeCells.getElementsByTagNameNS(spreadsheetNs, "mergeCell").length));
  };
  const cleanNumber = (value = "") => String(value || "").replace(/,/g, "").match(/-?\d+(?:\.\d+)?/)?.[0] || "";
  const truncated2 = (value = "") => { const number = Number(cleanNumber(value)); return Number.isFinite(number) && number !== 0 ? (Math.trunc(number * 100) / 100).toFixed(2) : ""; };
  const pingValue = (value = "") => `${truncated2(value)}坪`;
  const houseManagementFee = (value = "") => {
    const raw = String(value || "").trim();
    const amount = raw.replace(/元/g, "").replace(/[／/]\s*[年月]\s*/g, "").trim();
    const period = raw.match(/[／/]\s*([年月])/ )?.[1] || "月";
    return !amount || /^(?:無|\$?0(?:\.0+)?)$/.test(amount) ? "-/月" : `${amount}元/${period}`;
  };
  const houseCompletionDate = (value = "") => {
    const text = String(value || "").trim();
    const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (iso) return `${String(Number(iso[1]) - 1911).padStart(3, "0")}.${String(Number(iso[2])).padStart(2, "0")}.${String(Number(iso[3])).padStart(2, "0")}`;
    const roc = text.match(/^(\d{2,3})[.\/-](\d{1,2})[.\/-](\d{1,2})$/);
    return roc ? `${roc[1].padStart(3, "0")}.${roc[2].padStart(2, "0")}.${roc[3].padStart(2, "0")}` : text;
  };
  const houseBuildingAge = (value = "") => {
    const text = String(value || "").trim();
    const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    const roc = text.match(/^(\d{2,3})[.\/-](\d{1,2})[.\/-](\d{1,2})$/);
    const westernYear = iso ? Number(iso[1]) : roc ? Number(roc[1]) + 1911 : 0;
    const age = westernYear ? new Date().getFullYear() - westernYear : NaN;
    return Number.isFinite(age) && age >= 0 ? `約${age}年屋` : "";
  };
  const halfWidth = (value = "") => String(value || "").replace(/[０-９]/g, character => String.fromCharCode(character.charCodeAt(0) - 0xFEE0)).replace(/[～〜﹣－–—-]/g, "~");
  const currentFloorValue = (value = "") => {
    const normalized = halfWidth(value).replace(/\s+/g, "").replace(/^\u73fe\u6cc1/, "").trim();
    if (!normalized) return "";
    return normalized.startsWith("\u5728") ? normalized : `\u5728${normalized}`;
  };
  const taxValue = (value = "") => {
    const raw = String(value || "").trim();
    const number = cleanNumber(raw);
    if (!raw || (!number && !/[\d]/.test(raw))) return raw || "$0 依稅單為準";
    return !number || Number(number) === 0 ? "$0 依稅單為準" : `約$${Number(number).toLocaleString("en-US")}`;
  };
  const optionalPingValue = (value = "") => {
    const numeric = cleanNumber(value);
    return numeric && Number(numeric) === 0 ? "-" : pingValue(value);
  };
  const layoutNumber = (pattern: RegExp) => String(record.layout || "").match(pattern)?.[1] || "";
  const floorParts = String(record.floor || "").split(/[／/]/).map(value => value.trim()).filter(Boolean);
  const cleanNotes = colorSheetAttention(record.attentionNotes || "", "");
  const noteParts = [record.feature1, record.feature2, record.feature3, record.feature4].map(value => String(value || "").trim());
  const isLand = typeShort(record.type) === "土地" || /^(?:LG|LA)/i.test(record.propertyNo || "");
  const price = String(record.price || "").replace(/\s*萬(?:元)?\s*/g, "");
  const aboutMeter = (value = "") => { const number = cleanNumber(value); return number ? `約${Number(number)}米` : "－米"; };
  const cells: Record<string, string> = {
    A2: new Date().toLocaleString("zh-TW"), C2: developerFullNameText(record.developer) || "", D2: record.propertyNo || "",
    E2: displayRocDate(record.entrustStart || ""), F2: displayRocDate(record.entrustEnd || ""),
    G2: record.caseName || "", H2: record.address || "", I2: record.currentState || "", J2: record.key || "",
    K2: record.type || "", L2: price, M2: cleanNumber(record.registryBuildingPing || record.buildingPing), N2: cleanNumber(record.registryIndoorPing || record.indoorPing),
    O2: cleanNumber(record.landSharePing || record.landPing), P2: cleanNotes, Q2: cleanNumber(record.road), R2: cleanNumber(record.frontage),
    S2: cleanNumber(record.depth), T2: record.zoning || "", U2: [record.coverage, record.far].filter(Boolean).join("/"),
    V2: record.completionDate || record.builtYear || "", W2: layoutNumber(/(\d+)\s*房/), X2: layoutNumber(/(\d+)\s*廳/),
    Y2: layoutNumber(/(\d+)\s*衛(?:浴)?/), Z2: layoutNumber(/(\d+)\s*陽台/),
    AB2: record.titleFloor || floorParts[0] || "", AC2: record.currentFloor || floorParts[1] || floorParts[0] || "", AD2: record.buildingName || "", AE2: record.elevatorCount || "", AF2: record.unitsPerFloor || "",
    AG2: isLand ? record.direction || "" : "", AH2: isLand ? "" : record.direction || "", AI2: isLand ? "" : record.direction || "",
    AJ2: record.managementMethod || "", AK2: record.managementFee || "", AL2: record.parkingOwnership || "", AM2: record.parkingMethod || record.parkingType || record.parking || "", AR2: record.parkingNo || "",
    AS2: noteParts[0] || "", AT2: noteParts[1] || "", AU2: noteParts[2] || "", AV2: noteParts[3] || "",
    AW2: record.school || "", BA2: record.market || "", BB2: record.park || "", BC2: cleanNotes, BJ2: ""
  };
  Object.entries(cells).forEach(([address, value]) => setCell(address, value));
  zip.file(sheetPath, serializer.serializeToString(sheetDocument));

  // 開發欄位有時會只輸入名字後兩字並直接相連（例如「庭宇妤宸」）。
  // 下載表格時以通訊錄的全名比對，保留原輸入中的先後順序。
  const developerRaw = String(record.developer || "").replace(/\s/g, "");
  const matchedDeveloperNames = personnel.map(person => {
    const name = String(person.name || "").trim();
    if (!name) return null;
    const aliases = [name, name.slice(-2), ...(name === "王啟山" ? ["王總"] : [])].filter(alias => alias.length >= 2);
    const positions = aliases.map(alias => developerRaw.indexOf(alias)).filter(position => position >= 0);
    return positions.length ? { name, position: Math.min(...positions) } : null;
  }).filter((item): item is { name: string; position: number } => !!item).sort((a, b) => a.position - b.position).map(item => item.name);
  const developerNames = matchedDeveloperNames.length ? matchedDeveloperNames : developerNameLines(record.developer || "");
  const developerContactRows = developerNames.map(name => {
    const person = personnel.find(item => item.name === name || item.name.includes(name) || name.includes(item.name));
    return { name: person?.name || name, phone: person?.phone ? displayPhone(person.phone) : "" };
  });
  const developerNameWidth = Math.max(0, ...developerContactRows.map(item => Array.from(item.name).length));
  const developerContact = developerContactRows.map(item => `${item.name}${"　".repeat(Math.max(0, developerNameWidth - Array.from(item.name).length) + 1)}${item.phone || ""}`).join("\n");
  const developerNamesText = developerContactRows.map(item => item.name).join("\n");
  const developerPhonesText = developerContactRows.map(item => item.phone || "").join("\n");

  // Write the visible sales sheet with final values instead of leaving template
  // formulas such as =KEY文字!... in the cells.
  const visibleSheetPath = isLand ? "xl/worksheets/sheet8.xml" : "xl/worksheets/sheet3.xml";
  const visibleSheetXml = await zip.file(visibleSheetPath)?.async("string");
  if (visibleSheetXml) {
    const visibleDocument = parser.parseFromString(visibleSheetXml, "application/xml");
    if (!isLand) {
      ["A9:A11", "B9:F11", "G9:M11", "G7:J7", "K7:M7", "P44:R44"].forEach(ref => removeMergedRange(visibleDocument, ref));
      const mergeCells = visibleDocument.getElementsByTagNameNS(spreadsheetNs, "mergeCells")[0];
      ["A9:A11", "B9:F11", "G9:M11", "G7:J7", "K7:M7", "P44:R44"].forEach(ref => { const node = visibleDocument.createElementNS(spreadsheetNs, "mergeCell"); node.setAttribute("ref", ref); mergeCells.appendChild(node); });
      mergeCells.setAttribute("count", String(mergeCells.getElementsByTagNameNS(spreadsheetNs, "mergeCell").length));
      const allCells = Array.from(visibleDocument.getElementsByTagNameNS(spreadsheetNs, "c"));
      const copyStyle = (source: string, target: string) => { const sourceCell = allCells.find(cell => cell.getAttribute("r") === source); const targetCell = allCells.find(cell => cell.getAttribute("r") === target); const style = sourceCell?.getAttribute("s"); if (style && targetCell) targetCell.setAttribute("s", style); };
      copyStyle("C9", "G9"); copyStyle("L7", "K7");
    }
    const leftRows: number[] = [];
    const rightRows: number[] = [];
    leftRows.forEach(rowNumber => {
      removeMergedRange(visibleDocument, `F${rowNumber}:G${rowNumber}`);
      removeMergedRange(visibleDocument, `F${rowNumber}:J${rowNumber}`);
      removeMergedRange(visibleDocument, `F${rowNumber}:M${rowNumber}`);
      const wideValue = rowNumber >= 33 ? `D${rowNumber}:J${rowNumber}` : rowNumber >= 29 ? `D${rowNumber}:M${rowNumber}` : `D${rowNumber}:G${rowNumber}`;
      replaceMergedField(visibleDocument, `B${rowNumber}:E${rowNumber}`, `B${rowNumber}:C${rowNumber}`, wideValue, `C${rowNumber}`, `D${rowNumber}`);
    });
    rightRows.forEach(rowNumber => replaceMergedField(visibleDocument, `I${rowNumber}:L${rowNumber}`, `I${rowNumber}:J${rowNumber}`, `K${rowNumber}:M${rowNumber}`, `J${rowNumber}`, `K${rowNumber}`));
    const commonValues: Record<string, string> = isLand
      ? { G1: " 銷 售 資 料 表", B6: record.caseName || "", O6: record.address || "", D7: record.propertyNo || "", G7: "KEY編號:", K7: record.key || "", B9: "總價:", G9: `${price}${price ? "萬" : ""}` }
      : { G1: " 銷 售 資 料 表", B6: record.caseName || "", O6: record.address || "", D7: record.propertyNo || "", G7: "KEY編號:", K7: record.key || "", N7: `現況:${record.currentState || ""}`, A9: "價格", B9: "總價:", G9: `${price}${price ? "萬" : ""}` };
    const noParking = [record.parkingOwnership, record.parkingType, record.parkingMethod, record.parking].some(value => String(value || "").includes("無車位"));
    const directValues: Record<string, string> = isLand ? {
      ...commonValues,
      B13: "土地坪數:", F13: pingValue(record.landSharePing || record.landPing), H13: "", I13: "每坪單價:", M13: record.landPing && price ? `$${Math.round(Number(cleanNumber(price)) * 10000 / Number(cleanNumber(record.landSharePing || record.landPing) || 1)).toLocaleString("en-US")}` : "",
      B14: "臨路約:", F14: aboutMeter(record.road).replace(/^約/, "").replace(/^－/, "-"), I14: "座向:", M14: record.direction || "",
      B15: "面寬約:", F15: aboutMeter(record.frontage).replace(/^約/, "").replace(/^－/, "-"), I15: "使用分區:", M15: record.zoning || "",
      B16: "深度約:", F16: aboutMeter(record.depth).replace(/^約/, "").replace(/^－/, "-"), I16: "建蔽/容積:", M16: [record.coverage, record.far].filter(Boolean).join("/"),
      B19: "市場:", F19: record.market || "", B20: "公園:", F20: record.park || "", B21: "學校:", F21: schoolSummary(record),
      B23: "一般增值稅:", F23: taxValue(record.generalLandValueTax || ""), B24: "自用增值稅:", F24: taxValue(record.selfUseLandValueTax || ""),
      C26: noteParts[0], C27: noteParts[1], C28: noteParts[2], C29: noteParts[3], B30: cleanNotes, P33: developerContact,
    } : {
      ...commonValues, F13: pingValue(record.registryBuildingPing || record.buildingPing), M13: pingValue(record.landSharePing || record.landPing), F14: pingValue(record.registryIndoorPing || record.indoorPing), M14: pingValue(record.buildingOtherPing || record.basementPing),
      F15: pingValue(record.mainBuildingPing), M15: noParking ? "無車位" : record.parkingType || record.parking || "", F16: pingValue(record.auxiliaryBuildingPing), M16: noParking ? "" : record.parkingMethod || "", F17: pingValue(record.commonAreaPing), M17: noParking ? "" : record.parkingNo || "",
      F19: typeShort(record.type), M19: record.buildingName || "", F20: `${record.unitsPerFloor || ""}戶`, M20: `${record.elevatorCount || ""}部`, F21: record.managementMethod || "", M21: houseManagementFee(record.managementFee), F22: layoutForHouseWorkbook(record.layout || "", record.type),
      F23: record.titleFloor || floorParts[0] || "", M23: currentFloorValue(record.currentFloor || floorParts[1] || ""), F24: houseCompletionDate(record.completionDate || record.builtYear), M24: houseBuildingAge(record.completionDate || record.builtYear), F25: record.direction || "", M25: record.currentState || "",
      F27: aboutMeter(record.road), M27: [record.coverage, record.far].filter(Boolean).join("/"), F28: aboutMeter(record.frontage), M28: aboutMeter(record.depth), F29: record.market || "", F30: record.park || "", F31: schoolSummary(record),
      F33: taxValue(record.generalLandValueTax || ""), F34: taxValue(record.selfUseLandValueTax || ""), F35: "$- 依稅單為準", C37: noteParts[0], C38: noteParts[1], C39: noteParts[2], C40: noteParts[3], B41: cleanNotes, P44: developerContact,
    };
    // 注意事項內容格是 C30（土地）與 C41（房屋）；B 欄只是左側標題。
    directValues[isLand ? "C30" : "C41"] = cleanNotes;
    if (!isLand) {
      // 房屋 Excel：附屬、公設與建物其他坪數為 0 時直接顯示「-」。
      directValues.F16 = optionalPingValue(record.auxiliaryBuildingPing);
      directValues.F17 = optionalPingValue(record.commonAreaPing);
      directValues.M14 = optionalPingValue(record.buildingOtherPing || record.basementPing);
    }
    const featureAddresses = isLand ? ["C26", "C27", "C28", "C29"] : ["C37", "C38", "C39", "C40"];
    const longFeatureAddresses = new Set<string>();
    featureAddresses.forEach(address => { if (directValues[address]) { const normalized = String(directValues[address]).replace(/[\r\n]+/g, ""); if (Array.from(normalized).length > 27) longFeatureAddresses.add(address); directValues[address] = normalized; } });
    Object.entries(directValues).forEach(([address, value]) => setWorksheetCell(visibleDocument, address, value, isLand && address === "M15" ? 12 : 0));
    if (!isLand) {
      const f23 = Array.from(visibleDocument.getElementsByTagNameNS(spreadsheetNs, "c")).find(cell => cell.getAttribute("r") === "F23");
      const m23 = Array.from(visibleDocument.getElementsByTagNameNS(spreadsheetNs, "c")).find(cell => cell.getAttribute("r") === "M23");
      if (f23?.getAttribute("s") && m23) m23.setAttribute("s", f23.getAttribute("s") || "0");
    }
    // Separate 「約」 and the amount in tax rows: 「約」 stays at the left,
    // while the number keeps a centred field of its own.
    const taxRows = isLand
      ? [{ row: 23, value: taxValue(record.generalLandValueTax || "") }, { row: 24, value: taxValue(record.selfUseLandValueTax || "") }]
      : [{ row: 33, value: taxValue(record.generalLandValueTax || "") }, { row: 34, value: taxValue(record.selfUseLandValueTax || "") }];
    // 「約」直接和金額寫在同一個完整欄位，避免金額被窄欄縮小。
    taxRows.forEach(({ row, value }) => {
      const display = value.startsWith("約$") ? value.replace(/^約\$/, "約  ") : value;
      setWorksheetCell(visibleDocument, `F${row}`, display);
    });
    // The template already contains this company line as an editable text box.
    // Keep A4 empty so the cell text does not overlap the text box.
    setWorksheetCell(visibleDocument, "A4", "");
    setWorksheetRichCell(visibleDocument, "G9", `${price}${price ? "萬" : ""}`, { size: 36, font: "Microsoft JhengHei", bold: true, color: "FFFF0000" });
    if (!isLand) { setWorksheetRichCell(visibleDocument, "F13", directValues.F13 || "坪", { size: 16, bold: true }); setWorksheetRichCell(visibleDocument, "M15", directValues.M15 || "", { size: 14 }); }
    if (!isLand) setWorksheetRichCell(visibleDocument, "P44", developerContact, { size: 18, font: "DFKai-SB" });
    if (isLand) setWorksheetRichCell(visibleDocument, "P33", developerContact, { size: 12, font: "DFKai-SB" });
    const headerRow = Array.from(visibleDocument.getElementsByTagNameNS(spreadsheetNs, "row")).find(item => item.getAttribute("r") === "4");
    if (headerRow) { headerRow.setAttribute("ht", "20"); headerRow.setAttribute("customHeight", "1"); }
    const stylesXml = await zip.file("xl/styles.xml")?.async("string");
    if (stylesXml) {
      const stylesDocument = parser.parseFromString(stylesXml, "application/xml");
      const cellXfs = stylesDocument.getElementsByTagNameNS(spreadsheetNs, "cellXfs")[0];
      const fonts = stylesDocument.getElementsByTagNameNS(spreadsheetNs, "fonts")[0];
      const headerCell = Array.from(visibleDocument.getElementsByTagNameNS(spreadsheetNs, "c")).find(item => item.getAttribute("r") === "A4");
      const baseStyle = Number(headerCell?.getAttribute("s") || "0");
      const baseXf = cellXfs.children[baseStyle] || cellXfs.children[0];
      const newXf = baseXf.cloneNode(true) as Element;
      const headerFont = stylesDocument.createElementNS(spreadsheetNs, "font");
      const headerSize = stylesDocument.createElementNS(spreadsheetNs, "sz"); headerSize.setAttribute("val", "12"); headerFont.appendChild(headerSize);
      const headerName = stylesDocument.createElementNS(spreadsheetNs, "name"); headerName.setAttribute("val", "PMingLiU"); headerFont.appendChild(headerName);
      const headerFamily = stylesDocument.createElementNS(spreadsheetNs, "family"); headerFamily.setAttribute("val", "1"); headerFont.appendChild(headerFamily);
      const headerCharset = stylesDocument.createElementNS(spreadsheetNs, "charset"); headerCharset.setAttribute("val", "136"); headerFont.appendChild(headerCharset);
      fonts.appendChild(headerFont); fonts.setAttribute("count", String(fonts.children.length));
      newXf.setAttribute("fontId", String(fonts.children.length - 1)); newXf.setAttribute("applyFont", "1"); newXf.setAttribute("applyProtection", "1");
      let alignment = Array.from(newXf.getElementsByTagNameNS(spreadsheetNs, "alignment"))[0];
      if (!alignment) { alignment = stylesDocument.createElementNS(spreadsheetNs, "alignment"); newXf.appendChild(alignment); }
      alignment.removeAttribute("wrapText"); alignment.setAttribute("shrinkToFit", "1"); alignment.setAttribute("vertical", "center");
      let protection = Array.from(newXf.getElementsByTagNameNS(spreadsheetNs, "protection"))[0];
      if (!protection) { protection = stylesDocument.createElementNS(spreadsheetNs, "protection"); newXf.appendChild(protection); }
      protection.setAttribute("locked", "0"); protection.setAttribute("hidden", "0");
      cellXfs.appendChild(newXf); cellXfs.setAttribute("count", String(cellXfs.children.length));
      if (headerCell) headerCell.setAttribute("s", String(cellXfs.children.length - 1));
      if (isLand) {
        const mergeCells = visibleDocument.getElementsByTagNameNS(spreadsheetNs, "mergeCells")[0];
        const numberFormats = stylesDocument.getElementsByTagNameNS(spreadsheetNs, "numFmts")[0];
        const formatCode = '"約  "$#,##0';
        let formatId = Array.from(numberFormats?.getElementsByTagNameNS(spreadsheetNs, "numFmt") || []).find(item => item.getAttribute("formatCode") === formatCode)?.getAttribute("numFmtId") || "";
        if (!formatId && numberFormats) { const ids = Array.from(numberFormats.getElementsByTagNameNS(spreadsheetNs, "numFmt")).map(item => Number(item.getAttribute("numFmtId") || 163)); formatId = String(Math.max(163, ...ids) + 1); const format = stylesDocument.createElementNS(spreadsheetNs, "numFmt"); format.setAttribute("numFmtId", formatId); format.setAttribute("formatCode", formatCode); numberFormats.appendChild(format); numberFormats.setAttribute("count", String(numberFormats.children.length)); }
        taxRows.forEach(({ row, value }) => {
          const amount = Number(cleanNumber(value)); if (!value.startsWith("約") || !Number.isFinite(amount) || amount <= 0) return;
          removeMergedRange(visibleDocument, `F${row}:J${row}`);
          if (mergeCells) { const merge = visibleDocument.createElementNS(spreadsheetNs, "mergeCell"); merge.setAttribute("ref", `F${row}:J${row}`); mergeCells.appendChild(merge); mergeCells.setAttribute("count", String(mergeCells.getElementsByTagNameNS(spreadsheetNs, "mergeCell").length)); }
          setWorksheetNumberCell(visibleDocument, `F${row}`, amount);
          const cell = Array.from(visibleDocument.getElementsByTagNameNS(spreadsheetNs, "c")).find(item => item.getAttribute("r") === `F${row}`);
          if (!cell || !formatId) return;
          const sourceXf = cellXfs.children[Number(cell.getAttribute("s") || "0")] || cellXfs.children[0];
          const formattedXf = sourceXf.cloneNode(true) as Element;
          formattedXf.setAttribute("numFmtId", formatId); formattedXf.setAttribute("applyNumberFormat", "1");
          let alignment = Array.from(formattedXf.getElementsByTagNameNS(spreadsheetNs, "alignment"))[0]; if (!alignment) { alignment = stylesDocument.createElementNS(spreadsheetNs, "alignment"); formattedXf.appendChild(alignment); }
          alignment.setAttribute("horizontal", "center"); alignment.setAttribute("vertical", "center");
          cellXfs.appendChild(formattedXf); cell.setAttribute("s", String(cellXfs.children.length - 1));
        });
      }
      taxRows.forEach(({ row }) => {
        const applyTaxAlignment = (address: string, horizontal: "left" | "center") => {
          const cell = Array.from(visibleDocument.getElementsByTagNameNS(spreadsheetNs, "c")).find(item => item.getAttribute("r") === address);
          if (!cell) return;
          const sourceXf = cellXfs.children[Number(cell.getAttribute("s") || "0")] || cellXfs.children[0];
          const taxXf = sourceXf.cloneNode(true) as Element;
          let taxAlignment = Array.from(taxXf.getElementsByTagNameNS(spreadsheetNs, "alignment"))[0];
          if (!taxAlignment) { taxAlignment = stylesDocument.createElementNS(spreadsheetNs, "alignment"); taxXf.appendChild(taxAlignment); }
          taxAlignment.setAttribute("horizontal", horizontal); taxAlignment.setAttribute("vertical", "center"); taxAlignment.removeAttribute("indent");
          cellXfs.appendChild(taxXf); cell.setAttribute("s", String(cellXfs.children.length - 1));
        };
        applyTaxAlignment(`F${row}`, "center");
        ["F"].forEach(column => {
          const cell = Array.from(visibleDocument.getElementsByTagNameNS(spreadsheetNs, "c")).find(item => item.getAttribute("r") === `${column}${row}`);
          if (!cell) return;
          const sourceXf = cellXfs.children[Number(cell.getAttribute("s") || "0")] || cellXfs.children[0];
          const largerXf = sourceXf.cloneNode(true) as Element;
          const sourceFont = fonts.children[Number(sourceXf.getAttribute("fontId") || "0")] || fonts.children[0];
          const largerFont = sourceFont.cloneNode(true) as Element;
          let size = Array.from(largerFont.getElementsByTagNameNS(spreadsheetNs, "sz"))[0];
          if (!size) { size = stylesDocument.createElementNS(spreadsheetNs, "sz"); largerFont.appendChild(size); }
          size.setAttribute("val", "16");
          fonts.appendChild(largerFont); fonts.setAttribute("count", String(fonts.children.length));
          largerXf.setAttribute("fontId", String(fonts.children.length - 1)); largerXf.setAttribute("applyFont", "1");
          cellXfs.appendChild(largerXf); cell.setAttribute("s", String(cellXfs.children.length - 1));
        });
      });
      if (!isLand) {
        const priceCell = Array.from(visibleDocument.getElementsByTagNameNS(spreadsheetNs, "c")).find(item => item.getAttribute("r") === "G9");
        if (priceCell) {
          const priceBaseXf = cellXfs.children[Number(priceCell.getAttribute("s") || "0")] || cellXfs.children[0];
          const priceXf = priceBaseXf.cloneNode(true) as Element;
          let priceAlignment = Array.from(priceXf.getElementsByTagNameNS(spreadsheetNs, "alignment"))[0];
          if (!priceAlignment) { priceAlignment = stylesDocument.createElementNS(spreadsheetNs, "alignment"); priceXf.appendChild(priceAlignment); }
          priceAlignment.setAttribute("horizontal", "center"); priceAlignment.setAttribute("vertical", "center"); priceAlignment.removeAttribute("wrapText"); priceAlignment.removeAttribute("shrinkToFit"); priceAlignment.removeAttribute("justifyLastLine");
          cellXfs.appendChild(priceXf); priceCell.setAttribute("s", String(cellXfs.children.length - 1));
        }
      }
      if (isLand) {
        ["G23", "G24"].forEach(address => {
          const cell = Array.from(visibleDocument.getElementsByTagNameNS(spreadsheetNs, "c")).find(item => item.getAttribute("r") === address);
          if (!cell) return;
          const sourceXf = cellXfs.children[Number(cell.getAttribute("s") || "0")] || cellXfs.children[0];
          const amountXf = sourceXf.cloneNode(true) as Element;
          const sourceFont = fonts.children[Number(sourceXf.getAttribute("fontId") || "0")] || fonts.children[0];
          const amountFont = sourceFont.cloneNode(true) as Element;
          let amountSize = Array.from(amountFont.getElementsByTagNameNS(spreadsheetNs, "sz"))[0];
          if (!amountSize) { amountSize = stylesDocument.createElementNS(spreadsheetNs, "sz"); amountFont.appendChild(amountSize); }
          amountSize.setAttribute("val", "16");
          fonts.appendChild(amountFont); fonts.setAttribute("count", String(fonts.children.length));
          amountXf.setAttribute("fontId", String(fonts.children.length - 1)); amountXf.setAttribute("applyFont", "1");
          cellXfs.appendChild(amountXf); cell.setAttribute("s", String(cellXfs.children.length - 1));
        });
        const landZoningCell = Array.from(visibleDocument.getElementsByTagNameNS(spreadsheetNs, "c")).find(item => item.getAttribute("r") === "M15");
        if (landZoningCell) {
          const sourceXf = cellXfs.children[Number(landZoningCell.getAttribute("s") || "0")] || cellXfs.children[0];
          const font12Xf = sourceXf.cloneNode(true) as Element;
          const sourceFont = fonts.children[Number(sourceXf.getAttribute("fontId") || "0")] || fonts.children[0];
          const font12 = sourceFont.cloneNode(true) as Element;
          let size = Array.from(font12.getElementsByTagNameNS(spreadsheetNs, "sz"))[0];
          if (!size) { size = stylesDocument.createElementNS(spreadsheetNs, "sz"); font12.appendChild(size); }
          size.setAttribute("val", "12");
          fonts.appendChild(font12); fonts.setAttribute("count", String(fonts.children.length));
          font12Xf.setAttribute("fontId", String(fonts.children.length - 1)); font12Xf.setAttribute("applyFont", "1");
          cellXfs.appendChild(font12Xf); landZoningCell.setAttribute("s", String(cellXfs.children.length - 1));
        }
      }
      ["F33", "F34", "F35"].forEach(address => {
        const cell = Array.from(visibleDocument.getElementsByTagNameNS(spreadsheetNs, "c")).find(item => item.getAttribute("r") === address);
        if (!cell) return;
        const sourceXf = cellXfs.children[Number(cell.getAttribute("s") || "0")] || cellXfs.children[0];
        const centeredXf = sourceXf.cloneNode(true) as Element;
        let centeredAlignment = Array.from(centeredXf.getElementsByTagNameNS(spreadsheetNs, "alignment"))[0];
        if (!centeredAlignment) { centeredAlignment = stylesDocument.createElementNS(spreadsheetNs, "alignment"); centeredXf.appendChild(centeredAlignment); }
        centeredAlignment.setAttribute("horizontal", "center"); centeredAlignment.setAttribute("vertical", "center");
        cellXfs.appendChild(centeredXf); cell.setAttribute("s", String(cellXfs.children.length - 1));
      });
      featureAddresses.forEach(address => {
        const cell = Array.from(visibleDocument.getElementsByTagNameNS(spreadsheetNs, "c")).find(item => item.getAttribute("r") === address);
        if (!cell) return;
        const sourceXf = cellXfs.children[Number(cell.getAttribute("s") || "0")] || cellXfs.children[0];
        const wrappedXf = sourceXf.cloneNode(true) as Element;
        const sourceFont = fonts.children[Number(sourceXf.getAttribute("fontId") || "0")] || fonts.children[0];
        const featureFont = sourceFont.cloneNode(true) as Element;
        let featureSize = Array.from(featureFont.getElementsByTagNameNS(spreadsheetNs, "sz"))[0];
        if (!featureSize) { featureSize = stylesDocument.createElementNS(spreadsheetNs, "sz"); featureFont.appendChild(featureSize); }
        featureSize.setAttribute("val", longFeatureAddresses.has(address) ? "12" : "14");
        fonts.appendChild(featureFont); fonts.setAttribute("count", String(fonts.children.length));
        wrappedXf.setAttribute("fontId", String(fonts.children.length - 1)); wrappedXf.setAttribute("applyFont", "1");
        let wrappedAlignment = Array.from(wrappedXf.getElementsByTagNameNS(spreadsheetNs, "alignment"))[0];
        if (!wrappedAlignment) { wrappedAlignment = stylesDocument.createElementNS(spreadsheetNs, "alignment"); wrappedXf.appendChild(wrappedAlignment); }
        if (longFeatureAddresses.has(address)) wrappedAlignment.setAttribute("wrapText", "1"); else wrappedAlignment.removeAttribute("wrapText");
        wrappedAlignment.setAttribute("vertical", "center");
        cellXfs.appendChild(wrappedXf); cell.setAttribute("s", String(cellXfs.children.length - 1));
      });
      if (!isLand) {
        // 市場／公園／學校：每列固定 0.95 公分，內容過長時在格內換行。
        [29, 30, 31].forEach(rowNumber => {
          const rowNode = Array.from(visibleDocument.getElementsByTagNameNS(spreadsheetNs, "row")).find(item => item.getAttribute("r") === String(rowNumber));
          if (rowNode) { rowNode.setAttribute("ht", "33.82"); rowNode.setAttribute("customHeight", "1"); }
        });
        ["F29", "F30", "F31"].forEach(address => {
          const cell = Array.from(visibleDocument.getElementsByTagNameNS(spreadsheetNs, "c")).find(item => item.getAttribute("r") === address);
          if (!cell) return;
          const sourceXf = cellXfs.children[Number(cell.getAttribute("s") || "0")] || cellXfs.children[0];
          const wrappedXf = sourceXf.cloneNode(true) as Element;
          let alignment = Array.from(wrappedXf.getElementsByTagNameNS(spreadsheetNs, "alignment"))[0];
          if (!alignment) { alignment = stylesDocument.createElementNS(spreadsheetNs, "alignment"); wrappedXf.appendChild(alignment); }
          alignment.setAttribute("wrapText", "1"); alignment.setAttribute("vertical", "center"); alignment.setAttribute("horizontal", "left");
          cellXfs.appendChild(wrappedXf); cell.setAttribute("s", String(cellXfs.children.length - 1));
        });
      }
      [isLand ? "P33" : "P44"].forEach(address => {
        const cell = Array.from(visibleDocument.getElementsByTagNameNS(spreadsheetNs, "c")).find(item => item.getAttribute("r") === address);
        if (!cell) return;
        const sourceXf = cellXfs.children[Number(cell.getAttribute("s") || "0")] || cellXfs.children[0];
        const contactXf = sourceXf.cloneNode(true) as Element;
        const sourceFont = fonts.children[Number(sourceXf.getAttribute("fontId") || "0")] || fonts.children[0];
        const contactFont = sourceFont.cloneNode(true) as Element;
        let contactSize = Array.from(contactFont.getElementsByTagNameNS(spreadsheetNs, "sz"))[0];
        if (!contactSize) { contactSize = stylesDocument.createElementNS(spreadsheetNs, "sz"); contactFont.appendChild(contactSize); }
        contactSize.setAttribute("val", isLand ? "16" : "18");
        if (isLand && address === "P33") {
          let contactName = Array.from(contactFont.getElementsByTagNameNS(spreadsheetNs, "name"))[0];
          if (!contactName) { contactName = stylesDocument.createElementNS(spreadsheetNs, "name"); contactFont.appendChild(contactName); }
          contactName.setAttribute("val", "標楷體");
        }
        fonts.appendChild(contactFont); fonts.setAttribute("count", String(fonts.children.length));
        contactXf.setAttribute("fontId", String(fonts.children.length - 1)); contactXf.setAttribute("applyFont", "1");
        let contactAlignment = Array.from(contactXf.getElementsByTagNameNS(spreadsheetNs, "alignment"))[0];
        if (!contactAlignment) { contactAlignment = stylesDocument.createElementNS(spreadsheetNs, "alignment"); contactXf.appendChild(contactAlignment); }
        contactAlignment.setAttribute("horizontal", "center"); contactAlignment.removeAttribute("indent"); contactAlignment.setAttribute("vertical", "center"); contactAlignment.setAttribute("wrapText", "1");
        cellXfs.appendChild(contactXf); cell.setAttribute("s", String(cellXfs.children.length - 1));
        if (isLand && address === "P33") {
          const inline = Array.from(cell.getElementsByTagNameNS(spreadsheetNs, "is"))[0];
          const currentText = Array.from(cell.getElementsByTagNameNS(spreadsheetNs, "t")).map(node => node.textContent || "").join("");
          if (inline) {
            while (inline.firstChild) inline.removeChild(inline.firstChild);
            const run = visibleDocument.createElementNS(spreadsheetNs, "r");
            const runProperties = visibleDocument.createElementNS(spreadsheetNs, "rPr");
            const runFont = visibleDocument.createElementNS(spreadsheetNs, "rFont"); runFont.setAttribute("val", "標楷體");
            const runSize = visibleDocument.createElementNS(spreadsheetNs, "sz"); runSize.setAttribute("val", "16");
            runProperties.appendChild(runFont); runProperties.appendChild(runSize);
            const runText = visibleDocument.createElementNS(spreadsheetNs, "t"); runText.setAttribute("xml:space", "preserve"); runText.textContent = currentText;
            run.appendChild(runProperties); run.appendChild(runText); inline.appendChild(run);
          }
        }
      });
      cellXfs.setAttribute("count", String(cellXfs.children.length));
      zip.file("xl/styles.xml", serializer.serializeToString(stylesDocument));
    }
    // Write the physical row height again after all template adjustments.
    // This template's Excel display scale uses 33.82 pt for a 0.95 cm row.
    if (!isLand) [29, 30, 31].forEach(rowNumber => {
      const rowNode = Array.from(visibleDocument.getElementsByTagNameNS(spreadsheetNs, "row")).find(item => item.getAttribute("r") === String(rowNumber));
      if (rowNode) { rowNode.setAttribute("ht", "33.82"); rowNode.setAttribute("customHeight", "1"); }
    });
    (isLand ? [26, 27, 28, 29] : [37, 38, 39, 40]).forEach(rowNumber => {
      const rowNode = Array.from(visibleDocument.getElementsByTagNameNS(spreadsheetNs, "row")).find(item => item.getAttribute("r") === String(rowNumber));
      if (rowNode) { rowNode.setAttribute("ht", isLand ? "48.40" : "28.35"); rowNode.setAttribute("customHeight", "1"); }
    });
    Array.from(visibleDocument.getElementsByTagNameNS(spreadsheetNs, "f")).forEach(formula => {
      const cell = formula.parentNode as Element | null;
      if (!cell || cell.localName !== "c") return;
      while (cell.firstChild) cell.removeChild(cell.firstChild);
      cell.setAttribute("t", "inlineStr");
      const inline = visibleDocument.createElementNS(spreadsheetNs, "is");
      const textNode = visibleDocument.createElementNS(spreadsheetNs, "t"); textNode.textContent = "";
      inline.appendChild(textNode); cell.appendChild(inline);
    });
    zip.file(visibleSheetPath, serializer.serializeToString(visibleDocument));
  }

  const keyXml = await zip.file("xl/worksheets/sheet2.xml")?.async("string");
  if (keyXml) {
    const keyDocument = parser.parseFromString(keyXml, "application/xml");
    const keyRow = Array.from(keyDocument.getElementsByTagNameNS(spreadsheetNs, "row")).find(item => item.getAttribute("r") === "5");
    if (keyRow) {
      let contactCell = Array.from(keyRow.getElementsByTagNameNS(spreadsheetNs, "c")).find(item => item.getAttribute("r") === "F5");
      if (!contactCell) { contactCell = keyDocument.createElementNS(spreadsheetNs, "c"); contactCell.setAttribute("r", "F5"); keyRow.appendChild(contactCell); }
      while (contactCell.firstChild) contactCell.removeChild(contactCell.firstChild);
      contactCell.setAttribute("t", "inlineStr");
      const inline = keyDocument.createElementNS(spreadsheetNs, "is");
      const textNode = keyDocument.createElementNS(spreadsheetNs, "t");
      textNode.setAttribute("xml:space", "preserve"); textNode.textContent = developerContact;
      inline.appendChild(textNode); contactCell.appendChild(inline);
      zip.file("xl/worksheets/sheet2.xml", serializer.serializeToString(keyDocument));
    }
  }

  const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
  if (workbookXml) {
    const workbookDocument = parser.parseFromString(workbookXml, "application/xml");
    const targetName = isLand ? "土地資料表" : "房屋資料表";
    const sheets = Array.from(workbookDocument.getElementsByTagNameNS(spreadsheetNs, "sheet"));
    sheets.forEach(sheet => {
      if (sheet.getAttribute("name") === targetName) sheet.removeAttribute("state");
      else sheet.setAttribute("state", "hidden");
    });
    const view = workbookDocument.getElementsByTagNameNS(spreadsheetNs, "workbookView")[0];
    if (view) view.setAttribute("activeTab", isLand ? "7" : "2");
    const calc = workbookDocument.getElementsByTagNameNS(spreadsheetNs, "calcPr")[0];
    if (calc) { calc.setAttribute("calcMode", "auto"); calc.setAttribute("fullCalcOnLoad", "1"); calc.setAttribute("forceFullCalc", "1"); }
    zip.file("xl/workbook.xml", serializer.serializeToString(workbookDocument));
  }
  // The template uses drawing shapes as the photo/place-map frames.  Make every
  // drawing selectable and editable in the downloaded workbook so users can
  // place, move and resize their own pictures without unlocking the sheet.
  const drawingPaths = Object.keys(zip.files).filter(path => /^xl\/drawings\/drawing\d+\.xml$/i.test(path));
  for (const drawingPath of drawingPaths) {
    const drawingXml = await zip.file(drawingPath)?.async("string");
    if (!drawingXml) continue;
    const drawingDocument = parser.parseFromString(drawingXml, "application/xml");
    if (isLand && /\/drawing4\.xml$/i.test(drawingPath)) {
      const xdrNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
      const fixedFrames: Record<string, { width: number; height: number }> = {
        "矩形 9": { width: 9.61, height: 7.22 },
        "矩形 10": { width: 9.63, height: 6.94 },
        "矩形 11": { width: 9.61, height: 7.12 },
      };
      Array.from(drawingDocument.getElementsByTagNameNS(xdrNs, "twoCellAnchor")).forEach(anchor => {
        const nameNode = Array.from(anchor.getElementsByTagNameNS(xdrNs, "cNvPr")).find(node => fixedFrames[node.getAttribute("name") || ""]);
        if (!nameNode) return;
        const dimensions = fixedFrames[nameNode.getAttribute("name") || ""];
        const replacement = drawingDocument.createElementNS(xdrNs, "xdr:oneCellAnchor");
        const from = Array.from(anchor.children).find(node => node.localName === "from");
        if (!from) return;
        replacement.appendChild(from.cloneNode(true));
        const ext = drawingDocument.createElementNS(xdrNs, "xdr:ext");
        ext.setAttribute("cx", String(Math.round(dimensions.width * 360000)));
        ext.setAttribute("cy", String(Math.round(dimensions.height * 360000)));
        replacement.appendChild(ext);
        Array.from(anchor.children).filter(node => node.localName !== "from" && node.localName !== "to").forEach(node => replacement.appendChild(node.cloneNode(true)));
        anchor.parentNode?.replaceChild(replacement, anchor);
      });
      // Keep the three blue captions compact.  They are fixed-size shapes so
      // later row-height changes cannot stretch them or push the last frame
      // into the contact block.
      Array.from(drawingDocument.getElementsByTagNameNS(xdrNs, "twoCellAnchor")).forEach(anchor => {
        const captionNode = Array.from(anchor.getElementsByTagNameNS(xdrNs, "cNvPr")).find(node => ["13", "14", "15"].includes(node.getAttribute("id") || ""));
        if (!captionNode) return;
        const replacement = drawingDocument.createElementNS(xdrNs, "xdr:oneCellAnchor");
        const from = Array.from(anchor.children).find(node => node.localName === "from");
        if (!from) return;
        replacement.appendChild(from.cloneNode(true));
        const ext = drawingDocument.createElementNS(xdrNs, "xdr:ext");
        ext.setAttribute("cx", String(Math.round(3.49 * 360000)));
        ext.setAttribute("cy", String(Math.round(1.31 * 360000)));
        replacement.appendChild(ext);
        Array.from(anchor.children).filter(node => node.localName !== "from" && node.localName !== "to").forEach(node => replacement.appendChild(node.cloneNode(true)));
        anchor.parentNode?.replaceChild(replacement, anchor);
      });
      const lastFrame = Array.from(drawingDocument.getElementsByTagNameNS(xdrNs, "oneCellAnchor")).find(anchor =>
        Array.from(anchor.getElementsByTagNameNS(xdrNs, "cNvPr")).some(node => node.getAttribute("id") === "12")
      );
      const lastFrameOffset = lastFrame ? Array.from(lastFrame.getElementsByTagNameNS(xdrNs, "rowOff"))[0] : null;
      if (lastFrameOffset) lastFrameOffset.textContent = String(Math.max(0, Number(lastFrameOffset.textContent || "0") - 180000));
      const fixedLandObjects: Record<string, { x: number; y: number; width: number; height: number; caption?: boolean }> = {
        "10": { x: 4546386, y: 2154621, width: Math.round(9.61 * 360000), height: Math.round(7.22 * 360000) },
        "11": { x: 4537421, y: 5185821, width: Math.round(9.63 * 360000), height: Math.round(6.94 * 360000) },
        "12": { x: 4546386, y: 8116221, width: Math.round(9.61 * 360000), height: Math.round(7.12 * 360000) },
        "15": { x: 4528457, y: 1758621, width: Math.round(3.49 * 360000), height: Math.round(1 * 360000), caption: true },
        "14": { x: 4528457, y: 4789821, width: Math.round(3.49 * 360000), height: Math.round(1 * 360000), caption: true },
        "13": { x: 4528457, y: 7720221, width: Math.round(3.49 * 360000), height: Math.round(1 * 360000), caption: true },
      };
      Array.from(drawingDocument.documentElement.children).forEach(anchor => {
        const idNode = Array.from(anchor.getElementsByTagNameNS(xdrNs, "cNvPr")).find(node => fixedLandObjects[node.getAttribute("id") || ""]);
        if (!idNode) return;
        const config = fixedLandObjects[idNode.getAttribute("id") || ""];
        const replacement = drawingDocument.createElementNS(xdrNs, "xdr:absoluteAnchor");
        const pos = drawingDocument.createElementNS(xdrNs, "xdr:pos");
        pos.setAttribute("x", String(config.x)); pos.setAttribute("y", String(config.y));
        const ext = drawingDocument.createElementNS(xdrNs, "xdr:ext");
        ext.setAttribute("cx", String(config.width)); ext.setAttribute("cy", String(config.height));
        replacement.appendChild(pos); replacement.appendChild(ext);
        Array.from(anchor.children).filter(node => !["from", "to", "pos", "ext"].includes(node.localName)).forEach(node => replacement.appendChild(node.cloneNode(true)));
        anchor.parentNode?.replaceChild(replacement, anchor);
        const transform = Array.from(replacement.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "xfrm"))[0];
        if (transform) {
          const offset = Array.from(transform.children).find(node => node.localName === "off");
          const shapeExtent = Array.from(transform.children).find(node => node.localName === "ext");
          if (offset) { offset.setAttribute("x", String(config.x)); offset.setAttribute("y", String(config.y)); }
          if (shapeExtent) { shapeExtent.setAttribute("cx", String(config.width)); shapeExtent.setAttribute("cy", String(config.height)); }
        }
        if (config.caption) {
          const body = Array.from(replacement.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "bodyPr"))[0];
          if (body) body.setAttribute("anchor", "ctr");
          Array.from(replacement.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "pPr")).forEach(node => node.setAttribute("algn", "ctr"));
        }
      });
    }
    if (!isLand && /\/drawing2\.xml$/i.test(drawingPath)) {
      const xdrNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
      const positionFrame = Array.from(drawingDocument.getElementsByTagNameNS(xdrNs, "twoCellAnchor")).find(anchor =>
        Array.from(anchor.getElementsByTagNameNS(xdrNs, "cNvPr")).some(node => node.getAttribute("id") === "22")
      );
      if (positionFrame) {
        const replacement = drawingDocument.createElementNS(xdrNs, "xdr:oneCellAnchor");
        const from = Array.from(positionFrame.children).find(node => node.localName === "from");
        if (from) {
          replacement.appendChild(from.cloneNode(true));
          const ext = drawingDocument.createElementNS(xdrNs, "xdr:ext");
          ext.setAttribute("cx", "3480345");
          ext.setAttribute("cy", "2645112");
          replacement.appendChild(ext);
          Array.from(positionFrame.children).filter(node => node.localName !== "from" && node.localName !== "to").forEach(node => replacement.appendChild(node.cloneNode(true)));
          positionFrame.parentNode?.replaceChild(replacement, positionFrame);
        }
      }
      // Keep the complete right-side photo / layout / location block together.
      // Move all three blue captions and their white frames up by 0.3 cm.
      const up = 180000; // 0.5 cm in EMU
      const rightBlockIds = new Set(["22", "33", "35", "32", "34", "36"]);
      Array.from(drawingDocument.documentElement.children).forEach(anchor => {
        const idNode = Array.from(anchor.getElementsByTagNameNS(xdrNs, "cNvPr")).find(node => rightBlockIds.has(node.getAttribute("id") || ""));
        if (!idNode) return;
        const extraFrameUp: Record<string, number> = {
          "35": 72000,  // 2 照片框：再上 0.2 cm
          "22": 288000, // 6 位置圖框：累計再上 0.2 cm
        };
        const shift = up + (extraFrameUp[idNode.getAttribute("id") || ""] || 0);
        Array.from(anchor.getElementsByTagNameNS(xdrNs, "rowOff")).forEach(offset => offset.textContent = String(Math.max(0, Number(offset.textContent || "0") - shift)));
        const transform = Array.from(anchor.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "xfrm"))[0];
        const offset = transform && Array.from(transform.children).find(node => node.localName === "off");
        if (offset) offset.setAttribute("y", String(Math.max(0, Number(offset.getAttribute("y") || "0") - shift)));
      });
      // Blue captions use a fixed 1.5 cm height, matching the house sales
      // sheet reference and preventing row changes from stretching the labels.
      const houseCaptionIds = new Set(["32", "34", "36"]);
      Array.from(drawingDocument.documentElement.children).forEach(anchor => {
        const idNode = Array.from(anchor.getElementsByTagNameNS(xdrNs, "cNvPr")).find(node => houseCaptionIds.has(node.getAttribute("id") || ""));
        if (!idNode) return;
        const transform = Array.from(anchor.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "xfrm"))[0];
        const shapeExtent = transform && Array.from(transform.children).find(node => node.localName === "ext");
        const width = Number(shapeExtent?.getAttribute("cx") || "1256400");
        const from = Array.from(anchor.children).find(node => node.localName === "from");
        const captionShift: Record<string, number> = {
          "36": 756000, // 1 照片標題：再上 0.4 cm
          "34": -36000, // 3 格局圖標題：再下 0.2 cm
          "32": 72000,  // 5 位置圖標題：再上 0.3 cm
        };
        const shift = captionShift[idNode.getAttribute("id") || ""] || 0;
        Array.from(anchor.getElementsByTagNameNS(xdrNs, "rowOff")).forEach(offset => offset.textContent = String(Math.max(0, Number(offset.textContent || "0") - shift)));
        if (from && anchor.localName === "twoCellAnchor") {
          const replacement = drawingDocument.createElementNS(xdrNs, "xdr:oneCellAnchor");
          replacement.appendChild(from.cloneNode(true));
          const extent = drawingDocument.createElementNS(xdrNs, "xdr:ext");
          extent.setAttribute("cx", String(width)); extent.setAttribute("cy", String(1.5 * 360000));
          replacement.appendChild(extent);
          Array.from(anchor.children).filter(node => !["from", "to", "pos", "ext"].includes(node.localName)).forEach(node => replacement.appendChild(node.cloneNode(true)));
          const replacementTransform = Array.from(replacement.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "xfrm"))[0];
          const replacementExtent = replacementTransform && Array.from(replacementTransform.children).find(node => node.localName === "ext");
          if (replacementExtent) replacementExtent.setAttribute("cy", String(1.5 * 360000));
          anchor.parentNode?.replaceChild(replacement, anchor);
        }
        const target = anchor.localName === "twoCellAnchor" ? null : anchor;
        const targetTransform = target && Array.from(target.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "xfrm"))[0];
        const targetExtent = targetTransform && Array.from(targetTransform.children).find(node => node.localName === "ext");
        if (targetExtent) targetExtent.setAttribute("cy", String(1.5 * 360000));
      });
      // Use the same fixed caption-to-frame spacing as the now-correct land
      // sheet.  This keeps all six house objects stable even when Excel rows
      // or text sizes change.
      const fixedHouseObjects: Record<string, { x: number; y: number; width: number; height: number }> = {
        "35": { x: 4945403, y: 2558609, width: 3478714, height: 2606809 },
        "33": { x: 4958442, y: 5768789, width: 3480345, height: 2645112 },
        "22": { x: 4944588, y: 8969189, width: 3480345, height: 2645112 },
        "36": { x: 4936174, y: 2054609, width: 1335232, height: 468000 },
        "34": { x: 4936174, y: 5264789, width: 1335232, height: 468000 },
        "32": { x: 4922320, y: 8501189, width: 1335232, height: 468000 },
      };
      Array.from(drawingDocument.documentElement.children).forEach(anchor => {
        const idNode = Array.from(anchor.getElementsByTagNameNS(xdrNs, "cNvPr")).find(node => fixedHouseObjects[node.getAttribute("id") || ""]);
        if (!idNode) return;
        const config = fixedHouseObjects[idNode.getAttribute("id") || ""];
        const replacement = drawingDocument.createElementNS(xdrNs, "xdr:absoluteAnchor");
        const pos = drawingDocument.createElementNS(xdrNs, "xdr:pos"); pos.setAttribute("x", String(config.x)); pos.setAttribute("y", String(config.y));
        const extent = drawingDocument.createElementNS(xdrNs, "xdr:ext"); extent.setAttribute("cx", String(config.width)); extent.setAttribute("cy", String(config.height));
        replacement.appendChild(pos); replacement.appendChild(extent);
        Array.from(anchor.children).filter(node => !["from", "to", "pos", "ext"].includes(node.localName)).forEach(node => replacement.appendChild(node.cloneNode(true)));
        anchor.parentNode?.replaceChild(replacement, anchor);
        const transform = Array.from(replacement.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "xfrm"))[0];
        if (transform) {
          const offset = Array.from(transform.children).find(node => node.localName === "off");
          const shapeExtent = Array.from(transform.children).find(node => node.localName === "ext");
          if (offset) { offset.setAttribute("x", String(config.x)); offset.setAttribute("y", String(config.y)); }
          if (shapeExtent) { shapeExtent.setAttribute("cx", String(config.width)); shapeExtent.setAttribute("cy", String(config.height)); }
        }
      });
    }
    Array.from(drawingDocument.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", "clientData")).forEach(node => {
      node.setAttribute("fLocksWithSheet", "0");
      node.setAttribute("fPrintsWithSheet", "1");
    });
    Array.from(drawingDocument.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "picLocks")).forEach(node => {
      ["noMove", "noResize", "noSelect", "noEditPoints", "noChangeAspect"].forEach(attribute => node.removeAttribute(attribute));
    });
    Array.from(drawingDocument.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "spLocks")).forEach(node => {
      ["noMove", "noResize", "noSelect", "noEditPoints", "noTextEdit"].forEach(attribute => node.removeAttribute(attribute));
    });
    Array.from(drawingDocument.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "graphicFrameLocks")).forEach(node => node.parentNode?.removeChild(node));
    Array.from(drawingDocument.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", "cNvPr")).forEach(node => {
      node.removeAttribute("hidden"); node.removeAttribute("noSelect"); node.removeAttribute("locked");
    });
    zip.file(drawingPath, serializer.serializeToString(drawingDocument));
  }
  const worksheetPaths = Object.keys(zip.files).filter(path => /^xl\/worksheets\/sheet\d+\.xml$/i.test(path));
  for (const worksheetPath of worksheetPaths) {
    const worksheetXml = await zip.file(worksheetPath)?.async("string");
    if (!worksheetXml) continue;
    const worksheetDocument = parser.parseFromString(worksheetXml, "application/xml");
    Array.from(worksheetDocument.getElementsByTagNameNS(spreadsheetNs, "sheetProtection")).forEach(node => node.parentNode?.removeChild(node));
    Array.from(worksheetDocument.getElementsByTagNameNS(spreadsheetNs, "sheetView")).forEach(node => {
      if (worksheetPath === visibleSheetPath) node.setAttribute("tabSelected", "1");
      else node.removeAttribute("tabSelected");
    });
    zip.file(worksheetPath, serializer.serializeToString(worksheetDocument));
  }
  const unlockedWorkbookXml = await zip.file("xl/workbook.xml")?.async("string");
  if (unlockedWorkbookXml) {
    const unlockedWorkbookDocument = parser.parseFromString(unlockedWorkbookXml, "application/xml");
    Array.from(unlockedWorkbookDocument.getElementsByTagNameNS(spreadsheetNs, "workbookProtection")).forEach(node => node.parentNode?.removeChild(node));
    zip.file("xl/workbook.xml", serializer.serializeToString(unlockedWorkbookDocument));
  }
  // The sales sheet no longer needs VBA. Export a true .xlsx workbook so Excel
  // does not place the drawing placeholders behind macro-workbook restrictions.
  zip.remove("xl/vbaProject.bin");
  zip.remove("xl/calcChain.xml");
  const workbookRelsPath = "xl/_rels/workbook.xml.rels";
  const workbookRelsXml = await zip.file(workbookRelsPath)?.async("string");
  if (workbookRelsXml) {
    const relsDocument = parser.parseFromString(workbookRelsXml, "application/xml");
    Array.from(relsDocument.getElementsByTagName("Relationship")).filter(node => /(?:vbaProject|calcChain)/i.test(node.getAttribute("Type") || "") || /(?:vbaProject\.bin|calcChain\.xml)/i.test(node.getAttribute("Target") || "")).forEach(node => node.parentNode?.removeChild(node));
    zip.file(workbookRelsPath, serializer.serializeToString(relsDocument));
  }
  const contentTypesPath = "[Content_Types].xml";
  const contentTypesXml = await zip.file(contentTypesPath)?.async("string");
  if (contentTypesXml) {
    const contentTypesDocument = parser.parseFromString(contentTypesXml, "application/xml");
    Array.from(contentTypesDocument.getElementsByTagName("Override")).forEach(node => {
      if (/(?:vbaProject\.bin|calcChain\.xml)/i.test(node.getAttribute("PartName") || "")) node.parentNode?.removeChild(node);
      else if ((node.getAttribute("PartName") || "") === "/xl/workbook.xml") node.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
    });
    zip.file(contentTypesPath, serializer.serializeToString(contentTypesDocument));
  }
  const blob = await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  const filenameArea = String(record.area || areaCategory(record) || "").trim();
  const filenameAddress = String(record.address || "").replace(/^(?:台|臺)?[^市縣]+[市縣]/, "").replace(/^[^市縣區鄉鎮]+(?:區|鄉|鎮|市)/, "").trim();
  const filenameParts = [filenameArea, record.caseName, filenameAddress, developerFullNameText(record.developer)].map(value => String(value || "").trim()).filter(Boolean);
  link.download = `${(filenameParts.join("-") || record.propertyNo || "物件").replace(/[\\/:*?"<>|]/g, "-")}.xlsx`;
  document.body.appendChild(link); link.click(); link.remove();
  // Keep the blob URL alive for this local session. Edge may hand the temporary
  // download path to Excel a few seconds after the click; revoking immediately
  // can make Excel report that the downloaded file has already disappeared.
}

function printRecordDocument(record: RecordItem, kind: "color" | "cover") {
  const escapeHtml = (value = "") => String(value || "").replace(/[&<>"']/g, character => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[character] || character));
  const value = (key: string) => escapeHtml(key === "developer" ? (developerFullNameText(record.developer) || "—") : (record[key] || "—"));
  const date = (key: string) => escapeHtml(displayRocDate(record[key] || "") || "—");
  const photo = record.photos?.[0] || "";
  const common = `<style>@page{size:A4 portrait;margin:0}*{box-sizing:border-box}html,body{margin:0;background:#fff;color:#111;font-family:"Microsoft JhengHei",sans-serif}.page{width:210mm;height:297mm;padding:10mm;overflow:hidden}table{width:100%;border-collapse:collapse}th,td{border:1.2px solid #222;padding:4mm;text-align:center;vertical-align:middle}h1{margin:0;text-align:center;font-family:DFKai-SB,"標楷體",serif;letter-spacing:.35em}.blue{background:#078bc8;color:#fff}.red{color:#e93035}.small{font-size:10pt}.left{text-align:left}</style>`;
  const colorSheet = `<div class="page"><header><div style="display:flex;align-items:end;justify-content:space-between"><b style="font-size:24pt">台慶不動產</b><h1 style="font-size:32pt">銷售資料表</h1><b>台南文化崇明加盟店</b></div><p style="text-align:center;margin:2mm 0 3mm">連城不動產開發有限公司　台南市東區崇明路520號　06-3356699</p></header><table><tr><th class="blue" style="width:12mm">案名</th><td style="font-size:20pt;font-weight:700">${value("caseName")}</td><th class="blue" style="width:12mm">地址</th><td style="font-size:15pt">${value("address")}</td></tr></table><div style="display:flex;gap:12mm;margin:2mm 0;font-size:13pt"><span>物件編號：${value("propertyNo")}</span><span>KEY編號：${value("key")}</span><span>現況：${value("currentState")}</span></div><div style="display:grid;grid-template-columns:58% 42%;gap:4mm"><section><div style="font-size:22pt;margin:2mm 0">總價：<b class="red" style="font-size:36pt">${value("price")}萬</b></div><table><tr><th class="blue">面積</th><td class="left">建坪：${value("buildingPing")}坪<br/>室內坪：${value("indoorPing")}坪<br/>土地坪：${value("landPing")}坪</td></tr><tr><th class="blue">基本資料</th><td class="left">種類：${value("type")}<br/>格局：${value("layout")}<br/>樓層：${value("floor")}<br/>屋齡：${escapeHtml(ageOf(record))}<br/>朝向：${escapeHtml(directionShort(record.direction))}<br/>車位：${value("parking")}<br/>管理費：${value("managementFee")}</td></tr><tr><th class="blue">環境</th><td class="left">臨路：${value("road")}　面寬：${value("frontage")}　深度：${value("depth")}<br/>使用分區：${value("zoning")}　建蔽／容積：${value("coverage")}／${value("far")}</td></tr><tr><th class="blue">重點說明</th><td class="left" style="height:65mm">${value("notes")}</td></tr></table></section><aside><div class="blue" style="padding:3mm;text-align:center;font-size:18pt">照片</div><div style="height:105mm;border:1px solid #555;display:grid;place-items:center">${photo ? `<img src="${photo}" style="width:100%;height:100%;object-fit:cover">` : "尚無照片"}</div><div class="blue" style="padding:3mm;text-align:center;font-size:18pt;margin-top:4mm">案件資料</div><table><tr><th>開發</th><td>${value("developer")}</td></tr><tr><th>委託期間</th><td>${date("entrustStart")}～${date("entrustEnd")}</td></tr><tr><th>進案日期</th><td>${date("reportDate")}</td></tr></table></aside></div></div>`;
  const cover = `<div class="page" style="padding:14mm"><table style="height:270mm"><tr style="height:34mm"><td colspan="4"><h1 style="font-size:40pt">新進資料</h1></td></tr><tr><th style="width:22mm;font-size:18pt">案名</th><td colspan="3" style="font-size:21pt;font-weight:700">${value("caseName")}</td></tr><tr><th style="font-size:18pt">地址</th><td colspan="3" style="font-size:18pt">${value("address")}</td></tr><tr><th>經紀<br/>營業員</th><td style="font-size:18pt">開發　${value("developer")}</td><th>底價</th><td>${value("price")}萬</td></tr><tr><th rowspan="7" style="font-size:24pt">契<br/>約<br/>紀<br/>錄</th><th>契約編號</th><th>日期／委託起訖日</th><th>用途</th></tr><tr><td>${value("propertyNo")}</td><td>${date("entrustStart")}～${date("entrustEnd")}</td><td>開價${value("price")}萬</td></tr><tr><td>進案日期</td><td>${date("reportDate")}</td><td>—</td></tr><tr><td></td><td></td><td></td></tr><tr><td></td><td></td><td></td></tr><tr><td></td><td></td><td></td></tr><tr><td></td><td></td><td></td></tr><tr style="height:85mm"><th style="font-size:22pt">內<br/>附<br/>文<br/>件</th><td colspan="3" class="left" style="font-size:16pt;line-height:2">★主約　${value("propertyNo")}<br/>★契變　${value("contractChangeNo")}<br/>★原稿／草稿<br/>★土地權狀／建物權狀<br/>◎使用分區<br/>◎授權書</td></tr></table></div>`;
  const checked = (condition: boolean) => condition ? "☑" : "□";
  const bottomSource = record.coverBottomSource || "主約";
  const bottomText = record.coverBottomPrice ? `${escapeHtml(bottomSource)}${escapeHtml(record.coverBottomPrice)}萬` : "";
  const percentText = record.coverBottomPercent ? `${escapeHtml(record.coverPercentSource || "主約")}${escapeHtml(record.coverBottomPercent)}%` : "";
  const coverChangeNo = String(record.coverChangeNo || "").trim();
  const hasCoverChange = record.coverNoChange !== "無契變" && !!coverChangeNo && !isImportedDashPlaceholder(coverChangeNo);
  const changeBottomPrice = String(record.coverBottomPrice || "").replace(/\s*萬(?:元)?\s*/g, "").trim();
  const coverChangePurpose = hasCoverChange && changeBottomPrice ? `底價${escapeHtml(changeBottomPrice)}萬` : value("coverChangePurpose");
  const coverIsLand = typeShort(record.type) === "土地" || /^(?:LG|LA)/i.test(record.propertyNo || "");
  const coverZoningStatus = record.zoningDocumentStatus || (coverIsLand ? "" : "房屋不需要");
  const coverContractRows = false ? `<table class="contract-grid-v3"><colgroup><col style="width:25mm"><col style="width:11mm"><col style="width:11mm"><col style="width:11mm"><col style="width:58mm"><col></colgroup><tbody>
      <tr style="height:9mm"><th rowspan="2">契約編號</th><th colspan="3">契管</th><th rowspan="2">日期／委託起訖日</th><td class="report">進案日期：　${date("reportDate")}</td></tr>
      <tr style="height:9mm"><th>房管</th><th>照片</th><th>紙本</th><th>用途</th></tr>
      <tr style="height:12mm"><td>${value("propertyNo")}</td><td></td><td></td><td></td><td>${date("entrustStart")}～${date("entrustEnd")}</td><td>開價${value("price")}萬</td></tr>
      <tr style="height:12mm"><td>${record.coverNoChange === "無契變" ? "無契變" : value("coverChangeNo")}</td><td></td><td></td><td></td><td>${record.coverNoChange === "無契變" ? "-" : record.coverChangeDate ? date("coverChangeDate") : "-"}</td><td>${record.coverNoChange === "無契變" ? "-" : record.coverChangePurpose ? value("coverChangePurpose") : "-"}</td></tr>
      <tr style="height:10.8mm"><td></td><td></td><td></td><td></td><td></td><td></td></tr><tr style="height:10.8mm"><td></td><td></td><td></td><td></td><td></td><td></td></tr><tr style="height:10.8mm"><td></td><td></td><td></td><td></td><td></td><td></td></tr><tr style="height:10.8mm"><td></td><td></td><td></td><td></td><td></td><td></td></tr><tr style="height:10mm"><td></td><td></td><td></td><td></td><td></td><td></td></tr>
    </tbody></table>` : `<table class="contract-grid-v3"><colgroup><col style="width:36mm"><col style="width:58mm"><col></colgroup><tbody>
      <tr style="height:9mm"><th rowspan="2">契約編號</th><th rowspan="2">日期／委託起訖日</th><td class="report">進案日期：　${date("reportDate")}</td></tr>
      <tr style="height:9mm"><th>用途</th></tr>
      <tr style="height:12mm"><td>${value("propertyNo")}</td><td>${date("entrustStart")}～${date("entrustEnd")}</td><td>開價${value("price")}萬</td></tr>
      <tr style="height:12mm"><td>${record.coverNoChange === "無契變" ? "無契變" : (isImportedDashPlaceholder(record.coverChangeNo) ? "－" : value("coverChangeNo"))}</td><td>${record.coverNoChange === "無契變" || isImportedDashPlaceholder(record.coverChangeDate) ? "－" : date("coverChangeDate")}</td><td>${record.coverNoChange === "無契變" || isImportedDashPlaceholder(record.coverChangePurpose) && !hasCoverChange ? "－" : coverChangePurpose}</td></tr>
      <tr style="height:10.8mm"><td></td><td></td><td></td></tr><tr style="height:10.8mm"><td></td><td></td><td></td></tr><tr style="height:10.8mm"><td></td><td></td><td></td></tr><tr style="height:10.8mm"><td></td><td></td><td></td></tr><tr style="height:10mm"><td></td><td></td><td></td></tr>
    </tbody></table>`;
  const coverV2 = `<div class="page intake-cover"><style>
    .intake-cover{padding:9mm 12mm;font-family:"Microsoft JhengHei",sans-serif}.intake-cover h1{font-family:inherit;font-size:30pt;letter-spacing:.55em;margin:0 0 4mm;color:#000}.cover-line{display:grid;grid-template-columns:26mm 1fr;border-bottom:1.3px solid #222;min-height:15mm;align-items:center;font-size:17pt}.cover-line b{letter-spacing:.25em}.cover-line span{font-size:18pt}.cover-agent{display:grid;grid-template-columns:1fr 85mm;border-bottom:1.3px solid #222}.cover-agent-left{display:flex;align-items:center;gap:8mm;min-height:26mm;font-size:16pt}.cover-agent-left b{letter-spacing:.12em}.cover-bottom{border-left:1.3px solid #222;display:grid;grid-template-rows:1fr 1fr}.cover-bottom div{display:flex;align-items:center;padding:2mm 4mm;font-size:15pt}.cover-bottom div+div{border-top:1px solid #777}.cover-bottom strong{min-width:27mm}.contract-title,.document-title{margin:5mm 0 1.5mm;font-size:18pt;font-weight:900;letter-spacing:.18em}.contract-table th,.contract-table td{padding:2.5mm 2mm;font-size:12pt;border:1px solid #333}.contract-table th{background:#f2f2f2}.contract-table td:nth-child(1){width:36mm}.contract-table td:nth-child(2){width:64mm}.contract-table td:nth-child(3){text-align:left}.document-box{border:1.4px solid #222;padding:4mm 5mm;font-size:13pt;line-height:1.8}.document-box .stars{font-size:14pt;font-weight:800}.document-box .sub{padding-left:7mm}.document-box .square{font-family:"Microsoft JhengHei",sans-serif;font-size:15pt}.document-box hr{border:0;border-top:1px solid #777;margin:2.5mm 0}.cover-foot{margin-top:4mm;text-align:center;font-size:10pt}
  </style><h1>新 進 資 料</h1>
  <div class="cover-line"><b>案名</b><span>${value("caseName")}</span></div>
  <div class="cover-line"><b>地址</b><span>${value("address")}</span></div>
  <div class="cover-agent"><div class="cover-agent-left"><b>經紀營業員</b><span>開發　${value("developer")}</span></div><div class="cover-bottom"><div><strong>底價</strong><span>${bottomText || "＿＿＿＿＿＿"}</span></div><div><strong>％數</strong><span>${percentText || "＿＿＿＿＿＿"}</span></div></div></div>
  <div class="contract-title">契約紀錄</div><table class="contract-table"><thead><tr><th>契約編號</th><th>日期／委託起訖日</th><th>用途</th></tr></thead><tbody>
    <tr><td>${value("propertyNo")}</td><td>${date("entrustStart")} ～ ${date("entrustEnd")}</td><td>開價 ${value("price")} 萬</td></tr>
    <tr><td>${record.coverNoChange === "無契變" ? "無契變" : value("coverChangeNo")}</td><td>${date("coverChangeDate")}</td><td>${value("coverChangePurpose")}</td></tr>
  </tbody></table>
  <div class="document-title">內附文件之正本</div><div class="document-box">
    <div class="stars">★主約編號：${value("propertyNo")}　　★契變編號：${value("coverChangeNo")}</div>
    <div class="stars">★土地權狀 <span class="square">□</span> ${value("landTitleCount")} 張　／　建物權狀 <span class="square">□</span> ${value("buildingTitleCount")} 張　／　${checked(record.titleUndertaking === "有切結")} 切結</div>
    <hr><div>◎ 使用分區：如果銷售土地為空白，必須申請。</div>
    <div class="sub">${checked(record.zoningDocumentStatus === "房屋不需要")} 此物件為銷售房子，不需要</div>
    <div class="sub">● ${checked(record.zoningDocumentStatus === "土地已附正式分區")} 此物件為銷售土地，需要且已附正式分區</div>
    <div class="sub">● ${checked(record.zoningDocumentStatus === "謄本已標示不用附")} 此物件為銷售土地，謄本已標示──不用附上使用分區</div>
    <hr><div>◎ ★授權書：眼看「謄本」、「權狀」、「主約」是否一樣</div>
    <div class="sub">● ${checked(record.authorizationStatus === "無需要")} 無需要　／　${checked(record.authorizationStatus === "已附上歸檔")} 已附上歸檔（${checked(record.authorizationCopyType === "影本")} 影本　${checked(record.authorizationCopyType === "正本")} 正本）</div>
  </div><div class="cover-foot">以上資料及文件請逐項核對後歸檔</div></div>`;
  const coverV3 = `<div class="page intake-cover-v3"><style>
    .intake-cover-v3{padding:10mm;font-family:DFKai-SB,"標楷體",serif;color:#000}.cover-sheet{height:277mm;border:1.6px solid #000;display:flex;flex-direction:column}.cover-sheet *{box-sizing:border-box}.cover-title{height:32mm;flex:none;display:flex;align-items:center;justify-content:center;border-bottom:1.3px solid #000;font-size:45pt;letter-spacing:.25em;padding-left:.25em}.cover-row{display:grid;grid-template-columns:21mm 1fr;height:20mm;flex:none;border-bottom:1.3px solid #000}.cover-row .label,.agent-label,.vertical-label{display:flex;align-items:center;justify-content:center;border-right:1.3px solid #000;font-size:18pt}.cover-row .value{display:flex;align-items:center;justify-content:center;padding:1.5mm 4mm;font-size:19pt}.cover-agent-v3{height:20mm;flex:none;display:grid;grid-template-columns:21mm 25mm 1fr 53mm;border-bottom:1.3px solid #000}.agent-label{font-size:15pt;line-height:1.55;text-align:center}.agent-kind,.agent-name{display:flex;align-items:center;justify-content:center;border-right:1.3px solid #000;font-size:17pt}.agent-name{font-size:19pt;border-right:0}.agent-price{display:grid;grid-template-columns:21mm 1fr;align-items:stretch;border:2.2px solid #000;margin:0}.agent-price>b{display:flex;align-items:center;justify-content:center;font-size:18pt;white-space:nowrap}.agent-price-values{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:1mm;padding:1mm;font-size:16pt;line-height:1.15}.agent-price-values span{white-space:nowrap}.agent-price-values .no-price{font-size:13pt}.contract-wrap{height:104mm;flex:none;display:grid;grid-template-columns:21mm 1fr;border-bottom:1.3px solid #000}.vertical-label{font-size:21pt;line-height:1.75;writing-mode:vertical-rl;letter-spacing:.28em;padding-top:.28em}.document-label{font-size:24pt;writing-mode:horizontal-tb;letter-spacing:0;padding:2mm 0;flex-direction:column;justify-content:space-evenly;line-height:1}.contract-grid-v3{width:100%;height:100%;border-collapse:collapse;table-layout:fixed}.contract-grid-v3 th,.contract-grid-v3 td{border:1px solid #000;border-top:0;padding:1mm 2mm;text-align:center;vertical-align:middle;font-weight:400}.contract-grid-v3 tr>*:first-child{border-left:0}.contract-grid-v3 tr>*:last-child{border-right:0}.contract-grid-v3 tr:last-child>*{border-bottom:0}.contract-grid-v3 th{font-size:15pt}.contract-grid-v3 td{font-size:13pt}.contract-grid-v3 .report{text-align:left;padding-left:4mm}.document-wrap{flex:1;min-height:0;display:grid;grid-template-columns:21mm 1fr}.document-content{padding:2mm 3mm 0;font-size:13.5pt;line-height:1.75}.document-content .stars{font-size:14.5pt}.document-content .sub{padding-left:7mm}.document-content .square{font-family:"Microsoft JhengHei",sans-serif;font-size:15pt}.document-content .red{color:#f00}.document-content .form-box{display:inline-flex;width:7mm;height:7mm;border:1px solid #000;align-items:center;justify-content:center;vertical-align:middle;margin:0 1mm;font-family:"Microsoft JhengHei",sans-serif;font-size:13pt;line-height:1}
  </style><div class="cover-sheet">
    <div class="cover-title">新 進 資 料</div>
    <div class="cover-row"><div class="label">案名</div><div class="value">${value("caseName")}</div></div>
    <div class="cover-row"><div class="label">地址</div><div class="value">${value("address")}</div></div>
    <div class="cover-agent-v3"><div class="agent-label">經紀<br>營業員</div><div class="agent-kind">開　發</div><div class="agent-name">${value("developer")}</div><div class="agent-price"><b>底價：</b><div class="agent-price-values"><span class="${/無底價/.test(record.coverBottomPrice || "") ? "no-price" : ""}">${bottomText}</span><span>${percentText}</span></div></div></div>
    <div class="contract-wrap"><div class="vertical-label">契約紀錄</div>${coverContractRows}</div>
    <div class="document-wrap"><div class="vertical-label document-label"><span>內</span><span>附</span><span>文</span><span>件</span><span>之</span><span>正</span><span>本</span></div><div class="document-content">
      <div class="stars">★主約　<span class="red">${value("propertyNo")}</span>　★契變　<span class="red">${record.coverNoChange === "無契變" ? "無契變" : (isImportedDashPlaceholder(record.coverChangeNo) ? "－" : value("coverChangeNo"))}</span></div>
      <div class="stars">★原稿/草稿；</div>
      <div class="stars">★土地權狀×<span class="form-box">${escapeHtml(record.landTitleCount || "")}</span>張/建物權狀×<span class="form-box">${escapeHtml(record.buildingTitleCount || "")}</span>張/<span class="form-box">${record.titleUndertaking === "有切結" ? "✓" : ""}</span>切結</div>
      <div>◎ 使用分區:如果銷售土地為空白*必申請。</div>
      <div class="sub"><span class="form-box">${coverZoningStatus === "房屋不需要" ? "✓" : ""}</span>此物件為銷售房子不需要</div>
      <div class="sub">●此物件為銷售土地需要且已附(<span class="form-box">${coverZoningStatus === "土地已附正式分區" ? "✓" : ""}</span>正式分區)</div>
      <div class="sub">●此物件為銷售土地，謄本<span class="form-box">${coverZoningStatus === "謄本已標示不用附" ? "✓" : ""}</span>已標示--不用附上使用分區</div>
      <div>◎ ★授權書:眼看「謄本」、「權狀」、「主約」是否一樣</div>
      <div class="sub">● <span class="form-box">${record.authorizationStatus === "缺授" ? "✓" : ""}</span>缺授/<span class="form-box">${record.authorizationStatus === "無需要" ? "✓" : ""}</span>無需要/已附上歸檔(<span class="form-box">${record.authorizationCopyType === "影本" ? "✓" : ""}</span>影 本/<span class="form-box">${record.authorizationCopyType === "正本" ? "✓" : ""}</span>正本 )</div>
    </div></div>
  </div></div>`;
  const frame = document.createElement("iframe"); frame.style.position = "fixed"; frame.style.width = "0"; frame.style.height = "0"; frame.style.border = "0"; document.body.appendChild(frame); const printDocument = frame.contentDocument; if (!printDocument) { frame.remove(); return; } printDocument.open(); printDocument.write(`<!doctype html><html lang="zh-Hant"><head><meta charset="UTF-8"><title>${kind === "color" ? "彩色表" : "新進封面"}</title>${common}</head><body>${kind === "color" ? colorSheet : coverV3}</body></html>`); printDocument.close(); setTimeout(() => { frame.contentWindow?.focus(); frame.contentWindow?.print(); setTimeout(() => frame.remove(), 1000); }, 300);
}

function printKeySummaryTable(tableHtml: string) {
  const frame = document.createElement("iframe");
  frame.setAttribute("aria-hidden", "true");
  Object.assign(frame.style, { position: "fixed", right: "0", bottom: "0", width: "1px", height: "1px", border: "0", opacity: "0", pointerEvents: "none" });
  document.body.appendChild(frame);
  const printDocument = frame.contentDocument; const printWindow = frame.contentWindow;
  if (!printDocument || !printWindow) { frame.remove(); return alert("無法開啟列印畫面，請再試一次。"); }
  printDocument.open();
  printDocument.write(`<!doctype html><html lang="zh-Hant"><head><meta charset="UTF-8"><title>鑰匙總表</title><style>@page{size:A3 landscape;margin:.8cm .2cm .2cm}*{box-sizing:border-box}html,body{margin:0;padding:0;background:#fff;font-family:"Microsoft JhengHei",sans-serif;color:#111}h1{margin:0 0 1.5mm;text-align:center;font-size:40pt}table{width:100%;border-collapse:collapse;table-layout:fixed}th,td{height:11.4mm;padding:.5mm .8mm;border:1px solid #555;text-align:center;vertical-align:middle;font-size:14pt;line-height:1;overflow-wrap:anywhere}th{height:9mm;background:#e5eee9;font-size:14pt;font-weight:900}th:nth-child(1),td:nth-child(1),th:nth-child(6),td:nth-child(6){width:4.5%}th:nth-child(2),td:nth-child(2),th:nth-child(7),td:nth-child(7){width:6%}th:nth-child(3),td:nth-child(3),th:nth-child(8),td:nth-child(8){width:16%}th:nth-child(4),td:nth-child(4),th:nth-child(9),td:nth-child(9){width:16%}th:nth-child(5),td:nth-child(5),th:nth-child(10),td:nth-child(10){width:7.5%}td:nth-child(4),td:nth-child(9){font-size:12.5pt}.key-two-lines{display:flex;min-height:2em;flex-direction:column;align-items:center;justify-content:center;line-height:1}.key-two-lines>span{display:block;white-space:nowrap}tr{break-inside:avoid}.key-number{font-size:17pt;font-weight:900}</style></head><body><h1>鑰匙總表</h1>${tableHtml}</body></html>`);
  printDocument.close();
  const cleanup = () => setTimeout(() => frame.remove(), 1000);
  printWindow.addEventListener("afterprint", cleanup, { once: true });
  setTimeout(() => { printWindow.focus(); printWindow.print(); }, 300);
  setTimeout(() => { if (frame.isConnected) frame.remove(); }, 60000);
}

function KeySummary({ records }: { records: RecordItem[] }) {
  const byNumber = new Map<number, RecordItem>();
  records.forEach(record => { const match = String(record.key || "").match(/公司\s*[#＃]?\s*(\d+)/); if (match) byNumber.set(Number(match[1]), record); });
  const twoLineText = (value = "") => { const chars = Array.from(value); const lines = Array.from({ length: Math.ceil(chars.length / 15) }, (_, index) => chars.slice(index * 15, index * 15 + 15).join("")); return <span className="key-two-lines">{lines.slice(0, 2).map((line, index) => <span key={`${line}-${index}`}>{line}</span>)}</span>; };
  const keyCaseText = (value = "") => <span className="key-two-lines key-case-lines">{chunkText(value, 12).map((line, index) => <span key={`${line}-${index}`}>{line}</span>)}</span>;
  const actualDistrict = (record?: RecordItem) => {
    if (!record) return "";
    const address = String(record.address || "").replace(/臺/g, "台").trim();
    const fullDistrict = address.match(/^(.{2,3}?)(?:市|縣)([^市縣區]{1,4})區/);
    if (fullDistrict) return fullDistrict[1] === "台南" ? `${fullDistrict[2]}區` : `${fullDistrict[1]}${fullDistrict[2]}`;
    const district = address.match(/^([^市縣區]{1,4})區/)?.[1];
    if (district) return `${district}區`;
    const stored = String(record.area || "").replace(/臺/g, "台").trim();
    if (/^(?:外縣市|其他區|其他)$/.test(stored)) return "";
    const storedFull = stored.match(/^(.{2,3}?)(?:市|縣)([^市縣區]{1,4})區$/);
    if (storedFull) return storedFull[1] === "台南" ? `${storedFull[2]}區` : `${storedFull[1]}${storedFull[2]}`;
    return stored;
  };
  const cells = (number: number) => { const record = byNumber.get(number); const archivedLabel = record?.archived ? `${displayRocDate(record.archived)}${record.status || "已下架"}` : ""; const developers = developerNameLines(record?.developer || ""); return <><td className="key-number">{number}</td><td>{actualDistrict(record)}</td><td className="key-case">{keyCaseText(record?.caseName || "")}{archivedLabel && <small style={{ color: "#d71920", fontSize: "calc(.68em + 1.5pt)", fontWeight: 400, display: "block", whiteSpace: "nowrap" }}>{archivedLabel}</small>}</td><td className="key-address">{twoLineText(record?.address || "")}</td><td><span className="key-two-lines">{developers.map((name, index) => <span key={`${name}-${index}`}>{name}</span>)}</span></td></>; };
  const printSummary = () => { const table = document.querySelector(".key-summary-table") as HTMLTableElement | null; if (table) { const copy = table.cloneNode(true) as HTMLTableElement; copy.querySelectorAll("th").forEach(cell => cell.setAttribute("style", "background:#d9d9d9;color:#222")); printKeySummaryTable(copy.outerHTML); } };
  return <section className="content key-summary-page"><div className="list-head"><SectionTitle title="鑰匙總表" subtitle="委託中與封存物件，只要鑰匙欄仍標示公司#編號就會顯示"/><button type="button" className="primary key-print-button" onClick={printSummary}>列印鑰匙總表</button></div><div className="key-summary-wrap"><table className="key-summary-table"><thead><tr><th>編號</th><th>地區</th><th>案名</th><th>地址</th><th>開發</th><th>編號</th><th>地區</th><th>案名</th><th>地址</th><th>開發</th></tr></thead><tbody>{keySummaryLeft.map((leftNumber, index) => <tr key={leftNumber}>{cells(leftNumber)}{cells(keySummaryRight[index])}</tr>)}</tbody></table></div></section>;
}

function DailyActivity({ records, compact = false, onEdit }: { records: RecordItem[]; compact?: boolean; onEdit?: (record: RecordItem) => void }) {
  const [selectedDate, setSelectedDate] = useState(today());
  const [hiddenItems, setHiddenItems] = useState<string[]>(() => { try { return JSON.parse(localStorage.getItem(DAILY_HIDDEN_KEY) || "[]"); } catch { return []; } });
  useEffect(() => { localStorage.setItem(DAILY_HIDDEN_KEY, JSON.stringify(hiddenItems)); }, [hiddenItems]);
  useEffect(() => { const hide = (event: Event) => { const key = (event as CustomEvent<string>).detail; if (key) setHiddenItems(previous => [...new Set([...previous, key])]); }; window.addEventListener("hide-daily-item", hide); return () => window.removeEventListener("hide-daily-item", hide); }, []);
  const dateOnly = (value = "") => value.slice(0, 10);
  const adjacentDateLabel = (date: string) => `${displayRocDate(date)}（${["日", "一", "二", "三", "四", "五", "六"][new Date(`${date}T12:00:00`).getDay()]}）`;
  const previousDate = addDaysIso(selectedDate, -1);
  const nextDate = addDaysIso(selectedDate, 1);
  const rocDay = displayRocDate(selectedDate);
  const updateRecords = records.filter(record => !record.archived && !isExpired(record) && record.status === "委託中" && (dailyUpdateFields(record, selectedDate).length > 0 || dateOnly(record._restoredAt) === selectedDate)).map(record => { const changed = dailyUpdateFields(record, selectedDate); const restored = dateOnly(record._restoredAt) === selectedDate; return { ...record, caseNameNote: changed.length ? `更新：${dailyChangedLabels(changed).join("、")}` : "", _dailyHighlight: JSON.stringify(changed), ...(restored ? { _dailyAnnotation: `${shortRocMonthDay(selectedDate)}重新上架`, _dailyAnnotationType: "restored" } : {}) }; });
  const removedRecords = records.filter(record => !!record.archived && dateOnly(record._archiveActionDate || record.archived) === selectedDate).map(record => ({ ...record, _dailyAnnotation: `${displayRocDate(record.archived)}${record.status || "下架"}` }));
  const groups = [
    { key: "added", title: "新增物件", records: records.filter(record => dateOnly(record._dailyAddedAt || record.reportDate) === selectedDate) },
    { key: "updated", title: "更新物件", records: updateRecords },
    { key: "removed", title: "下架物件", records: removedRecords },
  ].map(group => ({ ...group, records: group.records.filter(record => !hiddenItems.includes(`${selectedDate}|${group.key}|${record.id}`)).map(record => compact ? record : { ...record, _dailyHideKey: `${selectedDate}|${group.key}|${record.id}` }) }));

  return <section className={`daily-activity${compact ? " compact" : ""}`} onClick={event => { if (!onEdit) return; const cell = (event.target as HTMLElement).closest("td.col-caseName"); if (!cell) return; const article = cell.closest("article.daily-card"); const row = cell.closest("tbody tr:not(.area-group-row)"); if (!article || !row) return; const groupIndex = Array.from(event.currentTarget.querySelectorAll("article.daily-card")).indexOf(article); const rowIndex = Array.from(article.querySelectorAll("tbody tr:not(.area-group-row)")).indexOf(row); const record = sortActiveRecords(groups[groupIndex]?.records || [])[rowIndex]; if (record) onEdit(record); }}>
    <div className="daily-activity-head">
      <div>
        <h2>{compact ? "每日物件動態" : "每日物件動態"}　{adjacentDateLabel(selectedDate)}</h2>
        {!compact && <p>查看每日新增、下架與到期物件</p>}
      </div>
      <div className="daily-date"><button type="button" className="daily-yesterday" onClick={() => setSelectedDate(previousDate)}>← {adjacentDateLabel(previousDate)}</button><button type="button" className="daily-next" disabled={selectedDate >= today()} onClick={() => setSelectedDate(nextDate)}>{adjacentDateLabel(nextDate)} →</button><button type="button" className={`daily-today${selectedDate === today() ? " selected" : ""}`} onClick={() => setSelectedDate(today())}>今天</button><input aria-label="選擇每日物件動態日期" type="date" value={selectedDate} max={today()} onChange={event => setSelectedDate(event.target.value)}/></div>
    </div>
    <div className="daily-cards">
      {groups.map(group => <article className={`daily-card ${group.key}`} key={group.key}>
        <header><h3>{group.title}</h3><b>{group.records.length}</b></header>
        {group.records.length ? <PropertyTable records={group.records} columns={compact ? dailyColumns : ["dailyHide", ...dailyColumns]} publicMode dailyMode onEdit={onEdit || (() => {})} onArchive={() => {}} onRestore={() => {}} onRemove={() => {}}/> : <div className="daily-empty">當日無{group.title}</div>}
      </article>)}
    </div>
  </section>;
}

const intakeEditFields = [
  ["時間戳記", "時間戳記"], ["表單填寫人", "表單填寫人"], ["開發１/開發２", "開發業務"], ["委託主約編號", "物件編號"], ["物件型態", "物件型態"],
  ["委託開始", "委託開始"], ["委託結束", "委託結束"], ["(物件)現況", "物件現況"], ["鑰匙位置", "鑰匙位置"], ["中人(介紹費)", "中人"],
  ["案名", "案名"], ["物件(完整)地址", "地址"],
  ["契約開價", "契約開價（萬）"], ["總建坪", "總建坪"], ["室內坪", "室內坪"], ["地坪", "地坪"],
  ["增建說明", "增建／坪數說明"], ["臨路", "臨路（米）"], ["面寬", "面寬（米）"], ["深度", "深度（米）"], ["使用分區", "使用分區"], ["建蔽率/容積率", "建蔽率／容積率"],
  ["特色說明1", "特色說明1"], ["特色說明2", "特色說明2"], ["特色說明3", "特色說明3"], ["特色說明4", "特色說明4"],
  ["鄰近國小", "鄰近國小"], ["鄰近國中", "鄰近國中"], ["鄰近高中", "鄰近高中"],
  ["鄰近大專", "鄰近大專"], ["市場/購物", "市場／購物"], ["公園綠地", "公園綠地"],
  ["注意事項", "注意事項"],
  ["建築完成日期", "建築完成日期"], ["格局 (房)", "房"], ["格局 (廳)", "廳"], ["格局 (衛浴)", "衛浴"], ["格局 (陽台)", "陽台"],
  ["有無電梯", "有無電梯"], ["權狀層數", "權狀層數"], ["透天請寫", "現況樓層"], ["大樓名稱", "大樓名稱"], ["電梯數", "電梯數"], ["每層戶數", "每層戶數"], ["警衛管理", "管理方式"], ["管理費", "管理費（月／元）"],
  ["車位", "車位"], ["車位編號", "車位編號"],
] as const;

const fullIntakeFields = new Set(["增建說明", "特色說明1", "特色說明2", "特色說明3", "特色說明4", "注意事項"]);
const quarterIntakeFields = new Set(["契約開價", "總建坪", "室內坪", "地坪"]);
const thirdIntakeFields = new Set(["鄰近國小", "鄰近國中", "鄰近高中", "鄰近大專", "市場/購物", "公園綠地"]);
const houseDetailFields = new Set(["建築完成日期", "格局 (房)", "格局 (廳)", "格局 (衛浴)", "格局 (陽台)", "有無電梯", "權狀層數", "透天請寫", "大樓名稱", "電梯數", "每層戶數", "警衛管理", "管理費", "車位", "車位編號"]);
const landHiddenIntakeFields = new Set(["建築完成日期", "格局 (房)", "格局 (廳)", "格局 (衛浴)", "格局 (陽台)", "有無電梯", "權狀層數", "透天請寫", "大樓名稱", "電梯數", "每層戶數", "警衛管理", "管理費", "車位", "車位編號"]);
const isLandIntakeDraft = (draft: IntakeData) => draft.propertyKind === "純土地" || /土地|農地|建地/.test([intakeValue(draft.values, "物件型態"), intakeValue(draft.values, "案名")].join(" "));
const intakeFieldClass = (key: string) => `field${fullIntakeFields.has(key) ? " full-field" : key === "案名" ? " case-field" : key === "物件(完整)地址" ? " address-field" : quarterIntakeFields.has(key) ? " quarter-field" : thirdIntakeFields.has(key) ? " third-field" : houseDetailFields.has(key) ? " house-detail-field" : ""}`;

const intakeDraftEditorTitle = (draft: IntakeData) => {
  const address = intakeValue(draft.values, "物件(完整)地址").trim();
  const withoutCity = address.replace(/^(?:臺灣省|台灣省)?[^縣市]{1,8}[縣市]/, "");
  const areaMatch = withoutCity.match(/^(.{1,8}區|.{1,8}[鄉鎮市])/);
  const area = areaMatch?.[1] || "未填地區";
  const shortAddress = (areaMatch ? withoutCity.slice(area.length) : withoutCity).trim() || "未填地址";
  const caseName = intakeValue(draft.values, "案名").trim() || "未命名案件";
  // 草稿檔名的開發姓名直接相連，例如「柯育婷蔡宇育」，不使用頓號。
  const developer = (developerFullNameText(intakeValue(draft.values, "開發１/開發２").trim()) || "未填開發").replace(/[、，,]/g, "");
  // 檔名的分隔符號使用半形 -；Windows 禁用符號（例如 |）也改為 -。
  const fileNamePart = (text: string) => String(text)
    .replace(/[－–—]/g, "-")
    .replace(/[\\/:*?"<>|／：＊？＂＜＞｜＼]/g, "-")
    .replace(/\u3000/g, " ");
  return [area, caseName, shortAddress, developer].map(fileNamePart).join("-");
};

function PreviousIntakePanel({ raw, setRaw, drafts, draft, selectDraft, deleteDraft, analyze, updateValue, clear, confirmIntake }: { raw: string; setRaw: (value: string) => void; drafts: IntakeData[]; draft: IntakeData | null; selectDraft: (id: string) => void; deleteDraft: (id: string) => void; analyze: () => void; updateValue: (key: string, value: string) => void; clear: () => void; confirmIntake: (id?: string) => void }) {
  return <section className="content intake-page"><div className="list-head"><SectionTitle title="進案草稿" subtitle="貼串列印後保留在列表；文件實際收到才按正式進案"/></div><article className="panel paste-panel"><label className="field"><span>貼上新的表單整串文字（可直接修改）</span><textarea className="intake-raw" value={raw} onChange={e => setRaw(e.target.value)} placeholder="請貼上標題列及一筆資料列…" rows={7}/></label><div className="paste-actions"><button onClick={clear}>清除貼串</button><button className="primary analyze-button" onClick={analyze}>加入草稿列表</button></div></article><div className="draft-list-head"><h3>已貼串進案列表</h3><span>{drafts.length} 筆待處理</span></div><div className="draft-list">{drafts.map(item => <article className={`draft-card ${draft?.id === item.id ? "selected" : ""}`} key={item.id}><div><span className={`kind-badge ${item.propertyKind === "純土地" ? "land" : ""}`}>{item.propertyKind}</span><h3>{intakeValue(item.values, "案名") || "未命名案件"}</h3><p>{intakeValue(item.values, "委託主約編號")}　{intakeValue(item.values, "物件(完整)地址")}</p><small>貼串時間：{new Date(item.createdAt).toLocaleString("zh-TW")}</small></div><div className="draft-actions"><button onClick={() => selectDraft(item.id)}>編輯</button><button onClick={() => { selectDraft(item.id); setTimeout(() => window.print(), 0); }}>列印</button><button className="primary" onClick={() => { selectDraft(item.id); setTimeout(confirmIntake, 0); }}>進案</button><button className="danger" onClick={() => deleteDraft(item.id)}>刪除</button></div></article>)}{!drafts.length && <div className="empty-drafts">尚無進案草稿，貼上表單後按「加入草稿列表」。</div>}</div>{draft && <><div className="draft-editor-head"><div><b>編輯：{intakeValue(draft.values, "案名")}</b><span>修改會自動保留</span></div><div><button className="pill-button" onClick={() => window.print()}>列印{draft.propertyKind}進案表</button><button className="pill-button primary" onClick={confirmIntake}>文件已收，正式進案</button></div></div><div className="intake-kind"><b>列印格式：</b><button className={draft.propertyKind === "房屋" ? "selected" : ""} onClick={() => updateValue("物件型態", intakeValue(draft.values, "物件型態").replace("土地", "房屋") || "房屋")}>房屋</button><button className={draft.propertyKind === "純土地" ? "selected" : ""} onClick={() => updateValue("物件型態", "土地")}>純土地</button><span>所有欄位都可修改後再列印</span></div><div className="intake-edit-grid">{intakeEditFields.filter(([key]) => !isLandIntakeDraft(draft) || !landHiddenIntakeFields.has(key)).map(([key, label]) => <label className={`field ${["案名", "物件(完整)地址", "增建說明", "注意事項"].includes(key) ? "wide-field" : ""}`} key={key}><span>{label}</span><input value={intakeValue(draft.values, key)} onChange={e => updateValue(key, e.target.value)}/></label>)}</div><PrintableIntake draft={draft}/></>}</section>;
}

function IntakePanel({ raw, setRaw, drafts, draft, selectDraft, deleteDraft, analyze, addManualDraft, updateValue, clear, confirmIntake, markPrintedForSales }: { raw: string; setRaw: (value: string) => void; drafts: IntakeData[]; draft: IntakeData | null; selectDraft: (id: string) => void; deleteDraft: (id: string) => void; analyze: () => void; addManualDraft: () => void; updateValue: (key: string, value: string) => void; clear: () => void; confirmIntake: (id?: string) => void; markPrintedForSales: (id: string) => void }) {
  const [showEntered, setShowEntered] = useState(false);
  const [featuresCopied, setFeaturesCopied] = useState(false);
  const printDraft = (id: string) => { selectDraft(id); setTimeout(() => { window.print(); selectDraft(""); }, 0); };
  const pendingDrafts = drafts.filter(item => !item.linkedRecordId);
  const enteredDrafts = drafts.filter(item => !!item.linkedRecordId);
  const copyFeatures = async () => {
    if (!draft) return;
    const text = [1, 2, 3, 4].map(number => `${number}.${intakeValue(draft.values, `特色說明${number}`).trim()}`).join("\n");
    try { await navigator.clipboard.writeText(text); }
    catch { const textarea = document.createElement("textarea"); textarea.value = text; document.body.appendChild(textarea); textarea.select(); document.execCommand("copy"); textarea.remove(); }
    setFeaturesCopied(true); window.setTimeout(() => setFeaturesCopied(false), 1800);
  };
  const isEntrustDate = (key: string) => key === "委託開始" || key === "委託結束";
  const isCompletionDate = (key: string) => key === "建築完成日期";
  const draftCard = (item: IntakeData) => <article className={`draft-card ${draft?.id === item.id ? "selected" : ""}`} key={item.id}>
    <div className="draft-summary"><span className={`kind-badge ${item.propertyKind === "純土地" ? "land" : ""}`}>{item.propertyKind === "純土地" ? "土地" : item.propertyKind}</span><button className="draft-case-link" onClick={() => selectDraft(item.id)}>{intakeValue(item.values, "案名") || "未命名案件"}</button><span>{intakeValue(item.values, "委託主約編號") || "—"}</span><span className="draft-address">{intakeValue(item.values, "物件(完整)地址") || "—"}</span><span>{developerFullNameText(intakeValue(item.values, "開發１/開發２")) || "—"}</span><small>{intakeValue(item.values, "時間戳記") || "—"}</small></div>
    <div className="draft-actions"><button onClick={() => printDraft(item.id)}>列印</button><button className={item.printedForSalesAt ? "draft-printed-recorded" : ""} title={item.printedForSalesAt ? "再點一次可取消列印紀錄" : "記錄已列印草稿"} onClick={() => markPrintedForSales(item.id)}>{item.printedForSalesAt ? `已列印草稿${displayRocDate(String(item.printedForSalesAt).slice(0, 10))}` : "已列印草稿"}</button><button className="primary" disabled={!!item.linkedRecordId} onClick={() => confirmIntake(item.id)}>{item.linkedRecordId ? "已進案" : "進案"}</button><button className="danger" onClick={() => deleteDraft(item.id)}>刪除</button></div>
  </article>;
  return <section className="content intake-page">
    <div className="list-head"><SectionTitle title="進案草稿" subtitle=""/><button className="manual-intake-button" onClick={addManualDraft}>＋ 手動新增案件</button></div>
    <article className="panel paste-panel"><label className="field"><textarea className="intake-raw" value={raw} onChange={e => setRaw(e.target.value)} rows={7}/></label><div className="paste-actions"><button onClick={clear}>清除貼串</button><button className="primary analyze-button" onClick={analyze}>加入草稿列表</button></div></article>
    <div className="draft-list-head"><h3>已貼串進案列表</h3><span>{pendingDrafts.length} 筆待處理</span></div>
    <div className="draft-list compact-drafts"><div className="draft-card draft-columns"><div className="draft-summary"><b>種類</b><b>案名</b><b>契約編號</b><b>地址</b><b>開發業務</b><b>時間戳記</b></div><span className="draft-action-title">操作</span></div>{pendingDrafts.map(draftCard)}{!pendingDrafts.length && <div className="empty-drafts">目前沒有待進案草稿。</div>}</div>
    <div className="draft-list-head entered-list-head" role="button" tabIndex={0} onClick={() => setShowEntered(value => !value)} onKeyDown={event => (event.key === "Enter" || event.key === " ") && setShowEntered(value => !value)}><h3>已進案列表</h3><div><span>{enteredDrafts.length} 筆已進案</span><b className="entered-toggle">{showEntered ? "收合 ▲" : "展開 ▼"}</b></div></div>
    {showEntered && <div className="draft-list compact-drafts entered-drafts"><div className="draft-card draft-columns"><div className="draft-summary"><b>種類</b><b>案名</b><b>契約編號</b><b>地址</b><b>開發業務</b><b>時間戳記</b></div><span className="draft-action-title">操作</span></div>{enteredDrafts.map(draftCard)}{!enteredDrafts.length && <div className="empty-drafts">目前沒有已進案資料。</div>}</div>}
    {draft && <>
      <div className="modal-backdrop intake-modal-backdrop" onMouseDown={event => event.target === event.currentTarget && selectDraft("")}><div className="modal intake-draft-modal">
        <div className="draft-editor-head"><div><b>　檔名: {intakeDraftEditorTitle(draft)}</b><span>{draft.linkedRecordId ? "已進案；修改會同步更新總表" : "修改會自動保留"}</span></div><div><button className="pill-button" onClick={() => window.print()}>列印{draft.propertyKind}進案表</button><button className="pill-button primary" disabled={!!draft.linkedRecordId} onClick={() => confirmIntake(draft.id)}>{draft.linkedRecordId ? "已進案" : "文件已收，正式進案"}</button><button className="pill-button" onClick={() => selectDraft("")}>存檔（關閉）</button></div></div>
        <div className="intake-kind"><b>列印格式：</b><button className={draft.propertyKind === "房屋" ? "selected" : ""} onClick={() => updateValue("物件型態", intakeValue(draft.values, "物件型態").replace("土地", "房屋") || "房屋")}>房屋</button><button className={draft.propertyKind === "純土地" ? "selected" : ""} onClick={() => updateValue("物件型態", "土地")}>土地</button><span>所有欄位都可修改後再列印</span></div>
        <div className="intake-edit-grid">{intakeEditFields.filter(([key]) => !isLandIntakeDraft(draft) || !landHiddenIntakeFields.has(key)).map(([key, label]) => { const rawValue = intakeValue(draft.values, key); const dateField = isEntrustDate(key); const completionDateField = isCompletionDate(key); const normalizedDate = normalizeDateInput(rawValue); return <label className={intakeFieldClass(key)} key={key}><span className={key === "特色說明1" ? "intake-feature-title" : ""}>{label}{key === "特色說明1" && <button type="button" className="copy-features-button" onClick={copyFeatures}>{featuresCopied ? "已複製" : "複製特色1～4"}</button>}</span><input inputMode={dateField || completionDateField ? "numeric" : undefined} placeholder={dateField ? "例如 115/8/6" : completionDateField ? "例如 079.11.17" : undefined} value={dateField ? displayRocDate(normalizedDate) : completionDateField ? displayRocDotDate(normalizedDate) : rawValue} onChange={e => updateValue(key, e.target.value)} onBlur={dateField || completionDateField ? e => updateValue(key, normalizeDateInput(e.target.value)) : undefined}/></label>; })}</div>
        <div className="intake-modal-preview"><PrintableIntake draft={draft}/></div>
      </div></div>
      <div className="print-output-only"><PrintableIntake draft={draft}/></div>
    </>}
  </section>;
}

function buildingAgeText(value: string) {
  const year = Number((value || "").trim().split(/[./年-]/)[0]);
  if (!year) return "";
  const rocYear = new Date().getFullYear() - 1911;
  return `${Math.max(0, rocYear - year)}年屋`;
}

function halfWidthDigits(value: string) {
  return value.replace(/[０-９]/g, digit => String.fromCharCode(digit.charCodeAt(0) - 0xFEE0));
}

function measurementNumber(value: string) {
  const match = halfWidthDigits(value || "").match(/\d+(?:\.\d+)?/);
  return match?.[0] || (value || "").replace(/[mM米.。\s]+$/g, "");
}

function PrintableIntake({ draft }: { draft: IntakeData }) {
  const v = draft.values;
  const val = (...keys: string[]) => { const value = intakeValue(v, ...keys); return keys.includes("開發１/開發２") ? (developerFullNameText(value) || "　") : (value || "　"); };
  const features = [1, 2, 3, 4].map(n => val(`特色說明${n}`));
  const completion = val("建築完成日期");
  const age = buildingAgeText(completion);
  const rawContractPrice = val("契約開價").trim();
  const contractPrice = /萬\s*$/.test(rawContractPrice) ? rawContractPrice : `${rawContractPrice} 萬`;
  const parkingTypes = intakeAll(v, "車位型態").filter(Boolean).join("／");
  const direction = [
    ["朝向 [土地朝]", val("朝向 [土地朝]", "朝向[土地朝]", "朝向{土地朝}")],
    ["朝向 [房屋朝]", val("朝向 [房屋朝]")],
    ["朝向 [大門朝]", val("朝向 [大門朝]")]
  ];
  return <section className="print-document">
    <article className={`print-page data-page excel-print ${draft.propertyKind === "純土地" ? "land-print" : "house-print"}`}>
      <div className="print-brand"><div className="print-logo"><b>台慶不動產</b><small>台南文化崇明加盟店</small></div><h1>物件資料表　送件審核</h1></div>
      <p className="company-line">連城不動產開發有限公司　店址：台南市東區崇明路520號　電話：06-3356699　傳真：06-2681166</p>
      <div className="case-band"><b>案名</b><strong>{val("案名")}</strong><b>屋址</b><span>{val("物件(完整)地址")}</span></div>
      <div className="hero-price"><b>價格</b><span>總　價：</span><strong>{contractPrice}</strong><em>□有符合契約簽訂開價</em></div>
      <div className="print-main"><div className="print-left">
        {draft.propertyKind === "純土地" ? <>
          <PrintSection title="面積"><div className="excel-lines"><p className="large-line">土地坪數：<strong>{val("地坪")}坪</strong></p></div></PrintSection>
          <PrintSection title="土地資訊"><div className="land-info-lines"><p>使用分區：<strong>{val("使用分區")}</strong></p><p>建蔽率／容積率：{val("建蔽率/容積率")}</p><p>座　向：{val("朝向 [土地朝]", "朝向[土地朝]", "朝向{土地朝}")}</p><p>臨　路：約 {measurementNumber(val("臨路"))} 米</p><p>面　寬：約 {measurementNumber(val("面寬"))} 米</p><p>深　度：約 {measurementNumber(val("深度"))} 米</p></div></PrintSection>
        </> : <>
          <PrintSection title="面積"><div className="excel-lines area-lines"><p className="print-data-pair"><span className="print-data-cell"><b>總建物坪數：</b><i><strong>{val("總建坪")}坪</strong></i></span><span className="print-data-cell print-data-right"><b>土地坪數：</b><i>{val("地坪")}坪</i></span></p><p className="print-data-pair"><span className="print-data-cell"><b>室內坪數：</b><i>{val("室內坪")}坪</i></span></p></div></PrintSection>
          <PrintSection title="增建說明"><div className="excel-lines addition-lines"><p>{val("增建說明")}</p></div></PrintSection>
          <PrintSection title="車位"><div className="excel-lines parking-lines"><p className="print-data-pair"><span className="print-data-cell"><b>車位產權：</b><i>{val("車位")}</i></span></p><p className="print-data-pair"><span className="print-data-cell"><b>車位型態：</b><i>{parkingTypes || "　"}</i></span><span className="print-data-cell print-data-right"><b>車位編號：</b><i>{val("車位編號")}</i></span></p></div></PrintSection>
          <PrintSection title="基本資料"><div className="excel-lines basic-lines">
            <p className="basic-pair"><span className="basic-cell"><b>物件類型：</b><i>{typeShort(val("物件型態"))}</i></span><span className="basic-cell right-cell building-name-cell"><b>大樓名稱：</b><i>{val("大樓名稱")}</i></span></p>
            <p className="basic-pair"><span className="basic-cell"><b>每層戶數：</b><i>{val("每層戶數")}</i></span><span className="basic-cell right-cell"><b>電梯數：</b><i>{val("電梯數")}</i></span></p>
            <p className="management-row"><span><span className="basic-cell"><b>管理方式：</b><i>{val("警衛管理")}</i></span><span className="basic-cell"><b>管理費：</b><i>{val("管理費")}</i></span></span><span className="basic-cell right-cell zoning-cell"><b><span className="zoning-top"><i>土</i><i>地</i></span><span className="zoning-bottom"><i>使</i><i>用</i><i>分</i><i>區</i></span></b><i>{val("使用分區")}</i></span></p>
            <p className="layout-line basic-cell"><b>格局：</b><i>{val("格局 (房)")}{val("格局 (廳)")}{val("格局 (衛浴)")}{val("格局 (陽台)")}</i></p>
            <p className="basic-pair floor-row"><span className="basic-cell"><b>現況樓層：</b><i>{halfWidthDigits(val("透天請寫"))}</i></span><span className="basic-cell right-cell"><b>權狀樓別：</b><i>{halfWidthDigits(val("權狀層數"))}</i></span></p>
            <p className="basic-pair"><span className="basic-cell"><b>建築完成日：</b><i>{completion}</i></span><span className="basic-cell right-cell"><b>屋齡：</b><i>{age}</i></span></p>
            <p className="direction-line"><b className="direction-label"><i>座</i><i>向</i><i>：</i></b><span className="direction-values">{direction.map(([label, value], index) => <span className="direction-item" key={label}>{label}：<small>{value}</small>{index < direction.length - 1 ? "　" : ""}</span>)}</span></p>
          </div></PrintSection>
          <PrintSection title="環境"><div className="excel-lines environment-lines"><p className="print-data-pair"><span className="print-data-cell"><b>臨路：</b><i>約 {val("臨路")}米</i></span><span className="print-data-cell print-data-right"><b>建蔽／容積：</b><i>{val("建蔽率/容積率")}</i></span></p><p className="print-data-pair"><span className="print-data-cell"><b>面寬：</b><i>約 {val("面寬")}米</i></span><span className="print-data-cell print-data-right"><b>深度：</b><i>約 {val("深度")}米</i></span></p></div></PrintSection>
        </>}
        <PrintSection title="重點說明" className="features-section"><div className="feature-lines">{features.map((feature, i) => <p key={i}><b>{i + 1}.</b><span>{feature}</span></p>)}</div></PrintSection>
      </div><div className="print-right">
        <PrintRightRow label="物件編號" value={val("委託主約編號")}/><PrintRightRow label="鑰匙位置" value={val("鑰匙位置")}/><PrintRightRow label="物件現況" value={val("(物件)現況")}/><PrintRightRow label="國小" value={val("鄰近國小")}/><PrintRightRow label="國中" value={val("鄰近國中")}/><PrintRightRow label="高中" value={val("鄰近高中")}/><PrintRightRow label="大專" value={val("鄰近大專")}/><PrintRightRow label="市場" value={val("市場/購物")}/><PrintRightRow label="公園" value={val("公園綠地")}/><hr/><PrintRightRow label="物件相片" value={val("當下進案文件 [物件相片]")}/>{draft.propertyKind === "房屋" && <PrintRightRow label="格局圖" value={val("進案文件 [格局圖]")}/>}<PrintRightRow label="現詢調" value={val("進案文件 [現詢調]")}/><PrintRightRow label="智能照片" value={val("進案文件 [物件照片 上傳系統]")}/>{draft.propertyKind === "房屋" && <><PrintRightRow label="智能主約" value={val("進案文件 [主約 契約拍照 上傳系統]")}/><PrintRightRow label="智能契變" value={val("進案文件 [契變 照片上傳系統]")}/></>}<div className="attention"><b>帶看注意事項：</b><span>中人：{val("中人(介紹費)")}</span><p>{val("注意事項")}</p></div>
      </div></div>
      <footer className="approval-footer excel-approval"><span>店東審核</span><span>開發1／開發2<br/><b>{val("開發１/開發２")}</b></span><span>以上<span className="red-text">資料無誤</span>簽名</span><span>業務交件日期：</span><span>助理收件日期：</span></footer>
      <small className="legal-note">◎以上資訊如有記載錯誤，一律依地政機關謄本登記簿為準。</small>
    </article>
    <ChecklistPageV2 val={val}/>
    <div className="preview-print-row"><button className="pill-button primary" onClick={() => window.print()}>列印</button></div>
  </section>;
}

function PrintRightRow({ label, value }: { label: string; value: string }) {
  const canWrap = ["國小", "國中", "高中", "大專", "市場", "公園"].includes(label);
  return <p className={`print-right-row${canWrap ? " print-right-row-wrap" : ""}`}><b>{label}：</b><span>{value}</span></p>;
}

function PreviousPrintableIntake({ draft }: { draft: IntakeData }) {
  const v = draft.values; const val = (...keys: string[]) => { const value = intakeValue(v, ...keys); return keys.includes("開發１/開發２") ? (developerFullNameText(value) || "　") : (value || "　"); };
  const features = [1,2,3,4].map(n => val(`特色說明${n}`));
  const parking = [val("車位"), ...intakeAll(v, "車位型態"), val("車位編號")].filter(x => x.trim()).join("／");
  return <section className="print-document"><article className={`print-page data-page ${draft.propertyKind === "純土地" ? "land-print" : "house-print"}`}><div className="print-brand"><div className="print-logo"><b>台慶不動產</b><small>台南文化崇明加盟店</small></div><h1>物件資料表　送件審核</h1></div><p className="company-line">連城不動產開發有限公司　店址：台南市東區崇明路520號　電話：06-3356699　傳真：06-2681166</p><div className="case-band"><b>案名</b><strong>{val("案名")}</strong><b>屋址</b><span>{val("物件(完整)地址")}</span></div><div className="hero-price"><b>價格</b><span>總　價：</span><strong>{val("契約開價")} 萬</strong><em>□有符合契約簽訂開價</em></div><div className="print-main"><div className="print-left">{draft.propertyKind === "純土地" ? <><PrintSection title="面積"><p className="large-line">土地坪數：<strong>{val("地坪")}坪</strong></p></PrintSection><PrintSection title="土地資訊"><p className="large-line">使用分區：<strong>{val("使用分區")}</strong></p><p>建蔽率／容積率：{val("建蔽率/容積率")}</p><p>座　向：{val("朝向 [土地朝]")}</p><p>臨　路：約 {val("臨路")} 米</p><p>面　寬：約 {val("面寬")} 米</p><p>深　度：約 {val("深度")} 米</p></PrintSection></> : <><PrintSection title="面積"><p>總建物坪數：<strong>{val("總建坪")}坪</strong>　｜土地坪數：{val("地坪")}坪</p><p>室內坪數：{val("室內坪")}坪</p></PrintSection><PrintSection title="增建說明"><p>{val("增建說明")}</p></PrintSection><PrintSection title="車位"><p>車位產權／型態／編號：{parking}</p></PrintSection><PrintSection title="基本資料"><p>物件類型：{val("物件型態")}　｜大樓名稱：{val("大樓名稱")}</p><p>每層戶數：{val("每層戶數")}　｜電梯數：{val("電梯數")}</p><p>管理方式：{val("警衛管理")}　｜管理費：{val("管理費")}</p><p>格局：{val("格局 (房)")}{val("格局 (廳)")}{val("格局 (衛浴)")}{val("格局 (陽台)")}</p><p>現況樓層：{val("透天請寫")}　｜權狀樓別：{val("權狀層數")}</p><p>建築完成日：{val("建築完成日期")}　｜屋齡：　　　　　</p><p>座向：{val("朝向 [房屋朝]", "朝向 [大門朝]")}</p></PrintSection><PrintSection title="環境"><p>臨路：約 {val("臨路")}米　｜建蔽／容積：{val("建蔽率/容積率")}</p><p>面寬：約 {val("面寬")}米　｜深度：約 {val("深度")}米</p></PrintSection></>}<PrintSection title="重點說明" className="features-section">{features.map((feature, i) => <p key={i}>{i + 1}.　{feature}</p>)}</PrintSection></div><div className="print-right"><p>物件編號：{val("委託主約編號")}</p><p>鑰匙位置：{val("鑰匙位置")}</p><p>物件現況：{val("(物件)現況")}</p><p>國小：{val("鄰近國小")}</p><p>國中：{val("鄰近國中")}</p><p>高中：{val("鄰近高中")}</p><p>大專：{val("鄰近大專")}</p><p>市場：{val("市場/購物")}</p><p>公園：{val("公園綠地")}</p><hr/><p>物件相片：{val("當下進案文件 [物件相片]")}</p>{draft.propertyKind === "房屋" && <p>格局圖：{val("進案文件 [格局圖]")}</p>}<p>現詢調：{val("進案文件 [現詢調]")}</p><p>智能照片：{val("進案文件 [物件照片 上傳系統]")}</p>{draft.propertyKind === "房屋" && <><p>智能主約：{val("進案文件 [主約 契約拍照 上傳系統]")}</p><p>智能契變：{val("進案文件 [契變 照片上傳系統]")}</p></>}<div className="attention"><b>帶看注意事項：</b><span>中人：{val("中人(介紹費)")}</span><p>{val("注意事項")}</p></div></div></div><footer className="approval-footer"><span>店東審核</span><span>開發1／開發2<br/><b>{val("開發１/開發２")}</b></span><span>以上資料無誤簽名</span><span>業務交件日期：</span><span>助理收件日期：</span></footer><small className="legal-note">◎以上資訊如有記載錯誤，一律依地政機關謄本登記簿為準。</small></article><ChecklistPage val={val}/></section>;
}

function PrintSection({ title, children, className = "" }: { title: string; children: React.ReactNode; className?: string }) { return <section className={`print-section ${className}`}><b className="vertical-title">{title}</b><div>{children}</div></section>; }

function ChecklistPageV2({ val }: { val: (...keys: string[]) => string }) {
  return <article className="print-page checklist-page">
    <div className="check-strip">案名：{val("案名")}　　標的物：{val("物件(完整)地址")}</div>
    <h1>請開發業務填寫檢查以下資訊</h1>
    <div className="check-price">
      <div><p><b>1.　底價：</b>寫在委託主約 P.2 特約____________萬</p><div className="choice-lines"><div><p className="price-indent">寫在契變____________萬</p><p className="price-indent">賣方口頭____________萬</p></div><span>擇一</span></div></div>
      <div><p><b>2.　%數：</b>寫在委託主約第一面______%</p><div className="choice-lines"><div><p className="price-indent">寫在契變______%</p><p className="price-indent">賣方口頭______%</p></div><span>擇一</span></div></div>
    </div>
    <ol className="check-list" start={3}>
      <li><b>中人</b>　□無，□有介紹人：____________元</li>
      <li className="note-item"><span><b>地號謄本</b>×_____筆　□紙本已附件　□曾經調閱請列印　□未調閱，請助理協助</span><span className="note-line">備註：____________________________________________________________</span></li>
      <li className="note-item"><span><b>建號謄本</b>×_____筆　□紙本已附件　□曾經調閱請列印　□未調閱，請助理協助</span><span className="note-line">備註：____________________________________________________________</span></li>
      <li><b>地籍圖</b>　□大樓：無須調　房屋／土地（□紙本已附件　□曾經調閱請列印　□未調閱）</li>
      <li><b>成果圖</b>　□土地：無須調　有建號（□紙本已附件　□曾經調閱請列印　□未調閱）</li>
      <li>檢查□ 委託人（賣方）- <b>說明書 是否已簽名</b></li>
      <li>檢查□ 委託人（賣方）- <b>個資法 是否已簽名</b></li>
      <li><b>建物權狀</b>共____張　□已附紙本　◆已傳_____LINE 請列印　◆賣方沒給開發寫無權狀切結□已附上</li>
      <li><b>土地權狀</b>共____張　□已附紙本　◆已傳_____LINE 請列印　◆賣方沒給開發寫無權狀切結□已附上</li>
    </ol>
    <div className="check-columns"><div><p className="column-heading"><b>12. 委託契約：</b>檢查是否有填寫</p><p>□ (1).　委託日期</p><p>□ (2).　廣告開價</p><p>□ (3).　承辦人開發簽名</p><p>□ (4).　契約現況勾選</p><b>13. 檢查□勾選委託主約 P2：是否一年內取得</b></div><div><p className="column-heading"><b>14. 賣方資訊：</b>檢查是否有填寫</p><div className="seller-info-columns"><div><p>□ (1).　委託人簽名</p><p>□ (2).　賣方 ID</p><p>□ (3).　賣方生日</p><p>□ (4).　賣方地址</p><p>□ (5).　賣方電話</p></div><div><p>□ (1).　代理人簽名</p><p>□ (2).　賣方 ID</p><p>□ (3).　代理人生日</p><p>□ (4).　代理人地址</p><p>□ (5).　代理人電話</p></div></div></div></div>
    <div className="check-tail"><p><b>15. 契變：</b>　□進案「無」契變<br/>　　　　　　□進案【有】契變（契變編號：______________－檢查契變上□承辦人開發簽名　□委託人賣方簽名）</p><p><b>16. 授權書：</b>(1)出售所有權人共_____人，本件賣方親簽<br/>　　　　　　(2)代理人：進案附上授權書（□正本共_____張／□影本共_____張）。□授權書後補<br/>　　Note：____________________________________________________________</p><p><b>17. 土地使用分區：</b>□本件房屋不需分區<br/>　　　　為都內土地：□進案已附　□曾經申請請列印　□本件為土地請助理申請<br/>　　　　都外土地：使用分區請填寫____________，使用地類別：____________</p><p><b>18. 進案其他文件：</b>____________________________________________________________</p></div>
    <div className="signature-box"><span>以上檢查資訊，填寫人簽名：</span><i aria-hidden="true"/></div>
    <div className="check-strip bottom">案名：{val("案名")}　　標的物：{val("物件(完整)地址")}</div>
  </article>;
}

function ChecklistPage({ val }: { val: (...keys: string[]) => string }) { return <article className="print-page checklist-page"><div className="check-strip">案名：{val("案名")}　　標的物：{val("物件(完整)地址")}</div><h1>請開發業務填寫檢查以下資訊</h1><div className="check-price"><div><b>1.　底價：</b>寫在委託主約 P.2 特約____________萬<br/>　　　　寫在契變____________萬<br/>　　　　賣方口頭____________萬　<span>擇一</span></div><div><b>2.　%數：</b>寫在委託主約第一面______%<br/>　　　　寫在契變______%<br/>　　　　賣方口頭______%　<span>擇一</span></div></div><ol className="check-list" start={3}><li><b>中人</b>　□無，□有介紹人：____________元</li><li><b>地號謄本</b>×_____筆　□紙本已附件　□曾經調閱請列印　□未調閱，請助理協助<br/>　備註：____________________________________________</li><li><b>建號謄本</b>×_____筆　□紙本已附件　□曾經調閱請列印　□未調閱，請助理協助<br/>　備註：____________________________________________</li><li><b>地籍圖</b>　□大樓：無須調　房屋／土地（□紙本已附件　□曾經調閱請列印　□未調閱）</li><li><b>成果圖</b>　□土地：無須調　有建號（□紙本已附件　□曾經調閱請列印　□未調閱）</li><li>檢查□ 委託人（賣方）- <b>說明書 是否已簽名</b></li><li>檢查□ 委託人（賣方）- <b>個資法 是否已簽名</b></li><li><b>建物權狀</b>共____張　□已附紙本　◆已傳_____LINE 請列印　◆賣方沒給開發寫無權狀切結□已附上</li><li><b>土地權狀</b>共____張　□已附紙本　◆已傳_____LINE 請列印　◆賣方沒給開發寫無權狀切結□已附上</li></ol><div className="check-columns"><div><b>12. 委託契約：</b>檢查是否有填寫<p>□ (1).　委託日期</p><p>□ (2).　廣告開價</p><p>□ (3).　承辦人開發簽名</p><p>□ (4).　契約現況勾選</p><b>13. 檢查□勾選委託主約 P2：是否一年內取得</b></div><div><b>14. 賣方資訊：</b>檢查是否有填寫<p>□ (1).　委託人簽名</p><p>□ (2).　賣方 ID</p><p>□ (3).　賣方生日</p><p>□ (4).　賣方地址</p><p>□ (5).　賣方電話</p></div></div><div className="check-tail"><p><b>15. 契變：</b>　□進案「無」契變<br/>　　　　　　□進案【有】契變（契變編號：________－檢查契變上□承辦人開發簽名　□委託人賣方簽名）</p><p><b>16. 授權書：</b>(1)出售所有權人共_____人，本件賣方親簽<br/>　　　　　　(2)代理人：進案附上授權書（□正本共_____張／□影本共_____張）。□授權書後補<br/>　　Note：____________________________________________</p><p><b>17. 土地使用分區：</b>□本件房屋不需分區<br/>　　　　為都內土地：□進案已附　□曾經申請請列印　□本件為土地請助理申請<br/>　　　　都外土地：使用分區請填寫____________，使用地類別：____________</p><p><b>18. 進案其他文件：</b>____________________________________________</p></div><div className="signature-box"><span>以上檢查資訊，填寫人簽名：</span><i aria-hidden="true"/></div><div className="check-strip bottom">案名：{val("案名")}　　標的物：{val("物件(完整)地址")}</div></article>; }

const numericPrice = (value = "") => { const match = String(value).replace(/,/g, "").match(/\d+(?:\.\d+)?/); return match ? Number(match[0]) : Number.NaN; };
const adjustedPriceLabel = (price = "", adjusted = "") => { const original = numericPrice(price), next = numericPrice(adjusted); return Number.isFinite(original) && Number.isFinite(next) ? next < original ? "降" : next > original ? "調" : "" : ""; };
function cellValue(r: RecordItem, key: string) { if (key === "age") return ageOf(r); if (key === "bookLocation") return locationOf(r); if (key === "coverageFar") return [r.coverage, r.far].filter(Boolean).join("／"); if (key === "status") return displayStatus(r); if (key === "archived") return displayRocDate(r.archived || (isExpired(r) ? today() : "")); if (dateKeys.has(key)) return displayRocDate(r[key]); if (key === "price") { const label = adjustedPriceLabel(r.price, r.reducedPrice); return label ? `${r.price || "—"}${label}${r.reducedPrice}` : r.price || ""; } if (key === "type") return typeShort(r.type); if (key === "layout") return typeShort(r.type) === "土地" ? "土地" : layoutFull(r.layout, r.type); if (key === "contractType") return contractShort(contractFromNo(r.propertyNo) || r.contractType); return r[key] || ""; }

function AdvancedFilterModal({ value, setValue, onClose, onReset }: { value: AdvancedFilter; setValue: (next: AdvancedFilter) => void; onClose: () => void; onReset: () => void }) {
  const change = (key: string, next: string) => setValue({ ...value, [key]: next });
  const text = (key: string, label: string, placeholder = "") => <label className="filter-field"><span>{label}</span><input value={value[key] || ""} placeholder={placeholder} onChange={event => change(key, event.target.value)}/></label>;
  const dateRange = (from: string, to: string, label: string) => <div className="filter-field range-field"><span>{label}</span><div><input value={value[from] || ""} aria-label={`${label}開始`} onChange={event => change(from, event.target.value)}/><i>至</i><input value={value[to] || ""} aria-label={`${label}結束`} onChange={event => change(to, event.target.value)}/></div></div>;
  const numberRange = (from: string, to: string, label: string, unit = "") => <div className="filter-field range-field"><span>{label}</span><div><input inputMode="decimal" value={value[from] || ""} placeholder="最小" onChange={event => change(from, event.target.value)}/><i>至</i><input inputMode="decimal" value={value[to] || ""} placeholder="最大" onChange={event => change(to, event.target.value)}/>{unit && <em>{unit}</em>}</div></div>;
  const select = (key: string, label: string, options: string[]) => <label className="filter-field"><span>{label}</span><select value={value[key] || ""} onChange={event => change(key, event.target.value)}><option value="">不限</option>{options.map(option => <option key={option} value={option}>{option === "filled" ? "已填" : option === "empty" ? "未填" : option}</option>)}</select></label>;
  const choices = (key: string, label: string, options: string[]) => <div className="filter-field choice-field"><span>{label}</span><div>{options.map(option => { const shown = option === "filled" ? "已填" : option === "empty" ? "未填" : option; return <label key={option}><input type="checkbox" checked={value[key] === option} onChange={() => change(key, value[key] === option ? "" : option)}/><b>{shown}</b></label>; })}</div></div>;
  const sites: [string, string][] = [["yes319", "YES319"], ["houseinfor", "HOUSE INFOR"], ["homeWeb", "我家網"], ["platform591", "591"], ["price5168", "5168"], ["windowAd", "櫥窗"], ["led", "LED"], ["goldExposure", "黃金曝光"]];
  return <div className="modal-backdrop"><div className="modal advanced-filter-modal"><div className="modal-head"><div><span>物件搜尋</span><h2>進階篩選</h2></div><button className="close" onClick={onClose}>×</button></div><div className="advanced-filter-body">
    <section><h3>搜尋來源</h3><div className="advanced-source-choices"><label><input type="checkbox" checked={value.sourceActive === "1"} onChange={event => change("sourceActive", event.target.checked ? "1" : "")}/><b>委託中物件</b></label><label><input type="checkbox" checked={value.sourceArchive === "1"} onChange={event => change("sourceArchive", event.target.checked ? "1" : "")}/><b>封存中</b></label></div></section>
    <section><h3>基本條件</h3><div className="advanced-filter-grid">{text("propertyNo", "物件編號")}{text("caseName", "案名")}{text("address", "地址")}{choices("type", "種類", ["透天", "華廈", "大樓", "公寓", "廠房", "土地"])}{text("direction", "朝向")}{text("developer", "開發業務")}{text("zoning", "使用分區")}{text("key", "鑰匙")}{text("currentState", "現況")}{choices("parking", "車位", ["坡道", "平面", "機械", "昇降", "庭院", "平移"])}{choices("salesBook", "銷售本", ["製作", "未製作", "完成"])}{choices("photoInfo", "照片", ["filled", "empty"])}</div></section>
    <section><h3>日期</h3><div className="advanced-filter-grid dates">{dateRange("entrustStartFrom", "entrustStartTo", "委託開始")}{dateRange("entrustEndFrom", "entrustEndTo", "委託結束")}{dateRange("reportFrom", "reportTo", "進案日期")}{dateRange("updateFrom", "updateTo", "更新日期")}{dateRange("groupFrom", "groupTo", "團看日期")}{dateRange("bookFrom", "bookTo", "物件本位置日期")}{dateRange("salesFrom", "salesTo", "銷售本日期")}</div></section>
    <section><h3>價格、坪數與樓層</h3><div className="advanced-filter-grid">{numberRange("priceFrom", "priceTo", "開價（萬）", "萬")}{numberRange("ageFrom", "ageTo", "屋齡", "年")}{text("floor", "所在樓層", "例：2")}{numberRange("indoorFrom", "indoorTo", "室內坪", "坪")}{numberRange("buildingFrom", "buildingTo", "建坪", "坪")}{numberRange("landFrom", "landTo", "土地坪", "坪")}{numberRange("roadFrom", "roadTo", "臨路", "米")}{numberRange("frontageFrom", "frontageTo", "面寬", "米")}{numberRange("depthFrom", "depthTo", "深度", "米")}{text("rooms", "格局－房", "例：3")}{text("halls", "格局－廳", "例：2")}{text("baths", "格局－衛", "例：2")}</div></section>
    <section><h3>網站編號</h3><p className="filter-help">勾選「已填」或「未填」，可和其他條件一起篩選。</p><div className="advanced-filter-grid website-filter-grid">{sites.map(([key, label]) => choices(key, label, ["filled", "empty"]))}</div></section>
  </div><div className="modal-foot"><button onClick={onReset}>清除全部條件</button><button className="primary" onClick={onClose}>套用篩選</button></div></div></div>;
}

function PropertyTable({ records, columns, publicMode = false, dailyMode = false, archiveMode = false, activeLead = false, expiryAnnotation = false, showAreaGroups = false, zoom = 100, onEdit, onArchive, onRestore, onRemove, onPptChange }: { records: RecordItem[]; columns: string[]; publicMode?: boolean; dailyMode?: boolean; archiveMode?: boolean; activeLead?: boolean; expiryAnnotation?: boolean; showAreaGroups?: boolean; zoom?: number; onEdit: (r: RecordItem) => void; onArchive: (r: RecordItem, status?: string) => void; onRestore: (r: RecordItem) => void; onRemove: (r: RecordItem) => void; onPptChange?: (r: RecordItem, patch: Partial<RecordItem>) => void }) {
  const [publicScrollLeft, setPublicScrollLeftState] = useState(0);
  const setPublicScrollLeft = (value: number) => setPublicScrollLeftState(value > 0 ? value + 81 : 0);
  const [hoveredPublicCase, setHoveredPublicCase] = useState("");
  const [hoveredPublicCaseTop, setHoveredPublicCaseTopState] = useState(0);
  const setHoveredPublicCaseTop = (value: number) => setHoveredPublicCaseTopState(Math.max(0, (value - 6) * zoom / 100 + 6));
  useEffect(() => {
    if (!publicMode) return;
    const rows = Array.from(document.querySelectorAll<HTMLTableRowElement>(".public-table tbody tr:not(.area-group-row)"));
    const cleanups = rows.map((row, index) => {
      const activate = () => {
        const record = records[index];
        if (!record) return;
        setHoveredPublicCase(record.caseName || record.propertyNo || "未命名案件");
        setHoveredPublicCaseTop(row.offsetTop + 6);
      };
      row.addEventListener("touchstart", activate, { passive: true });
      row.addEventListener("pointerdown", activate, { passive: true });
      return () => { row.removeEventListener("touchstart", activate); row.removeEventListener("pointerdown", activate); };
    });
    return () => cleanups.forEach(cleanup => cleanup());
  }, [publicMode, records, zoom]);
  const showPpt = !!onPptChange && !publicMode;
  const areaGroup = (record: RecordItem) => isRentalRecord(record) ? "租件" : areaCategory(record);
  const stickyCount = publicMode ? 0 : activeLead ? 7 : 4;
  const displayedRecords = publicMode ? sortActiveRecords(records) : records;
  if (publicMode) { records = displayedRecords; showAreaGroups = true; }
  return <div className={`table-wrap ${publicMode ? "public-table" : ""} ${dailyMode ? "daily-table" : ""} ${activeLead ? "active-lead-table" : ""}`} onScroll={event => publicMode && setPublicScrollLeft(event.currentTarget.scrollLeft)}>{publicMode && publicScrollLeft > 80 && hoveredPublicCase && <div className="public-floating-case-name" style={{ left: publicScrollLeft + 8, top: hoveredPublicCaseTop }}>{hoveredPublicCase}</div>}<table style={publicMode ? { zoom: zoom / 100 } : undefined}><thead><tr>{columns.map((k, i) => <th key={k} className={`${i < stickyCount ? `sticky sticky-${i}` : ""} col-${k}`}><ColumnLabel column={k}/></th>)}{showPpt && <><th className="ppt-col">加入PPT</th><th className="ppt-type-col">PPT分類</th></>}{!publicMode && records.some(r => r.archived || isExpired(r) || r.status !== "委託中") && <th className="actions-col">操作</th>}</tr></thead><tbody>{records.map((r, recordIndex) => { const highlighted: string[] = (() => { try { return JSON.parse(r._dailyHighlight || "[]"); } catch { return []; } })(); const isHighlighted = (key: string) => highlighted.includes(key) || (key === "caseName" && highlighted.includes("caseNameNote")) || (key === "coverageFar" && (highlighted.includes("coverage") || highlighted.includes("far"))); const archiveDate = r.archived || (isExpired(r) ? r.entrustEnd : ""); const archiveDateLabel = displayRocDate(archiveDate); const archiveLabel = `${archiveDateLabel}${r.archived ? (r.status || "下架") : "到期下架"}${r.status === "下架洽開發" && r.archiveReason ? `:${r.archiveReason}` : ""}`; const restoredLabel = r._restoredAt ? shortRocMonthDay(r._restoredAt) : ""; const expiryDays = daysUntil(r.entrustEnd); return <Fragment key={r.id}>{showAreaGroups && (recordIndex === 0 || areaGroup(records[recordIndex - 1]) !== areaGroup(r)) && <tr className="area-group-row"><td colSpan={columns.length + (showPpt ? 2 : 0)}><span>{areaGroup(r)}</span></td></tr>}<tr onMouseEnter={event => { if (publicMode) { setHoveredPublicCase(r.caseName || r.propertyNo || "未命名案件"); setHoveredPublicCaseTop(event.currentTarget.offsetTop + 6); } }} onMouseLeave={() => publicMode && setHoveredPublicCase("")}>{columns.map((k, i) => <td key={k} className={`${i < stickyCount ? `sticky sticky-${i}` : ""} col-${k} ${dateKeys.has(k) && k !== "groupViewDate" && !validDate(r[k]) ? "date-error" : ""} ${dailyMode && isHighlighted(k) ? "daily-updated-cell" : ""}`}>{k === "caseName" && publicMode ? <><span className="public-case-name">{chunkText(r.caseName || "—", 10).map((line, lineIndex) => <span key={`${line}-${lineIndex}`}>{line}</span>)}</span>{r.caseNameNote && <small className="case-name-note">{r.caseNameNote}</small>}{dailyMode && r._dailyAnnotation && <small className="daily-case-annotation">{r._dailyAnnotation}</small>}{expiryAnnotation && expiryDays >= 0 && expiryDays <= 30 && <small className="mine-expiry-annotation">提醒{displayRocDate(r.entrustEnd)}到期</small>}</> : k === "caseName" ? <><button className="case-link" onClick={() => onEdit(r)}>{chunkText(r.caseName || "—", activeLead ? 12 : 15).map((line, lineIndex) => <span className="case-name-line" key={`${line}-${lineIndex}`}>{line}</span>)}</button>{r.caseNameNote && <small className="case-name-note">{r.caseNameNote}</small>}{activeLead && restoredLabel && !String(r.caseNameNote || "").includes("重新上架") && <small className="restore-case-annotation">{restoredLabel}</small>}{archiveMode && <small className="archive-case-annotation">{archiveLabel}</small>}</> : k === "status" ? <span className={`status ${displayStatus(r) === "委託中" ? "live" : "off"}`}>{displayStatus(r)}</span> : <CellContent record={r} column={k}/>}</td>)}{showPpt && <><td className="ppt-col"><input type="checkbox" checked={pptIncluded(r)} onChange={e => onPptChange(r, { pptSelected: e.target.checked ? "1" : "0", pptCategory: r.pptCategory || (isCurrentPptWeek(r.reportDate) ? "本週進案" : "臨時新增") })}/></td><td className="ppt-type-col"><select value={pptCategory(r)} onChange={e => onPptChange(r, { pptCategory: e.target.value, pptSelected: "1" })}><option>本週進案</option><option>臨時新增</option></select></td></>}{!publicMode && (r.archived || isExpired(r) || r.status !== "委託中") && <td className="actions-col"><div className="row-actions"><button onClick={() => onRestore(r)}>恢復</button><button className="danger" onClick={() => onRemove(r)}>刪除</button></div></td>}</tr></Fragment>;})}</tbody></table></div>;
}

function ColumnLabel({ column }: { column: string }) {
  if (column === "dailyHide") return null;
  if (column === "price") return <>總價<small>（萬）</small></>;
  if (column === "coverageFar") return <>建蔽率<br/>容積率</>;
  if (column === "reportDate") return <>進案報件<br/>日期</>;
  if (column === "entrustPeriod") return <>委託開始<br/>委託結束</>;
  if (column === "bookLocation") return <>物件本<br/>位置</>;
  if (column === "zoning") return <>使用<br/>分區</>;
  if (column === "managementFee") return <>管理費</>;
  if (column === "road") return <>臨路</>;
  if (column === "frontage") return <>面寬</>;
  if (column === "depth") return <>深度</>;
  if (column === "houseinfor") return <>HOUSE<br/>INFOR</>;
  if (column === "windowAd") return <>櫥窗廣告<small>（只上專簽）</small></>;
  if (column === "led") return <>LED<small>（只上專簽）</small></>;
  if (column === "price5168") return <>5168<br/>實價網</>;
  if (column === "coverage") return <>建蔽率</>;
  if (column === "far") return <>容積率</>;
  if (column === "coverageFar") return <>建蔽率<br/>容積率</>;
  return <>{labels[column]}</>;
}

function CellContent({ record: r, column: k }: { record: RecordItem; column: string }) {
  if (k === "type") return <>{typeShort(r.type) || "—"}</>;
  if (k === "direction") return <>{directionShort(r.direction) || "—"}</>;
  if (k === "dailyHide") return r._dailyHideKey ? <button className="daily-hide-button" type="button" title="不顯示於前台每日物件動態" onClick={() => window.dispatchEvent(new CustomEvent("hide-daily-item", { detail: r._dailyHideKey }))}>刪</button> : null;
  if (k === "propertyNo" && r._dailyHideKey) return <span className="daily-property-number"><button type="button" title="不顯示於前台每日物件動態" onClick={() => window.dispatchEvent(new CustomEvent("hide-daily-item", { detail: r._dailyHideKey }))}>刪</button><span>{r.propertyNo || "—"}</span></span>;
  if (k === "propertyNo" && r._showEntrustEndInPropertyNo === "1") { const remaining = daysUntil(r.entrustEnd); const expiring = remaining >= 0 && remaining <= 30; return <span className="mine-property-number"><b>{r.propertyNo || "—"}</b><small className={expiring ? "expiring" : ""}>{displayRocDate(r.entrustEnd) || "—"}到期</small></span>; }
  if (k === "area") { const district = districtFromAddress(r.address) || String(r.area || "").trim(); const lineSize = Array.from(district).length === 4 ? 2 : 4; return <span className="area-lines">{chunkText(district, lineSize).map((line, index) => <span key={`${line}-${index}`}>{line}</span>)}</span>; }
  if (k === "housingRemoval") return r.housingDownDate ? <small className="archive-down-note">{displayRocDate(r.housingDownDate)}下架</small> : <span className="archive-down-pending">待下架</span>;
  if (k === "entrustPeriod") return <span className="two-line-value entrust-period-value"><span>{displayRocDate(r.entrustStart) || "—"}</span><span>{displayRocDate(r.entrustEnd) || "—"}</span></span>;
  if (["platform591", "price5168", "goldExposure", "yes319", "houseinfor", "homeWeb", "windowAd", "led"].includes(k)) {
    if (r[`${k}None`] === "1") return <>無</>;
    const website = websiteCellDisplay(r, k);
    return <span className="two-line-value website-list-value"><span>{website.value}</span>{website.notes.map((note, index) => <small className={note.includes("下架") ? "website-inline-note archive-down-note" : "website-inline-note"} key={`${note}-${index}`}>{note}</small>)}</span>;
  }
  if (k === "salesBook") return <span className="two-line-value archive-field-value"><span>{displayRocDate(r.salesBookDate || "") || "—"}</span>{(r.salesBookDate || r.salesBook) && <small>{String(r.salesBook || "").includes("更新") ? "更新" : "製作"}</small>}{r.salesBookDownDate && <small className="archive-down-note">{displayRocDate(r.salesBookDownDate)}下架</small>}</span>;
  if (k === "updateDate") return <>{displayRocDate(r.updateDate) || "—"}</>;
  if (k === "key") { const value = String(r.key || "").trim(); const landNoKey = value.match(/^土地\s*[（(]\s*無鑰匙\s*[）)]$/); if (landNoKey) return <span className="two-line-value land-no-key"><span>土地</span><small>（無鑰匙）</small></span>; const lines = chunkText(value, 5); const keyLength = Array.from(value).length; const sizeClass = keyLength > 10 ? " key-text-long" : keyLength > 5 ? " key-text-medium" : ""; return lines.length ? <span className={`floor-lines key-cell-value${sizeClass}`}>{lines.map((line, index) => <span key={`${line}-${index}`}>{line}</span>)}</span> : <>—</>; }
  if (k === "entrustEnd") { const remaining = daysUntil(r.entrustEnd); return <span className="two-line-value"><span className={remaining >= 0 && remaining <= 30 ? "expiry-date-text" : ""}>{displayRocDate(r.entrustEnd) || "—"}</span>{r.contractChangeNo && <small className="contract-change-number">{r.contractChangeNo}</small>}</span>; }
  if (k === "developer") { const names = developerNameLines(r.developer, developerPersonnelForDisplay); return names.length ? <span className="floor-lines developer-cell-value">{names.map((name, index) => <span key={`${name}-${index}`}>{name}</span>)}</span> : <>—</>; }
  if (k === "price") { const price = String(r.price || "").replace(/\s*萬(?:元)?\s*/g, "").trim(); const adjusted = String(r.reducedPrice || "").replace(/\s*萬(?:元)?\s*/g, "").trim(); const label = adjustedPriceLabel(price, adjusted); const dailyFields: string[] = (() => { try { return JSON.parse(r._dailyHighlight || "[]"); } catch { return []; } })(); return <span className="two-line-value"><b className={dailyFields.includes("price") ? "daily-price-original-updated" : ""}>{price || "—"}</b>{label && <span className={`price-reduction${dailyFields.includes("reducedPrice") ? " daily-price-reduction-updated" : ""}`}>{label}{adjusted}</span>}</span>; }
  if (k === "floor") { const titleFloor = String(r.titleFloor || "").trim(); const currentFloor = String(r.currentFloor || "").trim(); if (!titleFloor && !currentFloor) return <></>; const invalidFloorParts = String(r.floor || "").split(/[／/]/).map(value => value.replace(/\s/g, "").toUpperCase()).filter(Boolean); if (invalidFloorParts.length && invalidFloorParts.every(value => ["土地", "建地", "無", "0T", "農地", "遊樂區"].includes(value))) return <></>; const floor = floorShortFixed(r.floor); const parts = floor.match(/^(\d+T)(\/.*增建)$/); return parts ? <span className="two-line-value"><span>{parts[1]}</span><small>{parts[2]}</small></span> : <>{floor || "—"}</>; }
  if (k === "layout") return <>{layoutShort(r.layout, r.type) || "—"}</>;
  if (k === "parking") {
    if (typeShort(r.type) === "土地" || /^(?:LG|LA)/i.test(String(r.propertyNo || ""))) return <></>;
    const ownership = String(r.parkingOwnership || "").trim();
    if (ownership === "無車位" || /無車位/.test(String(r.parking || ""))) return <>無車位</>;
    const combined = String(r.parking || "");
    const method = String(r.parkingMethod || r.parkingType || "").trim() || ["坡道/平面", "坡道/機械", "昇降/平面", "昇降/機械", "獨立車庫", "庭院", "車庫", "平移/機械"].find(value => combined.includes(value)) || "";
    const methodShort: Record<string, string> = { "坡道/平面": "坡平", "坡道/機械": "坡機", "昇降/平面": "昇平", "昇降/機械": "昇機", "獨立車庫": "車庫", "庭院": "庭院", "車庫": "車位", "平移/機械": "平機" };
    const shortType = methodShort[method] || method;
    const parkingNo = String(r.parkingNo || "").trim().replace(/號$/g, "");
    if (!shortType && !parkingNo) return <>{parkingShort(combined) || "—"}</>;
    return <span className="two-line-value parking-list-value"><span>{shortType || "—"}</span>{parkingNo && <small>{parkingNo}</small>}</span>;
  }
  if (k === "indoorPing") return <>{String(r.registryIndoorPing || r.indoorPing || "").replace(/\s*坪\s*$/, "").trim() || "—"}</>;
  if (["buildingPing", "landPing"].includes(k)) return <>{String(r[k] || "").replace(/\s*坪\s*$/, "").trim() || "—"}</>;
  if (k === "zoning") { const chars = Array.from(r.zoning || ""); const lines = Array.from({ length: Math.ceil(chars.length / 5) }, (_, index) => chars.slice(index * 5, index * 5 + 5).join("")); return lines.length ? <span className="floor-lines">{lines.map((line, index) => <span key={`${line}-${index}`}>{line}</span>)}</span> : <>—</>; }
  if (k === "coverageFar") { const pairs = String(r.coverageCombined || "").split(/[、,，\n]+/).map(value => value.trim()).filter(value => /^\d+(?:\.\d+)?[\/／]\d+(?:\.\d+)?$/.test(value)); return pairs.length > 1 ? <span className="coverage-far-value coverage-pair-value">{pairs.map((pair, index) => <span key={`${pair}-${index}`}>{pair.replace("／", "/")}</span>)}</span> : <span className="coverage-far-value"><span>{r.coverage ? `${r.coverage.replace(/\s*%\s*$/, "")}%` : "—"}</span><span>{r.far ? `${r.far.replace(/\s*%\s*$/, "")}%` : "—"}</span></span>; }
  if (k === "age") { const text = ageOf(r); const parts = text.match(/^(.*?建)（(.*)）$/); return parts ? <span className="two-line-value"><b>{parts[1]}</b><small>（{parts[2]}）</small></span> : <>{text || "—"}</>; }
  if (k === "bookLocation") { const rocDate = displayRocDate(r.bookLocationDate || ""); const shortDate = rocDate ? rocDate.split("/").slice(-2).join("/") : ""; const location = r.bookLocationType || ""; return <span className="book-location-value"><b>{shortDate || "—"}{location || ""}</b>{r.bookDownDate && <small className="archive-down-note">{displayRocDate(r.bookDownDate)}下架</small>}</span>; }
  if (k === "road") { const values = String(r.road || "").match(/\d+(?:\.\d+)?/g) || []; return values.length ? <span className="floor-lines road-lines">{values.map((value, index) => <span className="measure-number" key={`${value}-${index}`}>{value}</span>)}<small>米</small></span> : <>—</>; }
  if (["frontage", "depth"].includes(k)) { const measure = String(r[k] || "").replace(/\s*米\s*$/, "").trim(); const isHouseZero = typeShort(r.type) !== "土地" && /^(?:0|0\.0+)$/.test(measure); const values = measure.match(/\d+(?:\.\d+)?/g) || []; return isHouseZero ? <>-</> : values.length ? <span className="two-line-value measure-value">{values.map((value, index) => <span className="measure-number" key={`${value}-${index}`}>{value}</span>)}<small>米</small></span> : measure ? <>{measure}</> : <>—</>; }
  if (["coverage", "far"].includes(k)) return r[k] ? <span className="two-line-value"><b>{r[k].replace(/\s*%\s*$/, "")}</b><small>%</small></span> : <>—</>;
  if (k === "managementFee") {
    const rawFee = String(r.managementFee || "").trim();
    const fee = rawFee.replace(/元/g, "").trim();
    if (!fee) return <>—</>;
    if (/^無(?:\$?0(?:\.0+)?)?(?:\s*[\/／]\s*月)?$/.test(fee) || /^\$?0(?:\.0+)?(?:\s*[\/／]\s*月)?$/.test(fee)) return <>無$0</>;
    if (/[\/／]\s*(?:年|月)$/.test(fee)) return <>{fee.replace(/／/g, "/")}</>;
    return <>{`${fee}/月`}</>;
  }
  if (k === "notes") {
    const internalNoteText = String(r.notes || "").trim();
    return <>{internalNoteText || "—"}</>;
  }
  return <>{cellValue(r, k) || "—"}</>;
}

function PropertyBookReview({ records, settings, openRequest, submit, openRecord }: { records: RecordItem[]; settings: Settings; openRequest: number; submit: (reviews: Array<{ record: RecordItem; status: string }>) => void; openRecord: (record: RecordItem) => void }) {
  const cycleStart = normalizeDateInput(settings.bookReviewCurrentDate || "") || "2026-07-30";
  const nextCheckDate = normalizeDateInput(settings.bookReviewNextDate || "") || "2026-09-30";
  const reviewCycle = today() >= cycleStart ? `book-${cycleStart}` : "";
  // 一輪物件本確認只需完成一次；確認日在本次起始日之後（例如 8/10 完成 7/30 這輪）即不再列入。
  const dueRecords = reviewCycle ? records.filter(record => {
    const confirmedDate = normalizeDateInput(record._bookReviewAt || record.bookLocationDate || "");
    return !record.archived && !isExpired(record) && (record.status || "委託中") === "委託中" && (!confirmedDate || confirmedDate < cycleStart);
  }) : [];
  const [reviewOpen, setReviewOpen] = useState(false); const [reviewVisible, setReviewVisible] = useState(true); const [values, setValues] = useState<Record<string, string>>({}); const [scannerOpen, setScannerOpen] = useState(false); const [scannedId, setScannedId] = useState(""); const [scannedIds, setScannedIds] = useState<string[]>([]); const [scanMessage, setScanMessage] = useState(""); const [archivedScanRecord, setArchivedScanRecord] = useState<RecordItem | null>(null); const videoRef = useRef<HTMLVideoElement | null>(null); const streamRef = useRef<MediaStream | null>(null); const scannedSetRef = useRef<Set<string>>(new Set());
  useEffect(() => { if (openRequest > 0) { setReviewVisible(true); setReviewOpen(true); } }, [openRequest]);
  const reportMobileResult = (id: number | undefined, ok: boolean, message: string) => { if (!id) return; void fetch("http://localhost:8765/api/mobile-qr-result", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ id, ok, message }) }).catch(() => {}); };
  const locateCode = (raw: string, mobileEventId?: number) => { const clean = String(raw || "").trim(); const record = records.find(item => clean.includes(item.propertyNo) || (item.address && clean.includes(item.address)) || (item.caseName && clean.includes(item.caseName))); if (!record) { const message = `刷到的 QR Code 不在目前物件本：${clean || "未讀到內容"}`; setScanMessage(`警示：${message}`); reportMobileResult(mobileEventId, false, message); alert(`警示：${message}`); return; } if (scannedSetRef.current.has(record.id)) { const message = `${record.propertyNo} 已掃描，請刷下一件`; setScanMessage(message); reportMobileResult(mobileEventId, false, message); return; } const archived = !!record.archived || isExpired(record) || (record.status || "委託中") !== "委託中"; if (archived) { const message = `${record.propertyNo}「${record.caseName || "未命名案件"}」已${record.status || "下架"}`; scannedSetRef.current.add(record.id); setScannerOpen(false); setReviewOpen(true); setArchivedScanRecord(record); setScanMessage(`通知：${message}。`); reportMobileResult(mobileEventId, false, message); return; } scannedSetRef.current.add(record.id); setScannedId(record.id); setScannedIds(previous => [...previous, record.id]); setScanMessage(`委託中物件，完成，下一件：${record.propertyNo}`); reportMobileResult(mobileEventId, true, "委託中物件，完成，下一件"); submit([{ record, status: record.bookLocationType || "架上" }]); };
  const decodeWithJsQR = (source: CanvasImageSource, width: number, height: number) => { const canvas = document.createElement("canvas"); canvas.width = width; canvas.height = height; const context = canvas.getContext("2d", { willReadFrequently: true }); if (!context) return ""; context.drawImage(source, 0, 0, width, height); const image = context.getImageData(0, 0, width, height); return jsQR(image.data, width, height, { inversionAttempts: "attemptBoth" })?.data || ""; };
  const decodeImage = async (file?: File) => { if (!file) return; try { const bitmap = await createImageBitmap(file); const Detector = (window as any).BarcodeDetector; let raw = ""; if (Detector) { const codes = await new Detector({ formats: ["qr_code"] }).detect(bitmap); raw = codes[0]?.rawValue || ""; } if (!raw) raw = decodeWithJsQR(bitmap, bitmap.width, bitmap.height); bitmap.close(); if (!raw) return alert("圖片中沒有辨識到QR Code，請換一張較清楚的照片"); locateCode(raw); } catch { alert("QR Code圖片讀取失敗，請換一張較清楚的照片"); } };
  // 手機掃描頁在區網服務（8765/8766），但總表可從 GitHub Pages 開啟；
  // 因此桌面總表固定向本機掃描服務讀取，掃描後才能即時寫回目前的總表資料。
  useEffect(() => { let stopped = false; const endpoint = "http://localhost:8765/api/mobile-qr"; const poll = async () => { try { const response = await fetch(endpoint, { cache: "no-store" }); if (!response.ok) return; const data = await response.json(); if (!stopped) (data.events || []).forEach((event: { id?: number; code?: string }) => event.code && locateCode(event.code, event.id)); } catch {} }; poll(); const timer = window.setInterval(poll, 900); return () => { stopped = true; clearInterval(timer); }; }, [records.map(record => `${record.id}:${record._bookReviewAt || ""}`).join("|")]);
  useEffect(() => { if (!scannerOpen) { streamRef.current?.getTracks().forEach(track => track.stop()); streamRef.current = null; return; } let stopped = false, timer = 0; const start = async () => { const Detector = (window as any).BarcodeDetector; try { const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" } }); streamRef.current = stream; if (!videoRef.current) return; videoRef.current.srcObject = stream; await videoRef.current.play(); const detector = Detector ? new Detector({ formats: ["qr_code"] }) : null; timer = window.setInterval(async () => { if (stopped || !videoRef.current || videoRef.current.readyState < 2) return; let raw = ""; if (detector) { const codes = await detector.detect(videoRef.current); raw = codes[0]?.rawValue || ""; } if (!raw) raw = decodeWithJsQR(videoRef.current, videoRef.current.videoWidth, videoRef.current.videoHeight); if (raw) locateCode(raw); }, 500); } catch { alert("無法開啟相機，請允許相機權限或改用上傳QR圖片"); setScannerOpen(false); } }; start(); return () => { stopped = true; if (timer) clearInterval(timer); streamRef.current?.getTracks().forEach(track => track.stop()); streamRef.current = null; }; }, [scannerOpen]);
  useEffect(() => { const close = (event: KeyboardEvent) => { if (event.key !== "Escape" || !reviewVisible) return; event.stopImmediatePropagation(); if (scannerOpen) setScannerOpen(false); else setReviewVisible(false); }; window.addEventListener("keydown", close, true); return () => window.removeEventListener("keydown", close, true); }, [reviewVisible, scannerOpen]);
  // 全數完成時按鈕與視窗都隱藏；下一輪日期到或有新待確認物件時才再次顯示。
  if (!reviewVisible || !dueRecords.length) return null;
  const statusOf = (record: RecordItem) => record.bookLocationType || "架上";
  const bookAreaOrder = ["北區", "東區", "中西區", "南區", "永康區", "安平區", "仁德區", "安南區", "其他區", "外縣市"];
  const groupedDueRecords = bookAreaOrder.map(area => ({ area, items: dueRecords.filter(record => areaCategory(record) === area) })).filter(group => group.items.length);
  return <div className="modal-backdrop book-review-reminder-backdrop"><section className={`book-review-panel ${reviewOpen ? "book-review-open" : "book-review-collapsed"}`}><div className="book-review-head" role="button" tabIndex={0} onClick={() => setReviewOpen(open => !open)} onKeyDown={event => (event.key === "Enter" || event.key === " ") && setReviewOpen(open => !open)}><div className="book-review-title"><b>物件本確認</b><span>本次確認：{displayRocDate(cycleStart)}　下次確認：{displayRocDate(nextCheckDate)}</span></div><div className="book-review-actions" onClick={event => event.stopPropagation()} onKeyDown={event => event.stopPropagation()}><strong>{dueRecords.length} 筆待確認</strong>{reviewOpen && <><button onClick={() => { setScanMessage(""); setArchivedScanRecord(null); setScannerOpen(true); }}>掃描QR Code</button><label className="file-button">上傳QR圖片<input type="file" accept="image/*" capture="environment" onChange={event => { decodeImage(event.target.files?.[0]); event.target.value = ""; }}/></label></>}<button className="close book-review-close" title="關閉" onClick={() => setReviewVisible(false)}>×</button></div></div>{reviewOpen && <>{archivedScanRecord && <div className="book-archived-scan-notice"><span>{archivedScanRecord.propertyNo}「{archivedScanRecord.caseName || "未命名案件"}」已{archivedScanRecord.status || "下架"}</span><button type="button" onClick={() => { setReviewVisible(false); openRecord(archivedScanRecord); }}>開啟案件編輯</button><button type="button" onClick={() => setArchivedScanRecord(null)}>留在掃描</button></div>}<div className="book-review-list">{groupedDueRecords.map(group => <Fragment key={group.area}><div className="book-review-area"><span>{group.area}</span><small>{group.items.length}筆</small></div>{group.items.map(record => <div id={`book-review-${record.id}`} className="book-review-row" key={record.id}><div><b>{record.caseName || record.propertyNo}</b><small className="book-record-line"><span>{record.propertyNo}</span><span>{record.address || "未填地址"}</span><span>開發：{developerFullNameText(record.developer, settings.personnel) || "未填寫"}</span></small></div><div className="book-review-side"><small className="book-current-note">{displayRocDate(record.bookLocationDate || "") || "未填日期"}　{record.bookLocationType || "未填位置"}</small><div className="book-review-confirm"><strong className="book-scan-status">待確認</strong><button type="button" onClick={() => submit([{ record, status: statusOf(record) }])}>今日確認</button></div></div></div>)}</Fragment>)}</div><div className="book-review-foot"><span>本次已掃描 {scannedIds.length} 筆</span></div></>}{scannerOpen && <div className="qr-scanner-backdrop"><div className="qr-scanner-modal"><div><b>連續掃描物件QR Code</b><button className="close" onClick={() => setScannerOpen(false)}>×</button></div><video ref={videoRef} playsInline muted/><strong className="qr-scan-count">已掃描 {scannedIds.length} 筆</strong><p>{scanMessage || "將QR Code放在畫面中央，掃完一件可直接刷下一件。"}</p><button className="primary qr-scan-done" onClick={() => setScannerOpen(false)}>掃描完成</button></div></div>}</section></div>;
}

function MonthlyConfirmationProgress({ records, people, lastLogins, close }: { records: RecordItem[]; people: Person[]; lastLogins: Record<string, string>; close: () => void }) {
  type ProgressSortKey = "person" | "total" | "completed" | "pending" | "lastLogin" | "latest" | "status";
  const [sort, setSort] = useState<{ key: ProgressSortKey; direction: "asc" | "desc" }>({ key: "pending", direction: "desc" });
  const activePeople = people.filter(person => person.status === "在職" && person.role !== "秘書");
  const progress = activePeople.map(person => {
    const assigned = records.filter(record => nameMatches(record.developer, person.name)).filter(record => {
      const date = normalizeDateInput(record.reportDate || "");
      return !!date && Math.floor((Date.parse(`${today()}T00:00:00`) - Date.parse(`${date}T00:00:00`)) / 86400000) >= 30;
    });
    const reportOf = (record: RecordItem) => { try { return (Object.values(JSON.parse(record._monthlyReports || "{}")) as any[]).filter(report => report.personId === person.id).sort((a, b) => String(b.reportedAt || "").localeCompare(String(a.reportedAt || "")))[0]; } catch { return undefined; } };
    const pending = assigned.filter(record => { const report = reportOf(record); if (!report) return true; if (report.status === "待確認") return String(report.dueDate || "") <= today(); const date = String(report.reportedAt || "").slice(0, 10); return !date || Math.floor((Date.parse(`${today()}T00:00:00`) - Date.parse(`${date}T00:00:00`)) / 86400000) >= 45; });
    const latest = assigned.map(reportOf).filter(Boolean).sort((a, b) => String(b.reportedAt || "").localeCompare(String(a.reportedAt || "")))[0];
    return { person, total: assigned.length, pending: pending.length, completed: Math.max(0, assigned.length - pending.length), latest: latest?.reportedAt || "", lastLogin: lastLogins[person.id] || "" };
  });
  const sortValue = (item: typeof progress[number], key: ProgressSortKey) => key === "person" ? item.person.name : key === "status" ? (item.pending ? "尚未完成" : "全數完成") : key === "lastLogin" || key === "latest" ? (item[key] ? new Date(item[key]).getTime() : 0) : item[key];
  const sortedProgress = [...progress].sort((a, b) => { const av = sortValue(a, sort.key), bv = sortValue(b, sort.key); const compared = typeof av === "number" && typeof bv === "number" ? av - bv : String(av).localeCompare(String(bv), "zh-TW"); return sort.direction === "asc" ? compared : -compared; });
  const sortBy = (key: ProgressSortKey) => setSort(current => current.key === key ? { key, direction: current.direction === "asc" ? "desc" : "asc" } : { key, direction: "asc" });
  const heading = (key: ProgressSortKey, label: string) => <button type="button" className={sort.key === key ? "selected" : ""} onClick={() => sortBy(key)}>{label}<small>{sort.key === key ? (sort.direction === "asc" ? "▲" : "▼") : "↕"}</small></button>;
  return <div className="modal-backdrop monthly-progress-backdrop"><section className="monthly-progress-modal"><header><div><b>45天確認進度</b><span>查看業務進入前台與回報進度；點選標題可排序</span></div><button className="close" onClick={close}>×</button></header><div className="monthly-progress-list"><div className="monthly-progress-row monthly-progress-labels">{heading("person", "業務")}{heading("total", "應確認")}{heading("completed", "已完成")}{heading("pending", "尚未回報")}{heading("lastLogin", "最後進入前台")}{heading("latest", "最後回報")}{heading("status", "狀態")}</div>{sortedProgress.map(item => <div className={`monthly-progress-row ${item.pending ? "has-pending" : "is-complete"}`} key={item.person.id}><strong>{item.person.name}</strong><span>{item.total} 件</span><span>{item.completed} 件</span><span>{item.pending} 件</span><span>{item.lastLogin ? new Date(item.lastLogin).toLocaleString("zh-TW") : "尚無進入紀錄"}</span><span>{item.latest ? new Date(item.latest).toLocaleString("zh-TW") : "尚無回報"}</span><em>{item.pending ? "尚未完成" : "全數完成"}</em></div>)}{!progress.length && <p className="monthly-progress-empty">尚未設定在職業務人員。</p>}</div><footer><button className="primary" onClick={close}>關閉</button></footer></section></div>;
}

function BusinessReportInbox({ records, resolve, archive }: { records: RecordItem[]; resolve: (record: RecordItem, reportKey: string, keepActive?: boolean) => void; archive: (record: RecordItem, status: string) => void }) {
  const [developerFilter, setDeveloperFilter] = useState("");
  const items = records.flatMap(record => {
    let reports: Record<string, any> = {}; try { reports = JSON.parse(record._monthlyReports || "{}"); } catch {}
    return Object.entries(reports).map(([key, report]) => ({ record, key, report })).filter(item => item.report?.status && !item.report.adminHandledAt);
  }).sort((a, b) => String(b.report.reportedAt || "").localeCompare(String(a.report.reportedAt || "")));
  if (!items.length) return null;
  const developers = Array.from(new Set(items.flatMap(item => developerNameLines(item.record.developer))));
  const shownItems = developerFilter ? items.filter(item => developerNameLines(item.record.developer).includes(developerFilter)) : items;
  const action = (item: { record: RecordItem; key: string; report: any }) => {
    if (item.report.status === "售出下架") return <button className="primary" onClick={() => archive(item.record, "售出下架")}>確認售出下架</button>;
    if (item.report.status === "下架洽開發") return <button className="primary" onClick={() => archive(item.record, "下架洽開發")}>確認下架</button>;
    if (item.report.status === "請跟開發業務2確認") return <small className="business-report-wait">等待另一位開發確認委託中</small>;
    return <button onClick={() => resolve(item.record, item.key, true)}>已查看，等待再次確認</button>;
  };
  return <section className="business-report-inbox"><header><div><div className="business-report-title-row"><b>業務回報待處理</b><div className="business-report-developer-filter">{developers.map(name => <button key={name} className={developerFilter === name ? "selected" : ""} onClick={() => setDeveloperFilter(current => current === name ? "" : name)}>#{name}</button>)}</div></div><span>業務回傳後，請在此完成下架或確認作業</span></div><strong>{shownItems.length}{developerFilter ? `／${items.length}` : ""} 筆</strong></header><div className="business-report-list">{shownItems.map(item => <article key={`${item.record.id}-${item.key}`}><div className="business-report-case"><div className="business-report-case-title"><b>{item.record.propertyNo || "—"}　{item.record.caseName || "未命名案件"}</b><small>開發：{developerFullNameText(item.record.developer) || "未填"}</small></div><small className="business-report-address">地址：{item.record.address || "未填地址"}</small><div className="business-report-action-row"><span>回報：<em>{item.report.status}</em>{item.report.reason ? `　原因：${item.report.reason}` : ""}</span><div className="business-report-action-meta"><small className="business-report-person">{item.report.personName || "業務人員"}　{new Date(item.report.reportedAt || Date.now()).toLocaleString("zh-TW")}</small>{action(item)}</div></div>{item.report.status === "待確認" && item.report.dueDate && <i>下次確認：{displayRocDate(item.report.dueDate)}</i>}</div></article>)}</div></section>;
}

function MonthlyPropertyReport({ records, person, submit }: { records: RecordItem[]; person: Person; submit: (record: RecordItem, status: string, reason: string) => void }) {
  const [choices, setChoices] = useState<Record<string, string>>({}); const [reasons, setReasons] = useState<Record<string, string>>({}); const [missingChoices, setMissingChoices] = useState<string[]>([]); const firstConfirmationDate = "2026-07-30";
  const reportOf = (record: RecordItem) => { try { const reports = Object.values(JSON.parse(record._monthlyReports || "{}")) as any[]; return reports.filter(report => report.personId === person.id).sort((a, b) => String(b.reportedAt || "").localeCompare(String(a.reportedAt || "")))[0]; } catch { return undefined; } };
  const oldEnoughRecords = records.filter(record => { const date = normalizeDateInput(record.reportDate || ""); return !!date && Math.floor((Date.parse(`${today()}T00:00:00`) - Date.parse(`${date}T00:00:00`)) / 86400000) >= 30; });
  const eligibleRecords = oldEnoughRecords.filter(record => { const report = reportOf(record); if (!report) return true; if (report.status === "待確認") return report.dueDate <= today(); const lastDate = String(report.reportedAt || "").slice(0, 10); return !lastDate || Math.floor((Date.parse(`${today()}T00:00:00`) - Date.parse(`${lastDate}T00:00:00`)) / 86400000) >= 45; });
  const overdue = eligibleRecords.filter(record => { const report = reportOf(record); return report?.status === "待確認" && report.dueDate <= today(); });
  useEffect(() => { if (overdue.length) setTimeout(() => alert(`有 ${overdue.length} 筆待確認物件已到期，請重新確認`), 100); }, [overdue.map(record => record.id).join("|")]);
  if (today() < firstConfirmationDate || !eligibleRecords.length) return null;
  const reportOptions = ["委託中", "請跟開發業務2確認", "售出下架", "下架洽開發", "待確認"];
  const submitAll = () => { const selected = eligibleRecords.filter(record => !!choices[record.id]); if (!selected.length) return alert("請至少選擇一筆物件回報"); const missingReason = selected.find(record => choices[record.id] === "下架洽開發" && !String(reasons[record.id] || "").trim()); if (missingReason) { setMissingChoices([missingReason.id]); setTimeout(() => document.getElementById(`report-${missingReason.id}`)?.scrollIntoView({ behavior: "smooth", block: "center" }), 50); return alert(`${missingReason.caseName || missingReason.propertyNo} 必須填寫下架原因`); } selected.forEach(record => submit(record, choices[record.id], reasons[record.id] || "")); setChoices({}); setReasons({}); setMissingChoices([]); };
  return <section className="monthly-report-panel"><div className="monthly-report-head"><div><b>每45天物件確認</b><span>{person.name}，最近45天已確認過的物件不會重複列出</span></div><strong>{eligibleRecords.length} 筆待處理</strong></div><div className="monthly-report-list">{eligibleRecords.map(record => { const report = reportOf(record), choice = choices[record.id] || ""; const options = developerNameLines(record.developer).length > 1 ? reportOptions : reportOptions.filter(option => option !== "請跟開發業務2確認"); return <div id={`report-${record.id}`} className={`monthly-report-row pill-report-row ${report?.status === "待確認" && report.dueDate <= today() ? "overdue" : ""} ${missingChoices.includes(record.id) ? "missing-choice" : ""}`} key={record.id}><div className="monthly-report-case"><b>{record.caseName || record.propertyNo}</b><small>{record.propertyNo}　{record.address}</small>{report?.status === "待確認" && <em>待確認已到期：{displayRocDate(report.dueDate)}</em>}</div><div className="monthly-report-pills">{options.map(option => <Fragment key={option}><button type="button" className={choice === option ? "selected" : ""} disabled={!!choice && choice !== option} onClick={() => { setChoices(previous => ({ ...previous, [record.id]: choice === option ? "" : option })); setMissingChoices(previous => previous.filter(id => id !== record.id)); }}>{option}</button>{option === "下架洽開發" && choice === option && <input className="inline-report-reason" value={reasons[record.id] || ""} onChange={event => { setReasons(previous => ({ ...previous, [record.id]: event.target.value })); if (event.target.value.trim()) setMissingChoices(previous => previous.filter(id => id !== record.id)); }} placeholder="請填寫下架原因"/>}</Fragment>)}</div></div>; })}</div><div className="monthly-report-submit"><button className="primary" onClick={submitAll}>送出回報</button></div></section>;
  return <section className="monthly-report-panel"><div className="monthly-report-head"><div><b>每45天物件確認</b><span>{person.name}，最近45天已確認過的物件不會重複列出</span></div><strong>{eligibleRecords.length} 筆待處理</strong></div><div className="monthly-report-list">{eligibleRecords.map(record => { const report = reportOf(record), choice = choices[record.id] || ""; return <div className={`monthly-report-row ${report?.status === "待確認" && report.dueDate <= today() ? "overdue" : ""}`} key={record.id}><div><b>{record.caseName || record.propertyNo}</b><small>{record.propertyNo}　{record.address}</small>{report?.status === "待確認" && <em>待確認已到期：{displayRocDate(report.dueDate)}</em>}</div><select value={choice} onChange={event => setChoices(previous => ({ ...previous, [record.id]: event.target.value }))}><option value="">選擇回報</option><option>委託中</option><option>請跟 B 業務確認</option><option>請跟開發業務確認</option><option>售出下架</option><option>下架洽開發</option><option>待確認</option></select>{choice === "下架洽開發" && <input value={reasons[record.id] || ""} onChange={event => setReasons(previous => ({ ...previous, [record.id]: event.target.value }))} placeholder="請填寫下架原因"/>}<button disabled={!choice || (choice === "下架洽開發" && !String(reasons[record.id] || "").trim())} onClick={() => submit(record, choice, reasons[record.id] || "")}>送出</button></div>; })}</div></section>;
}

function BusinessInventory({ records, people }: { records: RecordItem[]; people: Person[] }) {
  const [month, setMonth] = useState(today().slice(0, 7));
  const [monthlyDetailName, setMonthlyDetailName] = useState("");
  const [monthlyExpanded, setMonthlyExpanded] = useState(false);
  const [stockDetailName, setStockDetailName] = useState("");
  const monthAt = (offset: number) => { const [year, value] = today().slice(0, 7).split("-").map(Number); const date = new Date(year, value - 1 + offset, 1); return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`; };
  const previousMonth = monthAt(-1), currentMonth = monthAt(0);
  const shortMonth = (value: string) => `${Number(value.split("-")[1])}月`;
  const activeRecords = records.filter(record => !record.archived && !isExpired(record) && (record.status || "委託中") === "委託中");
  const knownNames = people.filter(person => person.role !== "秘書").map(person => String(person.name || "").trim()).filter(Boolean);
  const secretaryNames = people.filter(person => person.role === "秘書").map(person => String(person.name || "").trim()).filter(Boolean);
  const namesFor = (record: RecordItem) => {
    const raw = String(record.developer || "").replace(/[\s、，,／/+&和與]/g, "");
    const parsed = [...new Set(developerNameLines(raw, people).filter(name => !secretaryNames.includes(name)))];
    return parsed.length ? parsed : ["未歸屬業務"];
  };
  const allNames = Array.from(new Set([...knownNames, ...records.flatMap(record => namesFor(record))].filter(Boolean)));
  const monthMatches = (value = "") => normalizeDateInput(value).slice(0, 7) === month;
  const weightFor = (record: RecordItem, name: string) => {
    const names = namesFor(record);
    return names.includes(name) && names.length ? 1 / names.length : 0;
  };
  const sum = (items: RecordItem[], name: string) => items.reduce((total, record) => total + weightFor(record, name), 0);
  const isLand = (record: RecordItem) => /^L/i.test(String(record.propertyNo || "")) || String(record.type || "").includes("土地");
  const isExclusive = (record: RecordItem) => /^(EA|LA|EC)/i.test(String(record.propertyNo || "")) || String(contractFromNo(record.propertyNo) || record.contractType || "").includes("專約");
  const monthlyEntry = records.filter(record => monthMatches(record.reportDate));
  const monthlyArchive = records.filter(record => { const archiveDate = normalizeDateInput(record._archiveActionDate || record.archived); const reportDate = normalizeDateInput(record.reportDate); return Boolean(record.archived) && ((record.status || "委託中") !== "委託中" || Boolean(record._restoredAt)) && monthMatches(reportDate) && monthMatches(archiveDate) && reportDate <= archiveDate; });
  const numberText = (value: number) => Number.isInteger(value) ? String(value) : value.toFixed(2).replace(/0+$/, "").replace(/\.$/, "");
  const categoryTotals = (items: RecordItem[], name: string) => {
    const exclusiveHouse = sum(items.filter(record => !isRentalRecord(record) && isExclusive(record) && !isLand(record)), name);
    const exclusiveLand = sum(items.filter(record => !isRentalRecord(record) && isExclusive(record) && isLand(record)), name);
    const generalHouse = sum(items.filter(record => !isRentalRecord(record) && !isExclusive(record) && !isLand(record)), name);
    const generalLand = sum(items.filter(record => !isRentalRecord(record) && !isExclusive(record) && isLand(record)), name);
    const rental = sum(items.filter(isRentalRecord), name);
    return { exclusiveHouse, exclusiveLand, generalHouse, generalLand, rental, total: exclusiveHouse + exclusiveLand + generalHouse + generalLand + rental };
  };
  const monthlyRows = allNames.map(name => {
    const entered = categoryTotals(monthlyEntry, name), archived = categoryTotals(monthlyArchive, name);
    const total = { ...entered };
    return { name, entered, archived, total };
  }).filter(row => row.total.total > 0).sort((a, b) => b.total.total - a.total.total || a.name.localeCompare(b.name, "zh-TW"));
  const stockRows = allNames.map(name => {
    const exclusiveHouse = sum(activeRecords.filter(record => !isRentalRecord(record) && isExclusive(record) && !isLand(record)), name);
    const exclusiveLand = sum(activeRecords.filter(record => !isRentalRecord(record) && isExclusive(record) && isLand(record)), name);
    const generalHouse = sum(activeRecords.filter(record => !isRentalRecord(record) && !isExclusive(record) && !isLand(record)), name);
    const generalLand = sum(activeRecords.filter(record => !isRentalRecord(record) && !isExclusive(record) && isLand(record)), name);
    const rental = sum(activeRecords.filter(isRentalRecord), name);
    return { name, exclusiveHouse, exclusiveLand, generalHouse, generalLand, rental, total: exclusiveHouse + exclusiveLand + generalHouse + generalLand + rental };
  }).sort((a, b) => b.total - a.total || a.name.localeCompare(b.name, "zh-TW"));
  const rocMonth = (() => { const [year, value] = month.split("-"); return `${Number(year) - 1911}年${Number(value)}月`; })();
  const monthlyGrandTotal = monthlyRows.reduce((total, row) => total + row.total.total, 0);
  const stockGrandTotal = stockRows.reduce((total, row) => total + row.total, 0);
  const categoryText = (values: { exclusiveHouse: number; exclusiveLand: number; generalHouse: number; generalLand: number; rental: number; total: number }, totalOnly = false) => { const parts = ([['房專', values.exclusiveHouse], ['土專', values.exclusiveLand], ['房一', values.generalHouse], ['土一', values.generalLand], ['租', values.rental]] as const).filter(([, count]) => count > 0).map(([label, count]) => `${label}${numberText(count)}`); return totalOnly ? `共${numberText(values.total)}件` : parts.length ? `${parts.join(" +")}=${numberText(values.total)}` : "-"; };
  const categorySummary = (values: { exclusiveHouse: number; exclusiveLand: number; generalHouse: number; generalLand: number; rental: number; total: number }, totalOnly = false) => <div className="inventory-category-summary">{totalOnly ? <b>{categoryText(values, true)}</b> : <span>{categoryText(values)}</span>}</div>;
  const shortPropertyNo = (record: RecordItem) => String(record.propertyNo || "").match(/^[A-Za-z]{2}/)?.[0]?.toUpperCase() || String(record.propertyNo || "").slice(0, 2);
  const inventoryCaseLine = (record: RecordItem) => <><b>{shortPropertyNo(record)}　{record.caseName || "未命名案件"}</b><span>{record.address || "未填地址"}</span></>;
  const monthlyArchiveIds = new Set(monthlyArchive.map(record => record.id));
  const monthlyItemsFor = (name: string) => monthlyEntry.filter(record => weightFor(record, name) > 0).map(record => ({ record, action: monthlyArchiveIds.has(record.id) ? "本月進案／下架" : "本月進案", weight: weightFor(record, name) }));
  const monthlyDetailItems = monthlyDetailName ? monthlyItemsFor(monthlyDetailName) : [];
  const stockDetailRecords = stockDetailName ? activeRecords.filter(record => weightFor(record, stockDetailName) > 0).sort((a, b) => { const rank = (record: RecordItem) => /^[A-Z]A/i.test(String(record.propertyNo || "")) ? 0 : /^[A-Z]G/i.test(String(record.propertyNo || "")) ? 1 : 2; return rank(a) - rank(b) || String(a.propertyNo || "").localeCompare(String(b.propertyNo || ""), "zh-TW"); }) : [];
  const exportInventoryCanvas = (kind: "monthly" | "stock") => {
    const monthly = kind === "monthly"; const title = monthly ? `${rocMonth}進案統計　總件數 :${numberText(monthlyGrandTotal)}件` : `目前人員總件數表　總件數 :${numberText(stockGrandTotal)}件`;
    const headers = monthly ? ["人員", "本月進案", "本月下架", "本月件數"] : ["人員", "房專", "土專", "房一", "土一", "租件", "總件數"];
    const rows = monthly ? monthlyRows.map(row => [row.name, categoryText(row.entered), categoryText(row.archived), categoryText(row.total, true)]) : stockRows.map(row => [row.name, numberText(row.exclusiveHouse), numberText(row.exclusiveLand), numberText(row.generalHouse), numberText(row.generalLand), numberText(row.rental), `${numberText(row.total)}件`]);
    const widths = monthly ? [250, 760, 380, 280] : [250, 220, 220, 220, 220, 200, 240]; const margin = 44, titleHeight = 124, headerHeight = 88, rowHeight = 86, tableWidth = widths.reduce((sum, value) => sum + value, 0);
    const canvas = document.createElement("canvas"); canvas.width = (tableWidth + margin * 2) * 2; canvas.height = (titleHeight + headerHeight + Math.max(rows.length, 1) * rowHeight + margin * 2) * 2; const context = canvas.getContext("2d"); if (!context) return alert("圖片產生失敗"); context.scale(2, 2); context.fillStyle = "#ffffff"; context.fillRect(0, 0, canvas.width, canvas.height); context.textAlign = "center"; context.textBaseline = "middle"; context.fillStyle = "#183f34"; context.font = "bold 52px Microsoft JhengHei"; context.fillText(title, margin + tableWidth / 2, margin + titleHeight / 2);
    let x = margin, y = margin + titleHeight; context.font = "bold 34px Microsoft JhengHei"; headers.forEach((header, index) => { context.fillStyle = "#e7ece9"; context.fillRect(x, y, widths[index], headerHeight); context.strokeStyle = "#9fb0a8"; context.lineWidth = 1.8; context.strokeRect(x, y, widths[index], headerHeight); context.fillStyle = "#354b43"; context.fillText(header, x + widths[index] / 2, y + headerHeight / 2); x += widths[index]; });
    const shownRows = rows.length ? rows : [["目前沒有資料", ...headers.slice(1).map(() => "-")]]; shownRows.forEach((row, rowIndex) => { x = margin; y = margin + titleHeight + headerHeight + rowIndex * rowHeight; row.forEach((value, index) => { context.fillStyle = index === row.length - 1 ? "#fff1ed" : "#ffffff"; context.fillRect(x, y, widths[index], rowHeight); context.strokeStyle = "#bdc9c3"; context.lineWidth = 1.8; context.strokeRect(x, y, widths[index], rowHeight); context.fillStyle = index === row.length - 1 ? "#a83f3b" : "#17242a"; context.font = `${index === 0 || index === row.length - 1 ? "bold " : ""}${monthly ? 30 : 32}px Microsoft JhengHei`; context.fillText(String(value), x + widths[index] / 2, y + rowHeight / 2); x += widths[index]; }); });
    canvas.toBlob(blob => { if (!blob) return alert("圖片產生失敗"); const link = document.createElement("a"); const url = URL.createObjectURL(blob); link.href = url; link.download = `${monthly ? `${rocMonth}進案統計` : "目前人員總件數表"}.png`; document.body.appendChild(link); link.click(); setTimeout(() => { URL.revokeObjectURL(url); link.remove(); }, 3000); }, "image/png");
  };
  const exportPanelImage = (panelId: string, fileName: string) => {
    const source = document.getElementById(panelId); if (!source) return;
    const clone = source.cloneNode(true) as HTMLElement; clone.querySelectorAll("button").forEach(button => button.remove());
    const width = Math.max(source.scrollWidth, source.clientWidth), height = Math.max(source.scrollHeight, source.clientHeight);
    const css = Array.from(document.styleSheets).map(sheet => { try { return Array.from(sheet.cssRules).map(rule => rule.cssText).join("\n"); } catch { return ""; } }).join("\n");
    const body = document.createElement("div"); body.setAttribute("xmlns", "http://www.w3.org/1999/xhtml"); body.style.width = `${width}px`; body.style.background = "#ffffff"; const style = document.createElement("style"); style.textContent = css; body.append(style, clone);
    const markup = new XMLSerializer().serializeToString(body); const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}"><foreignObject width="100%" height="100%">${markup}</foreignObject></svg>`;
    const saveImage = (blob: Blob) => { const link = document.createElement("a"); const imageUrl = URL.createObjectURL(blob); link.href = imageUrl; link.download = `${fileName}.png`; link.style.display = "none"; document.body.appendChild(link); link.click(); setTimeout(() => { URL.revokeObjectURL(imageUrl); link.remove(); }, 3000); };
    const url = URL.createObjectURL(new Blob([svg], { type: "image/svg+xml;charset=utf-8" })); const image = new Image(); image.onload = () => { const canvas = document.createElement("canvas"); canvas.width = width * 2; canvas.height = height * 2; const context = canvas.getContext("2d"); if (!context) return alert("圖片產生失敗"); context.scale(2, 2); context.drawImage(image, 0, 0); URL.revokeObjectURL(url); canvas.toBlob(blob => { if (blob) saveImage(blob); else alert("圖片產生失敗"); }, "image/png"); }; image.onerror = () => { URL.revokeObjectURL(url); alert("圖片產生失敗，請再試一次"); }; image.src = url;
  };
  return <section className="content business-inventory-page">
    <div className="list-head"><SectionTitle title="物件庫存件數表" subtitle="共同開發案件依人數平均計算；未對應人員的案件列入未歸屬業務"/><div className="inventory-month-controls"><b>統計月份</b><button className={month === previousMonth ? "selected" : ""} onClick={() => setMonth(previousMonth)}>上月 {shortMonth(previousMonth)}</button><button className={month === currentMonth ? "selected" : ""} onClick={() => setMonth(currentMonth)}>本月 {shortMonth(currentMonth)}</button><label className="inventory-month"><span>選月份</span><input type="month" value={month} onChange={event => setMonth(event.target.value)}/></label></div></div>
    <div className="inventory-columns"><div className="inventory-panel monthly-inventory-panel" id="monthly-inventory-image">
      <div className="inventory-panel-title"><h3>{rocMonth}進案統計　總件數 :{numberText(monthlyGrandTotal)}件</h3><div className="inventory-title-actions"><button onClick={() => exportInventoryCanvas("monthly")}>產圖</button><button onClick={() => setMonthlyExpanded(value => !value)}>{monthlyExpanded ? "收合" : "展開"}</button></div></div>
      <table><colgroup><col style={{width:"16%"}}/><col style={{width:"46%"}}/><col style={{width:"20%"}}/><col style={{width:"18%"}}/></colgroup><thead><tr><th>人員</th><th>本月進案</th><th>本月下架</th><th>本月件數</th></tr></thead>
      <tbody>{monthlyRows.map(row => <Fragment key={row.name}><tr><td>{row.name}</td><td>{categorySummary(row.entered)}</td><td>{categorySummary(row.archived)}</td><td className="inventory-total"><button className="inventory-count-button" onClick={() => setMonthlyDetailName(row.name)}>{categorySummary(row.total, true)}</button></td></tr>{monthlyExpanded && <tr className="inventory-month-detail-row"><td colSpan={4}><div className="inventory-case-list inventory-inline-detail">{monthlyItemsFor(row.name).map(({ record, action, weight }, index) => <div key={`${record.id}-${action}-${index}`}><em>{action}</em>{inventoryCaseLine(record)}<strong>{numberText(weight)}件</strong></div>)}</div></td></tr>}</Fragment>)}{!monthlyRows.length && <tr><td colSpan={4} className="inventory-empty">本月目前沒有進案或下架紀錄</td></tr>}</tbody></table>
      <small>本月件數以案件去重計算；本月進案後又在同月下架，仍只計原本的 1 件或共同開發的分攤件數。重新上架不重複計入。</small>
    </div>
    <div className="inventory-panel stock-inventory-panel" id="stock-inventory-image">
      <div className="inventory-panel-title"><h3>目前人員總件數表　總件數 :{numberText(stockGrandTotal)}件</h3><button onClick={() => exportInventoryCanvas("stock")}>產圖</button></div>
      <table><thead><tr><th rowSpan={2}>人員</th><th colSpan={2}>專約</th><th colSpan={2}>一般約</th><th rowSpan={2}>租件</th><th rowSpan={2}>總件數</th></tr><tr><th>房屋</th><th>土地</th><th>房屋</th><th>土地</th></tr></thead>
      <tbody>{stockRows.map(row => <tr key={row.name}><td>{row.name}</td><td>{numberText(row.exclusiveHouse)}</td><td>{numberText(row.exclusiveLand)}</td><td>{numberText(row.generalHouse)}</td><td>{numberText(row.generalLand)}</td><td>{numberText(row.rental)}</td><td className="inventory-total"><button className="inventory-count-button" onClick={() => setStockDetailName(row.name)}>{numberText(row.total)}</button></td></tr>)}</tbody></table>
    </div></div>
    {monthlyDetailName && <div className="modal-backdrop"><div className="modal inventory-detail-modal"><div className="modal-head"><div><span>本月件數明細</span><h2>{monthlyDetailName}　總計 {numberText(monthlyRows.find(row => row.name === monthlyDetailName)?.total.total || 0)}件</h2></div><button className="close" onClick={() => setMonthlyDetailName("")}>×</button></div><div className="inventory-case-list monthly-detail-list">{monthlyDetailItems.map(({ record, action, weight }, index) => <div key={`${record.id}-${action}-${index}`}><em>{action}</em>{inventoryCaseLine(record)}<strong>{numberText(weight)}件</strong></div>)}</div><div className="modal-foot"><b>總計：{numberText(monthlyRows.find(row => row.name === monthlyDetailName)?.total.total || 0)}件</b><button onClick={() => setMonthlyDetailName("")}>完成</button></div></div></div>}
    {stockDetailName && <div className="modal-backdrop"><div className="modal inventory-detail-modal"><div className="modal-head"><div><span>目前總件數明細</span><h2>{stockDetailName}　{numberText(stockRows.find(row => row.name === stockDetailName)?.total || 0)}件</h2></div><button className="close" onClick={() => setStockDetailName("")}>×</button></div><div className="inventory-case-list monthly-detail-list stock-detail-list">{stockDetailRecords.map(record => <div key={record.id}>{inventoryCaseLine(record)}</div>)}</div><div className="modal-foot"><button onClick={() => setStockDetailName("")}>完成</button></div></div></div>}
  </section>;
}

function Field({ fieldKey, label, record, records, setRecord }: { fieldKey: string; label: string; record: RecordItem; records: RecordItem[]; setRecord: (r: RecordItem) => void }) {
  const value = record[fieldKey] || ""; const inputValue = fieldKey === "layout" ? (typeShort(record.type) === "土地" ? "" : layoutFull(value, record.type)) : value; const set = (v: string) => setRecord(fieldKey === "propertyNo" ? { ...record, propertyNo: v, contractType: contractFromNo(v) } : { ...record, [fieldKey]: v });
  const websiteInput = (key: string, siteLabel: string, expiryKey = "") => { const noneKey = `${key}None`; const none = record[noneKey] === "1"; const changeValue = (next: string) => { if (key === "windowAd" && next) { if (!/^\d+$/.test(next) || Number(next) < 1 || Number(next) > 15) return alert("櫥窗編號只能輸入 1～15"); const duplicate = records.find(item => item.id !== record.id && String(item.windowAd || "").trim() === next.trim() && item.windowAdNone !== "1"); if (duplicate) return alert(`櫥窗編號 ${next} 已由「${duplicate.caseName || duplicate.propertyNo}」使用`); } const detectedExpiry = expiryKey && !record[expiryKey] ? websiteEffectiveDate(next) : ""; setRecord({ ...record, [key]: next, ...(detectedExpiry ? { [expiryKey]: detectedExpiry } : {}) }); }; return <label className={`field website-field ${none ? "website-none" : ""}`}><span className="website-field-head"><b>{siteLabel}</b><i><input type="checkbox" checked={none} onChange={event => { if (event.target.checked && String(record[key] || "").trim() && !confirm(`${siteLabel}已有內容「${record[key]}」，確定要改成無並清除內容嗎？`)) return; setRecord({ ...record, [noneKey]: event.target.checked ? "1" : "", ...(event.target.checked ? { [key]: "", ...(expiryKey ? { [expiryKey]: "" } : {}) } : {}) }); }}/>無</i></span><input type="text" disabled={none} value={record[key] || ""} onChange={event => changeValue(event.target.value)} placeholder={none ? "不需刊登" : "輸入網站編號或註記"}/>{expiryKey && <><span className="website-expiry-label">{siteLabel}到期日期</span><input type="text" inputMode="numeric" disabled={none} value={displayRocDate(record[expiryKey] || websiteEffectiveDate(record[key] || ""))} onChange={event => setRecord({ ...record, [expiryKey]: event.target.value })} onBlur={event => setRecord({ ...record, [expiryKey]: normalizeDateInput(event.target.value) })} placeholder="例如 115/10/19"/></>}</label>; };
  if (["colorSheetHeader", "coverHeader", "websiteHeader"].includes(fieldKey)) return <div className="record-edit-section-title"><b>{label}</b><span>{fieldKey === "colorSheetHeader" ? "供彩色表 Excel 使用" : fieldKey === "coverHeader" ? "供列印新進資料封面使用" : "輸入各網站刊登編號"}</span></div>;
  if (fieldKey === "areaPaste") return <label className="field area-paste-field"><span>{label}</span><textarea rows={8} value={value} onChange={event => setRecord(parseAreaPaste(event.target.value, record))} placeholder="貼上建物面積、土地面積、建築完成日與主要建材整串文字，系統會自動抓取下方數字。"/><small>已抓取的內容仍可在下方個別修改。</small></label>;
  const pillChoices: Record<string, string[]> = {
    coverBottomSource: ["主約", "契變", "口頭"], coverPercentSource: ["主約", "契變", "口頭"], coverNoChange: ["無契變"], coverOriginalType: ["原稿", "草稿"], titleUndertaking: ["有切結", "無切結"],
    zoningDocumentStatus: ["房屋不需要", "土地已附正式分區", "謄本已標示不用附"], authorizationStatus: ["缺授", "無需要", "已附上歸檔"], authorizationCopyType: ["影本", "正本"],
    parkingOwnership: ["無車位", "停自有地", "固定車位", "車位另租", "抽籤決定", "先到先停"], parkingMethod: ["坡道/平面", "坡道/機械", "昇降/平面", "昇降/機械", "庭院", "車庫", "平移/機械"]
  };
  if (pillChoices[fieldKey]) { const displayedValue = fieldKey === "parkingMethod" && !value && pillChoices.parkingMethod.includes(record.parkingType || "") ? record.parkingType : value; const selectPill = (option: string) => fieldKey === "parkingMethod" ? setRecord({ ...record, parkingMethod: displayedValue === option ? "" : option, parkingType: "" }) : set(displayedValue === option ? "" : option); return <label className={`field cover-pill-field cover-pill-${fieldKey}`}><span>{label}</span><span className="pill-options">{pillChoices[fieldKey].map(option => <button type="button" className={displayedValue === option ? "active" : ""} onClick={() => selectPill(option)} key={option}>{option}</button>)}</span></label>; }
  if (fieldKey === "reducedPrice") return null;
  if (fieldKey === "caseNameNote") return <label className="field case-name-note-field"><span>{label}{record.caseNameNoteModifiedAt && <small>修改:{displayModifiedAt(record.caseNameNoteModifiedAt)}</small>}</span><input type="text" value={value} onChange={event => setRecord({ ...record, caseNameNote: event.target.value, caseNameNoteModifiedAt: new Date().toISOString() })}/></label>;
  if (fieldKey === "showingFollowUpDueDate") return <label className="field showing-follow-up-date-field"><span>{label}</span><input type="text" inputMode="numeric" value={displayRocDate(value)} onChange={event => setRecord({ ...record, showingFollowUpDueDate: event.target.value, showingFollowUp: event.target.value ? "暫停帶看／等待業務回覆" : "", showingFollowUpDate: event.target.value ? record.showingFollowUpDate || today() : "" })} onBlur={event => { const date = normalizeDateInput(event.target.value); setRecord({ ...record, showingFollowUpDueDate: date, showingFollowUp: date ? "暫停帶看／等待業務回覆" : "", showingFollowUpDate: date ? record.showingFollowUpDate || today() : "" }); }} placeholder="例如 115/8/15"/><small>清空日期即解除追蹤</small></label>;
  if (fieldKey === "completionDate") { const parsed = record.areaPaste ? parseAreaPaste(record.areaPaste, record) : record; const shownDate = value || parsed.completionDate || ""; const yearText = String(shownDate || record.builtYear || "").match(/\d{2,4}/)?.[0] || ""; const yearNumber = Number(yearText); const westernYear = yearNumber ? (yearNumber > 1911 ? yearNumber : yearNumber + 1911) : 0; const age = westernYear ? new Date().getFullYear() - westernYear : NaN; return <label className="field completion-age-field"><span>{label}{Number.isFinite(age) && age >= 0 && <small>約 {age} 年屋</small>}</span><input type="text" inputMode="numeric" value={displayRocDate(shownDate)} onChange={event => set(event.target.value)} onBlur={event => set(normalizeDateInput(event.target.value))} placeholder="例如 074.04.16"/></label>; }
  if (["buildingPing", "indoorPing", "landPing", "registryBuildingPing", "registryIndoorPing", "landSharePing"].includes(fieldKey)) { const parsed = record.areaPaste ? parseAreaPaste(record.areaPaste, record) : record; const compared = fieldKey === "buildingPing" ? parsed.registryBuildingPing : fieldKey === "indoorPing" ? parsed.registryIndoorPing : fieldKey === "landPing" ? parsed.landSharePing : fieldKey === "registryBuildingPing" ? record.buildingPing : fieldKey === "registryIndoorPing" ? record.indoorPing : record.landPing; const prefix = ["registryBuildingPing", "registryIndoorPing", "landSharePing"].includes(fieldKey) ? "進案" : "房管"; return <label className="field compared-ping-field"><span><b>{label}</b>{compared && <small>{prefix} {compared} 坪</small>}</span><input inputMode="decimal" value={value} onChange={event => set(event.target.value)}/></label>; }
  if (fieldKey === "price") return <div className="price-pair"><label className="field"><span>開價（萬）</span><input inputMode="decimal" value={record.price || ""} onChange={e => setRecord({ ...record, price: e.target.value, priceModifiedAt: new Date().toISOString() })}/></label><label className="field reduced-price-field"><span>降價/調價(萬){record.reducedPrice && <small>修改：{displayModifiedAt(record.reducedPriceModifiedAt || record.lastModifiedAt)}</small>}</span><input inputMode="decimal" value={record.reducedPrice || ""} onChange={e => setRecord({ ...record, reducedPrice: e.target.value, reducedPriceModifiedAt: e.target.value ? new Date().toISOString() : "" })}/></label></div>;
  if (fieldKey === "bookLocationNo") return null;
  if (fieldKey === "contractType") return <label className="field contract-field"><span>{label}<small>EG 房一、EA 房專、LG 土一、LA 土專、EB 租一、EC 租專、RG 預一、RA 預專</small></span><input value={contractFromNo(record.propertyNo) || value} readOnly placeholder="依物件編號自動判斷"/></label>;
  if (fieldKey === "coverage") { const combinedValue = record.coverageCombined || [record.coverage, record.far].filter(Boolean).join("/"); return <label className="field coverage-combined-field"><span>建蔽率%/容積率%</span><input type="text" inputMode="decimal" value={combinedValue} onChange={event => { const raw = event.target.value.replace(/／/g, "/"); const pairs = raw.split(/[、,，\n]+/).map(value => value.trim()).filter(Boolean).map(value => value.split("/")); const coverage = pairs.map(pair => pair[0] || "").filter(Boolean).join("/"); const far = pairs.map(pair => pair[1] || "").filter(Boolean).join("/"); setRecord({ ...record, coverageCombined: raw, coverage, far }); }} placeholder="例如 60/240；兩組可輸入 60/360、80/320"/></label>; }
  if (["platform591", "price5168", "goldExposure"].includes(fieldKey)) { const expiryKey = `${fieldKey}Expiry`; const siteLabel = fieldKey === "platform591" ? "591" : fieldKey === "price5168" ? "5168" : "黃金曝光"; return websiteInput(fieldKey, siteLabel, expiryKey); }
  if (fieldKey === "yes319") return <div className="website-stack">{websiteInput("yes319", "YES319")}{websiteInput("houseinfor", "HOUSE INFOR")}</div>;
  if (fieldKey === "windowAd") return <div className="website-stack">{websiteInput("windowAd", "櫥窗（專）")}{websiteInput("led", "LED（專）")}</div>;
  if (fieldKey === "homeWeb") return websiteInput("homeWeb", "我家網");
  if (["houseinfor", "led"].includes(fieldKey)) return null;
  if (fieldKey === "bookLocationType") return <div className="field combined-book-field"><span>物件本</span><div className="book-combined-row"><input type="text" inputMode="numeric" value={displayRocDate(record.bookLocationDate || "")} onChange={e => setRecord({ ...record, bookLocationDate: e.target.value })} onBlur={e => setRecord({ ...record, bookLocationDate: normalizeDateInput(e.target.value) })} placeholder="日期，例如 115/7/29"/><div className="pill-options">{["架上", "旁5"].map(o => <button type="button" key={o} className={value === o ? "selected" : ""} onClick={() => { const websiteKeys = ["platform591", "yes319", "houseinfor", "windowAd", "led", "homeWeb", "price5168", "goldExposure"]; const expiryKeys = ["platform591Expiry", "price5168Expiry", "goldExposureExpiry"]; if (o === "旁5") setRecord({ ...record, bookLocationType: o, ...Object.fromEntries(websiteKeys.map(key => [key, "旁5"])), ...Object.fromEntries(expiryKeys.map(key => [key, ""])) }); else setRecord({ ...record, bookLocationType: o, ...Object.fromEntries(websiteKeys.filter(key => record[key] === "旁5").map(key => [key, ""])) }); }}>{o}</button>)}</div></div></div>;
  if (fieldKey === "salesBook") { const selected = String(record.salesBook || "").includes("更新") ? "更新" : "製作"; return <div className="field combined-sales-field"><span>銷售本</span><div className="sales-combined-row"><input type="text" inputMode="numeric" value={displayRocDate(record.salesBookDate || "")} onChange={e => setRecord({ ...record, salesBookDate: e.target.value })} onBlur={e => setRecord({ ...record, salesBookDate: normalizeDateInput(e.target.value) })} placeholder="日期，例如 115/7/29"/><div className="pill-options sales-book-options">{["製作", "更新"].map(option => <button type="button" key={option} className={selected === option ? "selected" : ""} onClick={() => setRecord({ ...record, salesBook: option })}>{option}</button>)}</div></div></div>; }
  if (["notes", "additionNotes", "attentionNotes"].includes(fieldKey)) return <label className="field span-2"><span>{label}</span><textarea value={value} onChange={e => set(e.target.value)} rows={3}/></label>;
  if (fieldKey === "type") { const selectedType = selectOptions.type.find(option => typeShort(option) === typeShort(value)) || value; return <label className="field"><span>{label}</span><select value={selectedType} onChange={e => set(e.target.value)}><option value="">請選擇</option>{selectOptions.type.map(o => <option key={o} value={o}>{o}</option>)}</select></label>; }
  if (selectOptions[fieldKey]) return <label className="field"><span>{label}</span><select value={value} onChange={e => set(e.target.value)}><option value="">請選擇</option>{selectOptions[fieldKey].map(o => <option key={o}>{o}</option>)}</select></label>;
  if (fieldKey === "groupViewDate") return <label className="field"><span>{label}</span><input type="text" placeholder="例如：115年7月28日、已團看" value={value} onChange={e => set(e.target.value)}/><small>可輸入日期或中文註記</small></label>;
  if (fieldKey === "updateDate") return <label className="field update-date-field"><span>{label}</span><input type="text" inputMode="numeric" value={displayRocDate(value)} onChange={e => set(normalizeCompleteDateInput(e.target.value))} onBlur={e => set(normalizeDateInput(e.target.value))} placeholder="例如 115/7/29"/><small>{record.lastModifiedAt ? `最後修改:${displayModifiedAt(record.lastModifiedAt)}` : "尚無修改時間"}</small></label>;
  if (dateKeys.has(fieldKey)) return <label className={`field ${value && !validDate(normalizeDateInput(value)) ? "invalid" : ""}`}><span>{label}</span><input type="text" inputMode="numeric" placeholder="例如 7/24" value={displayRocDate(normalizeDateInput(value))} onChange={e => set(normalizeCompleteDateInput(e.target.value))} onBlur={e => set(normalizeDateInput(e.target.value))}/><small>可輸入 7/24，自動轉為民國日期</small></label>;
  return <label className="field"><span>{label}{fieldKey === "propertyNo" || fieldKey === "caseName" ? " *" : ""}</span><input type="text" inputMode={["price", "reducedPrice", "builtYear", "indoorPing", "buildingPing", "landPing", "coverage", "far"].includes(fieldKey) ? "decimal" : undefined} value={inputValue} onChange={e => set(e.target.value)}/>{fieldKey === "builtYear" && value && <small>目前換算：{ageOf(record)}</small>}</label>;
}

function SettingsPanel({ settings, setSettings, supabasePush, supabasePull, cloudSession, supabaseSignIn, supabaseSignOut }: { settings: Settings; setSettings: (s: Settings) => void; supabasePush: () => void; supabasePull: () => void; cloudSession: CloudSession | null; supabaseSignIn: (email: string, password: string, signUp?: boolean) => void; supabaseSignOut: () => void }) {
  const set = (k: keyof Settings, v: string) => setSettings({ ...settings, [k]: v });
  useEffect(() => {
    const personnelPanel = document.querySelector<HTMLElement>(".settings .personnel-panel");
    if (!personnelPanel) return;
    const panel = document.createElement("article");
    panel.className = "panel expiry-settings-panel";
    panel.innerHTML = `<h3>公司到期提醒</h3><p>591、5168 會在到期前 1 天提醒；經紀人會在到期前 1 個月提醒。</p><div class="form-grid"><label class="field"><span>591 到期日</span><input data-expiry-key="expiry591" placeholder="例如 115/8/31"></label><label class="field"><span>5168 到期日</span><input data-expiry-key="expiry5168" placeholder="例如 115/8/31"></label><label class="field"><span>經紀人到期日</span><input data-expiry-key="brokerExpiry" placeholder="例如 115/8/31"></label></div>`;
    panel.classList.remove("expiry-open");
    const toggleExpiryPanel = (event: MouseEvent) => { if ((event.target as HTMLElement).closest("input,select,textarea,button,label")) return; panel.classList.toggle("expiry-open"); };
    panel.addEventListener("click", toggleExpiryPanel);
    personnelPanel.parentElement?.insertBefore(panel, personnelPanel);
    const values: Record<string, string> = { expiry591: displayRocDate(settings.expiry591 || ""), expiry5168: displayRocDate(settings.expiry5168 || ""), brokerExpiry: displayRocDate(settings.brokerExpiry || "") };
    const handlers = Array.from(panel.querySelectorAll<HTMLInputElement>("[data-expiry-key]")).map(input => {
      input.value = values[input.dataset.expiryKey || ""] || "";
      const handler = () => setSettings({ ...settings, [input.dataset.expiryKey || ""]: normalizeDateInput(input.value) });
      input.addEventListener("blur", handler);
      return { input, handler };
    });
    return () => { panel.removeEventListener("click", toggleExpiryPanel); handlers.forEach(({ input, handler }) => input.removeEventListener("blur", handler)); panel.remove(); };
  }, [settings.expiry591, settings.expiry5168, settings.brokerExpiry]);
  useEffect(() => {
    const personnelPanel = document.querySelector<HTMLElement>(".settings .personnel-panel");
    const head = personnelPanel?.querySelector<HTMLElement>(".personnel-head");
    if (!personnelPanel || !head) return;
    personnelPanel.classList.remove("personnel-open");
    const togglePanel = (event: MouseEvent) => { if ((event.target as HTMLElement).closest("input,select,textarea,button,label")) return; personnelPanel.classList.toggle("personnel-open"); };
    personnelPanel.addEventListener("click", togglePanel);
    return () => { personnelPanel.removeEventListener("click", togglePanel); };
  }, []);
  const [cloudEmail, setCloudEmail] = useState("");
  const [cloudPassword, setCloudPassword] = useState("");
  const updatePerson = (id: string, patch: Partial<Person>) => setSettings({ ...settings, personnel: settings.personnel.map(p => p.id === id ? { ...p, ...patch } : p) });
  const personTextInput = (person: Person, field: "name" | "nationalId" | "phone", placeholder: string, type = "text") => {
    const normalize = (value: string) => field === "nationalId" ? value.toUpperCase() : value;
    const save = (value: string) => updatePerson(person.id, { [field]: normalize(value) } as Partial<Person>);
    return <input type={type} defaultValue={person[field] || ""} placeholder={placeholder}
      onChange={event => { if (!event.nativeEvent.isComposing) save(event.currentTarget.value); }}
      onCompositionEnd={event => save(event.currentTarget.value)}
      onBlur={event => save(event.currentTarget.value)} />;
  };
  const removePerson = (id: string) => {
    if (!confirm("確定刪除這位人員？刪除後序號會自動重新編排。")) return;
    const currentScrollTop = window.scrollY;
    const personnel = sortPeopleBySequence(settings.personnel.filter(person => person.id !== id)).map((person, index) => ({ ...person, sequence: String(index + 1) }));
    setSettings({ ...settings, personnel });
    requestAnimationFrame(() => requestAnimationFrame(() => window.scrollTo({ top: Math.min(currentScrollTop, Math.max(0, document.documentElement.scrollHeight - window.innerHeight)), behavior: "auto" })));
  };
  const addPerson = () => setSettings({ ...settings, personnel: [...settings.personnel, { id: newId(), sequence: String(settings.personnel.length + 1), name: "", nationalId: "", phone: "", role: "業務", status: "在職" }] });
  const activePeople = sortPeopleBySequence(settings.personnel.filter(p => p.status === "在職"));
  const formerPeople = sortPeopleBySequence(settings.personnel.filter(p => p.status === "離職"));
  const personRow = (p: Person) => <div className="person-row" key={p.id}><input className="person-sequence" type="number" min="1" value={p.sequence || ""} onChange={e => updatePerson(p.id, { sequence: e.target.value })} placeholder="序"/>{personTextInput(p, "name", "姓名")}{personTextInput(p, "nationalId", "身分證字號")}{personTextInput(p, "phone", "手機號碼", "tel")}<select value={p.role || "業務"} onChange={e => updatePerson(p.id, { role: e.target.value as Person["role"] })}><option>業務</option><option>秘書</option></select><select value={p.status} onChange={e => updatePerson(p.id, { status: e.target.value as Person["status"] })}><option>在職</option><option>離職</option></select><button className="danger" onClick={() => removePerson(p.id)}>刪除</button></div>;
  return <section className="settings content"><SectionTitle title="系統設定" subtitle=""/><details className="supabase"><summary><span>進階：Supabase 雲端同步</span><small>{cloudSession ? `已登入：${cloudSession.email || "雲端帳號"}（自動同步）` : "預設隱藏"}</small></summary><div className="supabase-body"><div className="warning">登入後，開啟系統會自動取得最新雲端資料；修改後等待 6 秒自動上傳。照片檔案不會上傳到雲端。</div><div className="form-grid"><label className="field"><span>Project URL</span><input value={settings.supabaseUrl} onChange={e => set("supabaseUrl", e.target.value)}/></label><label className="field"><span>Supabase Publishable key</span><input type="password" value={settings.supabaseKey} onChange={e => set("supabaseKey", e.target.value)} placeholder="貼上 anon / publishable key"/></label></div>{cloudSession ? <div className="backup-actions"><span className="cloud-auto-sync-status">已啟用自動下載與自動上傳</span><button onClick={supabaseSignOut}>登出雲端帳號</button></div> : <><div className="form-grid"><label className="field"><span>雲端登入 Email</span><input type="email" value={cloudEmail} onChange={e => setCloudEmail(e.target.value)}/></label><label className="field"><span>雲端登入密碼</span><input type="password" value={cloudPassword} onChange={e => setCloudPassword(e.target.value)}/></label></div><div className="backup-actions"><button onClick={() => void supabaseSignIn(cloudEmail, cloudPassword, false)}>登入</button><button className="primary" onClick={() => void supabaseSignIn(cloudEmail, cloudPassword, true)}>第一次使用：註冊雲端帳號</button></div></>}</div></details><details className="panel book-review-settings-panel"><summary>物件本確認日期</summary><div className="form-grid"><label className="field"><span>本次確認日期</span><input type="text" inputMode="numeric" value={displayRocDate(settings.bookReviewCurrentDate || "2026-07-30")} onChange={e => set("bookReviewCurrentDate", e.target.value)} onBlur={e => set("bookReviewCurrentDate", normalizeDateInput(e.target.value) || "2026-07-30")} placeholder="115/7/30"/></label><label className="field"><span>下次確認日期</span><input type="text" inputMode="numeric" value={displayRocDate(settings.bookReviewNextDate || "2026-09-30")} onChange={e => set("bookReviewNextDate", e.target.value)} onBlur={e => set("bookReviewNextDate", normalizeDateInput(e.target.value) || "2026-09-30")} placeholder="115/9/30"/></label></div></details><article className="panel personnel-panel personnel-open"><div className="personnel-head"><h3>人員設定</h3><button className="primary" onClick={addPerson}>＋ 新增人員</button></div><div className="personnel-table"><div className="person-row person-labels"><span>序</span><span>人員</span><span>身分證字號（前台密碼）</span><span>手機號碼</span><span>職務</span><span>狀態</span><span>操作</span></div>{activePeople.map(personRow)}{!activePeople.length && <div className="no-person">尚未新增在職人員</div>}</div>{formerPeople.length > 0 && <details className="former-people"><summary>離職人員（{formerPeople.length}）</summary><div className="personnel-table">{formerPeople.map(personRow)}</div></details>}</article></section>;
}

function SectionTitle({ title, subtitle }: { title: string; subtitle: string }) { return <div className="section-title"><h2>{title}</h2>{subtitle && <p>{subtitle}</p>}</div>; }
function Empty({ text }: { text: string }) { return <div className="empty"><span>⌂</span><h3>{text}</h3><p>新案件請先由進案草稿確認後正式進案。</p></div>; }
