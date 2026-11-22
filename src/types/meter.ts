export type MeterStatus =
  | 'รอชำระเงิน'
  | 'อยู่ระหว่างติดตั้ง'
  | 'ติดตั้งเสร็จสิ้น'
  | 'ไม่มีเส้นท่อ/ยกเลิก'
  | 'รับคำร้องแล้ว(รอสำรวจ)'
  | string;

export interface MeterRequest {
  /** 1-based row index in Google Sheets (row 1 is header, data starts at row 2) */
  rowIndex: number;
  /** Stable key for diffing across real-time syncs */
  uid: string;
  /** เลขที่ RQ หรือ เลขที่งาน */
  rqNumber: string;
  /** วันที่ยื่นคำร้อง / วันที่ (ISO YYYY-MM-DD or raw sheet date) */
  requestDate: string;
  /** ชื่อ - สกุล */
  fullName: string;
  /** บ้านเลขที่ (ถ้ามี) */
  houseNo: string;
  /** หมู่ที่ */
  moo: string;
  /** ตำบล */
  subdistrict: string;
  /** อำเภอ */
  district: string;
  /** จังหวัด */
  province: string;
  /** พื้นที่รวม (ตำบล-อำเภอ-จังหวัด) */
  locationSummary: string;
  /** เบอร์โทร */
  phone: string;
  /** ราคา / ยอดชำระ (บาท) */
  price: number;
  /** สถานะ */
  status: MeterStatus;
  /** พิกัด Latitude */
  latitude: number | null;
  /** พิกัด Longitude */
  longitude: number | null;
  /** หมายเหตุ */
  note: string;
  /** ลิงก์รูปหน้าบ้านใน Google Drive */
  housePhotoUrl?: string;
  /** Raw values keyed by exact sheet column header */
  raw: Record<string, string>;
}

export interface MeterFormPayload {
  rqNumber: string;
  requestDate: string;
  fullName: string;
  houseNo: string;
  moo: string;
  subdistrict: string;
  district: string;
  province: string;
  phone: string;
  price: number;
  status: string;
  latitude: number | null;
  longitude: number | null;
  note: string;
  housePhotoUrl?: string;
}

export interface SheetTabInfo {
  sheetId: number;
  title: string;
  rowCount?: number;
}

export interface SheetConfig {
  spreadsheetId: string;
  spreadsheetTitle?: string;
  sheetName: string;
  sheetId?: number;
  availableTabs: SheetTabInfo[];
  pollIntervalMs: number;
  appsScriptUrl?: string;
  driveFolderId?: string;
}

export type ConnectionMode = 'oauth' | 'apps_script' | 'public_csv' | 'disconnected' | 'error';

export type ChangeType = 'added' | 'updated' | 'deleted';

export interface FieldChangeDiff {
  field: string;
  from: string;
  to: string;
}

export interface DataChangeEvent {
  id: string;
  timestamp: number;
  type: ChangeType;
  recordUid: string;
  rqNumber: string;
  fullName: string;
  status: string;
  latitude: number | null;
  longitude: number | null;
  summary: string;
  diffs: FieldChangeDiff[];
  read: boolean;
}

export type MapLayerType = 'roadmap' | 'satellite' | 'hybrid' | 'terrain';

export interface RouteSummary {
  destinationName: string;
  destinationLat: number;
  destinationLng: number;
  distanceMeters: number;
  durationMillis: number;
  isEstimate?: boolean;
}

export interface CompletedBatchRecord {
  uid: string;
  rowIndex: number;
  rqNumber: string;
  jobNumber: string;
  waterUserNo: string;
  meterNumber: string;
  meterSize?: string;
  requestDate: string;
  fullName: string;
  houseNo: string;
  moo: string;
  subdistrict: string;
  district: string;
  province: string;
  locationSummary: string;
  phone: string;
  price: number;
  status: string;
  latitude: number | null;
  longitude: number | null;
  note: string;
}

export interface BatchMemoSigners {
  meterSizeText?: string;
  proposerName?: string;
  proposerTitle?: string;
  managerName?: string;
  managerTitle?: string;
  adminHeadName?: string;
  adminHeadTitle?: string;
  inspectorTitle?: string;
}

export interface CompletedBatch {
  id: string;
  spreadsheetId: string;
  batchCode: string;
  title: string;
  completedDate: string;
  inspectorName: string;
  note: string;
  exportFormats: Array<'excel' | 'pdf'>;
  statusUpdatedInSheet: boolean;
  includeTax?: boolean;
  memoSigners?: BatchMemoSigners;
  recordCount: number;
  totalAmount: number;
  mappedCount: number;
  records: CompletedBatchRecord[];
  createdAt: number;
}

export function parseLatLngPair(
  combined?: string | null,
  latStr?: string | null,
  lngStr?: string | null
): { lat: number | null; lng: number | null } {
  if (latStr && lngStr) {
    const lat = Number.parseFloat(String(latStr).trim());
    const lng = Number.parseFloat(String(lngStr).trim());
    if (Number.isFinite(lat) && Number.isFinite(lng) && Math.abs(lat) <= 90 && Math.abs(lng) <= 180) {
      return { lat, lng };
    }
  }

  if (combined) {
    const cleaned = String(combined).trim();
    if (cleaned) {
      const match = cleaned.match(/(-?\d+(?:\.\d+)?)\s*[,\s;/]+\s*(-?\d+(?:\.\d+)?)/);
      if (match) {
        const lat = Number.parseFloat(match[1]);
        const lng = Number.parseFloat(match[2]);
        if (Number.isFinite(lat) && Number.isFinite(lng) && Math.abs(lat) <= 90 && Math.abs(lng) <= 180) {
          return { lat, lng };
        }
      }
    }
  }

  return { lat: null, lng: null };
}

export function formatLatLngString(lat: number | null, lng: number | null): string {
  if (lat === null || lng === null || !Number.isFinite(lat) || !Number.isFinite(lng)) {
    return '';
  }
  return `${lat.toFixed(6)}, ${lng.toFixed(6)}`;
}

export function getStatusMeta(status: string): {
  label: string;
  color: string;
  textColor: string;
  dotClass: string;
  bgSoftClass: string;
  borderClass: string;
} {
  const s = (status || '').trim();
  if (s === 'ติดตั้งเสร็จสิ้น' || s === 'เสร็จสิ้น' || s === 'ดำเนินการแล้ว') {
    return {
      label: s || 'ติดตั้งเสร็จสิ้น',
      color: '#10b981',
      textColor: 'text-emerald-700 dark:text-emerald-400',
      dotClass: 'bg-emerald-500',
      bgSoftClass: 'bg-emerald-50 dark:bg-emerald-950/50',
      borderClass: 'border-emerald-200 dark:border-emerald-800/50',
    };
  }
  if (s === 'อยู่ระหว่างติดตั้ง' || s === 'กำลังดำเนินการ' || s === 'รอดำเนินการ') {
    return {
      label: s || 'อยู่ระหว่างติดตั้ง',
      color: '#d97706',
      textColor: 'text-amber-700 dark:text-amber-400',
      dotClass: 'bg-amber-500',
      bgSoftClass: 'bg-amber-50 dark:bg-amber-950/50',
      borderClass: 'border-amber-200 dark:border-amber-800/50',
    };
  }
  if (s === 'รอชำระเงิน' || s === 'ค้างชำระ') {
    return {
      label: s || 'รอชำระเงิน',
      color: '#e11d48',
      textColor: 'text-rose-700 dark:text-rose-400',
      dotClass: 'bg-rose-500',
      bgSoftClass: 'bg-rose-50 dark:bg-rose-950/50',
      borderClass: 'border-rose-200 dark:border-rose-800/50',
    };
  }
  if (s === 'ไม่มีเส้นท่อ/ยกเลิก' || s === 'ยกเลิก') {
    return {
      label: s || 'ไม่มีเส้นท่อ/ยกเลิก',
      color: '#64748b',
      textColor: 'text-slate-600 dark:text-slate-300',
      dotClass: 'bg-slate-400',
      bgSoftClass: 'bg-slate-100 dark:bg-slate-800/60',
      borderClass: 'border-slate-200 dark:border-slate-700',
    };
  }
  if (s === 'รับคำร้องแล้ว(รอสำรวจ)' || s === 'รอตรวจสอบ' || s === '') {
    return {
      label: 'รับคำร้องแล้ว(รอสำรวจ)',
      color: '#2271b3',
      textColor: 'text-[#2271B3] dark:text-[#5EA6E0]',
      dotClass: 'bg-[#2271B3]',
      bgSoftClass: 'bg-[#EDF5FC] dark:bg-[#2271B3]/15',
      borderClass: 'border-[#D6E7F6] dark:border-[#2271B3]/35',
    };
  }
  return {
    label: s,
    color: '#2271b3',
    textColor: 'text-[#2271B3] dark:text-[#5EA6E0]',
    dotClass: 'bg-[#2271B3]',
    bgSoftClass: 'bg-[#EDF5FC] dark:bg-[#2271B3]/15',
    borderClass: 'border-[#D6E7F6] dark:border-[#2271B3]/35',
  };
}
