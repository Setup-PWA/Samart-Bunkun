import { isInsideGoogleAppsScript } from './sheetsService';

export const HOUSE_PHOTO_FOLDER_ID = '1uE-I15oyqiqZb_wsuGQ5TfJXONZY_uY3';
export const HOUSE_PHOTO_FOLDER_URL =
  'https://drive.google.com/drive/folders/1uE-I15oyqiqZb_wsuGQ5TfJXONZY_uY3?usp=sharing';

export function extractDriveFolderId(input?: string | null): string {
  const trimmed = (input || '').trim();
  if (!trimmed) return HOUSE_PHOTO_FOLDER_ID;
  const folderMatch = trimmed.match(/\/folders\/([a-zA-Z0-9-_]+)/);
  if (folderMatch && folderMatch[1]) return folderMatch[1];
  const idParamMatch = trimmed.match(/[?&]id=([a-zA-Z0-9-_]+)/);
  if (idParamMatch && idParamMatch[1]) return idParamMatch[1];
  return trimmed;
}

export function buildDriveFolderUrl(folderId?: string | null): string {
  const cleanId = extractDriveFolderId(folderId);
  return `https://drive.google.com/drive/folders/${encodeURIComponent(cleanId)}?usp=sharing`;
}

export function extractDriveFileId(url?: string | null): string | null {
  const trimmed = (url || '').trim();
  if (!trimmed) return null;
  const fileMatch = trimmed.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (fileMatch && fileMatch[1]) return fileMatch[1];
  const idMatch = trimmed.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (idMatch && idMatch[1]) return idMatch[1];
  return null;
}

export function getDriveImagePreviewUrls(url?: string | null): string[] {
  const trimmed = (url || '').trim();
  if (!trimmed || !/^https?:\/\//i.test(trimmed)) return [];
  const fileId = extractDriveFileId(trimmed);
  if (fileId) {
    return [
      `https://drive.google.com/thumbnail?id=${encodeURIComponent(fileId)}&sz=w600`,
      `https://lh3.googleusercontent.com/d/${encodeURIComponent(fileId)}=w600`,
    ];
  }
  return [trimmed];
}

export interface HousePhotoPayload {
  name: string;
  mimeType: string;
  base64Data: string;
  previewDataUrl: string;
  sizeBytes: number;
}

export interface DriveUploadResult {
  fileId: string;
  fileName: string;
  webViewLink: string;
}

export type DriveAuthErrorCode =
  | 'missing_auth'
  | 'unauthorized'
  | 'insufficient_scope'
  | 'folder_forbidden';

export class DriveAuthError extends Error {
  code: DriveAuthErrorCode;
  constructor(code: DriveAuthErrorCode, message: string) {
    super(message);
    this.name = 'DriveAuthError';
    this.code = code;
  }
}

function sanitizeFilenameSegment(val: string): string {
  return (val || '')
    .trim()
    .replace(/[\\/:*?"<>|]+/g, '-')
    .replace(/\s+/g, '_')
    .slice(0, 48);
}

export function buildHousePhotoFileName(rqNumber: string, fullName: string): string {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  const hh = String(now.getHours()).padStart(2, '0');
  const mm = String(now.getMinutes()).padStart(2, '0');
  const ss = String(now.getSeconds()).padStart(2, '0');
  const tag = sanitizeFilenameSegment(rqNumber || fullName || 'คำขอ');
  return `รูปหน้าบ้าน_${tag}_${y}${m}${d}_${hh}${mm}${ss}.jpg`;
}

function readFileAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ''));
    reader.onerror = () => reject(new Error('ไม่สามารถอ่านไฟล์รูปภาพได้'));
    reader.readAsDataURL(file);
  });
}

function loadImageElement(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error('ไฟล์ที่เลือกไม่ใช่รูปภาพที่รองรับ'));
    img.src = src;
  });
}

function estimateBase64Bytes(b64: string): number {
  const padding = b64.endsWith('==') ? 2 : b64.endsWith('=') ? 1 : 0;
  return Math.max(0, Math.floor((b64.length * 3) / 4) - padding);
}

/**
 * Validates and compresses a selected house-front image (max 1600px, JPEG 0.85)
 * so phone photos upload reliably to Google Drive via both OAuth and Apps Script.
 */
export async function prepareHousePhotoFile(file: File): Promise<HousePhotoPayload> {
  const lowerName = (file.name || '').toLowerCase();
  const looksLikeImage =
    file.type.startsWith('image/') ||
    /\.(jpe?g|png|webp|gif|bmp|heic|heif)$/i.test(lowerName);
  if (!looksLikeImage) {
    throw new Error('กรุณาเลือกไฟล์รูปภาพ (JPG, PNG, WEBP) เท่านั้น');
  }

  const rawDataUrl = await readFileAsDataUrl(file);
  try {
    const img = await loadImageElement(rawDataUrl);
    const maxDim = 1600;
    const scale = Math.min(1, maxDim / Math.max(img.width || 1, img.height || 1));
    const targetW = Math.max(1, Math.round(img.width * scale));
    const targetH = Math.max(1, Math.round(img.height * scale));

    const canvas = document.createElement('canvas');
    canvas.width = targetW;
    canvas.height = targetH;
    const ctx = canvas.getContext('2d');
    if (!ctx) {
      throw new Error('Canvas unavailable');
    }
    ctx.fillStyle = '#FFFFFF';
    ctx.fillRect(0, 0, targetW, targetH);
    ctx.drawImage(img, 0, 0, targetW, targetH);

    const jpegDataUrl = canvas.toDataURL('image/jpeg', 0.85);
    const base64Data = jpegDataUrl.split(',')[1] || '';
    return {
      name: file.name || 'house_front.jpg',
      mimeType: 'image/jpeg',
      base64Data,
      previewDataUrl: jpegDataUrl,
      sizeBytes: estimateBase64Bytes(base64Data),
    };
  } catch {
    const commaIdx = rawDataUrl.indexOf(',');
    const base64Data = commaIdx !== -1 ? rawDataUrl.slice(commaIdx + 1) : '';
    if (!base64Data) {
      throw new Error('ไม่สามารถประมวลผลไฟล์รูปภาพนี้ได้ กรุณาลองเลือกไฟล์ JPG หรือ PNG');
    }
    return {
      name: file.name || 'house_front.jpg',
      mimeType: file.type || 'image/jpeg',
      base64Data,
      previewDataUrl: rawDataUrl,
      sizeBytes: file.size || estimateBase64Bytes(base64Data),
    };
  }
}

function base64ToUint8Array(base64Data: string): Uint8Array<ArrayBuffer> {
  const binaryString = atob(base64Data);
  const len = binaryString.length;
  const buffer = new ArrayBuffer(len);
  const bytes = new Uint8Array(buffer);
  for (let i = 0; i < len; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return bytes;
}

async function uploadViaGoogleDriveApi(
  accessToken: string,
  fileName: string,
  mimeType: string,
  base64Data: string,
  folderId: string
): Promise<DriveUploadResult> {
  const boundary = `pwaboundary_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
  const metadataPart = JSON.stringify({
    name: fileName,
    parents: [folderId],
  });

  const binaryBytes = base64ToUint8Array(base64Data);
  const preamble =
    `--${boundary}\r\n` +
    `Content-Type: application/json; charset=UTF-8\r\n\r\n` +
    `${metadataPart}\r\n` +
    `--${boundary}\r\n` +
    `Content-Type: ${mimeType || 'image/jpeg'}\r\n\r\n`;
  const epilogue = `\r\n--${boundary}--`;

  const bodyBlob = new Blob([preamble, binaryBytes, epilogue], {
    type: `multipart/related; boundary=${boundary}`,
  });

  const url =
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=id,name,webViewLink';
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': `multipart/related; boundary=${boundary}`,
    },
    body: bodyBlob,
  });

  if (!res.ok) {
    const errJson = await res.json().catch(() => ({}));
    const rawMsg = errJson?.error?.message || `HTTP ${res.status}`;
    const reasons: string[] = (errJson?.error?.errors || []).map((e: any) =>
      String(e?.reason || '')
    );

    if (res.status === 401) {
      throw new DriveAuthError(
        'unauthorized',
        'สิทธิ์การเชื่อมต่อ Google หมดอายุ กรุณากดเชื่อมสิทธิ์ Google ใหม่อีกครั้งเพื่ออัปโหลดรูปเข้า Drive'
      );
    }

    if (
      res.status === 403 &&
      (reasons.includes('insufficientPermissions') ||
        rawMsg.toLowerCase().includes('insufficient'))
    ) {
      throw new DriveAuthError(
        'insufficient_scope',
        'บัญชียังไม่ได้อนุญาตสิทธิ์ Google Drive กรุณากด "อนุญาตสิทธิ์ Google Drive" อีกครั้ง'
      );
    }

    if (res.status === 403 || res.status === 404) {
      throw new DriveAuthError(
        'folder_forbidden',
        `ไม่สามารถบันทึกเข้าโฟลเดอร์ Google Drive ปลายทางได้ (${rawMsg}) กรุณาตรวจสอบสิทธิ์การแก้ไขโฟลเดอร์`
      );
    }

    throw new Error(`อัปโหลดรูปเข้า Google Drive ไม่สำเร็จ (${rawMsg})`);
  }

  const data = await res.json();
  const fileId = String(data?.id || '');
  const webViewLink =
    String(data?.webViewLink || '') ||
    (fileId ? `https://drive.google.com/file/d/${fileId}/view` : buildDriveFolderUrl(folderId));

  return {
    fileId,
    fileName: String(data?.name || fileName),
    webViewLink,
  };
}

async function uploadViaAppsScript(
  appsScriptUrl: string,
  fileName: string,
  mimeType: string,
  base64Data: string,
  folderId: string
): Promise<DriveUploadResult> {
  const actionBody = {
    action: 'uploadHousePhoto',
    folderId,
    fileName,
    mimeType,
    base64Data,
  };

  if (isInsideGoogleAppsScript()) {
    const runner = (window as any).google?.script?.run;
    const res = await new Promise<any>((resolve, reject) => {
      runner
        .withSuccessHandler(resolve)
        .withFailureHandler((err: any) =>
          reject(new Error(err?.message || String(err) || 'Apps Script error'))
        )
        .handleApiAction(actionBody);
    });
    if (!res || res.success === false) {
      throw new Error(res?.message || 'Apps Script ไม่สามารถอัปโหลดรูปเข้า Google Drive ได้');
    }
    return {
      fileId: String(res.fileId || ''),
      fileName: String(res.fileName || fileName),
      webViewLink:
        String(res.webViewLink || '') ||
        (res.fileId
          ? `https://drive.google.com/file/d/${res.fileId}/view`
          : buildDriveFolderUrl(folderId)),
    };
  }

  const proxyRes = await fetch('/api/apps-script', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      appsScriptUrl: appsScriptUrl.trim(),
      payload: actionBody,
    }),
  });
  const json = await proxyRes.json().catch(() => null);
  if (!proxyRes.ok || (json && json.success === false)) {
    throw new Error(
      json?.message || `Apps Script ไม่สามารถอัปโหลดรูปได้ (HTTP ${proxyRes.status})`
    );
  }
  return {
    fileId: String(json?.fileId || ''),
    fileName: String(json?.fileName || fileName),
    webViewLink:
      String(json?.webViewLink || '') ||
      (json?.fileId
        ? `https://drive.google.com/file/d/${json.fileId}/view`
        : buildDriveFolderUrl(folderId)),
  };
}

export async function uploadHousePhotoToDrive(params: {
  photo: HousePhotoPayload;
  rqNumber: string;
  fullName: string;
  accessToken: string | null;
  appsScriptUrl?: string;
  folderId?: string;
}): Promise<DriveUploadResult> {
  const { photo, rqNumber, fullName, accessToken, appsScriptUrl, folderId } = params;
  const targetFolderId = extractDriveFolderId(folderId);
  const fileName = buildHousePhotoFileName(rqNumber, fullName);

  if (isInsideGoogleAppsScript()) {
    return uploadViaAppsScript('', fileName, photo.mimeType, photo.base64Data, targetFolderId);
  }

  if (accessToken) {
    return uploadViaGoogleDriveApi(
      accessToken,
      fileName,
      photo.mimeType,
      photo.base64Data,
      targetFolderId
    );
  }

  if (appsScriptUrl && appsScriptUrl.trim()) {
    try {
      return await uploadViaAppsScript(
        appsScriptUrl,
        fileName,
        photo.mimeType,
        photo.base64Data,
        targetFolderId
      );
    } catch (err: any) {
      const msg = String(err?.message || '');
      if (msg.includes('ไม่รู้จักคำสั่ง')) {
        throw new DriveAuthError(
          'missing_auth',
          'Apps Script ยังเป็นเวอร์ชันเดิมที่ยังไม่รองรับอัปโหลดรูป — กด "เชื่อมสิทธิ์ Google คลิกเดียว" หรืออัปเดต Code.gs ในเมนูตั้งค่าชีท'
        );
      }
      throw err;
    }
  }

  throw new DriveAuthError(
    'missing_auth',
    'กรุณากด "เชื่อมสิทธิ์ Google คลิกเดียว" เพื่ออนุญาตการบันทึกรูปหน้าบ้านเข้า Google Drive'
  );
}
