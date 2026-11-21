import React, { useState } from 'react';
import { ShieldCheck, X } from 'lucide-react';

interface AdminLoginModalProps {
  open: boolean;
  onClose: () => void;
  onSuccess: () => void;
}

const ADMIN_USERS: Array<{ u: string; p: string }> = [
  { u: 'Admin', p: '123456' },
  { u: '13367', p: '13367' },
  { u: '16479', p: '16479' },
  { u: '15563', p: '15563' },
  { u: '14356', p: '14356' },
  { u: '15687', p: '15687' },
  { u: '17191', p: '17191' },
  { u: '15676', p: '15676' },
];

export const AdminLoginModal: React.FC<AdminLoginModalProps> = ({
  open,
  onClose,
  onSuccess,
}) => {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');

  if (!open) return null;

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const u = username.trim();
    const p = password.trim();
    const matched = ADMIN_USERS.some((acc) => acc.u === u && acc.p === p);
    if (matched) {
      setError('');
      setUsername('');
      setPassword('');
      onSuccess();
      onClose();
    } else {
      setError('ชื่อผู้ใช้หรือรหัสผ่านผู้ดูแลระบบไม่ถูกต้อง');
    }
  };

  return (
    <div
      className="fixed inset-0 z-[2400] flex items-center justify-center bg-slate-900/60 dark:bg-black/70 backdrop-blur-[2px] p-4"
      role="dialog"
      aria-modal="true"
      aria-labelledby="admin-login-title"
    >
      <div className="bg-white dark:bg-[#111C2B] border border-slate-200 dark:border-[#213145] rounded-xl w-full max-w-md shadow-xl overflow-hidden">
        <div className="px-5 py-4 border-b border-[#1A5A8F] flex items-center justify-between bg-[#2271B3] dark:bg-[#1A5A8F] text-white">
          <div className="flex items-center gap-2.5">
            <ShieldCheck className="w-5 h-5 text-white/90" />
            <h3 id="admin-login-title" className="font-semibold text-base">
              เข้าสู่ระบบผู้ดูแล (Admin)
            </h3>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="text-white/75 hover:text-white p-1 rounded-lg transition-colors"
            aria-label="ปิดหน้าต่าง"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-5 space-y-4">
          <p className="text-xs text-slate-500 dark:text-slate-400">
            เข้าสู่ระบบผู้ดูแลเพื่อเปิดใช้งานฟอร์ม <strong className="text-slate-700 dark:text-slate-200">บันทึกคำขอใหม่</strong> แก้ไข และลบรายการคำขอมาตรวัดน้ำ
          </p>

          <div>
            <label className="block text-xs font-medium text-slate-700 dark:text-slate-300 mb-1.5">
              ชื่อผู้ใช้ (Username)
            </label>
            <input
              type="text"
              required
              autoFocus
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              placeholder="ระบุชื่อผู้ใช้ (เช่น Admin)"
              className="w-full bg-slate-50 dark:bg-[#0F1826] border border-slate-300 dark:border-[#263850] rounded-lg py-2 px-3 text-sm text-slate-900 dark:text-slate-100 placeholder:text-slate-400 dark:placeholder:text-slate-500 outline-none focus:bg-white dark:focus:bg-[#162336] focus:border-[#2271B3] dark:focus:border-[#5EA6E0] focus:ring-2 focus:ring-[#2271B3]/20"
            />
          </div>

          <div>
            <label className="block text-xs font-medium text-slate-700 dark:text-slate-300 mb-1.5">
              รหัสผ่าน (Password)
            </label>
            <input
              type="password"
              required
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder="ระบุรหัสผ่าน"
              className="w-full bg-slate-50 dark:bg-[#0F1826] border border-slate-300 dark:border-[#263850] rounded-lg py-2 px-3 text-sm text-slate-900 dark:text-slate-100 placeholder:text-slate-400 dark:placeholder:text-slate-500 outline-none focus:bg-white dark:focus:bg-[#162336] focus:border-[#2271B3] dark:focus:border-[#5EA6E0] focus:ring-2 focus:ring-[#2271B3]/20"
            />
          </div>

          {error && (
            <div className="text-xs text-rose-600 dark:text-rose-300 bg-rose-50 dark:bg-rose-950/40 border border-rose-200 dark:border-rose-800/60 rounded-lg px-3 py-2">
              {error}
            </div>
          )}

          <div className="flex items-center justify-end gap-2 pt-2">
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-2 text-xs font-medium text-slate-700 dark:text-slate-300 bg-white dark:bg-[#0F1826] border border-slate-300 dark:border-[#263850] rounded-lg hover:bg-slate-50 dark:hover:bg-[#162336] transition-colors"
            >
              ยกเลิก
            </button>
            <button
              type="submit"
              className="px-4 py-2 text-xs font-semibold text-white bg-[#2271B3] hover:bg-[#1A5A8F] dark:bg-[#3584C7] dark:hover:bg-[#2271B3] rounded-lg transition-colors"
            >
              เข้าสู่ระบบ
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};
