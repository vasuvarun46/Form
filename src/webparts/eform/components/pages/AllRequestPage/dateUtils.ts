export const formatDateDdMmYyyy = (iso?: string): string => {
  if (!iso) return "-";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return "-";
 
  // ✅ Manually apply UTC+7 (Bangkok) offset — SPFx runs server-side so local timezone doesn't apply
  const bangkokOffset = 7 * 60 * 60 * 1000; // 7 hours in ms
  const bangkokDate = new Date(d.getTime() + bangkokOffset);
 
  const dd = String(bangkokDate.getUTCDate()).padStart(2, "0");
  const mm = String(bangkokDate.getUTCMonth() + 1).padStart(2, "0");
  const yyyy = String(bangkokDate.getUTCFullYear());
 
  return `${dd}-${mm}-${yyyy}`;
};
 
export const yyyyMmDdToIso = (yyyyMmDd: string): string => {
  const [yyyy, mm, dd] = (yyyyMmDd || "").trim().split("-");
  if (!yyyy || !mm || !dd) throw new Error("Invalid date. Use yyyy-MM-dd");
  // ✅ Midnight Bangkok = 17:00 UTC
  return `${yyyy}-${mm}-${dd}T17:00:00Z`;
};

