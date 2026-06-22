import * as React from "react";
import { AadUserProfile, ISuparaItem } from "../CreateRequestService";

type Props = {
  userData: ISuparaItem[];
  card: string;
  aadProfile?: AadUserProfile | null;
};

function Info({ label, value }: { label: string; value: React.ReactNode }): JSX.Element {
  return (
    <div className="flex items-start text-[13px] pdf-thai min-w-0">
      <span className="text-gray-500 mr-2 whitespace-nowrap">{label}</span>
      <span className="text-gray-900 font-semibold whitespace-nowrap overflow-hidden text-ellipsis">
        {value}
      </span>
    </div>
  );
}

export default function GeneralInfoSection({ userData, card, aadProfile }: Props): JSX.Element {
  const u = userData[0];

  const displayName   = aadProfile?.displayName   || u?.Title              || "-";
  const jobPosition   = aadProfile?.jobTitle                               || "-";
  const department    = u?.Department?.Title       || aadProfile?.department || "-";

  return (
    <section className={`${card} p-4 sm:p-5`}>
      <h2 className="text-[13px] font-bold text-gray-900 mb-4">General Information</h2>
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-x-10 gap-y-4">
        <Info label="ชื่อผู้เบิก / Requestor Name :" value={displayName} />
        <Info label="รหัสพนักงาน / Employee ID :"    value={u?.EmployeeID || "-"} />
        <Info label="อีเมล / Email :"                value={u?.Email      || "-"} />

        <div className="lg:col-start-1">
          <Info label="ตำแหน่ง / Job Position :" value={jobPosition} />
        </div>
        <div className="lg:col-start-2">
          <Info label="แผนก / Department :"      value={department} />
        </div>
      </div>
    </section>
  );
}
