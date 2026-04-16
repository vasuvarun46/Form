// import * as React from "react";
// import { Row } from "./TeamService";

// export interface ActionConfirmModalProps {
//   kind: "approve" | "reject";
//   row: Row;
//   loading: boolean;
//   reason: string;
//   reasonError: string;
//   onReasonChange: (value: string) => void;
//   onCancel: () => void;
//   onConfirm: () => void;
  
// }

// export default function ActionConfirmModal(
//   props: ActionConfirmModalProps
// ): JSX.Element {

//   const title =
//     props.kind === "reject"
//       ? "Do you want to reject the request?"
//       : "Do you want to approve the request?";

//   const primaryText =
//     props.kind === "reject" ? "Reject" : "Approve";

//   const primaryClass =
//     props.kind === "reject"
//       ? "bg-red-600 hover:bg-red-700"
//       : "bg-green-600 hover:bg-green-700";

//   return (
//     <div className="fixed inset-0 bg-black/40 z-[9999] flex items-center justify-center">
//       <div className="bg-white w-[520px] rounded-xl shadow-xl p-6 relative">

//         <button
//           onClick={props.onCancel}
//           className="absolute right-4 top-4 text-gray-500"
//         >
//           ✕
//         </button>

//         <h2 className="text-lg font-bold text-center">{title}</h2>

//         <p className="text-center text-sm mt-2">
//           ID <strong>&quot;{props.row.id}&quot;</strong>
//         </p>

//         <div className="mt-5">
//           <label className="text-sm font-medium">
//             {props.kind === "reject" ? "Rejection reason *" : "Comment"}
//           </label>

//           <textarea
//             value={props.reason}
//             onChange={(e) => props.onReasonChange(e.target.value)}
//             disabled={props.loading}
//             maxLength={100}
//             rows={3}
//             placeholder={props.kind === "reject" ? "Please provide a reason for rejection..." : "Optional comment..."}
//             className="w-full border rounded-md mt-2 p-2 text-sm"
//           />
//           <div className="text-right text-xs text-gray-500 mt-1">
//             {props.reason.length}/100
//           </div>
//         </div>

//         <div className="flex justify-center gap-4 mt-6">
//           <button
//             onClick={props.onCancel}
//             className="px-4 py-2 bg-gray-200 rounded-md"
//           >
//             Cancel
//           </button>

//           <button
//             disabled={
//               props.loading ||
//               (props.kind === "reject" && !props.reason.trim())
//             }
//             onClick={props.onConfirm}
//             className={`px-4 py-2 text-white rounded-md ${primaryClass}`}
//           >
//             {primaryText}
//           </button>
//         </div>
//       </div>
//     </div>
//   );
// }
