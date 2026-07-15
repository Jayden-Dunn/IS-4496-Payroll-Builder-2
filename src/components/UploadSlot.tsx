import { useRef, useState } from 'react';

export type SlotStatus = 'empty' | 'ok' | 'error';

export interface SlotChip {
  status: SlotStatus;
  fileName: string;
  message: string;
}

interface UploadSlotProps {
  label: string;
  hint: string;
  optional?: boolean;
  chip: SlotChip;
  onFile: (file: File) => void;
}

export function UploadSlot({ label, hint, optional = false, chip, onFile }: UploadSlotProps): JSX.Element {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragOver, setDragOver] = useState(false);

  const border =
    chip.status === 'ok'
      ? 'border-green-500 bg-green-50'
      : chip.status === 'error'
        ? 'border-red-400 bg-red-50'
        : dragOver
          ? 'border-blue-400 bg-blue-50'
          : 'border-neutral-300 bg-white hover:border-neutral-400';

  return (
    <div
      className={`cursor-pointer rounded-md border border-dashed px-2.5 py-2 transition-colors ${border}`}
      onClick={() => inputRef.current?.click()}
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragOver(false);
        const file = e.dataTransfer.files[0];
        if (file) onFile(file);
      }}
      role="button"
      title={`${label} — drop a .xlsx or .csv here, or click to browse`}
    >
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) onFile(file);
          e.target.value = ''; // allow re-uploading the same file
        }}
      />
      <div className="flex items-center justify-between gap-2">
        <span className="text-xs font-semibold leading-tight">{label}</span>
        {optional && (
          <span className="shrink-0 rounded-full border border-neutral-300 bg-neutral-100 px-1.5 py-px text-[10px] uppercase tracking-wide text-neutral-500">
            optional
          </span>
        )}
      </div>
      <div className="mt-0.5 text-[11px] text-neutral-500">{hint}</div>
      {chip.status !== 'empty' && (
        <div
          className={`mt-1 rounded px-1.5 py-0.5 text-[11px] font-medium ${
            chip.status === 'ok' ? 'bg-green-100 text-green-800' : 'bg-red-100 text-red-800'
          }`}
        >
          <div className="truncate font-semibold">{chip.fileName}</div>
          <div>{chip.message}</div>
        </div>
      )}
    </div>
  );
}
