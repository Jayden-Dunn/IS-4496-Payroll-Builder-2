interface TopBarProps {
  canDownload: boolean;
  onDownload: () => void;
  onOpenSettings: () => void;
  downloadName: string;
  /** Round 4 §2.3 — "Export current filtered view" (Pivot by Crew filters). */
  exportFiltered: boolean;
  onExportFilteredChange: (v: boolean) => void;
  filtersActive: boolean;
}

export function TopBar({
  canDownload,
  onDownload,
  onOpenSettings,
  downloadName,
  exportFiltered,
  onExportFilteredChange,
  filtersActive,
}: TopBarProps): JSX.Element {
  return (
    <header className="flex shrink-0 items-center gap-4 border-b border-neutral-300 bg-white px-4 py-2 shadow-sm">
      <div className="flex items-baseline gap-3">
        <h1 className="text-lg font-bold tracking-tight">
          <span className="text-brand">%</span> of Completion Builder
        </h1>
        <span className="hidden text-xs text-neutral-500 sm:inline">
          AllPro Painters · weekly percent-completion payroll · runs entirely in your browser
        </span>
      </div>
      <div className="ml-auto flex items-center gap-2">
        <button
          type="button"
          className="btn"
          onClick={onOpenSettings}
          title="Estimator colors & order"
          aria-label="Settings"
        >
          ⚙ Settings
        </button>
        <label
          className={`flex items-center gap-1 text-[11px] ${filtersActive ? 'text-neutral-700' : 'text-neutral-400'}`}
          title="When on, the Pivot by Crew sheet exports exactly what the on-screen filters show. Off = full data."
        >
          <input
            type="checkbox"
            checked={exportFiltered}
            disabled={!filtersActive}
            onChange={(e) => onExportFilteredChange(e.target.checked)}
          />
          Export current filtered view
        </label>
        <button
          type="button"
          className="btn-primary"
          disabled={!canDownload}
          onClick={onDownload}
          title={canDownload ? `Download ${downloadName}` : 'Create the report first'}
        >
          ⬇ Download .xlsx
        </button>
      </div>
    </header>
  );
}
