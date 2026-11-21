import { Download } from 'lucide-react';
import type { ActiveTab } from '../types/attendance';

interface TopBarProps {
  activeTab: ActiveTab;
  onSelectTab: (tab: ActiveTab) => void;
  onExportExcel: () => void;
  isExporting: boolean;
  canExport: boolean;
}

const NAV_ITEMS: { id: ActiveTab; label: string }[] = [
  { id: 'builder', label: 'Builder' },
  { id: 'preview', label: 'Sheet Preview' },
  { id: 'analytics', label: 'Analytics' },
  { id: 'python', label: 'Python Code' },
];

export function TopBar({
  activeTab,
  onSelectTab,
  onExportExcel,
  isExporting,
  canExport,
}: TopBarProps) {
  return (
    <header className="sticky top-0 z-30 bg-white border-b border-slate-200">
      <div className="max-w-[1440px] mx-auto px-6 h-16 flex items-center justify-between gap-8">
        {/* Zone 1: Single text element wordmark */}
        <button
          type="button"
          onClick={() => onSelectTab('builder')}
          className="text-lg font-bold tracking-tight text-slate-900 whitespace-nowrap shrink-0 font-display focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-blue-600 cursor-pointer"
        >
          Biometric Attendance Studio
        </button>

        {/* Zone 2: 4 single-line text navigation links */}
        <nav aria-label="Primary" className="flex items-center gap-6 overflow-x-auto">
          {NAV_ITEMS.map((item) => {
            const isActive = activeTab === item.id;
            return (
              <button
                key={item.id}
                type="button"
                onClick={() => onSelectTab(item.id)}
                className={`relative py-5 text-sm font-medium transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-2 focus-visible:outline-blue-600 ${
                  isActive
                    ? 'text-blue-700'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                {item.label}
                {isActive && (
                  <span className="absolute inset-x-0 bottom-0 h-0.5 bg-blue-700" />
                )}
              </button>
            );
          })}
        </nav>

        {/* Zone 3: Single primary CTA action */}
        <div className="flex items-center shrink-0">
          <button
            type="button"
            onClick={onExportExcel}
            disabled={!canExport || isExporting}
            className="inline-flex items-center gap-2 px-4 py-2 text-sm font-medium text-white bg-blue-700 rounded-lg hover:bg-blue-800 disabled:opacity-50 disabled:cursor-not-allowed transition-colors whitespace-nowrap shrink-0 cursor-pointer focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-blue-600"
          >
            <Download className="w-4 h-4 shrink-0" />
            <span>{isExporting ? 'Exporting...' : 'Export Excel'}</span>
          </button>
        </div>
      </div>
    </header>
  );
}
