import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";

// Audit types that produce Excel output (shown in FS Reconciliation panel)
const RECONCILIATION_TYPES = [
  { value: "FD - Sampling", label: "FD - Sampling" },
  { value: "KYC",           label: "KYC" },
  { value: "CCIL",          label: "CCIL" },
  { value: "FOBO",          label: "FOBO" },
  { value: "IOA KPI1",      label: "IOA KPI1" },
  { value: "IOA KPI2",      label: "IOA KPI2" },
  { value: "IOA KPI3",      label: "IOA KPI3" },
  { value: "IOA KPI4",      label: "IOA KPI4" },
  { value: "Swift",         label: "Swift" },
];

/** Swift is the only type that requires two input files. */
export function needsTwoFiles(auditType: string): boolean {
  return auditType === "Swift";
}

interface ReconciliationTypeSelectorProps {
  auditType: string;
  onAuditTypeChange: (value: string) => void;
}

export function ReconciliationTypeSelector({
  auditType,
  onAuditTypeChange,
}: ReconciliationTypeSelectorProps) {
  return (
    <div className="space-y-4">
      <h3 className="kpmg-section-title">Audit Configuration</h3>

      <div className="space-y-2">
        <label htmlFor="recon-audit-type" className="text-sm font-medium">
          Audit Type
        </label>
        <Select value={auditType} onValueChange={onAuditTypeChange}>
          <SelectTrigger id="recon-audit-type" className="bg-card w-64">
            <SelectValue placeholder="Select audit type" />
          </SelectTrigger>
          <SelectContent>
            {RECONCILIATION_TYPES.map((t) => (
              <SelectItem key={t.value} value={t.value}>{t.label}</SelectItem>
            ))}
          </SelectContent>
        </Select>
      </div>
    </div>
  );
}
