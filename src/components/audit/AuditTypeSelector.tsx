import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";
import { Label } from "@/components/ui/label";

// Audit types that generate PPTX reports (shown in Report Automation panel)
const AUDIT_TYPES = [
  { value: "Internal Audit",        label: "Internal Audit" },
  { value: "Internal Audit-Zensar", label: "Internal Audit - Zensar" },
  { value: "Concurrent Audit",      label: "Concurrent Audit" },
  { value: "ICOFR",                 label: "ICOFR" },
];

const UTILITY_TYPES: Record<string, string[]> = {
  "Internal Audit":        ["Report"],
  "Internal Audit-Zensar": ["Report"],
  "Concurrent Audit":      ["Dashboard", "Report"],
  "ICOFR":                 ["Dashboard"],
};

const MONTHS = [
  "January","February","March","April","May","June",
  "July","August","September","October","November","December",
];

const YEARS = Array.from({ length: 11 }, (_, i) => String(2020 + i));

export interface AuditTypeSelectorProps {
  auditType: string;
  utilityType: string;
  reportType: string;
  month: string;
  year: string;
  onAuditTypeChange: (value: string) => void;
  onUtilityTypeChange: (value: string) => void;
  onReportTypeChange: (value: string) => void;
  onMonthChange: (value: string) => void;
  onYearChange: (value: string) => void;
}

export function AuditTypeSelector({
  auditType,
  utilityType,
  reportType,
  month,
  year,
  onAuditTypeChange,
  onUtilityTypeChange,
  onReportTypeChange,
  onMonthChange,
  onYearChange,
}: AuditTypeSelectorProps) {
  const utilityOptions = auditType ? (UTILITY_TYPES[auditType] ?? []) : [];
  const showReportType = auditType === "Concurrent Audit" && utilityType === "Report";
  const showMonthYear  = showReportType;

  return (
    <div className="space-y-4">
      <h3 className="kpmg-section-title">Audit Configuration</h3>

      <div className="grid grid-cols-2 gap-4">
        {/* Audit Type */}
        <div className="space-y-2">
          <label htmlFor="audit-type" className="text-sm font-medium">
            Audit Type
          </label>
          <Select
            value={auditType}
            onValueChange={(val) => { onAuditTypeChange(val); onUtilityTypeChange(""); }}
          >
            <SelectTrigger id="audit-type" className="bg-card">
              <SelectValue placeholder="Select audit type" />
            </SelectTrigger>
            <SelectContent>
              {AUDIT_TYPES.map((t) => (
                <SelectItem key={t.value} value={t.value}>{t.label}</SelectItem>
              ))}
            </SelectContent>
          </Select>
        </div>

        {/* Utility Type */}
        <div className="space-y-2">
          <label htmlFor="utility-type" className="text-sm font-medium">
            Utility Type
          </label>
          <Select
            value={utilityType}
            onValueChange={onUtilityTypeChange}
            disabled={utilityOptions.length === 0}
          >
            <SelectTrigger id="utility-type" className="bg-card">
              <SelectValue placeholder={auditType ? "Select utility type" : "Select audit type first"} />
            </SelectTrigger>
            <SelectContent>
              {utilityOptions.map((u) => (
                <SelectItem key={u} value={u}>{u}</SelectItem>
              ))}
            </SelectContent>
          </Select>
        </div>
      </div>

      {/* Report Type — only for Concurrent Audit + Report */}
      {showReportType && (
        <div className="space-y-2">
          <p className="text-sm font-medium">Report Type</p>
          <RadioGroup
            value={reportType}
            onValueChange={onReportTypeChange}
            className="flex gap-6"
          >
            {["Both", "Draft", "Final"].map((rt) => (
              <div key={rt} className="flex items-center gap-2">
                <RadioGroupItem value={rt} id={`rt-${rt}`} />
                <Label htmlFor={`rt-${rt}`} className="cursor-pointer">{rt}</Label>
              </div>
            ))}
          </RadioGroup>
        </div>
      )}

      {/* Month + Year — only for Concurrent Audit + Report */}
      {showMonthYear && (
        <div className="grid grid-cols-2 gap-4">
          <div className="space-y-2">
            <label htmlFor="report-month" className="text-sm font-medium">
              Report Issuance Month
            </label>
            <Select value={month} onValueChange={onMonthChange}>
              <SelectTrigger id="report-month" className="bg-card">
                <SelectValue placeholder="Select month" />
              </SelectTrigger>
              <SelectContent>
                {MONTHS.map((m) => (
                  <SelectItem key={m} value={m}>{m}</SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>

          <div className="space-y-2">
            <label htmlFor="report-year" className="text-sm font-medium">
              Year
            </label>
            <Select value={year} onValueChange={onYearChange}>
              <SelectTrigger id="report-year" className="bg-card">
                <SelectValue placeholder="Select year" />
              </SelectTrigger>
              <SelectContent>
                {YEARS.map((y) => (
                  <SelectItem key={y} value={y}>{y}</SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
        </div>
      )}
    </div>
  );
}
