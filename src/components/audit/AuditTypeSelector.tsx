import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";

interface AuditTypeSelectorProps {
  auditType: string;
  subAuditSector: string;
  onAuditTypeChange: (value: string) => void;
  onSubAuditSectorChange: (value: string) => void;
}

const auditTypes = [
  { value: "financial", label: "Financial Audit" },
  { value: "operational", label: "Operational Audit" },
  { value: "compliance", label: "Compliance Audit" },
  { value: "it", label: "IT Audit" },
  { value: "forensic", label: "Forensic Audit" },
  { value: "integrated", label: "Integrated Audit" },
];

const subAuditSectors: Record<string, { value: string; label: string }[]> = {
  financial: [
    { value: "banking", label: "Banking & Financial Services" },
    { value: "insurance", label: "Insurance" },
    { value: "investment", label: "Investment Management" },
    { value: "capital-markets", label: "Capital Markets" },
  ],
  operational: [
    { value: "supply-chain", label: "Supply Chain" },
    { value: "manufacturing", label: "Manufacturing" },
    { value: "retail", label: "Retail Operations" },
    { value: "healthcare", label: "Healthcare Operations" },
  ],
  compliance: [
    { value: "regulatory", label: "Regulatory Compliance" },
    { value: "sox", label: "SOX Compliance" },
    { value: "gdpr", label: "GDPR & Privacy" },
    { value: "aml", label: "AML/KYC Compliance" },
  ],
  it: [
    { value: "cybersecurity", label: "Cybersecurity" },
    { value: "infrastructure", label: "IT Infrastructure" },
    { value: "application", label: "Application Controls" },
    { value: "data", label: "Data Governance" },
  ],
  forensic: [
    { value: "fraud", label: "Fraud Investigation" },
    { value: "disputes", label: "Dispute Advisory" },
    { value: "corruption", label: "Anti-Corruption" },
    { value: "litigation", label: "Litigation Support" },
  ],
  integrated: [
    { value: "esg", label: "ESG Reporting" },
    { value: "sustainability", label: "Sustainability" },
    { value: "combined", label: "Combined Assurance" },
    { value: "multi-entity", label: "Multi-Entity" },
  ],
};

export function AuditTypeSelector({
  auditType,
  subAuditSector,
  onAuditTypeChange,
  onSubAuditSectorChange,
}: AuditTypeSelectorProps) {
  const availableSectors = auditType ? subAuditSectors[auditType] || [] : [];

  return (
    <div className="space-y-4">
      <h3 className="kpmg-section-title">Audit Configuration</h3>

      <div className="grid grid-cols-2 gap-4">
        <div className="space-y-2">
          <label htmlFor="audit-type" className="text-sm font-medium">
            Audit Type
          </label>
          <Select value={auditType} onValueChange={onAuditTypeChange}>
            <SelectTrigger id="audit-type" className="bg-card">
              <SelectValue placeholder="Select audit type" />
            </SelectTrigger>
            <SelectContent>
              {auditTypes.map((type) => (
                <SelectItem key={type.value} value={type.value}>
                  {type.label}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        </div>

        <div className="space-y-2">
          <label htmlFor="sub-sector" className="text-sm font-medium">
            Sub-Audit Sector
          </label>
          <Select
            value={subAuditSector}
            onValueChange={onSubAuditSectorChange}
            disabled={!auditType}
          >
            <SelectTrigger id="sub-sector" className="bg-card">
              <SelectValue placeholder={auditType ? "Select sector" : "Select audit type first"} />
            </SelectTrigger>
            <SelectContent>
              {availableSectors.map((sector) => (
                <SelectItem key={sector.value} value={sector.value}>
                  {sector.label}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        </div>
      </div>
    </div>
  );
}
