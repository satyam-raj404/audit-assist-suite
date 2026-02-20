import { useState, useCallback } from "react";
import { Upload, X, FileSpreadsheet, ArrowLeftRight, Play, RotateCcw, CheckCircle2, AlertCircle } from "lucide-react";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";
import { Progress } from "@/components/ui/progress";
import { toast } from "sonner";
import { AuditTypeSelector } from "@/components/audit/AuditTypeSelector";
import { ReconciliationFileSlot } from "./ReconciliationFileSlot";
import { ReconciliationResults } from "./ReconciliationResults";

interface ReconciliationFile {
  id: string;
  name: string;
  size: number;
}

interface ReconciliationPanelProps {
  onStatusChange: (status: "idle" | "running" | "success" | "error", message: string) => void;
}

export function ReconciliationPanel({ onStatusChange }: ReconciliationPanelProps) {
  const [file1, setFile1] = useState<ReconciliationFile | null>(null);
  const [file2, setFile2] = useState<ReconciliationFile | null>(null);
  const [auditType, setAuditType] = useState("");
  const [subAuditSector, setSubAuditSector] = useState("");
  const [isRunning, setIsRunning] = useState(false);
  const [progress, setProgress] = useState(0);
  const [results, setResults] = useState<null | { matched: number; mismatched: number; missing: number }>(null);

  const handleFileUpload = useCallback(
    (fileNum: 1 | 2) => (e: React.ChangeEvent<HTMLInputElement>) => {
      if (!e.target.files || !e.target.files[0]) return;

      const file = e.target.files[0];
      const ext = file.name.toLowerCase().split(".").pop();

      if (!["xlsx", "xls", "csv"].includes(ext || "")) {
        toast.error("Invalid file type", { description: "Please upload XLSX, XLS, or CSV files only." });
        return;
      }

      const newFile: ReconciliationFile = {
        id: crypto.randomUUID(),
        name: file.name,
        size: file.size,
      };

      if (fileNum === 1) setFile1(newFile);
      else setFile2(newFile);
      setResults(null);
    },
    []
  );

  const removeFile = useCallback((fileNum: 1 | 2) => {
    if (fileNum === 1) setFile1(null);
    else setFile2(null);
    setResults(null);
  }, []);

  const handleReset = () => {
    setFile1(null);
    setFile2(null);
    setAuditType("");
    setSubAuditSector("");
    setProgress(0);
    setResults(null);
    onStatusChange("idle", "Ready for reconciliation");
  };

  const runReconciliation = async () => {
    if (!file1 || !file2) return;

    setIsRunning(true);
    setProgress(0);
    setResults(null);
    onStatusChange("running", "Running reconciliation...");

    for (let i = 0; i <= 100; i += 2) {
      await new Promise((resolve) => setTimeout(resolve, 40));
      setProgress(i);
    }

    const simulatedResults = {
      matched: Math.floor(Math.random() * 500) + 200,
      mismatched: Math.floor(Math.random() * 50) + 10,
      missing: Math.floor(Math.random() * 20) + 5,
    };

    setResults(simulatedResults);
    setIsRunning(false);
    setProgress(100);
    onStatusChange("success", "Reconciliation completed");
    toast.success("Reconciliation Complete", {
      description: `Found ${simulatedResults.mismatched} discrepancies`,
    });
  };

  const canRun = file1 && file2 && auditType && subAuditSector && !isRunning;

  return (
    <div className="h-full flex flex-col">
      <div className="flex-1 overflow-y-auto p-6 space-y-6">
        {/* Audit Type & Sub Sector */}
        <AuditTypeSelector
          auditType={auditType}
          subAuditSector={subAuditSector}
          onAuditTypeChange={(value) => {
            setAuditType(value);
            setSubAuditSector("");
          }}
          onSubAuditSectorChange={setSubAuditSector}
        />

        {/* File Upload Section */}
        <div className="space-y-4 border-t border-border pt-6">
          <h3 className="kpmg-section-title">Select Files to Compare</h3>

          <div className="flex items-stretch gap-6">
            <ReconciliationFileSlot
              file={file1}
              label="Source File"
              onUpload={handleFileUpload(1)}
              onRemove={() => removeFile(1)}
            />

            <div className="flex items-center justify-center px-4">
              <div className="w-12 h-12 rounded-full bg-primary/10 flex items-center justify-center border border-primary/20">
                <ArrowLeftRight className="h-6 w-6 text-primary" />
              </div>
            </div>

            <ReconciliationFileSlot
              file={file2}
              label="Target File"
              onUpload={handleFileUpload(2)}
              onRemove={() => removeFile(2)}
            />
          </div>
        </div>

        {/* Progress Section */}
        {(isRunning || progress > 0) && (
          <div className="space-y-3 p-4 bg-card rounded border border-border">
            <div className="flex items-center justify-between">
              <span className="text-sm font-medium">Reconciliation Progress</span>
              <span className="text-sm text-muted-foreground">{Math.round(progress)}%</span>
            </div>
            <Progress value={progress} className="h-2" />
          </div>
        )}

        {/* Results */}
        {results && (
          <ReconciliationResults results={results} file1Name={file1?.name} file2Name={file2?.name} />
        )}
      </div>

      {/* Action Buttons */}
      <div className="shrink-0 p-4 border-t border-border bg-muted/30 flex items-center justify-between">
        <Button
          variant="ghost"
          onClick={handleReset}
          disabled={isRunning}
          className="text-muted-foreground"
        >
          <RotateCcw className="h-4 w-4 mr-2" />
          Reset
        </Button>

        <div className="flex gap-3">
          {results && (
            <Button variant="kpmg-outline">
              Export Results
            </Button>
          )}
          <Button
            variant="kpmg"
            onClick={runReconciliation}
            disabled={!canRun}
          >
            <Play className="h-4 w-4 mr-2" />
            Run Reconciliation
          </Button>
        </div>
      </div>
    </div>
  );
}
