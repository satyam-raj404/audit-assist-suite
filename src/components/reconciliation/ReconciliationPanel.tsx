import { useState, useCallback, useRef } from "react";
import { ArrowLeftRight, Play, RotateCcw } from "lucide-react";
import { Button } from "@/components/ui/button";
import { toast } from "sonner";
import { OutputFolderSelector } from "@/components/audit/OutputFolderSelector";
import { ReconciliationTypeSelector, needsTwoFiles } from "./ReconciliationTypeSelector";
import { ReconciliationFileSlot, type ReconFile } from "./ReconciliationFileSlot";
import { ProcessingDialog, type ProcessingStep } from "@/components/shared/ProcessingDialog";
import { api } from "@/lib/api";

const POLL_INTERVAL_MS = 2000;

const STEPS_TEMPLATE: ProcessingStep[] = [
  { id: "load",    label: "Loading input file(s)",     status: "pending" },
  { id: "process", label: "Running analysis",          status: "pending" },
  { id: "write",   label: "Writing output file",       status: "pending" },
];

interface ReconciliationPanelProps {
  onStatusChange: (status: "idle" | "running" | "success" | "error", message: string) => void;
}

export function ReconciliationPanel({ onStatusChange }: ReconciliationPanelProps) {
  const [auditType, setAuditType]     = useState("");
  const [file1, setFile1]             = useState<ReconFile | null>(null);
  const [file2, setFile2]             = useState<ReconFile | null>(null);
  const [outputPath, setOutputPath]   = useState("");

  const [showProcessing, setShowProcessing]  = useState(false);
  const [dialogSteps, setDialogSteps]        = useState<ProcessingStep[]>(STEPS_TEMPLATE);
  const [finalStatus, setFinalStatus]        = useState<"running" | "success" | "error">("running");
  const [reportPath, setReportPath]          = useState("");
  const pollRef = useRef<ReturnType<typeof setInterval> | null>(null);

  const twoFiles = needsTwoFiles(auditType);

  const canRun =
    Boolean(auditType) &&
    Boolean(file1) &&
    (!twoFiles || Boolean(file2)) &&
    Boolean(outputPath);

  const stopPolling = () => {
    if (pollRef.current) {
      clearInterval(pollRef.current);
      pollRef.current = null;
    }
  };

  // Reset file2 when switching away from Swift
  const handleAuditTypeChange = (val: string) => {
    setAuditType(val);
    if (!needsTwoFiles(val)) setFile2(null);
  };

  const handleRun = useCallback(async () => {
    if (!canRun || !file1) return;

    setDialogSteps(STEPS_TEMPLATE);
    setFinalStatus("running");
    setReportPath("");
    setShowProcessing(true);
    onStatusChange("running", "Running analysis…");

    try {
      const result = await api.startReconciliation({
        audit_type: auditType,
        source_file_id: file1.id,
        target_file_id: twoFiles && file2 ? file2.id : undefined,
        output_path: outputPath,
      });

      const reconId = result.id;

      pollRef.current = setInterval(async () => {
        try {
          const status = await api.getReconciliationStatus(reconId);

          // Derive step statuses from progress percentage
          const pct = status.progress;
          setDialogSteps([
            { id: "load",    label: "Loading input file(s)",   status: pct >= 25 ? "complete" : pct > 0 ? "running" : "pending" },
            { id: "process", label: "Running analysis",        status: pct >= 75 ? "complete" : pct >= 25 ? "running" : "pending" },
            { id: "write",   label: "Writing output file",     status: pct >= 100 ? "complete" : pct >= 75 ? "running" : "pending" },
          ]);

          if (status.status === "complete") {
            stopPolling();
            setFinalStatus("success");
            setReportPath(status.report_path || outputPath);
            onStatusChange("success", "Analysis completed");
            toast.success("Analysis Complete", {
              description: `Output saved to ${status.report_path || outputPath}`,
            });
          } else if (status.status === "error") {
            stopPolling();
            setFinalStatus("error");
            setReportPath(status.report_path || "");
            onStatusChange("error", "Analysis failed");
            toast.error("Analysis Failed");
          }
        } catch {
          stopPolling();
          setFinalStatus("error");
          onStatusChange("error", "Failed to get status");
        }
      }, POLL_INTERVAL_MS);

    } catch (err) {
      setFinalStatus("error");
      onStatusChange("error", "Failed to start analysis");
      toast.error("Failed to start", {
        description: err instanceof Error ? err.message : "Unknown error",
      });
    }
  }, [canRun, auditType, file1, file2, twoFiles, outputPath, onStatusChange]);

  const handleReset = () => {
    stopPolling();
    setAuditType("");
    setFile1(null);
    setFile2(null);
    setOutputPath("");
    setReportPath("");
    setDialogSteps(STEPS_TEMPLATE);
    setFinalStatus("running");
    onStatusChange("idle", "Ready for analysis");
  };

  return (
    <div className="h-full flex flex-col">
      <div className="flex-1 overflow-y-auto p-6 space-y-6">
        {/* Audit Type */}
        <ReconciliationTypeSelector
          auditType={auditType}
          onAuditTypeChange={handleAuditTypeChange}
        />

        {/* File Upload Section */}
        <div className="space-y-4 border-t border-border pt-6">
          <h3 className="kpmg-section-title">
            {twoFiles ? "Select Files to Reconcile" : "Select Input File"}
          </h3>

          {twoFiles ? (
            <div className="flex items-stretch gap-6">
              <ReconciliationFileSlot
                file={file1}
                label="File A — Reference"
                onFileChange={setFile1}
              />
              <div className="flex items-center justify-center px-4">
                <div className="w-12 h-12 rounded-full bg-primary/10 flex items-center justify-center border border-primary/20">
                  <ArrowLeftRight className="h-6 w-6 text-primary" />
                </div>
              </div>
              <ReconciliationFileSlot
                file={file2}
                label="File B — Reconciliation"
                onFileChange={setFile2}
              />
            </div>
          ) : (
            <ReconciliationFileSlot
              file={file1}
              label="Input Excel File"
              onFileChange={setFile1}
            />
          )}
        </div>

        {/* Output Folder */}
        <div className="border-t border-border pt-6">
          <OutputFolderSelector outputPath={outputPath} onOutputPathChange={setOutputPath} />
        </div>
      </div>

      {/* Action Buttons */}
      <div className="shrink-0 p-4 border-t border-border bg-muted/30 flex items-center justify-between">
        <Button
          variant="ghost"
          onClick={handleReset}
          className="text-muted-foreground"
        >
          <RotateCcw className="h-4 w-4 mr-2" />
          Reset
        </Button>

        <Button
          variant="kpmg"
          onClick={handleRun}
          disabled={!canRun}
        >
          <Play className="h-4 w-4 mr-2" />
          Run Analysis
        </Button>
      </div>

      <ProcessingDialog
        open={showProcessing}
        onOpenChange={(open) => {
          if (!open) stopPolling();
          setShowProcessing(open);
        }}
        steps={dialogSteps}
        finalStatus={finalStatus}
        outputPath={reportPath}
        title="Running Analysis"
      />
    </div>
  );
}
