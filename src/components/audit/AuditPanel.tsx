import { useState, useCallback, useRef } from "react";
import { Play, Square, RotateCcw } from "lucide-react";
import { Button } from "@/components/ui/button";
import { FileUpload, type UploadedFile } from "./FileUpload";
import { AuditTypeSelector } from "./AuditTypeSelector";
import { TemplateSelector } from "./TemplateSelector";
import { OutputFolderSelector } from "./OutputFolderSelector";
import { AuditProgress, type AuditStep } from "./AuditProgress";
import { WorkflowIllustration } from "./WorkflowIllustration";
import { ProcessingDialog, type ProcessingStep } from "@/components/shared/ProcessingDialog";
import { api } from "@/lib/api";
import { toast } from "sonner";

const POLL_INTERVAL_MS = 2000;

// Maps API step IDs to display steps
const STEPS_TEMPLATE: ProcessingStep[] = [
  { id: "parse",         label: "Parsing tracker file",        status: "pending" },
  { id: "validate",      label: "Validating input data",       status: "pending" },
  { id: "load_template", label: "Loading PPT template",        status: "pending" },
  { id: "populate",      label: "Populating template with data", status: "pending" },
  { id: "format",        label: "Applying formatting rules",   status: "pending" },
  { id: "save",          label: "Saving output file",          status: "pending" },
];

interface AuditPanelProps {
  onStatusChange: (status: "idle" | "running" | "success" | "error", message: string) => void;
}

export function AuditPanel({ onStatusChange }: AuditPanelProps) {
  // Selection state
  const [auditType, setAuditType]       = useState("");
  const [utilityType, setUtilityType]   = useState("");
  const [reportType, setReportType]     = useState("Both");
  const [month, setMonth]               = useState("");
  const [year, setYear]                 = useState("");

  // File state
  const [files, setFiles]               = useState<UploadedFile[]>([]);
  const [pptxTemplateId, setPptxTemplateId]     = useState("");
  const [pptxTemplateName, setPptxTemplateName] = useState("");
  const [pptxTemplatePath, setPptxTemplatePath] = useState("");
  const [outputPath, setOutputPath]     = useState("");

  // Run state
  const [showProcessing, setShowProcessing]   = useState(false);
  const [dialogSteps, setDialogSteps]         = useState<ProcessingStep[]>(STEPS_TEMPLATE);
  const [finalStatus, setFinalStatus]         = useState<"running" | "success" | "error">("running");
  const [reportPath, setReportPath]           = useState("");
  const pollRef = useRef<ReturnType<typeof setInterval> | null>(null);

  // Whether a PPTX template is required for the current selection
  const needsPptx = Boolean(auditType && utilityType);

  // Whether month + year are required (Concurrent Audit + Report)
  const needsMonthYear = auditType === "Concurrent Audit" && utilityType === "Report";

  const canRun =
    auditType &&
    utilityType &&
    files.length > 0 &&
    (!needsPptx || pptxTemplatePath) &&
    outputPath &&
    (!needsMonthYear || (month && year));

  const stopPolling = () => {
    if (pollRef.current) {
      clearInterval(pollRef.current);
      pollRef.current = null;
    }
  };

  const handleRun = useCallback(async () => {
    if (!canRun) return;

    setDialogSteps(STEPS_TEMPLATE);
    setFinalStatus("running");
    setReportPath("");
    setShowProcessing(true);
    onStatusChange("running", "Processing audit…");

    try {
      const result = await api.startAudit({
        audit_type: auditType,
        utility_type: utilityType,
        report_type: reportType,
        excel_file_id: files[0].id,
        pptx_path: pptxTemplatePath || undefined,
        month: month || undefined,
        year: year || undefined,
        output_path: outputPath,
      });

      const runId = result.id;

      pollRef.current = setInterval(async () => {
        try {
          const status = await api.getAuditStatus(runId);

          // Sync steps from API response
          if (status.steps?.length) {
            setDialogSteps(
              status.steps.map((s) => ({
                id: s.id,
                label: s.label,
                status: s.status as ProcessingStep["status"],
              }))
            );
          }

          if (status.status === "complete") {
            stopPolling();
            setFinalStatus("success");
            setReportPath(status.report_path || outputPath);
            onStatusChange("success", "Audit completed successfully");
            toast.success("Audit Complete", { description: `Saved to ${status.report_path || outputPath}` });
          } else if (status.status === "error") {
            stopPolling();
            setFinalStatus("error");
            setReportPath(status.report_path || "");
            onStatusChange("error", "Audit failed");
            toast.error("Audit Failed");
          }
        } catch {
          stopPolling();
          setFinalStatus("error");
          onStatusChange("error", "Failed to get status");
        }
      }, POLL_INTERVAL_MS);

    } catch (err) {
      setFinalStatus("error");
      onStatusChange("error", "Failed to start audit");
      toast.error("Failed to start audit", {
        description: err instanceof Error ? err.message : "Unknown error",
      });
    }
  }, [canRun, auditType, utilityType, reportType, files, pptxTemplatePath, month, year, outputPath, onStatusChange]);

  const handleStop = () => {
    stopPolling();
    setShowProcessing(false);
    onStatusChange("idle", "Audit cancelled by user");
    toast.info("Audit Cancelled");
  };

  const handleReset = () => {
    stopPolling();
    setFiles([]);
    setAuditType("");
    setUtilityType("");
    setReportType("Both");
    setMonth("");
    setYear("");
    setPptxTemplateId("");
    setPptxTemplateName("");
    setPptxTemplatePath("");
    setOutputPath("");
    setReportPath("");
    setDialogSteps(STEPS_TEMPLATE);
    setFinalStatus("running");
    onStatusChange("idle", "Ready to run audit");
  };

  return (
    <div className="h-full flex flex-col">
      <div className="flex-1 overflow-y-auto p-6 space-y-6">
        <WorkflowIllustration />

        <div className="border-t border-border pt-6">
          <AuditTypeSelector
            auditType={auditType}
            utilityType={utilityType}
            reportType={reportType}
            month={month}
            year={year}
            onAuditTypeChange={setAuditType}
            onUtilityTypeChange={setUtilityType}
            onReportTypeChange={setReportType}
            onMonthChange={setMonth}
            onYearChange={setYear}
          />
        </div>

        <div className="border-t border-border pt-6">
          <TemplateSelector
            auditType={auditType}
            utilityType={utilityType}
            selectedId={pptxTemplateId}
            selectedName={pptxTemplateName}
            onSelect={(id, name, path) => { setPptxTemplateId(id); setPptxTemplateName(name); setPptxTemplatePath(path); }}
          />
        </div>

        <div className="border-t border-border pt-6">
          <FileUpload onFilesChange={setFiles} />
        </div>

        <div className="border-t border-border pt-6">
          <OutputFolderSelector outputPath={outputPath} onOutputPathChange={setOutputPath} />
        </div>

        <AuditProgress
          steps={dialogSteps.map((s) => ({ id: s.id, label: s.label, status: s.status as AuditStep["status"] }))}
          progress={
            dialogSteps.filter((s) => s.status === "complete").length /
            Math.max(dialogSteps.length, 1) * 100
          }
          isRunning={finalStatus === "running" && showProcessing}
        />
      </div>

      <div className="shrink-0 p-4 border-t border-border bg-muted/30 flex items-center justify-between">
        <Button
          variant="ghost"
          onClick={handleReset}
          className="text-muted-foreground"
        >
          <RotateCcw className="h-4 w-4 mr-2" />
          Reset
        </Button>

        <div className="flex gap-3">
          {showProcessing && finalStatus === "running" ? (
            <Button variant="destructive" onClick={handleStop}>
              <Square className="h-4 w-4 mr-2" />
              Stop
            </Button>
          ) : (
            <Button
              variant="kpmg"
              onClick={handleRun}
              disabled={!canRun}
            >
              <Play className="h-4 w-4 mr-2" />
              Run
            </Button>
          )}
        </div>
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
        title="Generating Audit Report"
      />
    </div>
  );
}
