import { useState, useCallback } from "react";
import { Play, Square, RotateCcw } from "lucide-react";
import { Button } from "@/components/ui/button";
import { FileUpload } from "./FileUpload";
import { AuditTypeSelector } from "./AuditTypeSelector";
import { TemplateSelector } from "./TemplateSelector";
import { OutputFolderSelector } from "./OutputFolderSelector";
import { AuditProgress, type AuditStep } from "./AuditProgress";
import { WorkflowIllustration } from "./WorkflowIllustration";
import { ProcessingDialog } from "@/components/shared/ProcessingDialog";
import { toast } from "sonner";

interface UploadedFile {
  id: string;
  name: string;
  size: number;
  type: string;
}

const initialSteps: AuditStep[] = [
  { id: "validate", label: "Validating input files", status: "pending" },
  { id: "parse", label: "Parsing document content", status: "pending" },
  { id: "analyze", label: "Running audit analysis", status: "pending" },
  { id: "generate", label: "Generating report", status: "pending" },
  { id: "save", label: "Saving to output folder", status: "pending" },
];

interface AuditPanelProps {
  onStatusChange: (status: "idle" | "running" | "success" | "error", message: string) => void;
}

export function AuditPanel({ onStatusChange }: AuditPanelProps) {
  const [files, setFiles] = useState<UploadedFile[]>([]);
  const [auditType, setAuditType] = useState("");
  const [subAuditSector, setSubAuditSector] = useState("");
  const [selectedTemplate, setSelectedTemplate] = useState("");
  const [outputPath, setOutputPath] = useState("");
  const [isRunning, setIsRunning] = useState(false);
  const [progress, setProgress] = useState(0);
  const [steps, setSteps] = useState<AuditStep[]>(initialSteps);
  const [showProcessing, setShowProcessing] = useState(false);

  const canRunAudit = files.length > 0 && auditType && subAuditSector && selectedTemplate && outputPath;

  const simulateAudit = useCallback(async () => {
    setIsRunning(true);
    setProgress(0);
    setSteps(initialSteps);
    onStatusChange("running", "Processing audit...");

    const stepDurations = [800, 1200, 2000, 1500, 600];

    for (let i = 0; i < steps.length; i++) {
      setSteps((prev) =>
        prev.map((s, idx) =>
          idx === i ? { ...s, status: "running" } : s
        )
      );

      const startProgress = (i / steps.length) * 100;
      const endProgress = ((i + 1) / steps.length) * 100;
      const duration = stepDurations[i];
      const intervalTime = 50;
      const increments = duration / intervalTime;
      const progressIncrement = (endProgress - startProgress) / increments;

      for (let j = 0; j < increments; j++) {
        await new Promise((resolve) => setTimeout(resolve, intervalTime));
        setProgress((prev) => Math.min(prev + progressIncrement, endProgress));
      }

      setSteps((prev) =>
        prev.map((s, idx) =>
          idx === i ? { ...s, status: "complete" } : s
        )
      );
    }

    setIsRunning(false);
    setProgress(100);
    onStatusChange("success", "Audit completed successfully");
    toast.success("Audit Complete", {
      description: `Report saved to ${outputPath}`,
    });
  }, [onStatusChange, outputPath]);

  const handleStop = () => {
    setIsRunning(false);
    onStatusChange("idle", "Audit cancelled by user");
    toast.info("Audit Cancelled");
  };

  const handleReset = () => {
    setFiles([]);
    setAuditType("");
    setSubAuditSector("");
    setSelectedTemplate("");
    setOutputPath("");
    setProgress(0);
    setSteps(initialSteps);
    onStatusChange("idle", "Ready to run audit");
  };

  return (
    <div className="h-full flex flex-col">
      <div className="flex-1 overflow-y-auto p-6 space-y-6">
        {/* Workflow Illustration */}
        <WorkflowIllustration />

        <div className="border-t border-border pt-6">
          <AuditTypeSelector
            auditType={auditType}
            subAuditSector={subAuditSector}
            onAuditTypeChange={(value) => {
              setAuditType(value);
              setSubAuditSector("");
            }}
            onSubAuditSectorChange={setSubAuditSector}
          />
        </div>

        <div className="border-t border-border pt-6">
          <TemplateSelector
            selectedTemplate={selectedTemplate}
            onTemplateChange={setSelectedTemplate}
          />
        </div>

        {/* Upload Files moved AFTER Select Template */}
        <div className="border-t border-border pt-6">
          <FileUpload onFilesChange={setFiles} />
        </div>

        <div className="border-t border-border pt-6">
          <OutputFolderSelector
            outputPath={outputPath}
            onOutputPathChange={setOutputPath}
          />
        </div>

        <AuditProgress steps={steps} progress={progress} isRunning={isRunning} />
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
          {isRunning ? (
            <Button variant="destructive" onClick={handleStop}>
              <Square className="h-4 w-4 mr-2" />
              Stop
            </Button>
          ) : (
            <Button
              variant="kpmg"
              onClick={() => setShowProcessing(true)}
              disabled={!canRunAudit}
            >
              <Play className="h-4 w-4 mr-2" />
              Run
            </Button>
          )}
        </div>
      </div>

      <ProcessingDialog
        open={showProcessing}
        onOpenChange={setShowProcessing}
        onComplete={(success) => {
          if (success) {
            onStatusChange("success", "Audit completed successfully");
            toast.success("Audit Complete", { description: `Report saved to ${outputPath}` });
          } else {
            onStatusChange("error", "Audit failed");
            toast.error("Audit Failed", { description: "An error occurred during processing." });
          }
        }}
        title="Generating Audit Report"
      />
    </div>
  );
}
