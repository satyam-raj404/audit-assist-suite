import { useState, useCallback, useEffect } from "react";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogDescription,
} from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { CheckCircle2, XCircle, Loader2 } from "lucide-react";
import { cn } from "@/lib/utils";

export interface ProcessingStep {
  id: string;
  label: string;
  status: "pending" | "running" | "complete" | "error";
}

const defaultSteps: ProcessingStep[] = [
  { id: "parse", label: "Parsing tracker file", status: "pending" },
  { id: "load", label: "Loading Templates", status: "pending" },
  { id: "populate", label: "Populating Data", status: "pending" },
  { id: "risk", label: "Tagging Risks", status: "pending" },
  { id: "pages", label: "Updating Page numbers", status: "pending" },
  { id: "output", label: "Final Output", status: "pending" },
];

interface ProcessingDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  onComplete: (success: boolean) => void;
  title?: string;
}

export function ProcessingDialog({
  open,
  onOpenChange,
  onComplete,
  title = "Generating Document",
}: ProcessingDialogProps) {
  const [steps, setSteps] = useState<ProcessingStep[]>(defaultSteps);
  const [finalStatus, setFinalStatus] = useState<"running" | "success" | "error">("running");

  const runProcess = useCallback(async () => {
    setSteps(defaultSteps.map((s) => ({ ...s, status: "pending" })));
    setFinalStatus("running");

    const durations = [700, 900, 1200, 800, 600, 500];

    for (let i = 0; i < defaultSteps.length; i++) {
      setSteps((prev) =>
        prev.map((s, idx) => (idx === i ? { ...s, status: "running" } : s))
      );

      await new Promise((r) => setTimeout(r, durations[i]));

      // Simulate a random error (~10% chance on step 3+)
      const hasError = i >= 3 && Math.random() < 0.1;

      if (hasError) {
        setSteps((prev) =>
          prev.map((s, idx) => (idx === i ? { ...s, status: "error" } : s))
        );
        setFinalStatus("error");
        onComplete(false);
        return;
      }

      setSteps((prev) =>
        prev.map((s, idx) => (idx === i ? { ...s, status: "complete" } : s))
      );
    }

    setFinalStatus("success");
    onComplete(true);
  }, [onComplete]);

  useEffect(() => {
    if (open) {
      runProcess();
    }
  }, [open, runProcess]);

  const statusIcon = (status: ProcessingStep["status"]) => {
    switch (status) {
      case "complete":
        return <CheckCircle2 className="h-5 w-5 text-green-600" />;
      case "running":
        return <Loader2 className="h-5 w-5 text-primary animate-spin" />;
      case "error":
        return <XCircle className="h-5 w-5 text-destructive" />;
      default:
        return <div className="h-5 w-5 rounded-full border-2 border-muted-foreground/30" />;
    }
  };

  return (
    <Dialog open={open} onOpenChange={finalStatus !== "running" ? onOpenChange : undefined}>
      <DialogContent className="sm:max-w-md" onPointerDownOutside={(e) => { if (finalStatus === "running") e.preventDefault(); }}>
        <DialogHeader>
          <DialogTitle>{title}</DialogTitle>
          <DialogDescription>
            {finalStatus === "running"
              ? "Processing your request, please wait..."
              : finalStatus === "success"
              ? "Document generated successfully!"
              : "An error occurred during processing."}
          </DialogDescription>
        </DialogHeader>

        <div className="space-y-3 py-4">
          {steps.map((step) => (
            <div
              key={step.id}
              className={cn(
                "flex items-center gap-3 px-3 py-2 rounded-lg transition-colors",
                step.status === "running" && "bg-primary/5",
                step.status === "complete" && "bg-green-50",
                step.status === "error" && "bg-destructive/5"
              )}
            >
              {statusIcon(step.status)}
              <span
                className={cn(
                  "text-sm",
                  step.status === "pending" && "text-muted-foreground",
                  step.status === "running" && "text-foreground font-medium",
                  step.status === "complete" && "text-green-700",
                  step.status === "error" && "text-destructive font-medium"
                )}
              >
                {step.label}
              </span>
            </div>
          ))}
        </div>

        {/* Final status banner */}
        {finalStatus === "success" && (
          <div className="flex items-center gap-3 p-4 rounded-lg bg-green-50 border border-green-200">
            <CheckCircle2 className="h-6 w-6 text-green-600" />
            <div>
              <p className="text-sm font-semibold text-green-800">Success!</p>
              <p className="text-xs text-green-600">Your document has been generated successfully.</p>
            </div>
          </div>
        )}

        {finalStatus === "error" && (
          <div className="flex items-center gap-3 p-4 rounded-lg bg-destructive/5 border border-destructive/20">
            <XCircle className="h-6 w-6 text-destructive" />
            <div>
              <p className="text-sm font-semibold text-destructive">Error</p>
              <p className="text-xs text-destructive/80">Processing failed. Please try again.</p>
            </div>
          </div>
        )}

        {finalStatus !== "running" && (
          <div className="flex justify-end pt-2">
            <Button variant="kpmg" onClick={() => onOpenChange(false)}>
              Close
            </Button>
          </div>
        )}
      </DialogContent>
    </Dialog>
  );
}
