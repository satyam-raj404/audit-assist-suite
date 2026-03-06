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

interface ProcessingDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  steps: ProcessingStep[];
  finalStatus: "running" | "success" | "error";
  outputPath?: string;
  title?: string;
}

export function ProcessingDialog({
  open,
  onOpenChange,
  steps,
  finalStatus,
  outputPath,
  title = "Generating Document",
}: ProcessingDialogProps) {
  const isDone = finalStatus !== "running";

  const stepIcon = (status: ProcessingStep["status"]) => {
    switch (status) {
      case "complete": return <CheckCircle2 className="h-5 w-5 text-green-600" />;
      case "running":  return <Loader2 className="h-5 w-5 text-primary animate-spin" />;
      case "error":    return <XCircle className="h-5 w-5 text-destructive" />;
      default:         return <div className="h-5 w-5 rounded-full border-2 border-muted-foreground/30" />;
    }
  };

  return (
    <Dialog open={open} onOpenChange={isDone ? onOpenChange : undefined}>
      <DialogContent
        className="sm:max-w-md"
        onPointerDownOutside={(e) => { if (!isDone) e.preventDefault(); }}
      >
        <DialogHeader>
          <DialogTitle>{title}</DialogTitle>
          <DialogDescription>
            {finalStatus === "running"
              ? "Processing your request, please wait…"
              : finalStatus === "success"
              ? "Completed successfully!"
              : "An error occurred during processing."}
          </DialogDescription>
        </DialogHeader>

        <div className="space-y-3 py-4">
          {steps.map((step) => (
            <div
              key={step.id}
              className={cn(
                "flex items-center gap-3 px-3 py-2 rounded-lg transition-colors",
                step.status === "running"  && "bg-primary/5",
                step.status === "complete" && "bg-green-50",
                step.status === "error"    && "bg-destructive/5"
              )}
            >
              {stepIcon(step.status)}
              <span
                className={cn(
                  "text-sm",
                  step.status === "pending"  && "text-muted-foreground",
                  step.status === "running"  && "text-foreground font-medium",
                  step.status === "complete" && "text-green-700",
                  step.status === "error"    && "text-destructive font-medium"
                )}
              >
                {step.label}
              </span>
            </div>
          ))}
        </div>

        {finalStatus === "success" && (
          <div className="flex items-start gap-3 p-4 rounded-lg bg-green-50 border border-green-200">
            <CheckCircle2 className="h-6 w-6 text-green-600 shrink-0 mt-0.5" />
            <div>
              <p className="text-sm font-semibold text-green-800">Success!</p>
              {outputPath && (
                <p className="text-xs text-green-700 mt-1 break-all">Saved to: {outputPath}</p>
              )}
            </div>
          </div>
        )}

        {finalStatus === "error" && (
          <div className="flex items-center gap-3 p-4 rounded-lg bg-destructive/5 border border-destructive/20">
            <XCircle className="h-6 w-6 text-destructive" />
            <div>
              <p className="text-sm font-semibold text-destructive">Error</p>
              {outputPath && (
                <p className="text-xs text-destructive/80 mt-1 break-all">{outputPath}</p>
              )}
            </div>
          </div>
        )}

        {isDone && (
          <div className="flex justify-end pt-2">
            <Button variant="kpmg" onClick={() => onOpenChange(false)}>Close</Button>
          </div>
        )}
      </DialogContent>
    </Dialog>
  );
}
