import { Check, Circle, Loader2 } from "lucide-react";
import { Progress } from "@/components/ui/progress";
import { cn } from "@/lib/utils";

export interface AuditStep {
  id: string;
  label: string;
  status: "pending" | "running" | "complete" | "error";
}

interface AuditProgressProps {
  steps: AuditStep[];
  progress: number;
  isRunning: boolean;
}

export function AuditProgress({ steps, progress, isRunning }: AuditProgressProps) {
  if (!isRunning && progress === 0) {
    return null;
  }

  return (
    <div className="space-y-4 p-4 bg-muted/30 rounded border border-border">
      <div className="flex items-center justify-between">
        <h3 className="text-sm font-medium">Audit Progress</h3>
        <span className="text-sm text-muted-foreground">{Math.round(progress)}%</span>
      </div>

      <Progress value={progress} className="h-2" />

      <div className="space-y-2">
        {steps.map((step) => (
          <div key={step.id} className="flex items-center gap-3">
            <div
              className={cn(
                "w-5 h-5 rounded-full flex items-center justify-center shrink-0",
                step.status === "complete" && "bg-emerald-500",
                step.status === "running" && "bg-accent",
                step.status === "pending" && "bg-muted border border-border",
                step.status === "error" && "bg-destructive"
              )}
            >
              {step.status === "complete" && (
                <Check className="h-3 w-3 text-white" />
              )}
              {step.status === "running" && (
                <Loader2 className="h-3 w-3 text-white animate-spin" />
              )}
              {step.status === "pending" && (
                <Circle className="h-2 w-2 text-muted-foreground" />
              )}
              {step.status === "error" && (
                <span className="text-white text-xs">!</span>
              )}
            </div>
            <span
              className={cn(
                "text-sm",
                step.status === "complete" && "text-foreground",
                step.status === "running" && "text-foreground font-medium",
                step.status === "pending" && "text-muted-foreground",
                step.status === "error" && "text-destructive"
              )}
            >
              {step.label}
            </span>
          </div>
        ))}
      </div>
    </div>
  );
}
