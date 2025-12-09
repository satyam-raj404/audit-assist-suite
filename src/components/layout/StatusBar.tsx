import { Circle, Clock, CheckCircle2 } from "lucide-react";
import { cn } from "@/lib/utils";

interface StatusBarProps {
  status: "idle" | "running" | "success" | "error";
  message: string;
  timestamp?: string;
}

export function StatusBar({ status, message, timestamp }: StatusBarProps) {
  const getStatusConfig = () => {
    switch (status) {
      case "running":
        return {
          icon: Circle,
          className: "text-accent",
          iconClassName: "animate-pulse",
          label: "Processing",
        };
      case "success":
        return {
          icon: CheckCircle2,
          className: "text-emerald-600",
          iconClassName: "",
          label: "Complete",
        };
      case "error":
        return {
          icon: Circle,
          className: "text-destructive",
          iconClassName: "",
          label: "Error",
        };
      default:
        return {
          icon: Circle,
          className: "text-muted-foreground",
          iconClassName: "",
          label: "Ready",
        };
    }
  };

  const config = getStatusConfig();
  const StatusIcon = config.icon;

  return (
    <footer className="h-8 bg-muted border-t border-border flex items-center justify-between px-4 shrink-0">
      <div className="flex items-center gap-3">
        <div className={cn("flex items-center gap-2", config.className)}>
          <StatusIcon className={cn("h-3 w-3", config.iconClassName)} />
          <span className="text-xs font-medium">{config.label}</span>
        </div>
        <span className="text-xs text-muted-foreground">{message}</span>
      </div>

      {timestamp && (
        <div className="flex items-center gap-1.5 text-muted-foreground">
          <Clock className="h-3 w-3" />
          <span className="text-xs">{timestamp}</span>
        </div>
      )}
    </footer>
  );
}
