import { useState, useEffect, useCallback } from "react";
import { RefreshCw, ChevronDown, ChevronRight, AlertCircle, CheckCircle2, PlayCircle, Clock, Inbox } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { cn } from "@/lib/utils";
import { api, type UserLogEntry, type ErrorLogResponse } from "@/lib/api";

// ── Helpers ────────────────────────────────────────────────────────────────

const ACTION_LABELS: Record<string, string> = {
  ppt_start:               "Report generation started",
  ppt_complete:            "Report generation completed",
  reconciliation_start:    "Reconciliation started",
  reconciliation_complete: "Reconciliation completed",
};

const MODULE_LABELS: Record<string, string> = {
  ppt_automation: "Report",
  reconciliation: "Reconciliation",
};

function formatDate(iso: string) {
  const d = new Date(iso);
  return d.toLocaleString("en-IN", {
    day: "2-digit", month: "short", year: "numeric",
    hour: "2-digit", minute: "2-digit", hour12: true,
  });
}

function ActionIcon({ action }: { action: string }) {
  if (action.endsWith("_complete"))
    return <CheckCircle2 className="h-4 w-4 text-green-500 shrink-0" />;
  if (action.endsWith("_start"))
    return <PlayCircle className="h-4 w-4 text-primary shrink-0" />;
  return <Clock className="h-4 w-4 text-muted-foreground shrink-0" />;
}

// ── Run History tab ────────────────────────────────────────────────────────

function RunHistoryTab({ logs, loading, onRefresh }: {
  logs: UserLogEntry[];
  loading: boolean;
  onRefresh: () => void;
}) {
  return (
    <div className="flex flex-col h-full">
      <div className="flex items-center justify-between px-6 py-3 border-b border-border shrink-0">
        <p className="text-xs text-muted-foreground">{logs.length} entr{logs.length === 1 ? "y" : "ies"}</p>
        <Button variant="ghost" size="sm" onClick={onRefresh} disabled={loading}>
          <RefreshCw className={cn("h-3.5 w-3.5 mr-1.5", loading && "animate-spin")} />
          Refresh
        </Button>
      </div>

      <div className="flex-1 overflow-y-auto px-6 py-4">
        {loading && logs.length === 0 ? (
          <div className="flex items-center justify-center h-32 text-muted-foreground text-sm">
            Loading…
          </div>
        ) : logs.length === 0 ? (
          <EmptyState message="No run history yet. Runs will appear here after you execute an analysis." />
        ) : (
          <div className="space-y-2">
            {logs.map((log) => (
              <div
                key={log.id}
                className="flex items-start gap-3 p-3 rounded-lg border border-border bg-card hover:bg-muted/40 transition-colors"
              >
                <ActionIcon action={log.action} />

                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-2 flex-wrap">
                    <span className="text-sm font-medium text-foreground">
                      {ACTION_LABELS[log.action] ?? log.action}
                    </span>
                    <Badge variant="outline" className="text-xs">
                      {MODULE_LABELS[log.module] ?? log.module}
                    </Badge>
                  </div>

                  {log.details && (
                    <p className="text-xs text-muted-foreground mt-0.5 truncate">{log.details}</p>
                  )}

                  <div className="flex items-center gap-3 mt-1">
                    <span className="text-xs text-muted-foreground">{formatDate(log.created_at)}</span>
                    {log.run_id && (
                      <span className="text-xs text-muted-foreground font-mono">
                        #{log.run_id.slice(0, 8)}
                      </span>
                    )}
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

// ── Error Logs tab ─────────────────────────────────────────────────────────

function ErrorLogsTab({ logs, loading, onRefresh }: {
  logs: ErrorLogResponse[];
  loading: boolean;
  onRefresh: () => void;
}) {
  const [expanded, setExpanded] = useState<string | null>(null);

  return (
    <div className="flex flex-col h-full">
      <div className="flex items-center justify-between px-6 py-3 border-b border-border shrink-0">
        <p className="text-xs text-muted-foreground">
          {logs.length} error{logs.length !== 1 ? "s" : ""} (last 50)
        </p>
        <Button variant="ghost" size="sm" onClick={onRefresh} disabled={loading}>
          <RefreshCw className={cn("h-3.5 w-3.5 mr-1.5", loading && "animate-spin")} />
          Refresh
        </Button>
      </div>

      <div className="flex-1 overflow-y-auto px-6 py-4">
        {loading && logs.length === 0 ? (
          <div className="flex items-center justify-center h-32 text-muted-foreground text-sm">
            Loading…
          </div>
        ) : logs.length === 0 ? (
          <EmptyState message="No errors logged. Any script or server errors will appear here." />
        ) : (
          <div className="space-y-2">
            {logs.map((log) => {
              const isOpen = expanded === log.id;
              return (
                <div
                  key={log.id}
                  className="rounded-lg border border-destructive/30 bg-destructive/5 overflow-hidden"
                >
                  {/* Header row */}
                  <div className="flex items-start gap-3 p-3">
                    <AlertCircle className="h-4 w-4 text-destructive shrink-0 mt-0.5" />

                    <div className="flex-1 min-w-0">
                      <div className="flex items-center gap-2 flex-wrap">
                        <Badge variant="destructive" className="text-xs">
                          {MODULE_LABELS[log.module] ?? log.module}
                        </Badge>
                        <span className="text-xs text-muted-foreground">{formatDate(log.created_at)}</span>
                      </div>
                      <p className="text-sm text-foreground mt-1 break-words">{log.error_message}</p>
                    </div>

                    {log.stack_trace && (
                      <button
                        onClick={() => setExpanded(isOpen ? null : log.id)}
                        className="text-muted-foreground hover:text-foreground transition-colors shrink-0 mt-0.5"
                        title={isOpen ? "Hide trace" : "Show trace"}
                      >
                        {isOpen
                          ? <ChevronDown className="h-4 w-4" />
                          : <ChevronRight className="h-4 w-4" />}
                      </button>
                    )}
                  </div>

                  {/* Collapsible stack trace */}
                  {isOpen && log.stack_trace && (
                    <div className="border-t border-destructive/20 px-3 py-2 bg-muted/60">
                      <pre className="text-xs text-muted-foreground whitespace-pre-wrap break-words font-mono leading-relaxed">
                        {log.stack_trace}
                      </pre>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

// ── Empty state ────────────────────────────────────────────────────────────

function EmptyState({ message }: { message: string }) {
  return (
    <div className="flex flex-col items-center justify-center h-40 gap-3 text-center px-6">
      <Inbox className="h-8 w-8 text-muted-foreground/40" />
      <p className="text-sm text-muted-foreground">{message}</p>
    </div>
  );
}

// ── Main component ─────────────────────────────────────────────────────────

export function HistoryPanel() {
  const [runLogs, setRunLogs]       = useState<UserLogEntry[]>([]);
  const [errorLogs, setErrorLogs]   = useState<ErrorLogResponse[]>([]);
  const [loadingRuns, setLoadingRuns]     = useState(false);
  const [loadingErrors, setLoadingErrors] = useState(false);

  const userId: string | null = (() => {
    try {
      const stored = localStorage.getItem("user");
      return stored ? (JSON.parse(stored) as { id?: string }).id ?? null : null;
    } catch { return null; }
  })();

  const fetchRuns = useCallback(async () => {
    if (!userId) return;
    setLoadingRuns(true);
    try {
      const logs = await api.getUserLogs(userId);
      setRunLogs(logs);
    } catch {
      // silently fail — user may not be logged in yet
    } finally {
      setLoadingRuns(false);
    }
  }, [userId]);

  const fetchErrors = useCallback(async () => {
    setLoadingErrors(true);
    try {
      const logs = await api.getErrorLogs();
      setErrorLogs(logs);
    } catch {
      // silently fail
    } finally {
      setLoadingErrors(false);
    }
  }, []);

  useEffect(() => {
    fetchRuns();
    fetchErrors();
  }, [fetchRuns, fetchErrors]);

  return (
    <div className="h-full flex flex-col">
      <Tabs defaultValue="runs" className="flex-1 flex flex-col overflow-hidden">
        <div className="px-6 pt-4 shrink-0">
          <TabsList className="w-full grid grid-cols-2">
            <TabsTrigger value="runs">
              Run History
              {runLogs.length > 0 && (
                <Badge variant="secondary" className="ml-2 text-xs px-1.5 py-0 h-4">
                  {runLogs.length}
                </Badge>
              )}
            </TabsTrigger>
            <TabsTrigger value="errors">
              Error Logs
              {errorLogs.length > 0 && (
                <Badge variant="destructive" className="ml-2 text-xs px-1.5 py-0 h-4">
                  {errorLogs.length}
                </Badge>
              )}
            </TabsTrigger>
          </TabsList>
        </div>

        <TabsContent value="runs" className="flex-1 overflow-hidden mt-0 data-[state=active]:flex data-[state=active]:flex-col">
          <RunHistoryTab logs={runLogs} loading={loadingRuns} onRefresh={fetchRuns} />
        </TabsContent>

        <TabsContent value="errors" className="flex-1 overflow-hidden mt-0 data-[state=active]:flex data-[state=active]:flex-col">
          <ErrorLogsTab logs={errorLogs} loading={loadingErrors} onRefresh={fetchErrors} />
        </TabsContent>
      </Tabs>
    </div>
  );
}
