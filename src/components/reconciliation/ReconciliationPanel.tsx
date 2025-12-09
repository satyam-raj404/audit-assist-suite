import { useState, useCallback } from "react";
import { Upload, X, FileSpreadsheet, ArrowLeftRight, Play, RotateCcw, CheckCircle2, AlertCircle } from "lucide-react";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";
import { Progress } from "@/components/ui/progress";
import { toast } from "sonner";

interface ReconciliationFile {
  id: string;
  name: string;
  size: number;
}

const formatFileSize = (bytes: number) => {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
};

interface ReconciliationPanelProps {
  onStatusChange: (status: "idle" | "running" | "success" | "error", message: string) => void;
}

export function ReconciliationPanel({ onStatusChange }: ReconciliationPanelProps) {
  const [file1, setFile1] = useState<ReconciliationFile | null>(null);
  const [file2, setFile2] = useState<ReconciliationFile | null>(null);
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

      if (fileNum === 1) {
        setFile1(newFile);
      } else {
        setFile2(newFile);
      }
      setResults(null);
    },
    []
  );

  const removeFile = useCallback((fileNum: 1 | 2) => {
    if (fileNum === 1) {
      setFile1(null);
    } else {
      setFile2(null);
    }
    setResults(null);
  }, []);

  const handleReset = () => {
    setFile1(null);
    setFile2(null);
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

    // Simulate reconciliation process
    for (let i = 0; i <= 100; i += 2) {
      await new Promise((resolve) => setTimeout(resolve, 40));
      setProgress(i);
    }

    // Simulated results
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

  const canRun = file1 && file2 && !isRunning;

  const renderFileSlot = (
    fileNum: 1 | 2,
    file: ReconciliationFile | null,
    label: string
  ) => (
    <div className="flex-1">
      <p className="text-xs font-medium text-muted-foreground mb-2">{label}</p>
      {file ? (
        <div className="flex items-center gap-3 p-4 bg-card rounded border border-border">
          <FileSpreadsheet className="h-6 w-6 text-accent shrink-0" />
          <div className="flex-1 min-w-0">
            <p className="text-sm font-medium truncate text-foreground">{file.name}</p>
            <p className="text-xs text-muted-foreground">
              {formatFileSize(file.size)}
            </p>
          </div>
          <Button
            variant="ghost"
            size="icon"
            className="h-8 w-8 text-muted-foreground hover:text-destructive"
            onClick={() => removeFile(fileNum)}
            aria-label={`Remove ${file.name}`}
          >
            <X className="h-4 w-4" />
          </Button>
        </div>
      ) : (
        <label
          className={cn(
            "flex flex-col items-center justify-center p-6 border-2 border-dashed rounded cursor-pointer transition-colors",
            "border-border hover:border-accent/50 hover:bg-accent/5"
          )}
        >
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={handleFileUpload(fileNum)}
            className="sr-only"
          />
          <Upload className="h-8 w-8 text-muted-foreground mb-3" />
          <span className="text-sm text-foreground font-medium">Upload Excel File</span>
          <span className="text-xs text-muted-foreground mt-1">
            XLSX, XLS, CSV
          </span>
        </label>
      )}
    </div>
  );

  return (
    <div className="h-full flex flex-col">
      <div className="flex-1 overflow-y-auto p-6 space-y-6">
        {/* File Upload Section */}
        <div className="space-y-4">
          <h3 className="kpmg-section-title">Select Files to Compare</h3>

          <div className="flex items-stretch gap-6">
            {renderFileSlot(1, file1, "Source File")}

            <div className="flex items-center justify-center px-4">
              <div className="w-12 h-12 rounded-full bg-accent/20 flex items-center justify-center border border-accent/30">
                <ArrowLeftRight className="h-6 w-6 text-accent" />
              </div>
            </div>

            {renderFileSlot(2, file2, "Target File")}
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

        {/* Results Section */}
        {results && (
          <div className="space-y-4">
            <h3 className="kpmg-section-title">Reconciliation Results</h3>
            
            <div className="grid grid-cols-3 gap-4">
              <div className="p-4 bg-card rounded border border-border">
                <div className="flex items-center gap-3 mb-2">
                  <CheckCircle2 className="h-5 w-5 text-emerald-400" />
                  <span className="text-sm text-muted-foreground">Matched Records</span>
                </div>
                <p className="text-2xl font-bold text-emerald-400">{results.matched}</p>
              </div>

              <div className="p-4 bg-card rounded border border-border">
                <div className="flex items-center gap-3 mb-2">
                  <AlertCircle className="h-5 w-5 text-amber-400" />
                  <span className="text-sm text-muted-foreground">Mismatched</span>
                </div>
                <p className="text-2xl font-bold text-amber-400">{results.mismatched}</p>
              </div>

              <div className="p-4 bg-card rounded border border-border">
                <div className="flex items-center gap-3 mb-2">
                  <AlertCircle className="h-5 w-5 text-red-400" />
                  <span className="text-sm text-muted-foreground">Missing</span>
                </div>
                <p className="text-2xl font-bold text-red-400">{results.missing}</p>
              </div>
            </div>

            <div className="p-4 bg-card rounded border border-border">
              <p className="text-sm text-muted-foreground mb-3">Summary</p>
              <p className="text-sm text-foreground">
                Comparison of <span className="text-accent font-medium">{file1?.name}</span> and{" "}
                <span className="text-accent font-medium">{file2?.name}</span> completed.{" "}
                {results.mismatched + results.missing > 0 ? (
                  <span className="text-amber-400">
                    {results.mismatched + results.missing} discrepancies found that require attention.
                  </span>
                ) : (
                  <span className="text-emerald-400">All records matched successfully.</span>
                )}
              </p>
            </div>
          </div>
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
