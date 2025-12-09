import { useState, useCallback } from "react";
import { Upload, X, FileSpreadsheet, ArrowLeftRight } from "lucide-react";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";

interface ReconciliationFile {
  id: string;
  name: string;
  size: number;
}

interface ReconciliationSectionProps {
  onFilesChange: (file1: ReconciliationFile | null, file2: ReconciliationFile | null) => void;
}

const formatFileSize = (bytes: number) => {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
};

export function ReconciliationSection({ onFilesChange }: ReconciliationSectionProps) {
  const [file1, setFile1] = useState<ReconciliationFile | null>(null);
  const [file2, setFile2] = useState<ReconciliationFile | null>(null);

  const handleFileUpload = useCallback(
    (fileNum: 1 | 2) => (e: React.ChangeEvent<HTMLInputElement>) => {
      if (!e.target.files || !e.target.files[0]) return;

      const file = e.target.files[0];
      const ext = file.name.toLowerCase().split(".").pop();
      
      if (!["xlsx", "xls", "csv"].includes(ext || "")) {
        return;
      }

      const newFile: ReconciliationFile = {
        id: crypto.randomUUID(),
        name: file.name,
        size: file.size,
      };

      if (fileNum === 1) {
        setFile1(newFile);
        onFilesChange(newFile, file2);
      } else {
        setFile2(newFile);
        onFilesChange(file1, newFile);
      }
    },
    [file1, file2, onFilesChange]
  );

  const removeFile = useCallback(
    (fileNum: 1 | 2) => {
      if (fileNum === 1) {
        setFile1(null);
        onFilesChange(null, file2);
      } else {
        setFile2(null);
        onFilesChange(file1, null);
      }
    },
    [file1, file2, onFilesChange]
  );

  const renderFileSlot = (
    fileNum: 1 | 2,
    file: ReconciliationFile | null,
    label: string
  ) => (
    <div className="flex-1">
      <p className="text-xs font-medium text-muted-foreground mb-2">{label}</p>
      {file ? (
        <div className="flex items-center gap-3 p-3 bg-muted/50 rounded border border-border">
          <FileSpreadsheet className="h-5 w-5 text-primary shrink-0" />
          <div className="flex-1 min-w-0">
            <p className="text-sm font-medium truncate">{file.name}</p>
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
            "flex flex-col items-center justify-center p-4 border-2 border-dashed rounded cursor-pointer transition-colors",
            "border-border hover:border-muted-foreground/50 hover:bg-muted/30"
          )}
        >
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={handleFileUpload(fileNum)}
            className="sr-only"
          />
          <Upload className="h-6 w-6 text-muted-foreground mb-2" />
          <span className="text-sm text-muted-foreground">Upload Excel</span>
          <span className="text-xs text-muted-foreground/70 mt-1">
            XLSX, XLS, CSV
          </span>
        </label>
      )}
    </div>
  );

  return (
    <div className="space-y-4">
      <h3 className="kpmg-section-title">Reconciliation - Compare Excel Files</h3>

      <div className="flex items-stretch gap-4">
        {renderFileSlot(1, file1, "Source File")}

        <div className="flex items-center justify-center px-2">
          <div className="w-10 h-10 rounded-full bg-primary/10 flex items-center justify-center">
            <ArrowLeftRight className="h-5 w-5 text-primary" />
          </div>
        </div>

        {renderFileSlot(2, file2, "Target File")}
      </div>

      <p className="text-xs text-muted-foreground">
        Upload two Excel files to compare and identify discrepancies between datasets.
      </p>
    </div>
  );
}
