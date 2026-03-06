import { Upload, X, FileSpreadsheet } from "lucide-react";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";

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

interface ReconciliationFileSlotProps {
  file: ReconciliationFile | null;
  label: string;
  onUpload: (e: React.ChangeEvent<HTMLInputElement>) => void;
  onRemove: () => void;
}

export function ReconciliationFileSlot({ file, label, onUpload, onRemove }: ReconciliationFileSlotProps) {
  return (
    <div className="flex-1">
      <p className="text-xs font-medium text-muted-foreground mb-2">{label}</p>
      {file ? (
        <div className="flex items-center gap-3 p-4 bg-card rounded border border-border">
          <FileSpreadsheet className="h-6 w-6 text-primary shrink-0" />
          <div className="flex-1 min-w-0">
            <p className="text-sm font-medium truncate text-foreground">{file.name}</p>
            <p className="text-xs text-muted-foreground">{formatFileSize(file.size)}</p>
          </div>
          <Button
            variant="ghost"
            size="icon"
            className="h-8 w-8 text-muted-foreground hover:text-destructive"
            onClick={onRemove}
            aria-label={`Remove ${file.name}`}
          >
            <X className="h-4 w-4" />
          </Button>
        </div>
      ) : (
        <label
          className={cn(
            "flex flex-col items-center justify-center p-6 border-2 border-dashed rounded cursor-pointer transition-colors",
            "border-border hover:border-primary/50 hover:bg-primary/5"
          )}
        >
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={onUpload}
            className="sr-only"
          />
          <Upload className="h-8 w-8 text-muted-foreground mb-3" />
          <span className="text-sm text-foreground font-medium">Upload Excel File</span>
          <span className="text-xs text-muted-foreground mt-1">XLSX, XLS, CSV</span>
        </label>
      )}
    </div>
  );
}
