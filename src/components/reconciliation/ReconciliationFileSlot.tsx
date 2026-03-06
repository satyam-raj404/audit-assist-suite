import { useState } from "react";
import { Upload, X, FileSpreadsheet, Loader2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";
import { api } from "@/lib/api";
import { toast } from "sonner";

export interface ReconFile {
  id: string;   // server-assigned UUID
  name: string;
  size: number;
}

const formatFileSize = (bytes: number) => {
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
};

interface ReconciliationFileSlotProps {
  file: ReconFile | null;
  label: string;
  onFileChange: (file: ReconFile | null) => void;
}

export function ReconciliationFileSlot({ file, label, onFileChange }: ReconciliationFileSlotProps) {
  const [isUploading, setIsUploading] = useState(false);

  const handleFile = async (raw: File) => {
    const ext = raw.name.toLowerCase().split(".").pop();
    if (!["xlsx", "xls", "csv"].includes(ext || "")) {
      toast.error("Invalid file type", { description: "Please upload XLSX, XLS, or CSV." });
      return;
    }
    setIsUploading(true);
    try {
      const uploaded = await api.uploadFile(raw);
      onFileChange({ id: uploaded.id, name: uploaded.name, size: raw.size });
    } catch {
      toast.error(`Upload failed: ${raw.name}`);
    } finally {
      setIsUploading(false);
    }
  };

  const handleInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files?.[0]) handleFile(e.target.files[0]);
    e.target.value = "";
  };

  const handleRemove = async () => {
    if (file) {
      try { await api.deleteFile(file.id); } catch { /* best-effort */ }
    }
    onFileChange(null);
  };

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
            onClick={handleRemove}
            aria-label={`Remove ${file.name}`}
          >
            <X className="h-4 w-4" />
          </Button>
        </div>
      ) : (
        <label
          className={cn(
            "flex flex-col items-center justify-center p-6 border-2 border-dashed rounded cursor-pointer transition-colors",
            isUploading
              ? "border-primary/50 bg-primary/5"
              : "border-border hover:border-primary/50 hover:bg-primary/5"
          )}
        >
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={handleInput}
            className="sr-only"
            disabled={isUploading}
          />
          {isUploading ? (
            <>
              <Loader2 className="h-8 w-8 text-primary animate-spin mb-3" />
              <span className="text-sm text-muted-foreground">Uploading…</span>
            </>
          ) : (
            <>
              <Upload className="h-8 w-8 text-muted-foreground mb-3" />
              <span className="text-sm text-foreground font-medium">Upload Excel File</span>
              <span className="text-xs text-muted-foreground mt-1">XLSX, XLS, CSV</span>
            </>
          )}
        </label>
      )}
    </div>
  );
}
