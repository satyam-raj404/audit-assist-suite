import { useState, useCallback } from "react";
import { Upload, X, FileSpreadsheet, Loader2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";
import { api } from "@/lib/api";
import { toast } from "sonner";

export interface UploadedFile {
  id: string;   // server-assigned UUID
  name: string;
  size: number;
  type: string;
}

interface FileUploadProps {
  onFilesChange: (files: UploadedFile[]) => void;
}

const formatFileSize = (bytes: number) => {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
};

const isValidExcel = (file: File) => {
  const ext = file.name.toLowerCase().split(".").pop();
  return ["xlsx", "xls"].includes(ext || "");
};

export function FileUpload({ onFilesChange }: FileUploadProps) {
  const [files, setFiles] = useState<UploadedFile[]>([]);
  const [isDragOver, setIsDragOver] = useState(false);
  const [uploadingNames, setUploadingNames] = useState<string[]>([]);

  const uploadFiles = useCallback(
    async (rawFiles: File[]) => {
      const valid = rawFiles.filter(isValidExcel);
      if (valid.length < rawFiles.length) {
        toast.error("Invalid file type", { description: "Only .xlsx and .xls files are accepted." });
      }
      if (valid.length === 0) return;

      setUploadingNames(valid.map((f) => f.name));

      const uploaded: UploadedFile[] = [];
      for (const file of valid) {
        try {
          const res = await api.uploadFile(file);
          uploaded.push({ id: res.id, name: res.name, size: file.size, type: file.type });
        } catch {
          toast.error(`Upload failed: ${file.name}`);
        }
      }

      setUploadingNames([]);

      if (uploaded.length > 0) {
        const updated = [...files, ...uploaded];
        setFiles(updated);
        onFilesChange(updated);
        toast.success(`${uploaded.length} file${uploaded.length > 1 ? "s" : ""} uploaded`);
      }
    },
    [files, onFilesChange]
  );

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setIsDragOver(false);
      uploadFiles(Array.from(e.dataTransfer.files));
    },
    [uploadFiles]
  );

  const handleFileInput = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      if (!e.target.files) return;
      uploadFiles(Array.from(e.target.files));
      e.target.value = "";
    },
    [uploadFiles]
  );

  const removeFile = useCallback(
    async (id: string) => {
      try {
        await api.deleteFile(id);
      } catch {
        // best-effort delete
      }
      const updated = files.filter((f) => f.id !== id);
      setFiles(updated);
      onFilesChange(updated);
    },
    [files, onFilesChange]
  );

  const isUploading = uploadingNames.length > 0;

  return (
    <div className="space-y-4">
      <h3 className="kpmg-section-title">Upload Issue Tracker (Excel)</h3>

      <div
        onDragOver={(e) => { e.preventDefault(); setIsDragOver(true); }}
        onDragLeave={() => setIsDragOver(false)}
        onDrop={handleDrop}
        className={cn(
          "border-2 border-dashed rounded p-6 text-center transition-colors",
          isDragOver ? "border-primary bg-primary/5" : "border-border hover:border-muted-foreground/50"
        )}
      >
        {isUploading ? (
          <div className="flex flex-col items-center gap-2">
            <Loader2 className="h-8 w-8 text-primary animate-spin" />
            <p className="text-sm text-muted-foreground">
              Uploading {uploadingNames.join(", ")}…
            </p>
          </div>
        ) : (
          <>
            <Upload className="h-8 w-8 mx-auto text-muted-foreground mb-3" />
            <p className="text-sm text-muted-foreground mb-2">
              Drag and drop files here, or click to browse
            </p>
            <p className="text-xs text-muted-foreground/70 mb-4">Supported: XLSX, XLS</p>
            <label>
              <input
                type="file"
                multiple
                accept=".xlsx,.xls"
                onChange={handleFileInput}
                className="sr-only"
              />
              <Button variant="kpmg-outline" size="sm" asChild>
                <span className="cursor-pointer">Browse Files</span>
              </Button>
            </label>
          </>
        )}
      </div>

      {files.length > 0 && (
        <div className="space-y-2">
          {files.map((file) => (
            <div
              key={file.id}
              className="flex items-center gap-3 p-3 bg-muted/50 rounded border border-border"
            >
              <FileSpreadsheet className="h-5 w-5 text-primary shrink-0" />
              <div className="flex-1 min-w-0">
                <p className="text-sm font-medium truncate">{file.name}</p>
                <p className="text-xs text-muted-foreground">{formatFileSize(file.size)}</p>
              </div>
              <Button
                variant="ghost"
                size="icon"
                className="h-8 w-8 text-muted-foreground hover:text-destructive"
                onClick={() => removeFile(file.id)}
                aria-label={`Remove ${file.name}`}
              >
                <X className="h-4 w-4" />
              </Button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
