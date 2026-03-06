import { FolderOpen, ChevronRight } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";

interface OutputFolderSelectorProps {
  outputPath: string;
  onOutputPathChange: (path: string) => void;
}

export function OutputFolderSelector({
  outputPath,
  onOutputPathChange,
}: OutputFolderSelectorProps) {
  const handleBrowse = () => {
    // In a real app, this would open a folder picker dialog
    // For this demo, we'll simulate with a predefined path
    onOutputPathChange("C:\\KPMG\\AuditReports\\2024");
  };

  return (
    <div className="space-y-4">
      <h3 className="kpmg-section-title">Output Location</h3>

      <div className="flex gap-3">
        <div className="flex-1 relative">
          <FolderOpen className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-muted-foreground" />
          <Input
            value={outputPath}
            onChange={(e) => onOutputPathChange(e.target.value)}
            placeholder="Select output folder..."
            className="pl-10 bg-card"
          />
        </div>
        <Button variant="kpmg-outline" onClick={handleBrowse}>
          <span>Browse</span>
          <ChevronRight className="h-4 w-4" />
        </Button>
      </div>

      <p className="text-xs text-muted-foreground">
        Generated reports will be saved to this location with timestamp-based filenames.
      </p>
    </div>
  );
}
