import { useState } from "react";
import { FolderOpen, ChevronRight, Loader2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { api } from "@/lib/api";
import { toast } from "sonner";

interface OutputFolderSelectorProps {
  outputPath: string;
  onOutputPathChange: (path: string) => void;
}

export function OutputFolderSelector({
  outputPath,
  onOutputPathChange,
}: OutputFolderSelectorProps) {
  const [isBrowsing, setIsBrowsing] = useState(false);

  const handleBrowse = async () => {
    setIsBrowsing(true);
    try {
      const path = await api.browseFolder();
      if (path) {
        onOutputPathChange(path);
      }
    } catch {
      toast.error("Could not open folder picker", {
        description: "Make sure the backend server is running.",
      });
    } finally {
      setIsBrowsing(false);
    }
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
            placeholder="Select output folder or type path..."
            className="pl-10 bg-card"
          />
        </div>
        <Button variant="kpmg-outline" onClick={handleBrowse} disabled={isBrowsing}>
          {isBrowsing ? (
            <Loader2 className="h-4 w-4 animate-spin" />
          ) : (
            <>
              <span>Browse</span>
              <ChevronRight className="h-4 w-4" />
            </>
          )}
        </Button>
      </div>

      <p className="text-xs text-muted-foreground">
        Output files will be saved to this folder on the server.
      </p>
    </div>
  );
}
