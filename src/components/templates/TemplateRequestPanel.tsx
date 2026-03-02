import { useState, useRef } from "react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Upload, FileText, X, Send } from "lucide-react";
import { useToast } from "@/hooks/use-toast";
import { api } from "@/lib/api";

export function TemplateRequestPanel() {
  const [file, setFile] = useState<File | null>(null);
  const [spocEmail, setSpocEmail] = useState("");
  const [teamWide, setTeamWide] = useState<"yes" | "no">("no");
  const [isSubmitting, setIsSubmitting] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);
  const { toast } = useToast();

  const handleFileDrop = (e: React.DragEvent) => {
    e.preventDefault();
    const dropped = e.dataTransfer.files[0];
    if (dropped) setFile(dropped);
  };

  const handleFileInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selected = e.target.files?.[0];
    if (selected) setFile(selected);
  };

  const handleSubmit = async () => {
    if (!file) {
      toast({ title: "Missing file", description: "Please upload a template file.", variant: "destructive" });
      return;
    }
    if (!spocEmail.trim() || !spocEmail.includes("@")) {
      toast({ title: "Invalid email", description: "Please enter a valid SPOC email.", variant: "destructive" });
      return;
    }

    setIsSubmitting(true);
    try {
      const stored = localStorage.getItem("user");
      const user = stored ? JSON.parse(stored) : null;

      await api.submitTemplateRequest(file, spocEmail, teamWide === "yes", user?.id);
      toast({ title: "Request Submitted", description: "Your template request has been sent successfully." });
      setFile(null);
      setSpocEmail("");
      setTeamWide("no");
    } catch {
      toast({ title: "Error", description: "Failed to submit request.", variant: "destructive" });
    } finally {
      setIsSubmitting(false);
    }
  };

  const formatSize = (bytes: number) => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };

  return (
    <div className="p-6 overflow-auto h-full">
      <div className="max-w-2xl mx-auto space-y-6">
        <Card>
          <CardHeader>
            <CardTitle className="text-lg">Template Request</CardTitle>
            <CardDescription>
              Upload your template file and provide the SPOC details for review
            </CardDescription>
          </CardHeader>

          <CardContent className="space-y-6">
            {/* File Upload */}
            <div className="space-y-2">
              <Label>Template File</Label>
              <div
                onDragOver={(e) => e.preventDefault()}
                onDrop={handleFileDrop}
                className="border-2 border-dashed border-border rounded-lg p-6 text-center cursor-pointer hover:border-primary/50 transition-colors"
                onClick={() => inputRef.current?.click()}
              >
                <input
                  ref={inputRef}
                  type="file"
                  className="hidden"
                  accept=".xlsx,.xls,.docx,.pptx,.csv,.pdf"
                  onChange={handleFileInput}
                />
                {file ? (
                  <div className="flex items-center justify-center gap-3">
                    <FileText className="h-8 w-8 text-primary" />
                    <div className="text-left">
                      <p className="text-sm font-medium text-foreground">{file.name}</p>
                      <p className="text-xs text-muted-foreground">{formatSize(file.size)}</p>
                    </div>
                    <Button
                      variant="ghost"
                      size="icon"
                      className="ml-2"
                      onClick={(e) => {
                        e.stopPropagation();
                        setFile(null);
                      }}
                    >
                      <X className="h-4 w-4" />
                    </Button>
                  </div>
                ) : (
                  <div className="space-y-2">
                    <Upload className="h-8 w-8 text-muted-foreground mx-auto" />
                    <p className="text-sm text-muted-foreground">
                      Drag & drop or <span className="text-primary font-medium">browse</span>
                    </p>
                    <p className="text-xs text-muted-foreground">
                      Supported: XLSX, DOCX, PPTX, CSV, PDF
                    </p>
                  </div>
                )}
              </div>
            </div>

            {/* SPOC Email */}
            <div className="space-y-2">
              <Label htmlFor="spoc-email">SPOC Email</Label>
              <Input
                id="spoc-email"
                type="email"
                placeholder="spoc@kpmg.com"
                value={spocEmail}
                onChange={(e) => setSpocEmail(e.target.value)}
                disabled={isSubmitting}
              />
            </div>

            {/* Team-wide Radio */}
            <div className="space-y-3">
              <Label>Is this template used across your whole team?</Label>
              <RadioGroup
                value={teamWide}
                onValueChange={(v) => setTeamWide(v as "yes" | "no")}
                className="flex gap-6"
              >
                <div className="flex items-center gap-2">
                  <RadioGroupItem value="yes" id="team-yes" />
                  <Label htmlFor="team-yes" className="font-normal cursor-pointer">
                    Yes
                  </Label>
                </div>
                <div className="flex items-center gap-2">
                  <RadioGroupItem value="no" id="team-no" />
                  <Label htmlFor="team-no" className="font-normal cursor-pointer">
                    No
                  </Label>
                </div>
              </RadioGroup>
            </div>

            {/* Submit */}
            <Button
              variant="kpmg"
              className="w-full"
              onClick={handleSubmit}
              disabled={isSubmitting}
            >
              <Send className="h-4 w-4 mr-2" />
              {isSubmitting ? "Submitting…" : "Submit Request"}
            </Button>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
