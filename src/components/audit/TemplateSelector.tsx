import { useEffect, useState } from "react";
import { Eye, CheckCircle2, Presentation, AlertCircle, Loader2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { Badge } from "@/components/ui/badge";
import { cn } from "@/lib/utils";
import { api, type PptTemplate } from "@/lib/api";
import { toast } from "sonner";

interface TemplateSelectorProps {
  auditType: string;
  utilityType: string;
  selectedId: string;
  selectedName: string;
  onSelect: (id: string, name: string, path: string) => void;
}

export function TemplateSelector({
  auditType,
  utilityType,
  selectedId,
  onSelect,
}: TemplateSelectorProps) {
  const [templates, setTemplates] = useState<PptTemplate[]>([]);
  const [loading, setLoading] = useState(false);
  const [previewTemplate, setPreviewTemplate] = useState<PptTemplate | null>(null);

  const ready = Boolean(auditType && utilityType);

  useEffect(() => {
    if (!ready) {
      setTemplates([]);
      onSelect("", "", "");
      return;
    }
    let cancelled = false;
    setLoading(true);
    api
      .getTemplates(auditType, utilityType)
      .then((data) => { if (!cancelled) setTemplates(data); })
      .catch(() => { if (!cancelled) toast.error("Could not load templates"); })
      .finally(() => { if (!cancelled) setLoading(false); });
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [auditType, utilityType]);

  const handleSelect = (t: PptTemplate) => {
    if (!t.available) {
      toast.warning("Template file not found", {
        description: "Place template.pptx in backend/templates for this audit type.",
      });
      return;
    }
    if (selectedId === t.id) {
      onSelect("", "", "");
    } else {
      onSelect(t.id, t.name, t.path);
    }
  };

  return (
    <div className="space-y-4">
      <h3 className="kpmg-section-title">Select PPT Template</h3>

      {!ready ? (
        <div className="flex items-center gap-3 p-4 rounded border border-dashed border-border text-muted-foreground text-sm">
          <Presentation className="h-4 w-4 shrink-0" />
          Select an audit type and utility type to see available templates
        </div>
      ) : loading ? (
        <div className="flex items-center gap-3 p-4 rounded border border-dashed border-border text-muted-foreground text-sm">
          <Loader2 className="h-4 w-4 shrink-0 animate-spin" />
          Loading templates...
        </div>
      ) : templates.length === 0 ? (
        <div className="flex items-center gap-3 p-4 rounded border border-dashed border-border text-muted-foreground text-sm">
          <AlertCircle className="h-4 w-4 shrink-0" />
          No templates found for this combination
        </div>
      ) : (
        <div className="space-y-2">
          {templates.map((t) => {
            const isSelected = selectedId === t.id;
            return (
              <div
                key={t.id}
                onClick={() => handleSelect(t)}
                className={cn(
                  "group flex items-center gap-3 p-3 rounded border cursor-pointer transition-colors",
                  isSelected
                    ? "border-primary bg-primary/5"
                    : t.available
                    ? "border-border bg-card hover:border-primary/50 hover:bg-primary/5"
                    : "border-border bg-muted/30 opacity-60 cursor-not-allowed"
                )}
              >
                <div
                  className={cn(
                    "shrink-0 rounded p-1.5",
                    isSelected ? "bg-primary/10 text-primary" : "bg-muted text-muted-foreground"
                  )}
                >
                  <Presentation className="h-4 w-4" />
                </div>

                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-2">
                    <p className="text-sm font-medium truncate">{t.name}</p>
                    {!t.available && (
                      <Badge variant="outline" className="text-xs text-destructive border-destructive/40 shrink-0">
                        Missing
                      </Badge>
                    )}
                  </div>
                  <p className="text-xs text-muted-foreground line-clamp-1 mt-0.5">{t.description}</p>
                </div>

                <div className="flex items-center gap-1 shrink-0">
                  <Button
                    variant="ghost"
                    size="icon"
                    className="h-7 w-7 text-muted-foreground hover:text-foreground"
                    onClick={(e) => { e.stopPropagation(); setPreviewTemplate(t); }}
                    aria-label={"View details for " + t.name}
                  >
                    <Eye className="h-3.5 w-3.5" />
                  </Button>
                  {isSelected && <CheckCircle2 className="h-4 w-4 text-primary" />}
                </div>
              </div>
            );
          })}
        </div>
      )}

      <Dialog open={!!previewTemplate} onOpenChange={(open) => { if (!open) setPreviewTemplate(null); }}>
        <DialogContent className="max-w-md">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2">
              <Presentation className="h-5 w-5 text-primary" />
              {previewTemplate?.name}
            </DialogTitle>
          </DialogHeader>
          {previewTemplate && (
            <div className="space-y-4">
              <div className="flex flex-wrap gap-2">
                <Badge variant="secondary">{previewTemplate.audit_type}</Badge>
                <Badge variant="outline">{previewTemplate.utility_type}</Badge>
                {previewTemplate.available ? (
                  <Badge className="bg-green-500/10 text-green-600 border-green-500/30">Template available</Badge>
                ) : (
                  <Badge className="bg-destructive/10 text-destructive border-destructive/30">File missing</Badge>
                )}
              </div>

              <div className="space-y-1">
                <p className="text-xs font-medium text-muted-foreground uppercase tracking-wide">Description</p>
                <p className="text-sm text-foreground leading-relaxed">{previewTemplate.description}</p>
              </div>

              <div className="space-y-1">
                <p className="text-xs font-medium text-muted-foreground uppercase tracking-wide">Template path</p>
                <p className="text-xs font-mono bg-muted rounded p-2 break-all text-muted-foreground">{previewTemplate.path}</p>
              </div>

              {!previewTemplate.available && (
                <div className="flex items-start gap-2 p-3 rounded bg-destructive/5 border border-destructive/20 text-sm text-destructive">
                  <AlertCircle className="h-4 w-4 shrink-0 mt-0.5" />
                  <span>
                    Place <code className="font-mono text-xs">template.pptx</code> in the path shown above.
                  </span>
                </div>
              )}

              <Button
                className="w-full"
                variant={selectedId === previewTemplate.id ? "outline" : "kpmg"}
                disabled={!previewTemplate.available}
                onClick={() => { handleSelect(previewTemplate); setPreviewTemplate(null); }}
              >
                {selectedId === previewTemplate.id ? "Deselect Template" : "Use This Template"}
              </Button>
            </div>
          )}
        </DialogContent>
      </Dialog>
    </div>
  );
}
