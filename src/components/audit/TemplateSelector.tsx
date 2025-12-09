import { useState } from "react";
import { Eye, FileText, Check } from "lucide-react";
import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { cn } from "@/lib/utils";

interface Template {
  id: string;
  name: string;
  description: string;
  sections: string[];
}

interface TemplateSelectorProps {
  selectedTemplate: string;
  onTemplateChange: (templateId: string) => void;
}

const templates: Template[] = [
  {
    id: "standard",
    name: "Standard Audit Report",
    description: "Comprehensive audit report following KPMG standards",
    sections: ["Executive Summary", "Scope & Methodology", "Findings", "Recommendations", "Appendices"],
  },
  {
    id: "executive",
    name: "Executive Summary Only",
    description: "High-level summary for senior management",
    sections: ["Key Highlights", "Critical Findings", "Action Items"],
  },
  {
    id: "detailed",
    name: "Detailed Technical Report",
    description: "In-depth technical analysis with supporting data",
    sections: ["Technical Overview", "Data Analysis", "Risk Assessment", "Control Evaluation", "Detailed Findings", "Technical Recommendations"],
  },
  {
    id: "regulatory",
    name: "Regulatory Compliance Report",
    description: "Format aligned with regulatory requirements",
    sections: ["Compliance Overview", "Regulatory Mapping", "Gap Analysis", "Remediation Plan", "Evidence Documentation"],
  },
];

export function TemplateSelector({
  selectedTemplate,
  onTemplateChange,
}: TemplateSelectorProps) {
  const [previewTemplate, setPreviewTemplate] = useState<Template | null>(null);

  return (
    <div className="space-y-4">
      <h3 className="kpmg-section-title">Select Template</h3>

      <div className="grid grid-cols-2 gap-3">
        {templates.map((template) => (
          <div
            key={template.id}
            className={cn(
              "relative p-4 rounded border-2 cursor-pointer transition-all",
              "focus-within:ring-2 focus-within:ring-ring focus-within:ring-offset-2",
              selectedTemplate === template.id
                ? "border-primary bg-primary/5"
                : "border-border hover:border-muted-foreground/50"
            )}
            onClick={() => onTemplateChange(template.id)}
            onKeyDown={(e) => {
              if (e.key === "Enter" || e.key === " ") {
                e.preventDefault();
                onTemplateChange(template.id);
              }
            }}
            tabIndex={0}
            role="button"
            aria-pressed={selectedTemplate === template.id}
          >
            <div className="flex items-start gap-3">
              <div
                className={cn(
                  "shrink-0 w-5 h-5 rounded-full border-2 flex items-center justify-center mt-0.5",
                  selectedTemplate === template.id
                    ? "border-primary bg-primary"
                    : "border-muted-foreground/30"
                )}
              >
                {selectedTemplate === template.id && (
                  <Check className="h-3 w-3 text-primary-foreground" />
                )}
              </div>

              <div className="flex-1 min-w-0">
                <div className="flex items-center gap-2">
                  <FileText className="h-4 w-4 text-primary shrink-0" />
                  <span className="font-medium text-sm">{template.name}</span>
                </div>
                <p className="text-xs text-muted-foreground mt-1 line-clamp-2">
                  {template.description}
                </p>
              </div>

              <Button
                variant="ghost"
                size="icon"
                className="h-7 w-7 shrink-0"
                onClick={(e) => {
                  e.stopPropagation();
                  setPreviewTemplate(template);
                }}
                aria-label={`Preview ${template.name}`}
              >
                <Eye className="h-4 w-4" />
              </Button>
            </div>
          </div>
        ))}
      </div>

      {/* Preview Dialog */}
      <Dialog open={!!previewTemplate} onOpenChange={() => setPreviewTemplate(null)}>
        <DialogContent className="max-w-md">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2">
              <FileText className="h-5 w-5 text-primary" />
              {previewTemplate?.name}
            </DialogTitle>
          </DialogHeader>

          <div className="space-y-4 py-2">
            <p className="text-sm text-muted-foreground">
              {previewTemplate?.description}
            </p>

            <div>
              <h4 className="text-sm font-medium mb-2">Report Sections:</h4>
              <ul className="space-y-1">
                {previewTemplate?.sections.map((section, index) => (
                  <li
                    key={section}
                    className="flex items-center gap-2 text-sm text-muted-foreground"
                  >
                    <span className="w-5 h-5 rounded bg-muted flex items-center justify-center text-xs font-medium">
                      {index + 1}
                    </span>
                    {section}
                  </li>
                ))}
              </ul>
            </div>
          </div>
        </DialogContent>
      </Dialog>
    </div>
  );
}
