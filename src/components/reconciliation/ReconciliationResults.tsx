import { CheckCircle2, AlertCircle } from "lucide-react";

interface ReconciliationResultsProps {
  results: { matched: number; mismatched: number; missing: number };
  file1Name?: string;
  file2Name?: string;
}

export function ReconciliationResults({ results, file1Name, file2Name }: ReconciliationResultsProps) {
  return (
    <div className="space-y-4">
      <h3 className="kpmg-section-title">Reconciliation Results</h3>

      <div className="grid grid-cols-3 gap-4">
        <div className="p-4 bg-card rounded border border-border">
          <div className="flex items-center gap-3 mb-2">
            <CheckCircle2 className="h-5 w-5 text-emerald-600" />
            <span className="text-sm text-muted-foreground">Matched Records</span>
          </div>
          <p className="text-2xl font-bold text-emerald-600">{results.matched}</p>
        </div>

        <div className="p-4 bg-card rounded border border-border">
          <div className="flex items-center gap-3 mb-2">
            <AlertCircle className="h-5 w-5 text-amber-600" />
            <span className="text-sm text-muted-foreground">Mismatched</span>
          </div>
          <p className="text-2xl font-bold text-amber-600">{results.mismatched}</p>
        </div>

        <div className="p-4 bg-card rounded border border-border">
          <div className="flex items-center gap-3 mb-2">
            <AlertCircle className="h-5 w-5 text-red-600" />
            <span className="text-sm text-muted-foreground">Missing</span>
          </div>
          <p className="text-2xl font-bold text-red-600">{results.missing}</p>
        </div>
      </div>

      <div className="p-4 bg-card rounded border border-border">
        <p className="text-sm text-muted-foreground mb-3">Summary</p>
        <p className="text-sm text-foreground">
          Comparison of <span className="text-primary font-medium">{file1Name}</span> and{" "}
          <span className="text-primary font-medium">{file2Name}</span> completed.{" "}
          {results.mismatched + results.missing > 0 ? (
            <span className="text-amber-600">
              {results.mismatched + results.missing} discrepancies found that require attention.
            </span>
          ) : (
            <span className="text-emerald-600">All records matched successfully.</span>
          )}
        </p>
      </div>
    </div>
  );
}
