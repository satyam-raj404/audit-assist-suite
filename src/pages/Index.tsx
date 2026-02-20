import { useState } from "react";
import { TopNavbar } from "@/components/layout/TopNavbar";
import { Sidebar } from "@/components/layout/Sidebar";
import { StatusBar } from "@/components/layout/StatusBar";
import { AuditPanel } from "@/components/audit/AuditPanel";
import { ReconciliationPanel } from "@/components/reconciliation/ReconciliationPanel";
import { Switch } from "@/components/ui/switch";
import { Label } from "@/components/ui/label";

type AppStatus = "idle" | "running" | "success" | "error";

const viewTitles: Record<string, { title: string; description: string }> = {
  dashboard: {
    title: "Dashboard and Report Automation",
    description: "Upload documents, configure audit parameters, and generate automated reports",
  },
  reconciliation: {
    title: "Data Reconciliation",
    description: "Compare two Excel files to identify discrepancies and generate variance reports",
  },
  upload: {
    title: "Upload Files",
    description: "Upload documents for processing",
  },
  templates: {
    title: "Report Templates",
    description: "Manage and configure report templates",
  },
  output: {
    title: "Output Files",
    description: "View and manage generated reports",
  },
};

const Index = () => {
  const [status, setStatus] = useState<AppStatus>("idle");
  const [statusMessage, setStatusMessage] = useState("Ready");
  const [activeView, setActiveView] = useState("dashboard");
  const [isFS, setIsFS] = useState(true);

  const handleFSChange = (checked: boolean) => {
    setIsFS(checked);
    if (!checked && activeView === "reconciliation") {
      setActiveView("dashboard");
    }
  };

  const handleStatusChange = (newStatus: AppStatus, message: string) => {
    setStatus(newStatus);
    setStatusMessage(message);
  };

  const currentTime = new Date().toLocaleTimeString("en-US", {
    hour: "2-digit",
    minute: "2-digit",
  });

  const currentView = viewTitles[activeView] || viewTitles.dashboard;

  return (
    <div className="min-h-screen bg-muted flex items-center justify-center p-4">
      <div className="kpmg-container rounded-lg overflow-hidden flex flex-col border border-border bg-background">
        <TopNavbar />

        <div className="flex flex-1 overflow-hidden">
          <Sidebar activeView={activeView} onViewChange={setActiveView} isFS={isFS} />

          <main className="flex-1 bg-background overflow-hidden flex flex-col">
            <div className="p-6 border-b border-border flex items-center justify-between">
              <div>
                <h1 className="text-xl font-semibold text-foreground">
                  {currentView.title}
                </h1>
                <p className="text-sm text-muted-foreground mt-1">
                  {currentView.description}
                </p>
              </div>

              {/* FS / Non-FS Switch */}
              <div className="flex items-center gap-3 bg-muted rounded-lg px-4 py-2">
                <Label
                  htmlFor="fs-switch"
                  className={`text-sm font-semibold transition-colors ${!isFS ? "text-foreground" : "text-muted-foreground"}`}
                >
                  Non-FS
                </Label>
                <Switch
                  id="fs-switch"
                  checked={isFS}
                  onCheckedChange={handleFSChange}
                />
                <Label
                  htmlFor="fs-switch"
                  className={`text-sm font-semibold transition-colors ${isFS ? "text-primary" : "text-muted-foreground"}`}
                >
                  FS
                </Label>
              </div>
            </div>

            <div className="flex-1 overflow-hidden">
              {activeView === "dashboard" && (
                <AuditPanel onStatusChange={handleStatusChange} />
              )}
              {activeView === "reconciliation" && (
                <ReconciliationPanel onStatusChange={handleStatusChange} />
              )}
              {activeView === "upload" && (
                <div className="p-6 text-muted-foreground">Upload section coming soon...</div>
              )}
              {activeView === "templates" && (
                <div className="p-6 text-muted-foreground">Templates section coming soon...</div>
              )}
              {activeView === "output" && (
                <div className="p-6 text-muted-foreground">Output files section coming soon...</div>
              )}
            </div>
          </main>
        </div>

        <StatusBar status={status} message={statusMessage} timestamp={currentTime} />
      </div>
    </div>
  );
};

export default Index;
