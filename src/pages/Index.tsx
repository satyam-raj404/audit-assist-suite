import { useState } from "react";
import { TopNavbar } from "@/components/layout/TopNavbar";
import { Sidebar } from "@/components/layout/Sidebar";
import { StatusBar } from "@/components/layout/StatusBar";
import { AuditPanel } from "@/components/audit/AuditPanel";

type AppStatus = "idle" | "running" | "success" | "error";

const Index = () => {
  const [status, setStatus] = useState<AppStatus>("idle");
  const [statusMessage, setStatusMessage] = useState("Ready to run audit");

  const handleStatusChange = (newStatus: AppStatus, message: string) => {
    setStatus(newStatus);
    setStatusMessage(message);
  };

  const currentTime = new Date().toLocaleTimeString("en-US", {
    hour: "2-digit",
    minute: "2-digit",
  });

  return (
    <div className="min-h-screen bg-muted/50 flex items-center justify-center p-4">
      <div className="kpmg-container rounded-lg overflow-hidden flex flex-col border border-border">
        <TopNavbar />

        <div className="flex flex-1 overflow-hidden">
          <Sidebar />

          <main className="flex-1 bg-background overflow-hidden flex flex-col">
            <div className="p-6 border-b border-border bg-card">
              <h1 className="text-xl font-semibold text-foreground">
                Dashboard and Report Automation
              </h1>
              <p className="text-sm text-muted-foreground mt-1">
                Upload documents, configure audit parameters, and generate automated reports
              </p>
            </div>

            <div className="flex-1 overflow-hidden">
              <AuditPanel onStatusChange={handleStatusChange} />
            </div>
          </main>
        </div>

        <StatusBar status={status} message={statusMessage} timestamp={currentTime} />
      </div>
    </div>
  );
};

export default Index;
