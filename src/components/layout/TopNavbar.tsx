import { Bell, HelpCircle, User } from "lucide-react";
import { Button } from "@/components/ui/button";

export function TopNavbar() {
  return (
    <header className="h-14 bg-card border-b border-border flex items-center justify-between px-6 shrink-0">
      <div className="flex items-center gap-4">
        {/* KPMG Logo */}
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 bg-accent rounded flex items-center justify-center">
            <span className="text-accent-foreground font-bold text-sm">K</span>
          </div>
          <div className="text-foreground">
            <span className="font-bold text-lg tracking-tight">KPMG</span>
          </div>
        </div>
        <div className="h-6 w-px bg-border" />
        <span className="text-muted-foreground text-sm font-medium">
          Dashboard and Report Automation Utility
        </span>
      </div>

      <div className="flex items-center gap-1">
        <Button
          variant="ghost"
          size="icon"
          className="text-muted-foreground hover:text-foreground"
          aria-label="Notifications"
        >
          <Bell className="h-4 w-4" />
        </Button>
        <Button
          variant="ghost"
          size="icon"
          className="text-muted-foreground hover:text-foreground"
          aria-label="Help"
        >
          <HelpCircle className="h-4 w-4" />
        </Button>
        <div className="h-6 w-px bg-border ml-2" />
        <Button
          variant="ghost"
          size="sm"
          className="text-muted-foreground hover:text-foreground gap-2 ml-2"
        >
          <User className="h-4 w-4" />
          <span className="text-sm">User</span>
        </Button>
      </div>
    </header>
  );
}
