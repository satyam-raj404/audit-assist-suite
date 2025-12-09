import { Bell, HelpCircle, Settings, User } from "lucide-react";
import { Button } from "@/components/ui/button";

export function TopNavbar() {
  return (
    <header className="h-14 bg-primary flex items-center justify-between px-6 shrink-0">
      <div className="flex items-center gap-4">
        {/* KPMG Logo */}
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 bg-primary-foreground rounded flex items-center justify-center">
            <span className="text-primary font-bold text-xs">K</span>
          </div>
          <div className="text-primary-foreground">
            <span className="font-bold text-lg tracking-tight">KPMG</span>
          </div>
        </div>
        <div className="h-6 w-px bg-sidebar-border" />
        <span className="text-primary-foreground/90 text-sm font-medium">
          Dashboard and Report Automation Utility
        </span>
      </div>

      <div className="flex items-center gap-2">
        <Button
          variant="ghost"
          size="icon"
          className="text-primary-foreground/80 hover:text-primary-foreground hover:bg-sidebar-accent"
          aria-label="Notifications"
        >
          <Bell className="h-4 w-4" />
        </Button>
        <Button
          variant="ghost"
          size="icon"
          className="text-primary-foreground/80 hover:text-primary-foreground hover:bg-sidebar-accent"
          aria-label="Help"
        >
          <HelpCircle className="h-4 w-4" />
        </Button>
        <Button
          variant="ghost"
          size="icon"
          className="text-primary-foreground/80 hover:text-primary-foreground hover:bg-sidebar-accent"
          aria-label="Settings"
        >
          <Settings className="h-4 w-4" />
        </Button>
        <div className="h-6 w-px bg-sidebar-border ml-2" />
        <Button
          variant="ghost"
          size="sm"
          className="text-primary-foreground/90 hover:text-primary-foreground hover:bg-sidebar-accent gap-2 ml-2"
        >
          <User className="h-4 w-4" />
          <span className="text-sm">User</span>
        </Button>
      </div>
    </header>
  );
}
