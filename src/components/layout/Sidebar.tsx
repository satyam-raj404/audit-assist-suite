import { cn } from "@/lib/utils";
import {
  LayoutDashboard,
  Upload,
  FileText,
  FolderOpen,
  ArrowLeftRight,
} from "lucide-react";

interface SidebarItem {
  icon: React.ElementType;
  label: string;
  id: string;
}

const sidebarItems: SidebarItem[] = [
  { icon: LayoutDashboard, label: "Dashboard", id: "dashboard" },
  { icon: Upload, label: "Upload Files", id: "upload" },
  { icon: ArrowLeftRight, label: "Reconciliation", id: "reconciliation" },
  { icon: FileText, label: "Templates", id: "templates" },
  { icon: FolderOpen, label: "Output Files", id: "output" },
];

interface SidebarProps {
  activeView: string;
  onViewChange: (view: string) => void;
}

export function Sidebar({ activeView, onViewChange }: SidebarProps) {
  return (
    <aside className="w-48 bg-sidebar shrink-0 flex flex-col border-r border-sidebar-border">
      <nav className="flex-1 py-4">
        <ul className="space-y-1 px-2">
          {sidebarItems.map((item) => (
            <li key={item.id}>
              <button
                onClick={() => onViewChange(item.id)}
                className={cn(
                  "w-full flex items-center gap-3 px-3 py-2.5 rounded text-sm transition-colors",
                  "focus:outline-none focus-visible:ring-2 focus-visible:ring-sidebar-ring",
                  activeView === item.id
                    ? "bg-sidebar-accent text-accent"
                    : "text-sidebar-foreground/70 hover:bg-sidebar-accent/50 hover:text-sidebar-foreground"
                )}
              >
                <item.icon className="h-4 w-4" />
                <span>{item.label}</span>
              </button>
            </li>
          ))}
        </ul>
      </nav>

      <div className="p-4 border-t border-sidebar-border">
        <p className="text-xs text-sidebar-foreground/50">Version 1.0.0</p>
      </div>
    </aside>
  );
}
