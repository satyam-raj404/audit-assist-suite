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
  requiresFS?: boolean;
}

const sidebarItems: SidebarItem[] = [
  { icon: LayoutDashboard, label: "Dashboard", id: "dashboard" },
  { icon: Upload, label: "Upload Files", id: "upload" },
  { icon: ArrowLeftRight, label: "Reconciliation", id: "reconciliation", requiresFS: true },
  { icon: FileText, label: "Templates", id: "templates" },
  { icon: FolderOpen, label: "Output Files", id: "output" },
];

interface SidebarProps {
  activeView: string;
  onViewChange: (view: string) => void;
  isFS: boolean;
}

export function Sidebar({ activeView, onViewChange, isFS }: SidebarProps) {
  return (
    <aside className="w-48 bg-card shrink-0 flex flex-col border-r border-border">
      <nav className="flex-1 py-4">
        <ul className="space-y-1 px-2">
        {sidebarItems
          .filter((item) => !(item.requiresFS && !isFS))
          .map((item) => (
              <li key={item.id}>
                <button
                  onClick={() => onViewChange(item.id)}
                  className={cn(
                    "w-full flex items-center gap-3 px-3 py-2.5 rounded text-sm transition-colors",
                    "focus:outline-none focus-visible:ring-2 focus-visible:ring-ring",
                    activeView === item.id
                      ? "bg-primary/10 text-primary font-medium"
                      : "text-muted-foreground hover:bg-muted hover:text-foreground"
                  )}
                >
                  <item.icon className="h-4 w-4" />
                  <span>{item.label}</span>
                </button>
              </li>
            ))}
        </ul>
      </nav>

      <div className="p-4 border-t border-border">
        <p className="text-xs text-muted-foreground">Version 1.0.0</p>
      </div>
    </aside>
  );
}
