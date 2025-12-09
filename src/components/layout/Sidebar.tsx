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
  active?: boolean;
}

const sidebarItems: SidebarItem[] = [
  { icon: LayoutDashboard, label: "Dashboard", active: true },
  { icon: Upload, label: "Upload Files" },
  { icon: ArrowLeftRight, label: "Reconciliation" },
  { icon: FileText, label: "Templates" },
  { icon: FolderOpen, label: "Output Files" },
];

export function Sidebar() {
  return (
    <aside className="w-48 bg-sidebar shrink-0 flex flex-col">
      <nav className="flex-1 py-4">
        <ul className="space-y-1 px-2">
          {sidebarItems.map((item) => (
            <li key={item.label}>
              <button
                className={cn(
                  "w-full flex items-center gap-3 px-3 py-2.5 rounded text-sm transition-colors",
                  "focus:outline-none focus-visible:ring-2 focus-visible:ring-sidebar-ring",
                  item.active
                    ? "bg-sidebar-accent text-sidebar-accent-foreground"
                    : "text-sidebar-foreground/80 hover:bg-sidebar-accent/50 hover:text-sidebar-foreground"
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
        <p className="text-xs text-sidebar-foreground/60">Version 1.0.0</p>
      </div>
    </aside>
  );
}
