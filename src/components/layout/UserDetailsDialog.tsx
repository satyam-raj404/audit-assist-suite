import { useEffect, useState } from "react";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogDescription,
} from "@/components/ui/dialog";
import { Label } from "@/components/ui/label";
import type { UserProfile } from "@/pages/UserOnboarding";

interface UserDetailsDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
}

export function UserDetailsDialog({ open, onOpenChange }: UserDetailsDialogProps) {
  const [profile, setProfile] = useState<UserProfile | null>(null);

  useEffect(() => {
    if (open) {
      const stored = localStorage.getItem("userProfile");
      if (stored) {
        setProfile(JSON.parse(stored));
      }
    }
  }, [open]);

  const fields: { key: keyof UserProfile; label: string }[] = [
    { key: "teamName", label: "Team Name" },
    { key: "partnerName", label: "Partner's Name" },
    { key: "clientName", label: "Client Name" },
    { key: "sector", label: "Sector" },
    { key: "employeeId", label: "Employee ID" },
  ];

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-md">
        <DialogHeader>
          <DialogTitle>User Details</DialogTitle>
          <DialogDescription>Your profile information</DialogDescription>
        </DialogHeader>

        {profile ? (
          <div className="space-y-4 py-2">
            {fields.map((f) => (
              <div key={f.key} className="space-y-1">
                <Label className="text-xs text-muted-foreground">{f.label}</Label>
                <p className="text-sm font-medium text-foreground">
                  {profile[f.key] || "—"}
                </p>
              </div>
            ))}
          </div>
        ) : (
          <p className="text-sm text-muted-foreground py-4">
            No profile information found. Please complete the onboarding form.
          </p>
        )}
      </DialogContent>
    </Dialog>
  );
}
