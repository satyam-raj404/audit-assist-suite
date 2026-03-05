import { useEffect, useState } from "react";
import { User } from "lucide-react";
import { Label } from "@/components/ui/label";
import type { UserProfile } from "@/pages/UserOnboarding";

const fields: { key: keyof UserProfile; label: string }[] = [
  { key: "teamName", label: "Team Name" },
  { key: "partnerName", label: "Partner's Name" },
  { key: "clientName", label: "Client Name" },
  { key: "sector", label: "Sector" },
  { key: "employeeId", label: "Employee ID" },
];

export function UserDetailsSection() {
  const [profile, setProfile] = useState<UserProfile | null>(null);

  useEffect(() => {
    const stored = localStorage.getItem("userProfile");
    if (stored) {
      setProfile(JSON.parse(stored));
    }
  }, []);

  return (
    <div className="p-6 h-full overflow-auto">
      <div className="max-w-lg">
        <div className="flex items-center gap-3 mb-6">
          <div className="w-10 h-10 rounded-full bg-primary/10 flex items-center justify-center">
            <User className="h-5 w-5 text-primary" />
          </div>
          <div>
            <h2 className="text-lg font-semibold text-foreground">User Details</h2>
            <p className="text-sm text-muted-foreground">Your profile information</p>
          </div>
        </div>

        {profile ? (
          <div className="grid gap-4">
            {fields.map((f) => (
              <div key={f.key} className="space-y-1 p-3 rounded-lg bg-muted/50 border border-border">
                <Label className="text-xs text-muted-foreground">{f.label}</Label>
                <p className="text-sm font-medium text-foreground">
                  {profile[f.key] || "—"}
                </p>
              </div>
            ))}
          </div>
        ) : (
          <p className="text-sm text-muted-foreground">
            No profile information found. Please complete the onboarding form.
          </p>
        )}
      </div>
    </div>
  );
}
