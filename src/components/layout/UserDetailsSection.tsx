import { useEffect, useState } from "react";
import { User, Pencil, X, Check, Loader2 } from "lucide-react";
import { Label } from "@/components/ui/label";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { toast } from "sonner";
import { api } from "@/lib/api";
import type { UserProfile } from "@/pages/UserOnboarding";

const fields: { key: keyof UserProfile; label: string; placeholder: string }[] = [
  { key: "teamName",    label: "Team Name",       placeholder: "e.g. Audit & Assurance" },
  { key: "partnerName", label: "Partner's Name",  placeholder: "e.g. John Smith" },
  { key: "clientName",  label: "Client Name",     placeholder: "e.g. Acme Corp" },
  { key: "sector",      label: "Sector",          placeholder: "e.g. Financial Services" },
  { key: "employeeId",  label: "Employee ID",     placeholder: "e.g. EMP-12345" },
];

export function UserDetailsSection() {
  const [profile, setProfile] = useState<UserProfile | null>(null);
  const [editing, setEditing]  = useState(false);
  const [draft, setDraft]      = useState<UserProfile | null>(null);
  const [isSaving, setIsSaving] = useState(false);

  useEffect(() => {
    const stored = localStorage.getItem("userProfile");
    if (stored) setProfile(JSON.parse(stored));
  }, []);

  const startEdit = () => {
    setDraft(profile ? { ...profile } : {
      teamName: "", partnerName: "", clientName: "", sector: "", employeeId: "",
    });
    setEditing(true);
  };

  const cancelEdit = () => {
    setEditing(false);
    setDraft(null);
  };

  const handleSave = async () => {
    if (!draft) return;

    const empty = Object.entries(draft).find(([, v]) => !v.trim());
    if (empty) {
      toast.error("Please fill in all fields.");
      return;
    }

    setIsSaving(true);
    try {
      const stored = localStorage.getItem("user");
      const user = stored ? JSON.parse(stored) : null;
      if (user?.id) {
        await api.saveUserProfile({
          user_id: user.id,
          team_name: draft.teamName,
          partner_name: draft.partnerName,
          client_name: draft.clientName,
          sector: draft.sector,
          employee_id: draft.employeeId,
        });
      }
      localStorage.setItem("userProfile", JSON.stringify(draft));
      setProfile(draft);
      setEditing(false);
      setDraft(null);
      toast.success("Profile updated.");
    } catch {
      toast.error("Could not save profile. Please try again.");
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className="p-6 h-full overflow-auto">
      <div className="max-w-lg">
        <div className="flex items-center justify-between mb-6">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-full bg-primary/10 flex items-center justify-center">
              <User className="h-5 w-5 text-primary" />
            </div>
            <div>
              <h2 className="text-lg font-semibold text-foreground">User Details</h2>
              <p className="text-sm text-muted-foreground">Your profile information</p>
            </div>
          </div>

          {!editing && profile && (
            <Button variant="ghost" size="sm" onClick={startEdit} className="text-muted-foreground">
              <Pencil className="h-4 w-4 mr-1" />
              Edit
            </Button>
          )}
        </div>

        {editing && draft ? (
          <div className="grid gap-4">
            {fields.map((f) => (
              <div key={f.key} className="space-y-1">
                <Label htmlFor={`edit-${f.key}`} className="text-xs text-muted-foreground">
                  {f.label}
                </Label>
                <Input
                  id={`edit-${f.key}`}
                  placeholder={f.placeholder}
                  value={draft[f.key]}
                  onChange={(e) => setDraft((prev) => prev ? { ...prev, [f.key]: e.target.value } : prev)}
                  disabled={isSaving}
                />
              </div>
            ))}

            <div className="flex gap-2 mt-2">
              <Button variant="kpmg" size="sm" onClick={handleSave} disabled={isSaving}>
                {isSaving ? <Loader2 className="h-4 w-4 animate-spin mr-1" /> : <Check className="h-4 w-4 mr-1" />}
                Save
              </Button>
              <Button variant="ghost" size="sm" onClick={cancelEdit} disabled={isSaving}>
                <X className="h-4 w-4 mr-1" />
                Cancel
              </Button>
            </div>
          </div>
        ) : profile ? (
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
          <div className="space-y-4">
            <p className="text-sm text-muted-foreground">
              No profile information found. Please complete the onboarding form.
            </p>
            <Button variant="kpmg" size="sm" onClick={startEdit}>
              <Pencil className="h-4 w-4 mr-1" />
              Add Details
            </Button>
          </div>
        )}
      </div>
    </div>
  );
}
