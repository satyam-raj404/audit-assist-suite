import { useState } from "react";
import { useNavigate } from "react-router-dom";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  Card,
  CardContent,
  CardDescription,
  CardFooter,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import { useToast } from "@/hooks/use-toast";
import { api } from "@/lib/api";

export interface UserProfile {
  teamName: string;
  partnerName: string;
  clientName: string;
  sector: string;
  employeeId: string;
}

const UserOnboarding = () => {
  const [form, setForm] = useState<UserProfile>({
    teamName: "",
    partnerName: "",
    clientName: "",
    sector: "",
    employeeId: "",
  });
  const [isLoading, setIsLoading] = useState(false);
  const navigate = useNavigate();
  const { toast } = useToast();

  const handleChange = (field: keyof UserProfile, value: string) => {
    setForm((prev) => ({ ...prev, [field]: value }));
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();

    const empty = Object.entries(form).find(([, v]) => !v.trim());
    if (empty) {
      toast({
        title: "Validation Error",
        description: "Please fill in all fields.",
        variant: "destructive",
      });
      return;
    }

    setIsLoading(true);
    try {
      const stored = localStorage.getItem("user");
      const user = stored ? JSON.parse(stored) : null;

      if (user?.id) {
        await api.saveUserProfile({
          user_id: user.id,
          team_name: form.teamName,
          partner_name: form.partnerName,
          client_name: form.clientName,
          sector: form.sector,
          employee_id: form.employeeId,
        });
      }

      // Also keep in localStorage for quick access
      localStorage.setItem("userProfile", JSON.stringify(form));
      toast({ title: "Profile Saved", description: "Welcome aboard!" });
      navigate("/");
    } catch {
      toast({
        title: "Error",
        description: "Something went wrong. Please try again.",
        variant: "destructive",
      });
    } finally {
      setIsLoading(false);
    }
  };

  const fields: { key: keyof UserProfile; label: string; placeholder: string }[] = [
    { key: "teamName", label: "Team Name", placeholder: "e.g. Audit & Assurance" },
    { key: "partnerName", label: "Partner's Name", placeholder: "e.g. John Smith" },
    { key: "clientName", label: "Client Name", placeholder: "e.g. Acme Corp" },
    { key: "sector", label: "Sector", placeholder: "e.g. Financial Services" },
    { key: "employeeId", label: "Employee ID", placeholder: "e.g. EMP-12345" },
  ];

  return (
    <div className="min-h-screen bg-background flex items-center justify-center p-4">
      <div className="w-full max-w-lg space-y-6">
        <div className="flex flex-col items-center gap-3">
          <div className="w-14 h-14 bg-primary rounded-lg flex items-center justify-center shadow-lg">
            <span className="text-primary-foreground font-bold text-2xl">K</span>
          </div>
          <div className="text-center">
            <h1 className="text-2xl font-bold text-foreground tracking-tight">
              Complete Your Profile
            </h1>
            <p className="text-sm text-muted-foreground mt-1">
              Please fill in your details to continue
            </p>
          </div>
        </div>

        <Card className="border-border shadow-lg bg-card/95">
          <CardHeader className="space-y-1">
            <CardTitle className="text-xl">User Details</CardTitle>
            <CardDescription>
              This information helps us personalise your experience
            </CardDescription>
          </CardHeader>

          <form onSubmit={handleSubmit}>
            <CardContent className="space-y-4">
              {fields.map((f) => (
                <div key={f.key} className="space-y-2">
                  <Label htmlFor={f.key}>{f.label}</Label>
                  <Input
                    id={f.key}
                    placeholder={f.placeholder}
                    value={form[f.key]}
                    onChange={(e) => handleChange(f.key, e.target.value)}
                    disabled={isLoading}
                  />
                </div>
              ))}
            </CardContent>

            <CardFooter>
              <Button
                type="submit"
                className="w-full"
                variant="kpmg"
                disabled={isLoading}
              >
                {isLoading ? "Saving…" : "Continue"}
              </Button>
            </CardFooter>
          </form>
        </Card>
      </div>
    </div>
  );
};

export default UserOnboarding;
