import { useState } from "react";
import { useNavigate } from "react-router-dom";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from "@/components/ui/card";
import { Checkbox } from "@/components/ui/checkbox";
import { useToast } from "@/hooks/use-toast";
import { ParticleBackground } from "@/components/particles/ParticleBackground";

const Login = () => {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [rememberMe, setRememberMe] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const navigate = useNavigate();
  const { toast } = useToast();

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();

    if (!email.trim() || !password.trim()) {
      toast({
        title: "Validation Error",
        description: "Please enter both email and password.",
        variant: "destructive",
      });
      return;
    }

    setIsLoading(true);

    try {
      await new Promise((resolve) => setTimeout(resolve, 1000));
      toast({
        title: "Login Successful",
        description: "Welcome back!",
      });
      navigate("/");
    } catch {
      toast({
        title: "Login Failed",
        description: "Invalid email or password. Please try again.",
        variant: "destructive",
      });
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-background flex items-center justify-center p-4 relative overflow-hidden">
      <ParticleBackground />

      <div className="w-full max-w-md space-y-6 relative z-10">
        {/* Logo / Branding */}
        <div className="flex flex-col items-center gap-3">
          <div className="w-14 h-14 bg-primary rounded-lg flex items-center justify-center shadow-lg">
            <span className="text-primary-foreground font-bold text-2xl">K</span>
          </div>
          <div className="text-center">
            <h1 className="text-2xl font-bold text-foreground tracking-tight">KPMG</h1>
            <p className="text-sm text-muted-foreground mt-1">
              Dashboard &amp; Report Automation Utility
            </p>
          </div>
        </div>

        {/* Login Card */}
        <Card className="border-border shadow-lg backdrop-blur-sm bg-card/95">
          <CardHeader className="space-y-1">
            <CardTitle className="text-xl">Sign in</CardTitle>
            <CardDescription>
              Enter your credentials to access the platform
            </CardDescription>
          </CardHeader>

          <form onSubmit={handleLogin}>
            <CardContent className="space-y-4">
              <div className="space-y-2">
                <Label htmlFor="email">Email</Label>
                <Input
                  id="email"
                  type="email"
                  placeholder="name@kpmg.com"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  autoComplete="email"
                  disabled={isLoading}
                />
              </div>

              <div className="space-y-2">
                <div className="flex items-center justify-between">
                  <Label htmlFor="password">Password</Label>
                  <button
                    type="button"
                    className="text-xs text-primary hover:underline"
                    onClick={() =>
                      toast({
                        title: "Password Reset",
                        description: "Please contact your administrator.",
                      })
                    }
                  >
                    Forgot password?
                  </button>
                </div>
                <Input
                  id="password"
                  type="password"
                  placeholder="••••••••"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  autoComplete="current-password"
                  disabled={isLoading}
                />
              </div>

              <div className="flex items-center gap-2">
                <Checkbox
                  id="remember"
                  checked={rememberMe}
                  onCheckedChange={(checked) => setRememberMe(checked === true)}
                  disabled={isLoading}
                />
                <Label htmlFor="remember" className="text-sm font-normal cursor-pointer">
                  Remember me
                </Label>
              </div>
            </CardContent>

            <CardFooter>
              <Button type="submit" className="w-full" variant="kpmg" disabled={isLoading}>
                {isLoading ? "Signing in…" : "Sign in"}
              </Button>
            </CardFooter>
          </form>
        </Card>

        <p className="text-xs text-center text-muted-foreground">
          © {new Date().getFullYear()} KPMG. All rights reserved.
        </p>
      </div>
    </div>
  );
};

export default Login;
