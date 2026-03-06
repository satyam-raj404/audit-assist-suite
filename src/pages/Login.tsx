import { useState } from "react";
import { useNavigate } from "react-router-dom";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from "@/components/ui/card";
import { Checkbox } from "@/components/ui/checkbox";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { useToast } from "@/hooks/use-toast";
import { api } from "@/lib/api";

const Login = () => {
  const [activeTab, setActiveTab] = useState("login");

  // Login state
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [rememberMe, setRememberMe] = useState(false);
  const [isLoading, setIsLoading] = useState(false);

  // Register state
  const [regName, setRegName] = useState("");
  const [regEmail, setRegEmail] = useState("");
  const [regPassword, setRegPassword] = useState("");
  const [regConfirmPassword, setRegConfirmPassword] = useState("");

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
      const user = await api.login(email, password);
      localStorage.setItem("user", JSON.stringify(user));
      toast({
        title: "Login Successful",
        description: "Welcome back!",
      });
      navigate(localStorage.getItem("userProfile") ? "/" : "/onboarding");
    } catch (err: any) {
      toast({
        title: "Login Failed",
        description: err.message || "Invalid email or password. Please try again.",
        variant: "destructive",
      });
    } finally {
      setIsLoading(false);
    }
  };

  const handleRegister = async (e: React.FormEvent) => {
    e.preventDefault();

    if (!regName.trim() || !regEmail.trim() || !regPassword.trim() || !regConfirmPassword.trim()) {
      toast({
        title: "Validation Error",
        description: "Please fill in all fields.",
        variant: "destructive",
      });
      return;
    }

    if (regPassword !== regConfirmPassword) {
      toast({
        title: "Validation Error",
        description: "Passwords do not match.",
        variant: "destructive",
      });
      return;
    }

    if (regPassword.length < 8) {
      toast({
        title: "Validation Error",
        description: "Password must be at least 8 characters.",
        variant: "destructive",
      });
      return;
    }

    setIsLoading(true);

    try {
      await api.register(regName, regEmail, regPassword);
      toast({
        title: "Account Created",
        description: "Your account has been created successfully. Please sign in.",
      });
      setActiveTab("login");
      setEmail(regEmail);
      setRegName("");
      setRegEmail("");
      setRegPassword("");
      setRegConfirmPassword("");
    } catch (err: any) {
      toast({
        title: "Registration Failed",
        description: err.message || "Could not create account. Please try again.",
        variant: "destructive",
      });
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-background flex items-center justify-center p-4">
      <div className="w-full max-w-md space-y-6">
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

        {/* Login / Register Card */}
        <Card className="border-border shadow-lg backdrop-blur-sm bg-card/95">
          <Tabs value={activeTab} onValueChange={setActiveTab}>
            <CardHeader className="space-y-3">
              <TabsList className="grid w-full grid-cols-2">
                <TabsTrigger value="login">Sign In</TabsTrigger>
                <TabsTrigger value="register">Create Account</TabsTrigger>
              </TabsList>
            </CardHeader>

            {/* Sign In Tab */}
            <TabsContent value="login" className="mt-0">
              <form onSubmit={handleLogin}>
                <CardContent className="space-y-4">
                  <CardDescription>
                    Enter your credentials to access the platform
                  </CardDescription>
                  <div className="space-y-2">
                    <Label htmlFor="login-email">Email</Label>
                    <Input
                      id="login-email"
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
                      <Label htmlFor="login-password">Password</Label>
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
                      id="login-password"
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
            </TabsContent>

            {/* Create Account Tab */}
            <TabsContent value="register" className="mt-0">
              <form onSubmit={handleRegister}>
                <CardContent className="space-y-4">
                  <CardDescription>
                    Create a new account to get started
                  </CardDescription>
                  <div className="space-y-2">
                    <Label htmlFor="reg-name">Full Name</Label>
                    <Input
                      id="reg-name"
                      type="text"
                      placeholder="John Doe"
                      value={regName}
                      onChange={(e) => setRegName(e.target.value)}
                      autoComplete="name"
                      disabled={isLoading}
                    />
                  </div>

                  <div className="space-y-2">
                    <Label htmlFor="reg-email">Email</Label>
                    <Input
                      id="reg-email"
                      type="email"
                      placeholder="name@kpmg.com"
                      value={regEmail}
                      onChange={(e) => setRegEmail(e.target.value)}
                      autoComplete="email"
                      disabled={isLoading}
                    />
                  </div>

                  <div className="space-y-2">
                    <Label htmlFor="reg-password">Password</Label>
                    <Input
                      id="reg-password"
                      type="password"
                      placeholder="Min. 8 characters"
                      value={regPassword}
                      onChange={(e) => setRegPassword(e.target.value)}
                      autoComplete="new-password"
                      disabled={isLoading}
                    />
                  </div>

                  <div className="space-y-2">
                    <Label htmlFor="reg-confirm-password">Confirm Password</Label>
                    <Input
                      id="reg-confirm-password"
                      type="password"
                      placeholder="••••••••"
                      value={regConfirmPassword}
                      onChange={(e) => setRegConfirmPassword(e.target.value)}
                      autoComplete="new-password"
                      disabled={isLoading}
                    />
                  </div>
                </CardContent>

                <CardFooter>
                  <Button type="submit" className="w-full" variant="kpmg" disabled={isLoading}>
                    {isLoading ? "Creating account…" : "Create Account"}
                  </Button>
                </CardFooter>
              </form>
            </TabsContent>
          </Tabs>
        </Card>

        <p className="text-xs text-center text-muted-foreground">
          © {new Date().getFullYear()} KPMG. All rights reserved.
        </p>
      </div>
    </div>
  );
};

export default Login;
