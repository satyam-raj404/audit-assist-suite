// API client for communicating with FastAPI backend
const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || "http://localhost:8000";

async function request<T>(endpoint: string, options?: RequestInit): Promise<T> {
  const res = await fetch(`${API_BASE_URL}${endpoint}`, {
    headers: {
      "Content-Type": "application/json",
      ...options?.headers,
    },
    ...options,
  });

  if (!res.ok) {
    const error = await res.json().catch(() => ({ detail: res.statusText }));
    throw new Error(error.detail || "API request failed");
  }

  return res.json();
}

// ---- Types ----

export interface PptTemplate {
  id: string;
  name: string;
  description: string;
  audit_type: string;
  utility_type: string;
  path: string;
  available: boolean;
}

export interface PPTAutomationRequest {
  audit_type: string;
  utility_type?: string;
  report_type?: string;
  excel_file_id: string;
  pptx_file_id?: string;
  pptx_path?: string;
  month?: string;
  year?: string;
  output_path: string;
  user_id?: string;
}

export interface AuditResult {
  id: string;
  status: "pending" | "running" | "complete" | "error";
  progress: number;
  steps: { id: string; label: string; status: string }[];
  report_path?: string;
}

export interface ReconciliationRequest {
  audit_type: string;
  source_file_id: string;
  target_file_id?: string;
  output_path: string;
  user_id?: string;
}

export interface ReconciliationResult {
  id: string;
  status: "pending" | "running" | "complete" | "error";
  progress: number;
  matched: number;
  mismatched: number;
  missing: number;
  report_path?: string;
}

export interface UploadedFileResponse {
  id: string;
  name: string;
  size: number;
  type: string;
}

export interface UserProfile {
  user_id: string;
  team_name: string;
  partner_name: string;
  client_name: string;
  sector: string;
  employee_id: string;
}

export interface UserProfileResponse extends UserProfile {
  id: string;
  created_at: string;
}

export interface AuthUser {
  id: string;
  name: string;
  email: string;
}

export interface ErrorLogEntry {
  user_id?: string;
  module: string;
  error_message: string;
  stack_trace?: string;
}

export interface ErrorLogResponse {
  id: string;
  user_id?: string;
  module: string;
  error_message: string;
  stack_trace?: string;
  created_at: string;
}

export interface UserLogEntry {
  id: string;
  user_id?: string;
  action: string;
  module: string;
  run_id?: string;
  details?: string;
  created_at: string;
}

// ---- API Functions ----

export const api = {
  // File uploads
  async uploadFile(file: File): Promise<UploadedFileResponse> {
    const formData = new FormData();
    formData.append("file", file);

    const res = await fetch(`${API_BASE_URL}/api/files/upload`, {
      method: "POST",
      body: formData,
    });

    if (!res.ok) {
      const error = await res.json().catch(() => ({ detail: res.statusText }));
      throw new Error(error.detail || "Upload failed");
    }

    return res.json();
  },

  async deleteFile(fileId: string): Promise<void> {
    await request(`/api/files/${fileId}`, { method: "DELETE" });
  },

  // PPT Automation
  async startAudit(data: PPTAutomationRequest): Promise<AuditResult> {
    return request("/api/ppt/start", {
      method: "POST",
      body: JSON.stringify(data),
    });
  },

  async getAuditStatus(auditId: string): Promise<AuditResult> {
    return request(`/api/ppt/${auditId}/status`);
  },

  async getTemplates(auditType: string, utilityType: string): Promise<PptTemplate[]> {
    const params = new URLSearchParams({ audit_type: auditType, utility_type: utilityType });
    return request(`/api/ppt-templates?${params}`);
  },

  async cancelAudit(auditId: string): Promise<void> {
    await request(`/api/ppt/${auditId}/cancel`, { method: "POST" });
  },

  // Reconciliation
  async startReconciliation(data: ReconciliationRequest): Promise<ReconciliationResult> {
    return request("/api/reconciliation/start", {
      method: "POST",
      body: JSON.stringify(data),
    });
  },

  async getReconciliationStatus(reconciliationId: string): Promise<ReconciliationResult> {
    return request(`/api/reconciliation/${reconciliationId}/status`);
  },

  async exportReconciliationResults(reconciliationId: string): Promise<Blob> {
    const res = await fetch(`${API_BASE_URL}/api/reconciliation/${reconciliationId}/export`);
    if (!res.ok) throw new Error("Export failed");
    return res.blob();
  },

  // Auth
  async login(email: string, password: string): Promise<AuthUser> {
    return request("/api/auth/login", {
      method: "POST",
      body: JSON.stringify({ email, password }),
    });
  },

  async register(name: string, email: string, password: string): Promise<AuthUser> {
    return request("/api/auth/register", {
      method: "POST",
      body: JSON.stringify({ name, email, password }),
    });
  },

  // User Profile (Onboarding)
  async saveUserProfile(profile: UserProfile): Promise<UserProfileResponse> {
    return request("/api/users/profile", {
      method: "POST",
      body: JSON.stringify(profile),
    });
  },

  async getUserProfile(userId: string): Promise<UserProfileResponse> {
    return request(`/api/users/profile/${userId}`);
  },

  // Template Requests
  async submitTemplateRequest(file: File, spocEmail: string, teamWide: boolean, userId?: string): Promise<unknown> {
    const formData = new FormData();
    formData.append("file", file);
    formData.append("spoc_email", spocEmail);
    formData.append("team_wide", String(teamWide));
    if (userId) formData.append("user_id", userId);

    const res = await fetch(`${API_BASE_URL}/api/templates/request`, {
      method: "POST",
      body: formData,
    });

    if (!res.ok) {
      const error = await res.json().catch(() => ({ detail: res.statusText }));
      throw new Error(error.detail || "Template request failed");
    }

    return res.json();
  },

  // Error Logging
  async logError(entry: ErrorLogEntry): Promise<void> {
    await request("/api/errors/log", {
      method: "POST",
      body: JSON.stringify(entry),
    });
  },

  // User Logs
  async getUserLogs(userId: string): Promise<UserLogEntry[]> {
    return request(`/api/logs/${userId}`);
  },

  // Error Logs
  async getErrorLogs(limit = 50): Promise<ErrorLogResponse[]> {
    return request(`/api/errors?limit=${limit}`);
  },

  // Folder browser (opens native OS dialog on the server machine)
  async browseFolder(): Promise<string> {
    const res = await request<{ path: string }>("/api/files/browse-folder");
    return res.path;
  },

  // Health check
  async healthCheck(): Promise<{ status: string }> {
    return request("/api/health");
  },
};
