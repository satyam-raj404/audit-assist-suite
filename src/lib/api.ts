// API client for communicating with FastAPI backend
// Update API_BASE_URL to point to your running FastAPI server

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

export interface AuditRequest {
  audit_type: string;
  sub_audit_sector: string;
  template_id: string;
  output_path: string;
  file_ids: string[];
}

export interface AuditResult {
  id: string;
  status: "pending" | "running" | "complete" | "error";
  progress: number;
  steps: { id: string; label: string; status: string }[];
  report_path?: string;
}

export interface ReconciliationRequest {
  source_file_id: string;
  target_file_id: string;
}

export interface ReconciliationResult {
  id: string;
  status: "pending" | "running" | "complete" | "error";
  progress: number;
  matched: number;
  mismatched: number;
  missing: number;
}

export interface UploadedFileResponse {
  id: string;
  name: string;
  size: number;
  type: string;
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

  // Audit
  async startAudit(data: AuditRequest): Promise<AuditResult> {
    return request("/api/audit/start", {
      method: "POST",
      body: JSON.stringify(data),
    });
  },

  async getAuditStatus(auditId: string): Promise<AuditResult> {
    return request(`/api/audit/${auditId}/status`);
  },

  async cancelAudit(auditId: string): Promise<void> {
    await request(`/api/audit/${auditId}/cancel`, { method: "POST" });
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

  // Health check
  async healthCheck(): Promise<{ status: string }> {
    return request("/api/health");
  },
};
