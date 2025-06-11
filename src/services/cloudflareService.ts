/**
 * Service for interacting with LiraDocs Cloudflare Worker API
 * Handles email ID storage, retrieval, and removal operations
 */

const WORKER_URL = "https://liradocs-email-transfer.imran-71e.workers.dev";

export interface CloudflareResponse {
  success?: boolean;
  error?: string;
  data?: any;
}

export class CloudflareService {
  /**
   * Get all email IDs for a specific client
   */
  static async getEmailIds(clientId: string): Promise<string[]> {
    try {
      const requestUrl = `${WORKER_URL}/emails/${clientId}`;
      const response = await fetch(requestUrl);
      
      if (!response.ok) {
        const errorText = await response.text();
        console.error(`Worker error: ${response.status}`, errorText);
        throw new Error(`Worker responded with ${response.status}: ${errorText}`);
      }
      
      const result: CloudflareResponse = await response.json();
      return result.data || [];
    } catch (error) {
      console.error("Error getting email IDs:", error);
      throw error;
    }
  }

  /**
   * Store an email ID for a specific client
   */
  static async storeEmailId(clientId: string, emailId: string): Promise<CloudflareResponse> {
    try {
      const requestUrl = `${WORKER_URL}/emails`;
      const requestOpts = {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ clientId, emailId })
      };
      
      const response = await fetch(requestUrl, requestOpts);
      
      if (!response.ok) {
        const errorText = await response.text();
        console.error(`Worker error: ${response.status}`, errorText);
        throw new Error(`Worker responded with ${response.status}: ${errorText}`);
      }
      
      return await response.json();
    } catch (error) {
      console.error("Error storing email ID:", error);
      throw error;
    }
  }

  /**
   * Remove an email ID for a specific client
   */
  static async removeEmailId(clientId: string, emailId: string): Promise<CloudflareResponse> {
    try {
      const requestUrl = `${WORKER_URL}/emails/${clientId}/${emailId}`;
      const requestOpts = { method: "DELETE" };
      
      const response = await fetch(requestUrl, requestOpts);
      
      if (!response.ok) {
        const errorText = await response.text();
        console.error(`Worker error: ${response.status}`, errorText);
        throw new Error(`Worker responded with ${response.status}: ${errorText}`);
      }
      
      return await response.json();
    } catch (error) {
      console.error("Error removing email ID:", error);
      throw error;
    }
  }

  /**
   * Check if an email ID exists for a specific client
   */
  static async isEmailTransferred(clientId: string, emailId: string): Promise<boolean> {
    try {
      const emailIds = await this.getEmailIds(clientId);
      return emailIds.includes(emailId);
    } catch (error) {
      console.error("Error checking email transfer status:", error);
      return false;
    }
  }
} 