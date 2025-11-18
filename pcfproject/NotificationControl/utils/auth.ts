// utils/auth.ts
// MSAL.js authentication logic for PCF control



import { PublicClientApplication, AuthenticationResult } from "@azure/msal-browser";
import { NotificationContext } from "./api";

let msalInstance: PublicClientApplication | null = null;

export function initializeMsal(clientId: string, tenantId: string, redirectUri: string): PublicClientApplication {
  const msalConfig = {
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
      redirectUri
    }
  };
  msalInstance = new PublicClientApplication(msalConfig);
  return msalInstance;
}

function getMsalInstance(): PublicClientApplication {
  if (!msalInstance) throw new Error("MSAL instance not initialized. Call initializeMsal(context) first.");
  return msalInstance;
}

// Initiate login redirect
export function loginWithRedirect(scopes: string[] = ["User.Read"]) {
  getMsalInstance().loginRedirect({ scopes });
}

// Handle redirect response and store token
export async function handleRedirect(): Promise<string | null> {
  const response = await getMsalInstance().handleRedirectPromise();
  if (response && response.accessToken) {
    localStorage.setItem("graphToken", response.accessToken);
    return response.accessToken;
  }
  return null;
}

// Get token from localStorage
export function getStoredToken(): string | null {
  return localStorage.getItem("graphToken");
}

// Acquire token silently (if possible)
export async function acquireTokenSilent(scopes: string[] = ["User.Read"]): Promise<string | null> {
  const accounts = getMsalInstance().getAllAccounts();
  if (accounts.length === 0) return null;
  try {
    const response: AuthenticationResult = await getMsalInstance().acquireTokenSilent({
      account: accounts[0],
      scopes
    });
    localStorage.setItem("graphToken", response.accessToken);
    return response.accessToken;
  } catch (error) {
    return null;
  }
}

// Utility: Get token (try silent, fallback to stored)
export async function getToken(scopes: string[] = ["User.Read"]): Promise<string | null> {
  let token = await acquireTokenSilent(scopes);
  if (!token) token = getStoredToken();
  return token;
}
