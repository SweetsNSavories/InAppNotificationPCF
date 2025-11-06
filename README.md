# In-App Notification PCF Control

## Environment Variables


To use Microsoft Graph authentication, set the following environment variables in a `.env` file in the project root:

```env
InAppNotif_App_Tenant_Id=your-tenant-id-here
InAppNotif_App_Client_Id=your-client-id-here
```

You can copy `.env.example` to `.env` and fill in your values.

## Build & Run

```
npm install
npm run build
```

## Publishing to Gallery

Before publishing, remove any hardcoded secrets and ensure `.env` is excluded from source control.
