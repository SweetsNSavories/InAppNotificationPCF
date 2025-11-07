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

# InAppNotificationPCF

## Overview
A robust Power Apps Component Framework (PCF) control for Dataverse, designed to deliver, display, and manage in-app notifications. This control supports secure environment variable lookup, publisher-agnostic logic, recipient resolution, and integrates with Microsoft Graph for advanced scenarios. It is built for easy adoption by developers and customers.

For more information about in-app notifications in Dataverse, see [Microsoft Docs: Send in-app notifications within model-driven apps](https://learn.microsoft.com/power-apps/developer/model-driven-apps/clientapi/send-in-app-notifications).

## Screenshots

### Notification List View
![Notification List](images/notification-list.png)
*Sent notifications grouped by title and body, showing count and recipient access.*

### Notification Details View
![Notification Details](images/notification-details.png)
*Detail view with graceful fallback for missing recipients.*

### Create Notification Form
![Notification Form](images/notification-form.png)
*Comprehensive form with multiple recipient selection options: System Users, Teams, Queues, and Outlook DLs.*

## Features
- **Notification Delivery:** Send notifications to users or groups in Dataverse.
- **Recipient Resolution:** Fetch and display recipient names using systemuser IDs.
- **Environment Variable Lookup:** Secure, publisher-agnostic configuration.
- **Microsoft Graph Integration:** Authenticate and fetch data from Microsoft Graph using MSAL.js.
- **Robust UI:** Modern, responsive React components with Fluent UI styling.
- **Error Handling:** Graceful fallback for missing recipients and robust client-side logic.

## Project Structure
```
NotificationControl/
  components/
    NotificationDetails.tsx      # Displays notification details and recipient info
    NotificationForm.tsx         # Main form for creating and sending notifications
    NotificationList.tsx         # Lists all notifications and handles navigation
    NotificationForm.css         # Styles for the notification form
    NotificationList.css         # Styles for the notification list
  context/
    NotificationContext.tsx      # React context for notification state
  hooks/
    useNotifications.ts          # Custom hook for notification logic
  utils/
    api.ts                       # Core notification logic, environment variable lookup, Graph API integration
    auth.ts                      # Authentication helpers for MSAL.js
  ControlManifest.Input.xml      # PCF control manifest (input)
  ControlManifest.xml            # PCF control manifest
  index.ts                       # Entry point for the control
```

## Component Details & Usage
### NotificationDetails.tsx
- **Purpose:** Displays the details of a notification, including title, body, icon, type, and recipient information.
- **Props:**
  - `notification`: The notification object to display.
  - `showSystemUsers`: Whether to show recipient info.
  - `onBack`: Handler to return to the notification list.
  - `context`: Dataverse context for API calls.
- **Key Functions:**
  - `fetchNames`: Fetches recipient names from Dataverse using systemuser IDs. Handles missing recipients gracefully by showing "No recipient assigned".
- **Adoption:**
  - Use in detail view to show notification info and recipient names. Handles all edge cases for missing or undefined recipients.

### NotificationForm.tsx
- **Purpose:** Main form for creating and sending notifications. Centralizes environment variable and authentication logic.
- **Props:**
  - `context`: Dataverse context for environment variable lookup and authentication.
- **Key Functions:**
  - Handles user input, MSAL authentication, and notification submission.
- **Features:**
  - Multiple recipient selection options: System Users, Teams, Queues, and Outlook DLs
  - Required fields: Title and Body
  - Optional settings: Icon Type (Info, Success, Error, Warning) and Notification Type (Toast, Banner)
- **Adoption:**
  - Use as the entry point for users to create and send notifications. Integrates with Microsoft Graph for advanced scenarios.
  
![Notification Form Example](images/notification-form.png)

### NotificationList.tsx
- **Purpose:** Displays a list of notifications and allows navigation to detail view. Notifications are grouped by title and body for efficiency.
- **Props:**
  - `notifications`: Array of notification objects.
  - `onSelect`: Handler to select a notification for detail view.
- **Key Features:**
  - **Grouping:** Notifications are grouped by title and body because the same notification is sent separately to each recipient in Dataverse. This prevents duplicate entries in the list view.
  - **Lazy Loading:** Recipients are loaded on-demand (lazy loaded) when you click to view details, reducing initial load time and improving performance.
  - **Date Filtering:** Only notifications from the last 7-14 days are loaded to keep the list manageable and avoid costly aggregation queries on large datasets.
- **Adoption:**
  - Use to provide users with an overview of all notifications. Passes context and props to child components.

### NotificationContext.tsx
- **Purpose:** Provides global notification state and actions to components via React context.
- **Adoption:**
  - Use to share notification state and actions across components.

### useNotifications.ts
- **Purpose:** Custom React hook for fetching, sending, and managing notifications.
- **Adoption:**
  - Use in components to access notification logic and state.

### api.ts
- **Purpose:** Contains core logic for notification delivery, environment variable lookup, recipient resolution, and Microsoft Graph API integration.
- **Key Functions:**
  - `getDLMemberObjectIds`, `getSystemUserIdsByObjectIds`, `getSystemUserNamesByIds`: Utility functions for recipient resolution.
- **Adoption:**
  - Use for all backend logic and API calls related to notifications.

### auth.ts
- **Purpose:** Handles authentication logic using MSAL.js for Microsoft Graph API access.
- **Adoption:**
  - Use to authenticate users and obtain tokens for Graph API calls.

## How to Adopt This Control
1. **Import the control into your Dataverse environment.**
2. **Configure environment variables for publisher-agnostic setup.**
3. **Grant users read privileges on the `appnotifications` entity.** Users receiving in-app notifications must have read privilege on the `appnotifications` table in Dataverse to view their notifications.
4. **Use NotificationForm to create and send notifications.**
5. **Display notifications using NotificationList and NotificationDetails.**
6. **Integrate with Microsoft Graph by configuring Azure AD app registration and MSAL.js.**
7. **Customize styles using the provided CSS files.**

## Prerequisites: Environment Variables & Attaching the Control

### 1. Environment Variables Setup
To make your PCF control publisher-agnostic and configurable, you should use Dataverse environment variables. These allow you to store configuration values (such as Azure AD app registration details, notification settings, or API endpoints) outside your code.

**Steps:**
- Go to Power Platform Admin Center or your Dataverse environment.
- Navigate to **Solutions** > your solution > **Environment Variables**.
- Create environment variables for:
  - Azure AD Client ID (for MSAL/Graph integration)
  - Notification settings (e.g., default DL, sender, etc.)
  - Any other config your control needs
- Reference these variables in your PCF control using the Dataverse WebAPI or context object.

**Example:**
```typescript
const clientId = context.parameters.envClientId.raw;
```

### 2. Attaching the Control to a Form
This PCF control can be attached to **any field on any form** in Dataverse. The control supports the following field types:
- Single Line of Text
- Email
- Phone
- URL
- Multiple Lines of Text
- Whole Number

**Important:** The field value itself is not used by the control—it only serves as a placeholder for the control's UI. You can place this control on any entity (User, Account, Contact, custom entities, etc.) based on your requirements.

**Steps:**
1. Go to **Power Apps** > **Tables** > Select your desired table (e.g., **System User**, **Account**, **Contact**).
2. Navigate to **Forms** and edit the form where you want to show notifications.
3. Find or add a field that matches one of the supported types listed above.
4. Click the field, then select **Change Control** > **Add Control**.
5. Choose your PCF control (e.g., `InAppNotification.NotificationControl`).
6. Configure the control properties as needed (e.g., bind environment variables, set display options).
7. Save and publish the form.

**Result:**
- The control will appear on the form, showing all in-app notifications for the current environment.
- Users can view, create, and manage notifications directly from any form where the control is placed.

## Prerequisites: Outlook DL Selection & Graph API Access

### 1. Outlook Distribution List (DL) Selection
To enable users to select Outlook Distribution Lists (DLs) for notifications, your control integrates with Microsoft Graph API. This requires:
- Azure AD app registration with delegated permissions to read DLs and users.
- Environment variable(s) to store the Azure AD Client ID and other config.

**Steps:**
- Register an app in Azure AD (portal.azure.com > Azure Active Directory > App registrations).
- Add delegated permissions for `Group.Read.All`, `User.Read`, and any other required Graph scopes.
- Store the Client ID and other config in Dataverse environment variables.
- Configure your control to use these variables for Graph API calls.

### 2. App Registration & Graph API Access
- The app registration must allow access to Microsoft Graph for reading users and groups.
- Redirect URI should be set for SPA (Single Page Application) and implicit grant enabled.
- The control uses MSAL.js to authenticate and acquire tokens for Graph API.

**Security Note:**
- By default, the control stores the authentication token in the browser (local/session storage) for convenience and seamless user experience.
- **If customers want to avoid storing tokens in the browser:**
  - They can disable persistent authentication or use a different MSAL configuration.
  - This may require users to re-authenticate more frequently and could impact usability.
  - Document this option in your deployment guide and provide configuration instructions.

**Example MSAL config:**
```typescript
const msalConfig = {
  auth: {
    clientId: clientIdFromEnvVar,
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin
  },
  cache: {
    cacheLocation: "sessionStorage", // or "localStorage"
    storeAuthStateInCookie: false
  }
};
```

## Redirect URI Requirement for MSAL Silent Authentication

### Why You Need a Redirect URI
For MSAL.js to perform silent authentication (acquire tokens without user interaction), your Azure AD app registration must include a valid redirect URI. In Dataverse/Power Apps, this is often set to an empty HTML web resource.

**Purpose:**
- The redirect URI is where MSAL will redirect the browser after authentication.
- For silent authentication, MSAL uses an iframe and the redirect URI must be a valid, accessible page in your environment.
- An empty HTML web resource is commonly used because it loads quickly and does not display content.

### How to Set Up
1. **Create an empty HTML web resource in Dataverse:**
   - Go to Power Apps > Solutions > Add > Web Resource.
   - Name it (e.g., `msal-redirect.html`).
   - Content can be just:
     ```html
     <html><head></head><body></body></html>
     ```
   - Save and publish.
2. **Add the web resource URL as a redirect URI in Azure AD app registration:**
   - Go to Azure Portal > Azure Active Directory > App registrations > Your app.
   - Under **Authentication**, add the web resource URL (e.g., `https://<org>.crm.dynamics.com/WebResources/msal-redirect.html`) as a redirect URI for SPA.
   - Enable **Access tokens** and **ID tokens** under Implicit grant.

**Important:**
When adding your web resource redirect URI as a SPA in Azure AD app registration, ensure you enable both **Access token** and **ID token** under Implicit grant. This allows MSAL.js to acquire tokens for authentication and API access, enabling seamless user experience and secure integration with Microsoft Graph.

### Example
```plaintext
Redirect URI: https://yourorg.crm.dynamics.com/WebResources/msal-redirect.html
```

### Why This Matters
- Without a valid redirect URI, MSAL cannot complete silent authentication and users may be prompted to sign in more often.
- The empty HTML web resource acts as a safe landing page for token acquisition.

## Using Dataverse Teams for In-App Notifications

If your organization configures on-floor teams as Dataverse Teams (with members assigned in Dataverse), supervisors can leverage this feature to send targeted in-app notifications to their team members.

### How It Works
- **Dataverse Team:** A group entity in Dataverse that can have multiple users as members. Teams can represent departments, on-floor groups, or any logical unit.
- **Supervisor Use Case:** If you are a supervisor and your team is set up as a Dataverse Team, you can select the team in the notification form and send in-app notifications to all its members at once.

### Benefits
- **Targeted Communication:** Easily notify all team members about important updates, tasks, or alerts.
- **Efficient Workflow:** No need to select individual users; simply select the team and send the notification.
- **Integration:** The PCF control fetches team members from Dataverse and ensures notifications are delivered to each member.

### Example Scenario
> A supervisor wants to notify their on-floor team about a shift change. The team is configured as a Dataverse Team. The supervisor selects the team in the notification form and sends the message. All team members receive the notification instantly in their Dataverse environment.

### How to Set Up
1. Ensure your teams are created and configured in Dataverse (Power Apps > Teams).
2. Assign users as members to each team.
3. Use the notification form in the PCF control to select a team and send notifications.

---

This approach streamlines communication and ensures all relevant users are informed efficiently. For more details, see [Microsoft Docs: Manage teams in Dataverse](https://learn.microsoft.com/power-platform/admin/manage-teams).

## Using Queue Selection for Workstream Notifications

If your organization uses queues and workstreams (common in Customer Service or Omnichannel scenarios), you can leverage queue selection to send in-app notifications to all agents associated with a specific workstream.

### How It Works
- **Queue:** A Dataverse entity that holds work items and is associated with agents who can work on those items.
- **Workstream:** A collection of queues and routing rules that define how work is distributed to agents.
- **Agent Use Case:** Supervisors or admins can select a queue in the notification form to send in-app notifications to all agents assigned to that queue's workstream.

### Benefits
- **Targeted Communication:** Notify all agents working on a specific queue or workstream about important updates, new assignments, or urgent issues.
- **Efficient Workflow:** No need to manually identify and select agents; simply select the queue and the control will resolve all associated agents.
- **Integration:** The PCF control fetches queue members from Dataverse and ensures notifications are delivered to each agent.

### Example Scenario
> A supervisor wants to notify all agents working on the "Support Queue" about a critical system update. The supervisor selects the queue in the notification form and sends the message. All agents associated with that queue's workstream receive the notification instantly.

### How to Set Up
1. Ensure your queues and workstreams are configured in Dataverse (Power Apps > Queues).
2. Assign agents to queues or workstreams.
3. Use the notification form in the PCF control to select a queue and send notifications to all associated agents.

---

This approach is especially useful for Customer Service and Omnichannel environments where agents are organized by workstreams. For more details, see [Microsoft Docs: Manage queues in Dataverse](https://learn.microsoft.com/dynamics365/customer-service/set-up-queues-manage-activities-cases).

---

For more details, see the [Microsoft Docs: Use environment variables in Dataverse](https://learn.microsoft.com/power-apps/maker/data-platform/environment-variables), [Add PCF controls to forms](https://learn.microsoft.com/power-apps/developer/component-framework/add-custom-controls-forms-views), [Microsoft Docs: Register an app with Azure AD](https://learn.microsoft.com/azure/active-directory/develop/quickstart-register-app), and [MSAL.js configuration options](https://learn.microsoft.com/azure/active-directory/develop/msal-js-initializing-client-applications).

## Developer Notes
- All components are documented with JSDoc comments for easy understanding.
- Error handling and fallback logic are implemented for robust user experience.
- The codebase is modular and easy to extend for new notification types or integrations.

## Contributing
- Fork the repository and create a pull request for improvements or bug fixes.
- Please document new components and functions using JSDoc comments and update the README as needed.

## License
MIT

---

For more details on each component, see the inline JSDoc comments in the source files. For questions or support, open an issue on GitHub.
