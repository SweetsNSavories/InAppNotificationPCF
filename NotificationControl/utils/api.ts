// Search Outlook DLs via Microsoft Graph API
export async function searchOutlookDLs(query: string, accessToken: string): Promise<{ key: string; name: string; email: string }[]> {
    const url = `https://graph.microsoft.com/v1.0/groups?$filter=mailEnabled eq true and securityEnabled eq false and startswith(mail,'${query}')&$select=id,displayName,mail`;
    const resp = await fetch(url, {
        headers: { Authorization: `Bearer ${accessToken}` }
    });
    const data = await resp.json();
    return (data.value || []).map((g: { id: string; displayName: string; mail: string }) => ({ key: g.id, name: g.displayName, email: g.mail }));
}
// Batch resolve systemuserids to names
export async function getSystemUserNamesByIds(userIds: string[], context: NotificationContext): Promise<Record<string, string>> {
    if (!userIds.length) return {};
    const filter = userIds.map(id => `systemuserid eq '${id}'`).join(' or ');
    const result = await context.webAPI.retrieveMultipleRecords("systemuser", `?$select=systemuserid,fullname&$filter=${filter}`);
    const map: Record<string, string> = {};
    for (const u of result.entities as { systemuserid: string; fullname: string }[]) {
        map[u.systemuserid] = u.fullname;
    }
    return map;
}
export interface NotificationContext {
    webAPI: {
        retrieveMultipleRecords: (entity: string, query: string, top?: number) => Promise<{ entities: unknown[] }>;
        createRecord?: (entity: string, data: object) => Promise<unknown>;
    };
    page?: {
        getClientUrl: () => string;
        // ...other properties if needed
    };
}
// ...existing code...
export async function resolveNotificationTargets(
    recipients: { type: "user" | "team" | "queue" | "dl"; id: string; email?: string }[],
    context: NotificationContext,
    accessToken?: string
): Promise<string[]> {
    const userGuids: string[] = [];
    for (const recipient of recipients) {
        if (recipient.type === "user") {
            userGuids.push(recipient.id);
        } else if (recipient.type === "team") {
            // Get all active members of the team
            const teamMembers = await context.webAPI.retrieveMultipleRecords(
                "systemuser",
                `?$select=systemuserid&$filter=teamid eq '${recipient.id}' and isdisabled eq false`
            );
            userGuids.push(...(teamMembers.entities as { systemuserid: string }[]).map(u => u.systemuserid));
        } else if (recipient.type === "queue") {
            // Get all users associated with the queue
            const queueMembers = await context.webAPI.retrieveMultipleRecords(
                "systemuser",
                `?$select=systemuserid&$filter=queueid eq '${recipient.id}' and isdisabled eq false`
            );
            userGuids.push(...(queueMembers.entities as { systemuserid: string }[]).map(u => u.systemuserid));
        } else if (recipient.type === "dl" && accessToken) {
            // Get DL member Azure AD objectIds via Graph API
            const objectIds = await getDLMemberObjectIds(recipient.id, accessToken);
            // Map objectIds to systemuser GUIDs
            const guids = await getSystemUserIdsByObjectIds(objectIds, context);
            userGuids.push(...guids);
        }
    }
    // Remove duplicates
    return Array.from(new Set(userGuids));
}
// ...existing code...



interface NotificationEntity {
    title: string;
    body: string;
    data?: string;
    createdon: string;
    ownerid: string;
}

export async function fetchNotificationsByRecipient(context: NotificationContext, userId: string): Promise<{ title: string; body: string; data?: string; createdon: string; recipients: { id: string; name: string }[] }[]> {
    // Calculate date 14 days ago
    const fourteenDaysAgo = new Date(Date.now() - 14 * 24 * 60 * 60 * 1000).toISOString();
    // Filter on _createdby_value to get notifications created by this user
    const filter = `_createdby_value eq ${userId} and createdon ge ${fourteenDaysAgo}`;
    const result = await context.webAPI.retrieveMultipleRecords(
        "appnotification",
        `?$select=title,body,data,createdon,ownerid&$filter=${filter}&$orderby=createdon desc&$top=100`
    );
    // Map notifications to display format
    return (result.entities as NotificationEntity[]).map(n => ({
        title: n.title,
        body: n.body,
        data: n.data,
        createdon: n.createdon,
        recipients: [{ id: n.ownerid, name: n.ownerid }], // Optionally resolve name if needed
    }));
}

// Group sent notifications by title/body for NotificationList
export async function fetchSentNotificationsGrouped(
    context: NotificationContext,
    userId: string
): Promise<{ title: string; body: string; recipients: { name: string }[] }[]> {
    // Fetch all notifications created by this user in last 14 days
    const notifications = await fetchNotificationsByRecipient(context, userId);
    // Group by title and body, collect all ownerids for each group
    interface GroupedNotification {
        title: string;
        body: string;
        ownerids: string[];
    }
    const groups: Record<string, GroupedNotification> = {};
    for (const n of notifications) {
        const key = `${n.title}|${n.body}`;
        if (!groups[key]) {
            groups[key] = { title: n.title, body: n.body, ownerids: [] };
        }
        // Collect all ownerids for this notification
        groups[key].ownerids.push(...n.recipients.map((r: { id: string }) => r.id));
    }
    // Remove duplicate ownerids in each group
    let allOwnerIds = Array.from(new Set(Object.values(groups).flatMap(g => g.ownerids)));
    allOwnerIds = allOwnerIds.filter(id => !!id && id !== 'undefined');
    const idToName = await getSystemUserNamesByIds(allOwnerIds, context);
    // Return grouped notifications with resolved recipient names
    return Object.values(groups).map(g => ({
        title: g.title,
        body: g.body,
        count: Array.from(new Set(g.ownerids)).length,
        recipients: Array.from(new Set(g.ownerids)).map(id => ({ name: idToName[id] || id })),
    }));
}

// Get DL members via Graph API
// Get DL members' Azure AD objectIds via Graph API
export async function getDLMemberObjectIds(dlId: string, accessToken: string): Promise<string[]> {
    const url = `https://graph.microsoft.com/v1.0/groups/${dlId}/members?$select=id`;
    const resp = await fetch(url, {
        headers: { Authorization: `Bearer ${accessToken}` }
    });
    const data = await resp.json();
    return (data.value || []).map((m: { id: string }) => m.id).filter(Boolean);
}

// Match emails to systemuser IDs in Dataverse
// Match Azure AD objectIds to systemuser IDs in Dataverse
export async function getSystemUserIdsByObjectIds(objectIds: string[], context: NotificationContext): Promise<string[]> {
    if (objectIds.length === 0) return [];
    if (!context.webAPI?.retrieveMultipleRecords) return [];
    const filter = objectIds.map((id: string) => `azureactivedirectoryobjectid eq ${id}`).join(' or ');
    const result = await context.webAPI.retrieveMultipleRecords("systemuser", `?$select=systemuserid,azureactivedirectoryobjectid&$filter=${filter}`);
    return (result.entities as { systemuserid: string }[]).map(u => u.systemuserid);
}

// Async search for system users
export async function searchSystemUsers(query: string, context: NotificationContext): Promise<{ key: string; name: string }[]> {
    if (!context.webAPI?.retrieveMultipleRecords) return [];
    const result = await context.webAPI.retrieveMultipleRecords("systemuser", `?$select=systemuserid,fullname&$filter=contains(fullname,'${query}') and isdisabled eq false`, 5);
    return (result.entities as { systemuserid: string; fullname: string }[]).map(u => ({ key: u.systemuserid, name: u.fullname }));
}

// Async search for teams
export async function searchTeams(query: string, context: NotificationContext): Promise<{ key: string; name: string }[]> {
    if (!context.webAPI?.retrieveMultipleRecords) return [];
    const result = await context.webAPI.retrieveMultipleRecords("team", `?$select=teamid,name&$filter=contains(name,'${query}')`, 5);
    return (result.entities as { teamid: string; name: string }[]).map(t => ({ key: t.teamid, name: t.name }));
}

// Async search for queues
export async function searchQueues(query: string, context: NotificationContext): Promise<{ key: string; name: string }[]> {
    if (!context.webAPI?.retrieveMultipleRecords) return [];
    const result = await context.webAPI.retrieveMultipleRecords("queue", `?$select=queueid,name&$filter=contains(name,'${query}')`, 5);
    return (result.entities as { queueid: string; name: string }[]).map(q => ({ key: q.queueid, name: q.name }));
}

export interface InAppNotification {
    targets: string[]; // user/team/queue IDs
    title: string;
    body: string;
    iconType?: "Info" | "Success" | "Error" | "Warning";
    notificationType?: "Toast" | "Banner";
    expiry?: Date;
}

export function sendInAppNotification(notification: InAppNotification, context: NotificationContext): Promise<void> {
    // Use createRecord for appnotification, matching official schema and values
    return new Promise((resolve, reject) => {
        if (!context.webAPI?.createRecord) {
            reject("context.webAPI.createRecord is not available. Make sure you are running in a Model-driven app context.");
            return;
        }
        const iconTypeMap: Record<string, number> = {
            Info: 100000000,
            Success: 100000001,
            Error: 100000002,
            Warning: 100000003
        };
        const toastTypeMap: Record<string, number> = {
            Toast: 200000000,
            Banner: 200000001
        };
        const promises = notification.targets.map(targetId => {
            const entity = {
                title: notification.title,
                body: notification.body,
                icontype: iconTypeMap[notification.iconType ?? "Info"],
                toasttype: toastTypeMap[notification.notificationType ?? "Toast"],
                // TTLInSeconds is not a standard field, only include if required by your schema
                ...(notification.expiry ? { ttl: Math.floor((notification.expiry.getTime() - Date.now()) / 1000) } : {}),
                ...(targetId ? { "ownerid@odata.bind": `/systemusers(${targetId})` } : {})
            };
            if (context.webAPI.createRecord) {
                return context.webAPI.createRecord("appnotification", entity);
            } else {
                return Promise.reject("createRecord is not available");
            }
        });
        Promise.all(promises)
            .then(() => resolve())
            .catch(reject);
    });
}
