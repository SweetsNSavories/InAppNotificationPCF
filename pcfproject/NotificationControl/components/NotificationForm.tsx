// Type for environment variable value entity
interface EnvVarValueEntity {
		value: string;
		environmentvariabledefinition?: {
				schemaname?: string;
		};
}
// Helper to fetch agents for a selected queue
interface WorkstreamQueueEntity { _msdyn_workstreamid_value: string; }
interface WorkstreamUserEntity { _msdyn_userid_value: string; }
interface SystemUserEntity { systemuserid: string; fullname: string; }

async function fetchAgentsForQueue(queueId: string, context: { webAPI: unknown }): Promise<{ id: string; name: string }[]> {
	// Type assertion for webAPI
	const webAPI = (context.webAPI as {
		retrieveMultipleRecords: (entity: string, query: string) => Promise<{ entities: unknown[] }>;
	});
	// 1. Get workstreams for the queue
	const wsResult = await webAPI.retrieveMultipleRecords(
		"msdyn_omnichannelworkstreamqueue",
		`?$filter=_msdyn_queueid_value eq ${queueId}&$select=_msdyn_workstreamid_value`
	);
	const workstreamEntities = wsResult.entities as WorkstreamQueueEntity[];
	const workstreamIds = workstreamEntities.map(ws => ws._msdyn_workstreamid_value);
	const agentIds = new Set<string>();
	// 2. For each workstream, get agents
	for (const workstreamId of workstreamIds) {
		const agentResult = await webAPI.retrieveMultipleRecords(
			"msdyn_omnichannelworkstreamuser",
			`?$filter=_msdyn_workstreamid_value eq ${workstreamId}&$select=_msdyn_userid_value`
		);
		const agentEntities = agentResult.entities as WorkstreamUserEntity[];
		agentEntities.forEach(a => agentIds.add(a._msdyn_userid_value));
	}
	// 3. Get agent details
	if (agentIds.size === 0) return [];
	const idsFilter = Array.from(agentIds).map(id => `systemuserid eq ${id}`).join(' or ');
	const userResult = await webAPI.retrieveMultipleRecords(
		"systemuser",
		`?$select=systemuserid,fullname&$filter=${idsFilter}`
	);
	const userEntities = userResult.entities as SystemUserEntity[];
	return userEntities.map(u => ({ id: u.systemuserid, name: u.fullname }));
}

import * as React from "react";
import { TextField, Dropdown, DefaultButton, MessageBar, MessageBarType, TagPicker, ITag } from "@fluentui/react";
import { sendInAppNotification, InAppNotification, searchSystemUsers, searchTeams, searchQueues, searchOutlookDLs, getDLMemberObjectIds, getSystemUserIdsByObjectIds } from "../utils/api";
import { PublicClientApplication, AuthenticationResult } from "@azure/msal-browser";

const iconOptions = [
	{ key: "Info", text: "Info" },
	{ key: "Success", text: "Success" },
	{ key: "Error", text: "Error" },
	{ key: "Warning", text: "Warning" },
];

const typeOptions = [
	{ key: "Toast", text: "Toast" },
	{ key: "Banner", text: "Banner" },
];

import { IInputs } from "../generated/ManifestTypes";

interface NotificationFormProps {
	context: ComponentFramework.Context<IInputs>;
	onBack?: () => void;
}

const NotificationForm: React.FC<NotificationFormProps> = ({ context, onBack }) => {
	// Suggestion and change handlers
	const resolveUserSuggestions = async (filter: string): Promise<ITag[]> => {
		if (!filter) return [];
		const results = await searchSystemUsers(filter, context);
		return results.map((u: { key: string; name: string }) => ({ key: u.key, name: u.name }));
	};

	const handleTeamChange = async (items: ITag[] | undefined) => {
		if (!items) {
			setTeams([]);
			return;
		}
		// If any item has a GUID as name, replace with correct name from Dataverse
		const updated = await Promise.all(items.map(async item => {
			if (item.name && item.name !== item.key) return item;
			// Try to resolve name from Dataverse
			const result = await context.webAPI.retrieveMultipleRecords("team", `?$select=name&$filter=teamid eq '${item.key}'`);
			const name = result.entities[0]?.name || item.key;
			return { key: item.key, name };
		}));
		setTeams(updated);
	};

	const resolveTeamSuggestions = async (filter: string): Promise<ITag[]> => {
		if (!filter) return [];
		// Query Dataverse for teams matching filter, always return name and id
		const results = await searchTeams(filter, context);
		return results.map((t: { key: string; name: string }) => ({ key: t.key, name: t.name }));
	};

	const resolveQueueSuggestions = async (filter: string): Promise<ITag[]> => {
		if (!filter) return [];
		const results = await searchQueues(filter, context);
		return results.map((q: { key: string; name: string }) => ({ key: q.key, name: q.name }));
	};
	const [users, setUsers] = React.useState<ITag[]>([]);
	const [teams, setTeams] = React.useState<ITag[]>([]);
	// ...existing code...
		const [graphToken, setGraphToken] = React.useState<string>("");
		const [graphAuthError, setGraphAuthError] = React.useState<string>("");
		const [tenantId, setTenantId] = React.useState<string>("");
		const [clientId, setClientId] = React.useState<string>("");
		const [message, setMessage] = React.useState<string>("");
		const [messageType, setMessageType] = React.useState<MessageBarType>(MessageBarType.info);
		const [loading, setLoading] = React.useState(false);
		const [title, setTitle] = React.useState<string>("");
		const [body, setBody] = React.useState<string>("");
		const [iconType, setIconType] = React.useState<string>("Info");
		const [notificationType, setNotificationType] = React.useState<string>("Toast");
		const [queues, setQueues] = React.useState<ITag[]>([]);
		const [dls, setDLs] = React.useState<ITag[]>([]);
	// ...existing code...

	React.useEffect(() => {
		async function fetchEnvVarsAndInitMsal() {
			try {
				// Step 1: Get all environment variable definitions
				const defsResult = await context.webAPI.retrieveMultipleRecords(
					"environmentvariabledefinition",
					"?$select=environmentvariabledefinitionid,schemaname"
				);
				let tenantDefId = "";
				let clientDefId = "";
			for (const def of defsResult.entities as { schemaname?: string; environmentvariabledefinitionid?: string }[]) {
					const schema = (def.schemaname || "").toLowerCase();
					if (schema.includes("_inappnotif_app_tenant_id")) tenantDefId = def.environmentvariabledefinitionid ?? "";
					if (schema.includes("_inappnotif_app_client_id")) clientDefId = def.environmentvariabledefinitionid ?? "";
				}
				// Step 2: Get value for each found definition
				let tenantIdVal = "";
				let clientIdVal = "";
				if (tenantDefId) {
					const tenantValResult = await context.webAPI.retrieveMultipleRecords(
						"environmentvariablevalue",
						`?$select=value&$filter=_environmentvariabledefinitionid_value eq ${tenantDefId}`
					);
					tenantIdVal = (tenantValResult.entities[0] as { value?: string })?.value ?? "";
				}
				if (clientDefId) {
					const clientValResult = await context.webAPI.retrieveMultipleRecords(
						"environmentvariablevalue",
						`?$select=value&$filter=_environmentvariabledefinitionid_value eq ${clientDefId}`
					);
					clientIdVal = (clientValResult.entities[0] as { value?: string })?.value ?? "";
				}
				setTenantId(tenantIdVal);
				setClientId(clientIdVal);

				// Step 3: Initialize MSAL and get token
				if (tenantIdVal && clientIdVal) {
								const redirectUri = `${window.location.origin}/WebResources/msdyn_showDeprecationPrompt.html`;
					const msalConfig = {
						auth: {
							clientId: clientIdVal,
							authority: `https://login.microsoftonline.com/${tenantIdVal}`,
							redirectUri
						}
					};
					const msalInstance = new PublicClientApplication(msalConfig);
					const loginRequest = { scopes: ["User.Read", "Group.Read.All"] };
					setGraphAuthError("");
					try {
						await msalInstance.initialize();
						const accounts = msalInstance.getAllAccounts();
						let token = "";
						if (accounts.length > 0) {
							const response: AuthenticationResult = await msalInstance.acquireTokenSilent({
								...loginRequest,
								account: accounts[0],
							});
							token = response.accessToken;
						} else {
							const response: AuthenticationResult = await msalInstance.loginPopup(loginRequest);
							token = response.accessToken;
						}
						setGraphToken(token);
					} catch (err: unknown) {
						setGraphAuthError("Graph sign-in failed: " + ((err as Error)?.message || String(err)));
					}
				}
			} catch (err) {
				// Optionally handle error
			}
		}
		fetchEnvVarsAndInitMsal();
	}, [context]);

	const resolveDLSuggestions = async (filter: string): Promise<ITag[]> => {
		if (!filter || !graphToken) return [];
		const results = await searchOutlookDLs(filter, graphToken);
		return results.map((dl: { key: string; name: string }) => ({ key: dl.key, name: dl.name }));
	};

	async function expandTargetsAsync(): Promise<string[]> {
		const allUserIds: string[] = [...users.map(u => String(u.key))];
		try {
			for (const team of teams) {
				if (context.webAPI?.retrieveMultipleRecords) {
					try {
						const result = await context.webAPI.retrieveMultipleRecords("teammembership", `?$select=systemuserid&$filter=teamid eq '${team.key}'`);
						allUserIds.push(...(result.entities as { systemuserid: string }[]).map(m => m.systemuserid));
					} catch (err) {
						setMessage("Error resolving team members: " + (err instanceof Error ? err.message : String(err)));
						setMessageType(MessageBarType.error);
					}
				}
			}
			for (const queue of queues) {
				try {
					// Use Omnichannel workstream/agent logic
					const agents = await fetchAgentsForQueue(String(queue.key), context);
					allUserIds.push(...agents.map(a => a.id));
				} catch (err) {
					setMessage("Error resolving queue agents: " + (err instanceof Error ? err.message : String(err)));
					setMessageType(MessageBarType.error);
				}
			}
			for (const dl of dls) {
				if (graphToken) {
					try {
						const dlId = String(dl.key);
						console.log('[DL] Processing DL:', dlId, dl);
						const objectIds = await getDLMemberObjectIds(dlId, graphToken);
						console.log('[DL] Member objectIds:', objectIds);
						const userIds = await getSystemUserIdsByObjectIds(objectIds, context);
						console.log('[DL] Matched systemuser IDs:', userIds);
						allUserIds.push(...userIds);
					} catch (err) {
						setMessage("Error resolving DL members: " + (err instanceof Error ? err.message : String(err)));
						setMessageType(MessageBarType.error);
						console.error('[DL] Error resolving DL members:', err);
					}
				} else {
					console.warn('[DL] No graphToken available for DL processing');
				}
			}
		} catch (err) {
			setMessage("Error during auto resolve: " + (err instanceof Error ? err.message : String(err)));
			setMessageType(MessageBarType.error);
		}
		return Array.from(new Set(allUserIds.map(id => String(id))));
	}

	const handleSend = async () => {
		setLoading(true);
		setMessage("");
		try {
			const targets = await expandTargetsAsync();
			if (targets.length === 0) throw "No targets selected.";
			let allSuccess = true;
			const errorMessages: string[] = [];
			for (const userId of targets) {
				try {
					const notification: InAppNotification = {
						targets: [userId],
						title,
						body,
						iconType: iconType as "Info" | "Success" | "Error" | "Warning",
						notificationType: notificationType as "Toast" | "Banner",
					};
					await sendInAppNotification(notification, context);
				} catch (err) {
					allSuccess = false;
					errorMessages.push("Error sending notification to user: " + userId + ". " + (err instanceof Error ? err.message : String(err)));
				}
			}
			if (allSuccess) {
				setMessage("Notifications sent successfully!");
				setMessageType(MessageBarType.success);
				setTitle("");
				setBody("");
				setUsers([]); setTeams([]); setQueues([]); setDLs([]);
			} else {
				setMessage(errorMessages.join("\n"));
				setMessageType(MessageBarType.error);
			}
		} catch (err) {
			setMessage("Error sending notification: " + (err instanceof Error ? err.message : String(err)));
			setMessageType(MessageBarType.error);
		}
		setLoading(false);
	};

		// Debugging: Show error if env vars missing
		let debugError = "";
		if (!tenantId) debugError += "Tenant ID environment variable is missing or not loaded. ";
		if (!clientId) debugError += "Client ID environment variable is missing or not loaded. ";
		if (!graphToken && tenantId && clientId) debugError += "MSAL authentication did not succeed. No Graph token. ";

		return (
			<div className="notification-form-container">
				<h3>Send In-App Notification</h3>
				{debugError && <MessageBar messageBarType={MessageBarType.error}>{debugError}</MessageBar>}
				{message && <MessageBar messageBarType={messageType}>{message}</MessageBar>}
				<TagPicker
					label="System Users"
					selectedItems={users}
					onChange={(items: ITag[] | undefined) => setUsers(items ?? [])}
					onResolveSuggestions={resolveUserSuggestions}
					pickerSuggestionsProps={{ suggestionsHeaderText: "Select users" }}
				/>
				<TagPicker
					label="Teams"
					selectedItems={teams}
					onChange={handleTeamChange}
					onResolveSuggestions={resolveTeamSuggestions}
					pickerSuggestionsProps={{ suggestionsHeaderText: "Select teams" }}
				/>
				<TagPicker
					label="Queues"
					selectedItems={queues}
					onChange={(items: ITag[] | undefined) => setQueues(items ?? [])}
					onResolveSuggestions={resolveQueueSuggestions}
					pickerSuggestionsProps={{ suggestionsHeaderText: "Select queues" }}
				/>
				<TagPicker
					label="Outlook DLs"
					selectedItems={dls}
					onChange={(items: ITag[] | undefined) => setDLs(items ?? [])}
					onResolveSuggestions={resolveDLSuggestions}
					pickerSuggestionsProps={{ suggestionsHeaderText: "Select DLs" }}
					disabled={!graphToken}
				/>
				{graphAuthError && <MessageBar messageBarType={MessageBarType.error}>{graphAuthError}</MessageBar>}
				<TextField label="Title" value={title} onChange={(_ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) => setTitle(v || "")} required />
				<TextField label="Body" value={body} onChange={(_ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v?: string) => setBody(v || "")} multiline rows={3} required />
				<Dropdown label="Icon Type" options={iconOptions} selectedKey={iconType} onChange={(_ev: React.FormEvent<HTMLDivElement>, option?: import("@fluentui/react").IDropdownOption) => setIconType(option?.key?.toString() ?? "Info")} />
				<Dropdown label="Notification Type" options={typeOptions} selectedKey={notificationType} onChange={(_ev: React.FormEvent<HTMLDivElement>, option?: import("@fluentui/react").IDropdownOption) => setNotificationType(option?.key?.toString() ?? "Toast")} />
						<div className="notification-form-btn-row">
							<DefaultButton text="Send Notification" onClick={handleSend} disabled={loading || !title || !body} />
							{onBack && <DefaultButton text="Cancel" onClick={onBack} />}
						</div>
			</div>
		);
};

export default NotificationForm;
