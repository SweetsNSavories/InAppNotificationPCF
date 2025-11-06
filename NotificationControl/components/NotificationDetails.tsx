
import * as React from "react";


export interface NotificationDetailsProps {
	notification: {
		title: string;
		body: string;
		iconType?: "Info" | "Success" | "Error" | "Warning";
		notificationType?: "Toast" | "Banner";
		recipients?: Recipient[];
	};
	readOnly?: boolean;
	showSystemUsers?: boolean;
	onBack?: () => void;
}
import { IInputs } from "../generated/ManifestTypes";
import { NotificationListProps } from "./NotificationList";

export interface Recipient {
	id?: string; // systemuserid
	name?: string;
}

export interface NotificationDetailsProps {
	notification: {
		title: string;
		body: string;
		iconType?: "Info" | "Success" | "Error" | "Warning";
		notificationType?: "Toast" | "Banner";
		recipients?: Recipient[];
	};
	readOnly?: boolean;
	showSystemUsers?: boolean;
	onBack?: () => void;
	context: ComponentFramework.Context<IInputs>;
}

const NotificationDetails: React.FC<NotificationDetailsProps> = ({ notification, readOnly, showSystemUsers, onBack, context }) => {
	const handleRecipientClick = () => {
		console.log("Recipient list clicked");
		setShowNames(true);
	};
	const [recipientNames, setRecipientNames] = React.useState("");
	const [showNames, setShowNames] = React.useState(false);

	// Debug: log click events and state
	React.useEffect(() => {
		console.log("NotificationDetails mounted", notification);
	}, []);

	React.useEffect(() => {
		async function fetchNames() {
			// Use only id for systemuserid
			const validRecipients = (notification.recipients || []).filter(r => r && r.id);
			const ids = validRecipients.map(r => r.id);
			let result;
			if (ids.length) {
				if (context && context.webAPI && context.webAPI.retrieveMultipleRecords) {
					result = await context.webAPI.retrieveMultipleRecords("systemuser", `?$select=systemuserid,fullname&$filter=${ids.map(id => `systemuserid eq '${id}'`).join(' or ')}`);
				} else if (
					window.parent &&
					// @ts-expect-error: Xrm is only available in Dataverse runtime
					(window.parent).Xrm &&
					// @ts-expect-error: Xrm.WebApi is only available in Dataverse runtime
					(window.parent).Xrm.WebApi &&
					// @ts-expect-error: Xrm.WebApi.retrieveMultipleRecords is only available in Dataverse runtime
					(window.parent).Xrm.WebApi.retrieveMultipleRecords
				) {
					// @ts-expect-error: Xrm.WebApi.retrieveMultipleRecords is only available in Dataverse runtime
					result = await (window.parent).Xrm.WebApi.retrieveMultipleRecords("systemuser", `?$select=systemuserid,fullname&$filter=${ids.map(id => `systemuserid eq '${id}'`).join(' or ')}`);
				}
				if (result && result.entities) {
					const namesMap: Record<string, string> = {};
					for (const u of result.entities) {
						namesMap[u.systemuserid] = u.fullname;
					}
					setRecipientNames(ids.map(id => (id ? namesMap[id] || id : "")).filter(Boolean).join(", "));
				} else {
					setRecipientNames("");
				}
			} else {
				setRecipientNames("");
			}
		}
		if (showNames) {
			fetchNames();
		}
	}, [notification, context, showNames]);

	return (
		<div className="notification-details">
			<h4>{notification.title}</h4>
			<div>{notification.body}</div>
			{notification.iconType && <div>Icon: {notification.iconType}</div>}
			{notification.notificationType && <div>Type: {notification.notificationType}</div>}
			{showSystemUsers && (
				(() => {
					const validRecipients = (notification.recipients || []).filter(r => r && r.id);
					if (validRecipients.length === 0) {
						return (
							<div className="systemuser-container">
								<label>System Users</label>
								<div className="systemuser-none">No recipient assigned</div>
							</div>
						);
					}
					return (
						<div className="systemuser-container">
							<label htmlFor="systemuser-textbox">System Users</label>
							<div
								id="systemuser-clickable"
								className="systemuser-clickable"
								style={{ cursor: "pointer", color: "#0078d4", textDecoration: "underline" }}
								onClick={handleRecipientClick}
								title="Click to show recipient names"
							>
								{showNames ? (
									<input
										id="systemuser-textbox"
										type="text"
										readOnly
										value={recipientNames}
										className="systemuser-textbox"
										title="List of users who received the notification"
										placeholder="Recipients"
									/>
								) : (
									<span>Show Recipients</span>
								)}
							</div>
						</div>
					);
				})()
			)}
			{onBack && (
				<button type="button" onClick={onBack} className="notification-back-btn">Back to List</button>
			)}
		</div>
	);
};

export default NotificationDetails;
