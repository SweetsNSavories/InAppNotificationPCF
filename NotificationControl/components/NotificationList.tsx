import { IInputs } from "../generated/ManifestTypes";
import * as React from "react";
import { DetailsList, IColumn, SelectionMode, IconButton, Panel, PrimaryButton, Stack } from "@fluentui/react";
import { fetchSentNotificationsGrouped } from "../utils/api";
import "./NotificationList.css";
import NotificationForm from "./NotificationForm";
import NotificationDetails from "./NotificationDetails";

interface Recipient {
	name: string;
}

export interface NotificationGroup {
	title: string;
	body: string;
	count: number;
	recipients: Recipient[];
}

export interface NotificationListProps {
	context: ComponentFramework.Context<IInputs>;
}

export const NotificationList: React.FC<NotificationListProps> = ({ context }) => {
	const [groups, setGroups] = React.useState<NotificationGroup[]>([]);
	const [view, setView] = React.useState<'list' | 'new' | 'details'>('list');
	const [detailsGroup, setDetailsGroup] = React.useState<NotificationGroup | null>(null);

	React.useEffect(() => {
		// Use context.userSettings.userId if available, otherwise fallback to empty string
	const userId = (context as ComponentFramework.Context<IInputs>)?.userSettings?.userId || "";
		fetchSentNotificationsGrouped(context, userId)
			.then((result: { title: string; body: string; recipients: { name: string }[] }[]) => {
				console.log('[NotificationList] Raw fetchSentNotificationsGrouped result:', result);
				const mapped: NotificationGroup[] = result.map((g) => ({
					title: g.title,
					body: g.body,
					count: g.recipients ? g.recipients.length : 0,
					recipients: g.recipients ?? [],
				}));
				console.log('[NotificationList] Mapped groups:', mapped);
				setGroups(mapped);
				return mapped;
			})
			.catch((err) => {
				console.error('[NotificationList] Error fetching notifications:', err);
				setGroups([]);
			});
	}, [context]);

	const columns: IColumn[] = [
		{ key: "title", name: "Title", fieldName: "title", minWidth: 120, maxWidth: 200, isResizable: true },
		{ key: "body", name: "Body", fieldName: "body", minWidth: 200, maxWidth: 400, isResizable: true },
		{ key: "count", name: "Count", fieldName: "count", minWidth: 60, maxWidth: 80 },
		{
			key: "recipients", name: "Recipients", minWidth: 100, maxWidth: 200, onRender: (item: NotificationGroup) => (
				<IconButton iconProps={{ iconName: "People" }} title="Show Recipients" onClick={() => {
					setDetailsGroup(item);
					setView('details');
				}} />
			)
		}
	];

	// Handle row click to open details view
	const onRowClick = (item?: NotificationGroup): void => {
		if (item) {
			setDetailsGroup(item);
			setView('details');
		}
	};

	// Render details view (read-only)
	const renderDetailsView = () => (
		detailsGroup && (
			<NotificationDetails
				notification={detailsGroup}
				context={context}
				readOnly={true}
				showSystemUsers={true}
				onBack={() => setView('list')}
			/>
		)
	);

	// Render new notification form
	const renderNewForm = () => (
		<NotificationForm
			context={context}
			onBack={() => setView('list')}
		/>
	);

	// Render list view
	const renderListView = () => (
			<Stack tokens={{ childrenGap: 16 }}>
				<Stack horizontal horizontalAlign="space-between">
					<span className="notification-list-title">Sent Notifications</span>
					<PrimaryButton iconProps={{ iconName: "Add" }} text="New Notification" onClick={() => setView('new')} />
				</Stack>
				<DetailsList
					items={groups}
					columns={columns}
					selectionMode={SelectionMode.none}
					onActiveItemChanged={onRowClick}
				/>
			</Stack>
	);

	return (
		<div>
			{view === 'list' && renderListView()}
			{view === 'new' && renderNewForm()}
			{view === 'details' && renderDetailsView()}
		</div>
	);
};
