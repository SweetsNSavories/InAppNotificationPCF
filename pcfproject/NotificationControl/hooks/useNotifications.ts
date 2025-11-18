
import * as React from "react";

export interface Notification {
	key: number;
	title: string;
	status: "Sent" | "Draft";
}

export function useNotifications(viewMode?: "all" | "sent" | "drafts") {
	const [items, setItems] = React.useState<Notification[]>([]);
	const [loading, setLoading] = React.useState(false);

	// Simulate initial fetch and filtering
	React.useEffect(() => {
		setLoading(true);
		setTimeout(() => {
			let notifications: Notification[] = Array.from({ length: 5 }, (_, i) => ({
				key: i,
				title: `Notification ${i + 1}`,
				status: i % 2 === 0 ? "Sent" : "Draft",
			}));
			if (viewMode === "sent") notifications = notifications.filter(n => n.status === "Sent");
			if (viewMode === "drafts") notifications = notifications.filter(n => n.status === "Draft");
			setItems(notifications);
			setLoading(false);
		}, 500);
	}, [viewMode]);

	// Refresh handler
	const refresh = () => {
		setLoading(true);
		setTimeout(() => {
			setItems(prev => [...prev, { key: prev.length, title: `Notification ${prev.length + 1}`, status: "Sent" }]);
			setLoading(false);
		}, 500);
	};

	return { items, loading, refresh };
}


