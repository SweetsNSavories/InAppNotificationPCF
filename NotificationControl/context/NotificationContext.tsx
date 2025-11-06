
import * as React from "react";
import { Notification } from "../hooks/useNotifications";
import { IInputs } from "../generated/ManifestTypes";

export interface NotificationContextType {
	notifications: Notification[];
	setNotifications: React.Dispatch<React.SetStateAction<Notification[]>>;
	notificationCount: number;
	context: ComponentFramework.Context<IInputs>;
}

export const NotificationContext = React.createContext<NotificationContextType | undefined>(undefined);

export const NotificationProvider: React.FC<{ children: React.ReactNode; context: ComponentFramework.Context<IInputs> }> = ({ children, context }) => {
	const [notifications, setNotifications] = React.useState<Notification[]>([]);
	const notificationCount = notifications.length;
	return (
		<NotificationContext.Provider value={{ notifications, setNotifications, notificationCount, context }}>
			{children}
		</NotificationContext.Provider>
	);
};

export function useNotificationContext() {
	const contextValue = React.useContext(NotificationContext);
	if (!contextValue) throw new Error("useNotificationContext must be used within a NotificationProvider");
	return contextValue;
}


