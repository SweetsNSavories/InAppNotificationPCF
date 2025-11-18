
// React/Fluent UI integration

import * as React from "react";
import { createRoot, Root } from "react-dom/client";
import { NotificationList } from "./components/NotificationList";
import { NotificationProvider } from "./context/NotificationContext";
import { IInputs, IOutputs } from "./generated/ManifestTypes";

// Simple error boundary for debugging blank screens
class ErrorBoundary extends React.Component<{ children: React.ReactNode }, { error: unknown }> {
    constructor(props: { children: React.ReactNode }) {
        super(props);
        this.state = { error: null };
    }
    static getDerivedStateFromError(error: unknown) {
        return { error };
    }
    componentDidCatch(error: unknown, info: unknown) {
        // Optionally log error
        // console.error(error, info);
    }
    render() {
        if (this.state.error) {
            return React.createElement(
                'div',
                { style: { color: 'red', padding: 16 } },
                React.createElement('b', {}, 'Error: '),
                String(this.state.error)
            );
        }
        return this.props.children;
    }
}

export class NotificationControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private container!: HTMLDivElement;
    private root: Root | undefined;

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.container = container;
        this.root = createRoot(container);
        // Extract notifications dataset rows
    const dataset = (context as ComponentFramework.Context<IInputs> & { datasets?: { notifications?: { sortedRecordIds: string[]; records: Record<string, { getValue: (field: string) => string }> } } }).datasets?.notifications;
    let notificationRows: { key: string; title: string; status: string }[] = [];
        if (dataset && dataset.sortedRecordIds) {
            notificationRows = dataset.sortedRecordIds.map((id: string) => {
                const record = dataset.records[id];
                return {
                    key: id,
                    title: record.getValue("title"),
                    status: record.getValue("status"),
                };
            });
        }
        // Pass only manifest-defined props to NotificationList
        const props = {
            fieldValue: context.parameters.fieldValue.raw || "",
            notifications: notificationRows,
        };
        // Wrap NotificationList with NotificationProvider for global state
        this.root.render(
                React.createElement(
                    ErrorBoundary,
                    { children: React.createElement(
                        NotificationProvider,
                        { context, children: React.createElement(NotificationList, { ...props, context }) }
                    ) }
                )
            );
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // TODO: Pass updated props to React component if needed
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        if (this.root) {
            this.root.unmount();
        }
    }
}
