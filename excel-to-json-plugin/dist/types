export namespace Editor {
    export namespace Dialog {
        export function select(options: {
            title: string;
            path?: string;
            type?: 'directory' | 'file';
            filters?: Array<{ name: string; extensions: string[] }>;
            multi?: boolean;
            properties?: Array<'openFile' | 'openDirectory' | 'multiSelections'>;
        }): Promise<{ filePaths: string[]; canceled: boolean }>;
        
        export function warn(message: string): void;
    }

    export namespace Message {
        export function send(packageName: string, message: string, ...args: any[]): void;
        export function request(packageName: string, message: string, ...args: any[]): Promise<any>;
    }

    export namespace Panel {
        export function open(name: string): void;

        interface PanelOptions {
            template: string;
            style?: string;
            listeners?: {
                show?: () => void;
                hide?: () => void;
            };
            $?: {
                [key: string]: any;
            };
            methods?: {
                [key: string]: (...args: any[]) => any;
            };
            ready?: () => void;
            update?: (...args: any[]) => void | Promise<void>;
        }

        export function define(options: PanelOptions): void;
    }

    export namespace Project {
        export const path: string;
    }
}

export interface EditorComponent {
    $: Record<string, any>;
    shadowRoot?: ShadowRoot;
    update?(): Promise<void>;
    ready?(): void;
}
