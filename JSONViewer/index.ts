import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { initializeIcons } from '@fluentui/react/lib/Icons';
import '@fluentui/react/dist/css/fabric.css';

initializeIcons();

enum TableStyleOptions {
  Nested = 1,
  Table = 2,
}

export class AnotherJSONForm implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private container: HTMLDivElement;
    private jsonString: string = "";
    private labelPosition: string;
    private errorMessage: string;
    private viewMode: 'form' | 'json';
    private allowSwitch: boolean;
    private jsonData: any;
    private notifyOutputChanged: () => void;
    private maxLabelWidth: number = 0;
    private searchBox: HTMLInputElement;
    private tableStyle: 'Nested' | 'Table';
    private viewContainer: HTMLDivElement; // Add this property

    constructor() { }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ) {
        this.container = container;
        this.viewContainer = document.createElement('div'); // Initialize view container
        this.container.appendChild(this.viewContainer); // Append view container to main container

        this.notifyOutputChanged = notifyOutputChanged;
        this.jsonString = context.parameters.jsonString.raw || "";
        this.labelPosition = context.parameters.labelPosition.raw || "Left";
        this.errorMessage = context.parameters.errorMessage.raw || "Invalid JSON format.";
        this.viewMode = context.parameters.viewMode.raw === "json" ? 'json' : 'form';
        this.allowSwitch = context.parameters.allowSwitch.raw === "true";

        // Adjusted tableStyle mapping
        const tableStyleValue = context.parameters.tableStyle.raw;
        if (tableStyleValue === TableStyleOptions.Table) {
            this.tableStyle = 'Table';
        } else {
            this.tableStyle = 'Nested';
        }

        try {
            this.jsonData = JSON.parse(this.jsonString);
        } catch (e) {
            console.error("Invalid JSON string during init:", e);
            this.jsonData = {};
        }

        this.calculateMaxLabelWidth();
        this.renderControl();
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        const newJsonString = context.parameters.jsonString.raw || "";

        if (newJsonString !== this.jsonString) {
            this.jsonString = newJsonString;
            try {
                this.jsonData = JSON.parse(this.jsonString);
            } catch (e) {
                console.error("Invalid JSON string during updateView:", e);
                this.jsonData = {};
            }
        }

        this.labelPosition = context.parameters.labelPosition.raw || "Left";
        this.errorMessage = context.parameters.errorMessage.raw || "Invalid JSON format.";
        this.viewMode = context.parameters.viewMode.raw === "json" ? 'json' : 'form';
        this.allowSwitch = context.parameters.allowSwitch.raw === "true";

        // Map the OptionSet value to the corresponding style
        const tableStyleValue = context.parameters.tableStyle.raw;
        if (tableStyleValue === TableStyleOptions.Table) {
            this.tableStyle = 'Table';
        } else {
            this.tableStyle = 'Nested';
        }

        this.calculateMaxLabelWidth();
        this.renderControl();
    }

    public getOutputs(): IOutputs {
        return { jsonString: JSON.stringify(this.jsonData) };
    }

    public destroy(): void { }

    private calculateMaxLabelWidth(): void {
        this.maxLabelWidth = 0;

        for (const key in this.jsonData) {
            if (Object.prototype.hasOwnProperty.call(this.jsonData, key)) {
                const labelWidth = this.getTextWidth(key);
                if (labelWidth > this.maxLabelWidth) {
                    this.maxLabelWidth = labelWidth;
                }
            }
        }

        // Set the max label width as a CSS variable
        this.container.style.setProperty('--max-label-width', `${this.maxLabelWidth}px`);
    }

    private getTextWidth(text: string): number {
        const canvas = document.createElement("canvas");
        const context = canvas.getContext("2d");
        if (context) {
            context.font = "14px Arial";
            return context.measureText(text).width;
        }
        return 0;
    }

    private renderControl(): void {
        // Clear the view container (not the main container)
        this.viewContainer.innerHTML = '';

        // Create or update the switch button if allowed
        if (this.allowSwitch) {
            // Check if the button already exists to avoid duplicates
            let buttonContainer = this.container.querySelector('.button-container') as HTMLDivElement;
            if (!buttonContainer) {
                buttonContainer = document.createElement('div');
                buttonContainer.className = 'button-container';

                const changeButton = document.createElement('button');
                changeButton.className = 'change-view-button';
                changeButton.addEventListener('click', () => {
                    this.viewMode = this.viewMode === 'form' ? 'json' : 'form';
                    this.renderControl();
                });

                buttonContainer.appendChild(changeButton);
                this.container.insertBefore(buttonContainer, this.viewContainer); // Insert before the view content
            }

            // Update the button text based on the current view mode
            const changeButton = buttonContainer.querySelector('.change-view-button') as HTMLButtonElement;
            changeButton.textContent = this.viewMode === 'form' ? 'Switch to JSON View' : 'Switch to Form View';
        }

        // Render the appropriate view into the view container
        if (this.viewMode === 'form') {
            this.renderFormView();
        } else {
            this.renderJsonView();
        }
    }

    private renderFormView(): void {
        // Clear the view container
        this.viewContainer.innerHTML = '';

        if (!this.jsonData) {
            this.renderErrorMessage();
            return;
        }

        const formContainer = document.createElement('div');
        formContainer.className = 'form-container';

        const createControl = (key: string, value: any): HTMLElement => {
            if (typeof value === 'object' && value !== null) {
                if (Array.isArray(value) && this.tableStyle === 'Table') {
                    // Render the array as a table
                    return this.createTable(key, value);
                } else {
                    // Render nested object
                    const controlContainer = document.createElement('div');
                    controlContainer.className = 'control-container';

                    const label = document.createElement('label');
                    label.textContent = key;
                    label.className = 'control-label';

                    const nestedContainer = document.createElement('div');
                    nestedContainer.className = 'nested-container';

                    for (const nestedKey in value) {
                        if (Object.prototype.hasOwnProperty.call(value, nestedKey)) {
                            const nestedControl = createControl(nestedKey, value[nestedKey]);
                            nestedContainer.appendChild(nestedControl);
                        }
                    }

                    controlContainer.appendChild(label);
                    controlContainer.appendChild(nestedContainer);

                    return controlContainer;
                }
            } else {
                // Render primitive value
                const controlContainer = document.createElement('div');
                controlContainer.className = 'control-container';

                const label = document.createElement('label');
                label.textContent = key;
                label.className = 'control-label';

                const input = document.createElement('input');
                input.type = 'text';
                input.value = value;
                input.className = 'control-input';

                controlContainer.appendChild(label);
                controlContainer.appendChild(input);

                return controlContainer;
            }
        };

        for (const key in this.jsonData) {
            if (Object.prototype.hasOwnProperty.call(this.jsonData, key)) {
                const value = this.jsonData[key];
                const control = createControl(key, value);
                formContainer.appendChild(control);
            }
        }

        this.viewContainer.appendChild(formContainer); // Append to view container
    }

    private renderJsonView(): void {
        // Clear the view container
        this.viewContainer.innerHTML = '';

        if (!this.jsonData) {
            this.renderErrorMessage();
            return;
        }

        const jsonTreeContainer = document.createElement('div');
        jsonTreeContainer.className = 'json-tree-container';

        // Generate JSON tree
        const tree = this.createJsonTree(this.jsonData);

        // Append tree to container
        jsonTreeContainer.appendChild(tree);

        this.viewContainer.appendChild(jsonTreeContainer); // Append to view container
    }

    private renderJsonTree(): void {
        // Clear existing JSON tree
        const existingJsonTreeContainer = this.container.querySelector(".json-tree-container");
        if (existingJsonTreeContainer) {
            existingJsonTreeContainer.remove();
        }

        // Create new JSON tree container
        const jsonTreeContainer = document.createElement("div");
        jsonTreeContainer.className = "json-tree-container";

        if (!this.jsonData) {
            this.renderErrorMessage();
            return;
        }

        // Generate JSON tree
        const tree = this.createJsonTree(this.jsonData);

        // Append tree to container
        jsonTreeContainer.appendChild(tree);
        this.container.appendChild(jsonTreeContainer);
    }

    private renderErrorMessage(): void {
        // Clear the view container
        this.viewContainer.innerHTML = '';

        const errorMessageContainer = document.createElement('div');
        errorMessageContainer.className = 'error-message';
        errorMessageContainer.textContent = this.errorMessage;

        this.viewContainer.appendChild(errorMessageContainer); // Append to view container
    }

    private createControl(key: string, value: any): HTMLElement {
        const controlContainer = document.createElement('div');
        controlContainer.className = 'control-container';

        const label = document.createElement('label');
        label.textContent = key;
        label.className = 'control-label';

        if (typeof value === 'object' && value !== null) {
            if (Array.isArray(value) && this.tableStyle === 'Table') {
                // Render the array as a table
                const tableElement = this.createTable(key, value);
                controlContainer.appendChild(label);
                controlContainer.appendChild(tableElement);
            } else {
                // Render nested object
                const nestedContainer = this.createNestedControl(value);
                controlContainer.appendChild(label);
                controlContainer.appendChild(nestedContainer);
            }
        } else {
            // Render primitive value
            const input = document.createElement('input');
            input.type = 'text';
            input.value = value !== undefined ? value : '';
            input.className = 'control-input';

            // Add event listener to handle changes
            input.addEventListener('input', (event) => {
                const newValue = (event.target as HTMLInputElement).value;
                // Update the value in the data model if needed
                this.notifyOutputChanged();
            });

            controlContainer.appendChild(label);
            controlContainer.appendChild(input);
        }

        return controlContainer;
    }

    private createValueControl(value: any, onChange: (newValue: any) => void): HTMLInputElement | HTMLTextAreaElement {
        let valueElement: HTMLInputElement | HTMLTextAreaElement;
        if (typeof value === "string") {
            if (value.length < 100) {
                valueElement = document.createElement("input");
                valueElement.type = "text";
            } else {
                valueElement = document.createElement("textarea");
                valueElement.rows = 3;
            }
        } else if (typeof value === "number") {
            valueElement = document.createElement("input");
            valueElement.type = "number";
        } else if (typeof value === "boolean") {
            valueElement = document.createElement("input");
            valueElement.type = "checkbox";
            (valueElement as HTMLInputElement).checked = value;
        } else if (value instanceof Date) {
            valueElement = document.createElement("input");
            valueElement.type = "datetime-local";
            valueElement.value = new Date(value).toISOString().slice(0, 16);
        } else {
            valueElement = document.createElement("input");
            valueElement.type = "text";
        }

        if (valueElement.type !== "checkbox") {
            valueElement.value = value;
        }

        valueElement.className = "ms-TextField";
        valueElement.style.flex = "1";

        valueElement.addEventListener("input", (event) => {
            let newValue: any;
            if (valueElement.type === "checkbox") {
                newValue = (event.target as HTMLInputElement).checked;
            } else if (valueElement.type === "number") {
                newValue = parseFloat((event.target as HTMLInputElement).value);
            } else {
                newValue = (event.target as HTMLInputElement).value;
            }
            onChange(newValue);
        });

        return valueElement;
    }

    private createJsonTree(data: any, key?: string, parentElement?: HTMLElement, level: number = 0): HTMLElement {
        const container = document.createElement('div');
        container.className = 'json-node';

        const wrapper = document.createElement('div');
        wrapper.className = 'json-wrapper';

        // Indentation based on the level
        wrapper.style.marginLeft = `${level * 20}px`;

        // Handle keys
        if (key !== undefined) {
            const keySpan = document.createElement('span');
            keySpan.className = 'json-key';
            keySpan.textContent = `"${key}": `;
            wrapper.appendChild(keySpan);
        }

        if (data && typeof data === 'object' && data !== null) {
            const isArray = Array.isArray(data);
            const openingBracket = isArray ? '[' : '{';
            const closingBracket = isArray ? ']' : '}';

            // Collapse control
            const collapseControl = document.createElement('span');
            collapseControl.className = 'json-toggle';
            collapseControl.textContent = '-';
            collapseControl.style.cursor = 'pointer';

            wrapper.appendChild(collapseControl);

            // Opening bracket
            const openBracket = document.createElement('span');
            openBracket.className = 'json-bracket';
            openBracket.textContent = openingBracket;
            wrapper.appendChild(openBracket);

            // Line break and indent
            const lineBreak = document.createElement('br');
            wrapper.appendChild(lineBreak);

            // Child nodes container
            const childrenContainer = document.createElement('div');
            childrenContainer.className = 'json-children';

            // Recursively create child nodes
            const childKeys = Object.keys(data);
            childKeys.forEach((childKey, index) => {
                const childNode = this.createJsonTree(
                    data[childKey],
                    isArray ? undefined : childKey,
                    childrenContainer,
                    level + 1
                );
                childrenContainer.appendChild(childNode);

                // Comma after each item except the last
                if (index < childKeys.length - 1) {
                    const comma = document.createElement('span');
                    comma.textContent = ',';
                    childNode.appendChild(comma);
                }
            });

            wrapper.appendChild(childrenContainer);

            // Closing bracket with indentation
            const closingWrapper = document.createElement('div');
            closingWrapper.style.marginLeft = `${level * 20}px`;
            const closeBracket = document.createElement('span');
            closeBracket.className = 'json-bracket';
            closeBracket.textContent = closingBracket;
            closingWrapper.appendChild(closeBracket);

            wrapper.appendChild(closingWrapper);

            // Expand/Collapse functionality
            collapseControl.addEventListener('click', () => {
                const isCollapsed = childrenContainer.style.display === 'none';
                childrenContainer.style.display = isCollapsed ? 'block' : 'none';
                closingWrapper.style.display = isCollapsed ? 'block' : 'none';
                collapseControl.textContent = isCollapsed ? '-' : '+';
            });
        } else {
            // Primitive values
            const valueSpan = document.createElement('span');
            valueSpan.className = `json-value ${typeof data}`;
            valueSpan.textContent =
                typeof data === 'string' ? `"${data}"` : String(data);
            wrapper.appendChild(valueSpan);
        }

        container.appendChild(wrapper);
        return container;
    }

    private highlightMatches(element: HTMLElement, searchValue: string): void {
        if (!searchValue) return;

        const regex = new RegExp(`(${searchValue})`, 'gi');

        const traverse = (node: HTMLElement) => {
            if (node.childNodes.length > 0) {
                node.childNodes.forEach((child: Node) => {
                    if (child.nodeType === Node.TEXT_NODE) {
                        const text = child.textContent || '';
                        if (regex.test(text)) {
                            const highlightedText = text.replace(regex, '<span class="highlight">$1</span>');
                            const fragment = document.createElement('span');
                            fragment.innerHTML = highlightedText;
                            node.replaceChild(fragment, child);
                        }
                    } else if (child.nodeType === Node.ELEMENT_NODE) {
                        traverse(child as HTMLElement);
                    }
                });
            }
        };

        traverse(element);
    }

    private createTable(key: string, dataArray: any[]): HTMLElement {
        const controlContainer = document.createElement('div');
        controlContainer.className = 'control-container';

        const label = document.createElement('label');
        label.textContent = key;
        label.className = 'control-label';

        const tableContainer = document.createElement('div');
        tableContainer.className = 'table-container';

        const table = document.createElement('table');
        table.className = 'data-table';

        // Create table header
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');

        // Get all unique keys from dataArray
        const keysSet = new Set<string>();
        dataArray.forEach(item => {
            Object.keys(item).forEach(k => keysSet.add(k));
        });
        const keys = Array.from(keysSet);

        keys.forEach(colKey => {
            const th = document.createElement('th');
            th.textContent = colKey;
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);
        table.appendChild(thead);

        // Create table body
        const tbody = document.createElement('tbody');
        dataArray.forEach((item, rowIndex) => {
            const row = document.createElement('tr');
            keys.forEach(colKey => {
                const td = document.createElement('td');

                const cellValue = item[colKey];
                if (typeof cellValue === 'object' && cellValue !== null) {
                    if (Array.isArray(cellValue) && this.tableStyle === 'Table') {
                        // Handle nested arrays recursively
                        const nestedTable = this.createTable(colKey, cellValue);
                        td.appendChild(nestedTable);
                    } else {
                        // Render nested object
                        const nestedControl = this.createNestedControl(cellValue);
                        td.appendChild(nestedControl);
                    }
                } else {
                    // Create input box for primitive values
                    const cellInput = document.createElement('input');
                    cellInput.type = 'text';
                    cellInput.value = cellValue !== undefined ? cellValue : '';
                    cellInput.className = 'control-input';

                    // Add event listener to update dataArray
                    cellInput.addEventListener('input', (event) => {
                        const newValue = (event.target as HTMLInputElement).value;
                        dataArray[rowIndex][colKey] = newValue;
                        this.notifyOutputChanged();
                    });
                    td.appendChild(cellInput);
                }

                row.appendChild(td);
            });
            tbody.appendChild(row);
        });
        table.appendChild(tbody);

        tableContainer.appendChild(table);

        controlContainer.appendChild(label);
        controlContainer.appendChild(tableContainer);

        return controlContainer;
    }

    private createNestedControl(value: any): HTMLElement {
        const nestedContainer = document.createElement('div');
        nestedContainer.className = 'nested-container';

        for (const key in value) {
            if (Object.prototype.hasOwnProperty.call(value, key)) {
                const control = this.createControl(key, value[key]);
                nestedContainer.appendChild(control);
            }
        }

        return nestedContainer;
    }
}
