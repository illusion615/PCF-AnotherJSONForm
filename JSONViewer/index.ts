import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { initializeIcons } from '@fluentui/react/lib/Icons';
import '@fluentui/react/dist/css/fabric.css';

initializeIcons();

export class JSONViewer implements ComponentFramework.StandardControl<IInputs, IOutputs> {
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

    constructor() {}

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ) {
        this.container = container;
        this.notifyOutputChanged = notifyOutputChanged;
        this.jsonString = context.parameters.jsonString.raw || "";
        this.labelPosition = context.parameters.labelPosition.raw || "Left";
        this.errorMessage = context.parameters.errorMessage.raw || "Invalid JSON format.";
        this.viewMode = context.parameters.viewMode.raw === "json" ? 'json' : 'form';
        this.allowSwitch = context.parameters.allowSwitch.raw === "true";

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

        this.calculateMaxLabelWidth();
        this.renderControl();
    }

    public getOutputs(): IOutputs {
        return { jsonString: JSON.stringify(this.jsonData) };
    }

    public destroy(): void {}

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
        this.container.innerHTML = "";
        this.container.style.overflow = "auto";
        this.container.style.height = "100%";

        if (this.allowSwitch) {
            const buttonContainer = document.createElement("div");
            buttonContainer.style.display = "flex";
            buttonContainer.style.justifyContent = "flex-end";
            buttonContainer.style.marginBottom = "10px";

            const changeButton = document.createElement("button");
            changeButton.className = "ms-Button";
            changeButton.style.display = "flex";
            changeButton.style.alignItems = "center";
            changeButton.type = "button";

            const iconElement = document.createElement("i");
            iconElement.className = "ms-Icon ms-Icon--Refresh";
            iconElement.setAttribute("aria-hidden", "true");
            iconElement.style.fontSize = "16px";
            iconElement.style.marginRight = "5px";

            const textNode = document.createTextNode("Change View");

            changeButton.appendChild(iconElement);
            changeButton.appendChild(textNode);

            changeButton.addEventListener("click", () => {
                this.viewMode = this.viewMode === 'json' ? 'form' : 'json';
                this.renderControl();
            });

            buttonContainer.appendChild(changeButton);
            this.container.appendChild(buttonContainer);
        }

        if (this.viewMode === 'form') {
            this.renderFormView();
        } else {
            this.renderJsonView();
        }
    }

    private renderFormView(): void {
        if (!this.jsonData) {
            this.renderErrorMessage();
            return;
        }

        for (const key in this.jsonData) {
            if (Object.prototype.hasOwnProperty.call(this.jsonData, key)) {
                const value = this.jsonData[key];
                const control = this.createControl(key, value, this.jsonData);
                this.container.appendChild(control);
            }
        }
    }

    private renderJsonView(): void {
        const searchContainer = document.createElement("div");
        searchContainer.style.marginBottom = "10px";

        this.searchBox = document.createElement("input");
        this.searchBox.type = "text";
        this.searchBox.placeholder = "Search...";
        this.searchBox.className = "search-box"; // Apply the search-box class
        this.searchBox.addEventListener("input", () => this.renderJsonTree());

        searchContainer.appendChild(this.searchBox);
        this.container.appendChild(searchContainer);

        this.renderJsonTree();
    }

    private renderJsonTree(): void {
        const existingJsonTreeContainer = this.container.querySelector(".json-tree-container");
        if (existingJsonTreeContainer) {
            this.container.removeChild(existingJsonTreeContainer);
        }

        const jsonTreeContainer = document.createElement("div");
        jsonTreeContainer.className = "json-tree-container";

        if (!this.jsonData) {
            this.renderErrorMessage();
            return;
        }

        const tree = this.createJsonTree(this.jsonData);
        this.highlightMatches(tree, this.searchBox.value);
        jsonTreeContainer.appendChild(tree);

        this.container.appendChild(jsonTreeContainer);
    }

    private highlightJson(json: string): string {
        const searchValue = this.searchBox ? this.searchBox.value : "";

        // Escape HTML special characters
        json = json.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

        // Syntax highlighting
        json = json.replace(
          /("(\\u[a-zA-Z0-9]{4}|\\[^u]|[^\\"])*"(\s*:)?|\b(true|false|null)\b|\b\d+(\.\d+)?\b)/g,
          (match) => {
            let cls = 'number';
            if (/^"/.test(match)) {
              if (/:$/.test(match)) {
                cls = 'key';
              } else {
                cls = 'string';
              }
            } else if (/true|false/.test(match)) {
              cls = 'boolean';
            } else if (/null/.test(match)) {
              cls = 'null';
            }

            // Highlight search term
            if (searchValue && match.toLowerCase().includes(searchValue.toLowerCase())) {
              return `<span class="${cls} highlight">${match}</span>`;
            }

            return `<span class="${cls}">${match}</span>`;
          }
        );

        return json;
    }

    private renderErrorMessage(): void {
        const errorMessageContainer = document.createElement("div");
        errorMessageContainer.style.display = "flex";
        errorMessageContainer.style.justifyContent = "center";
        errorMessageContainer.style.alignItems = "center";
        errorMessageContainer.style.height = "100%";
        errorMessageContainer.style.color = "red";
        errorMessageContainer.style.fontSize = "16px";
        errorMessageContainer.textContent = this.errorMessage;

        this.container.appendChild(errorMessageContainer);
    }

    private createControl(label: string, value: any, parentObject: any): HTMLDivElement {
        const controlContainer = document.createElement("div");
        controlContainer.style.display = "flex";
        controlContainer.style.flexDirection = this.labelPosition === "Top" ? "column" : "row";
        controlContainer.style.alignItems = "flex-start";
        controlContainer.style.marginBottom = "8px";
        controlContainer.style.flex = "1"; // Make the control container fill the available space

        const labelElement = document.createElement("label");
        labelElement.textContent = label;
        labelElement.className = "ms-Label";
        labelElement.style.marginRight = this.labelPosition === "Top" ? "0" : "8px";
        labelElement.style.width = this.labelPosition === "Top" ? "auto" : `${this.maxLabelWidth}px`;
        labelElement.style.whiteSpace = "nowrap";
        labelElement.style.textAlign = "left";

        controlContainer.appendChild(labelElement);

        if (typeof value === "object" && !Array.isArray(value)) {
            // Handle nested object
            const groupContainer = document.createElement("div");
            groupContainer.style.marginLeft = this.labelPosition === "Top" ? "0" : "20px";
            groupContainer.style.flex = "1"; // Allow the group container to grow
            groupContainer.style.display = "flex";
            groupContainer.style.flexDirection = "column";

            for (const key in value) {
                if (Object.prototype.hasOwnProperty.call(value, key)) {
                    const nestedControl = this.createControl(key, value[key], value);
                    groupContainer.appendChild(nestedControl);
                }
            }
            controlContainer.appendChild(groupContainer);
        } else if (Array.isArray(value)) {
            // Handle array
            const tableContainer = document.createElement("div");
            tableContainer.style.overflow = "auto";
            tableContainer.style.flex = "1"; // Allow the table container to grow
            tableContainer.style.width = "100%";

            const table = document.createElement("table");
            table.className = "ms-Table";
            table.style.width = "100%";
            table.style.height = "100%"; // Make the table use full height

            // Create table header
            const thead = document.createElement("thead");
            const headerRow = document.createElement("tr");
            headerRow.className = "ms-Table-row";

            // Dynamically create table headers based on the first element's keys
            const firstElement = value[0];
            for (const key in firstElement) {
                if (Object.prototype.hasOwnProperty.call(firstElement, key)) {
                    const headerCell = document.createElement("th");
                    headerCell.className = "ms-Table-cell";
                    headerCell.textContent = key;
                    headerRow.appendChild(headerCell);
                }
            }
            thead.appendChild(headerRow);
            table.appendChild(thead);

            // Create table body
            const tbody = document.createElement("tbody");
            value.forEach((item, index) => {
                const row = document.createElement("tr");
                row.className = "ms-Table-row";

                for (const key in item) {
                    if (Object.prototype.hasOwnProperty.call(item, key)) {
                        const cell = document.createElement("td");
                        cell.className = "ms-Table-cell";

                        const valueElement = this.createValueControl(item[key], (newValue) => {
                            item[key] = newValue;
                            this.notifyOutputChanged();
                        });
                        cell.appendChild(valueElement);
                        row.appendChild(cell);
                    }
                }
                tbody.appendChild(row);
            });
            table.appendChild(tbody);

            tableContainer.appendChild(table);
            controlContainer.appendChild(tableContainer);
        } else {
            // Handle primitive value
            const valueElement = this.createValueControl(value, (newValue) => {
                parentObject[label] = newValue;
                this.notifyOutputChanged();
            });
            valueElement.style.flex = "1";
            controlContainer.appendChild(valueElement);
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

    private createJsonTree(data: any, key?: string): HTMLElement {
        const container = document.createElement('div');
        container.className = 'json-node';

        const wrapper = document.createElement('div');

        // Check if data is an object or array
        if (typeof data === 'object' && data !== null) {
            const isArray = Array.isArray(data);
            const openingBracket = isArray ? '[' : '{';
            const closingBracket = isArray ? ']' : '}';

            // Toggle element for expand/collapse
            const toggle = document.createElement('span');
            toggle.className = 'json-toggle';
            toggle.textContent = '- ';
            toggle.addEventListener('click', () => {
                const isHidden = content.style.display === 'none';
                content.style.display = isHidden ? 'block' : 'none';
                toggle.textContent = isHidden ? '- ' : '+ ';
            });

            // Key
            const keySpan = document.createElement('span');
            keySpan.className = 'json-key';
            if (key !== undefined) {
                keySpan.textContent = `"${key}": `;
            }

            // Brackets
            const openBracket = document.createElement('span');
            openBracket.className = 'json-bracket';
            openBracket.textContent = openingBracket;

            const closeBracket = document.createElement('span');
            closeBracket.className = 'json-bracket';
            closeBracket.textContent = closingBracket;

            // Content container
            const content = document.createElement('div');
            content.style.display = 'block';
            content.style.paddingLeft = '20px';

            // Recursively create child nodes
            for (const childKey in data) {
                if (Object.prototype.hasOwnProperty.call(data, childKey)) {
                    const childNode = this.createJsonTree(data[childKey], childKey);
                    content.appendChild(childNode);
                }
            }

            // Assemble the node
            wrapper.appendChild(toggle);
            wrapper.appendChild(keySpan);
            wrapper.appendChild(openBracket);
            wrapper.appendChild(document.createElement('br'));
            wrapper.appendChild(content);
            wrapper.appendChild(closeBracket);
        } else {
            // For primitive values
            const keySpan = document.createElement('span');
            keySpan.className = 'json-key';
            if (key !== undefined) {
                keySpan.textContent = `"${key}": `;
            }

            const valueSpan = document.createElement('span');
            const dataType = typeof data;
            valueSpan.className = `json-value ${dataType}`;
            valueSpan.textContent = dataType === 'string' ? `"${data}"` : String(data);

            wrapper.appendChild(keySpan);
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
}
