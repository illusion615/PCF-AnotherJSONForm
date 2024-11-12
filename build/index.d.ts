/// <reference types="powerapps-component-framework" />
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import '@fluentui/react/dist/css/fabric.css';
export declare class AnotherJSONForm implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private container;
    private jsonString;
    private labelPosition;
    private errorMessage;
    private viewMode;
    private allowSwitch;
    private jsonData;
    private notifyOutputChanged;
    private maxLabelWidth;
    private searchBox;
    private tableStyle;
    private viewContainer;
    constructor();
    init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void;
    updateView(context: ComponentFramework.Context<IInputs>): void;
    getOutputs(): IOutputs;
    destroy(): void;
    private calculateMaxLabelWidth;
    private getTextWidth;
    private renderControl;
    private renderFormView;
    private renderJsonView;
    private renderJsonTree;
    private renderErrorMessage;
    private createControl;
    private createValueControl;
    private createJsonTree;
    private highlightMatches;
    private createTable;
    private createNestedControl;
}
