import { IInputs, IOutputs } from "./generated/ManifestTypes";
const columns = ["header", "status"];
const linkParameterName = "link";

export class DisplayStatusGrid implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private contextObj: ComponentFramework.Context<IInputs>;
    private mainContainer: HTMLDivElement;
    private gridContainer: HTMLDivElement;
    private refreshControlButton: HTMLButtonElement;

	constructor()
	{
	}

	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
        context.mode.trackContainerResize(true);

        this.mainContainer = document.createElement("div");
        this.gridContainer = document.createElement("div");
        this.gridContainer.classList.add("DisplayStatusControl_grid-container");

        this.refreshControlButton = document.createElement("button");
        this.refreshControlButton.setAttribute("type", "button");
        this.refreshControlButton.innerText = context.resources.getString("PCF_DisplayStatusControl_Refresh_ButtonLabel");
        this.refreshControlButton.classList.add("DisplayStatusControl_RefreshButton_Style");
        this.refreshControlButton.addEventListener("click", this.refreshControlButtonClick.bind(this));

        this.mainContainer.appendChild(this.gridContainer);
        this.mainContainer.appendChild(this.refreshControlButton);
        this.mainContainer.classList.add("DisplayStatusControl_main-container");
        container.appendChild(this.mainContainer);
	}
    
	public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        this.contextObj = context;
        while (this.gridContainer.firstChild) {
            this.gridContainer.removeChild(this.gridContainer.firstChild);
        }
        this.gridContainer.appendChild(this.drawDisplayStatusGrid(context.parameters.jsonGrid.raw));
	}

	public getOutputs(): IOutputs
	{
		return {};
	}

	public destroy(): void
	{
    }

    private refreshControlButtonClick(event: Event): void {
        alert("Refresh Tracker (Does nothing yet)");
    }

    private clickThroughButtonClick(event: Event): void {
        var clickedElement = event.target as HTMLElement;
        var linkToOpen = clickedElement.getAttribute(linkParameterName);
        if (linkToOpen == null) {
            console.log("No link to open.");
            return;
        }
        alert("Open Link: " + linkToOpen);
        this.contextObj.navigation.openUrl(linkToOpen);
    }

    private drawDisplayStatusGrid(jsonValue: string): HTMLDivElement {
        let gridBody: HTMLDivElement = document.createElement("div");
        var obj = null;
        try {
            obj = JSON.parse(jsonValue);
        }
        catch (err) {
            console.log(err);
        }
        if (obj != null && obj.statusList != null && obj.statusList.length > 0) {
            for (let status of obj.statusList) {
                let gridRecord: HTMLDivElement = document.createElement("div");
                gridRecord.setAttribute(linkParameterName, status[linkParameterName]);
                gridRecord.classList.add("DisplayStatusControl_grid-item");
                gridRecord.addEventListener("click", this.clickThroughButtonClick.bind(this));
                gridRecord.style.backgroundColor = status.colour;
                
                for (var i = 0; i < columns.length; i++) {
                    let paragraphElement = document.createElement("p");
                    paragraphElement.classList.add("DisplayStatusControl_grid-text");
                    var parameterName = columns[i];
                    paragraphElement.textContent = status[parameterName];
                    gridRecord.appendChild(paragraphElement);
                }

                gridBody.appendChild(gridRecord);
            }
        }
        else {
            let noRecordLabel: HTMLDivElement = document.createElement("div");
            noRecordLabel.classList.add("DisplayStatusControl_grid-norecords");
            noRecordLabel.style.width = this.contextObj.mode.allocatedWidth - 25 + "px";
            noRecordLabel.innerText = this.contextObj.resources.getString("PCF_DisplayStatusControl_No_Record_Found");            
            gridBody.appendChild(noRecordLabel);
        }

        return gridBody;
    }
}