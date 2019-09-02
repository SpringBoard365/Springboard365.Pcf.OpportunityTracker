import { IInputs, IOutputs } from "./generated/ManifestTypes";
const linkParameterName = "link";

export class DisplayStatusGrid implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private contextObj: ComponentFramework.Context<IInputs>;
    private mainContainer: HTMLDivElement;
    private gridContainer: HTMLDivElement;
    private refreshControlButton: HTMLButtonElement;
    private entityId: string;
    private globalOptionSetName: string;
    private statusEntityName: string;
    private relationshipName: string;
    private nameFieldName: string;
    private statusFieldName: string;
    private dateFieldName: string;
    private linkFieldName: string;
    private colourPalletteObject: ColourPallette = new ColourPallette();

	constructor()
	{
	}

	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
        this.initiateProperties(context);
        
        this.mainContainer = document.createElement("div");
        this.gridContainer = document.createElement("div");
        this.gridContainer.classList.add("DisplayStatusControl_grid-container");
        this.mainContainer.appendChild(this.gridContainer);

        if (context.parameters.showRefreshButton.raw == "Yes") {
            this.refreshControlButton = document.createElement("button");
            this.refreshControlButton.setAttribute("type", "button");
            this.refreshControlButton.innerText = context.resources.getString("PCF_DisplayStatusControl_Refresh_ButtonLabel");
            this.refreshControlButton.classList.add("DisplayStatusControl_RefreshButton_Style");
            this.refreshControlButton.addEventListener("click", this.refreshControlButtonClick.bind(this));
            this.mainContainer.appendChild(this.refreshControlButton);
        }
    
        this.mainContainer.classList.add("DisplayStatusControl_main-container");
        container.appendChild(this.mainContainer);
    }
    
    public initiateProperties(context: ComponentFramework.Context<IInputs>) {
        context.mode.trackContainerResize(true);
        // @ts-ignore
        this.entityId = context.mode.contextInfo.entityId;
        this.globalOptionSetName = context.parameters.globalOptionSetName.raw;
        this.statusEntityName = context.parameters.statusEntityName.raw;
        this.relationshipName = context.parameters.relationshipName.raw;
        this.nameFieldName = context.parameters.nameFieldName.raw;
        this.statusFieldName = context.parameters.statusFieldName.raw;
        this.dateFieldName = context.parameters.dateFieldName.raw;
        this.linkFieldName = context.parameters.linkFieldName.raw;
    }
    
	public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        this.contextObj = context;
        this.DisplayStatuses(context);
	}

	public getOutputs(): IOutputs
	{
		return {};
	}

	public destroy(): void
	{
    }

    public DisplayStatuses(context: ComponentFramework.Context<IInputs>)
    {
        var thisCtrl = this;
        this.colourPalletteObject = this.GetGlobalOptionSetColourPallette(context);
        var query: string = "?$filter=" + encodeURIComponent(this.relationshipName + " eq " + this.entityId);
        var entityName: string = this.statusEntityName;
        var maxPageSize: number = 50;
        context.webAPI.retrieveMultipleRecords(entityName, query, maxPageSize).then(
            function(result){
                var dataToReturn: Status = new Status();
                result.entities.forEach(function(record) {
                    if (record == null)
                    {
                        return;
                    }
                    var thisRecord = record;
                    console.log(thisRecord[thisCtrl.nameFieldName]);
                    var statusName =  thisRecord[thisCtrl.statusFieldName];
                    var statusColour = thisCtrl.colourPalletteObject.pallette.find(element => element.status == statusName);
                    var recordObject: StatusObject = {
                        name: thisRecord[thisCtrl.nameFieldName],
                        status: statusName,
                        date: thisRecord[thisCtrl.dateFieldName],
                        colour: statusColour != null ? statusColour.colour : "#000000",
                        link: thisRecord[thisCtrl.linkFieldName]
                    };
                    dataToReturn.statusList.push(recordObject);
                });
                while (thisCtrl.gridContainer.firstChild) {
                    thisCtrl.gridContainer.removeChild(thisCtrl.gridContainer.firstChild);
                }
                thisCtrl.gridContainer.appendChild(thisCtrl.drawDisplayStatusGrid(dataToReturn));
            },
            function(error){
                context.navigation.openAlertDialog(
                { 
                    text: "Error: " + error.errorCode + " - " + error.message
                });
            }
        );
    }

    private refreshControlButtonClick(event: Event): void
    {
        this.DisplayStatuses(this.contextObj);
    }

    private clickThroughButtonClick(event: Event): void {
        var clickedElement = event.target as HTMLElement;
        var linkToOpen = clickedElement.getAttribute(linkParameterName);
        if (linkToOpen == null) {
            console.log("No link to open.");
            return;
        }
        this.contextObj.navigation.openUrl(linkToOpen);
    }

    private drawDisplayStatusGrid(obj: Status): HTMLDivElement
    {
        let gridBody: HTMLDivElement = document.createElement("div");
        if(obj == null || obj.statusList == null || obj.statusList.length < 1)
        {
            let noRecordLabel: HTMLDivElement = document.createElement("div");
            noRecordLabel.classList.add("DisplayStatusControl_grid-norecords");
            noRecordLabel.style.width = this.contextObj.mode.allocatedWidth - 25 + "px";
            noRecordLabel.innerText = this.contextObj.resources.getString("PCF_DisplayStatusControl_No_Record_Found");            
            gridBody.appendChild(noRecordLabel);

            return gridBody;
        }

        for (let status of obj.statusList) {
            let gridRecord: HTMLDivElement = document.createElement("div");
            gridRecord.setAttribute(linkParameterName, status.link);
            gridRecord.classList.add("DisplayStatusControl_grid-item");
            gridRecord.addEventListener("click", this.clickThroughButtonClick.bind(this));
            gridRecord.style.backgroundColor = status.colour;

            let headerElement = document.createElement("p");
            headerElement.classList.add("DisplayStatusControl_grid-text");
            headerElement.textContent = status.name;
            gridRecord.appendChild(headerElement);

            let statusElement = document.createElement("p");
            statusElement.classList.add("DisplayStatusControl_grid-text");
            statusElement.textContent = status.status;
            gridRecord.appendChild(statusElement);

            let dateElement = document.createElement("p");
            dateElement.classList.add("DisplayStatusControl_grid-text");
            dateElement.textContent = status.date;
            gridRecord.appendChild(dateElement);

            gridBody.appendChild(gridRecord);
        }

        return gridBody;
    }

    private GetGlobalOptionSetColourPallette(context: ComponentFramework.Context<IInputs>):ColourPallette
    {
        var colourPallette: ColourPallette = new ColourPallette();
        let requestUrl = "/api/data/v9.1/GlobalOptionSetDefinitions(Name='" + this.globalOptionSetName + "')";
	
		let request = new XMLHttpRequest();
		request.open("GET", requestUrl, true);
		request.setRequestHeader("OData-MaxVersion", "4.0");
		request.setRequestHeader("OData-Version", "4.0");
		request.setRequestHeader("Accept", "application/json");
		request.setRequestHeader("Content-Type", "application/json; charset=utf-8");
		request.onreadystatechange = () => {
			if (request.readyState === 4) {
				request.onreadystatechange = null;
				if (request.status === 200) {
                    debugger;
					let entityMetadata = JSON.parse(request.response);
                    let options = entityMetadata.Options;
					for (var i = 0; i < options.length; i++) {
                        colourPallette.pallette.push({ status: options[i].Label.UserLocalizedLabel.Label, colour: options[i].Color, value: options[i].Value});			
					}
                }
                else {
					let errorText = request.responseText;
					console.log(errorText);
				}
			}
		};
        request.send();
        
        return colourPallette;
    }
}
class Status {
    statusList : StatusObject[] = new Array() 
}

class StatusObject {
    name: string
    date: string
    status: string
    colour: string
    link:string
}

class ColourPallette {
    pallette: StatusColour[] = new Array()
}

class StatusColour {
    status: string
    colour: string
    value: string
}