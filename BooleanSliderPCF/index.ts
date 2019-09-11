import {IInputs, IOutputs} from "./generated/ManifestTypes";
import { number, any } from "prop-types";

export class BooleanSliderPCF implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	//value of the field that is used and stored in the control
	private _value: boolean;
	
	//Reference to the PCF context input object
	private _context: ComponentFramework.Context<IInputs>;
	// Event Handler 'refreshData' reference
	private _refreshData: EventListenerOrEventListenerObject;

	// PCF framework delegate which will be assigned to this object which would be called whenever any update happens. 
	private _notifyOutputChanged: () => void;
	
	// label element created as part of this control
	//private sliderLabel: HTMLLabelElement;

	//on/off slider input element
	private _slider:HTMLInputElement;

	//contains all the elements for the control
	private _container: HTMLDivElement;

	/**
	 * Empty constructor.
	 */
	constructor()
	{

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		//control initialization code
		this._context = context;
        this._container = document.createElement("div");
        this._container.setAttribute("class", "divpadding");
		this._value = context.parameters.BooleanSlider.raw as boolean;
		this._notifyOutputChanged = notifyOutputChanged;
        this._refreshData = this.refreshData.bind(this);

        //Get option defaults
        // @ts-ignore
        let boolOptions: any[] = context.parameters.BooleanSlider.attributes.Options;

		this._slider = document.createElement("input");
		this._slider.setAttribute("id","pcfSlider");
		this._slider.setAttribute("type","checkbox");
		this._slider.addEventListener("change",this._refreshData);
		this._slider.setAttribute("checked", this._value.toString());

		var sliderSpan: HTMLSpanElement;
		sliderSpan = document.createElement("span");
        sliderSpan.setAttribute("class", "slider");

        var sliderTable: HTMLTableElement;
        sliderTable = document.createElement("table");

        //True Value Settings
        var tdTrue: HTMLTableDataCellElement;
        tdTrue = document.createElement("td");
        tdTrue.innerHTML = boolOptions[1]["Label"] != "" ? boolOptions[1]["Label"] :"Hello";
        //use d365 true colour
        this.changeCSS("input:checked + .slider", "background-color", boolOptions[1]["Color"] != "" ? boolOptions[1]["Color"] : "#2196F3");

        //False Value Settings
        var tdFalse: HTMLTableDataCellElement;
        tdFalse = document.createElement("td");
        tdFalse.innerHTML = boolOptions[0]["Label"] !="" ? boolOptions[0]["Label"]:"Goodbye";
        //use d365 false colour
        this.changeCSS(".slider", "background-color", boolOptions[0]["Color"] != "" ? boolOptions[0]["Color"] : "#ccc");

        sliderTable.appendChild(tdTrue);
        sliderTable.appendChild(tdFalse);
        sliderSpan.appendChild(sliderTable);

        //Associate all elements under a single element
		var sliderLabel: HTMLLabelElement;
		sliderLabel = document.createElement("label");
		sliderLabel.setAttribute("class","switch");
		sliderLabel.appendChild(this._slider);
		sliderLabel.appendChild(sliderSpan);
		container.appendChild(this._container);
        container.appendChild(sliderLabel);
	}

    public changeCSS(typeAndClass: string, newRule: string, newValue: string): void {

        //find the stylesheet to update
        var styleList: StyleSheetList = document.styleSheets;
        let thisCSS: any;
        for (let i = 0; i < styleList.length; i++) {
            let hreftext: string = styleList[i].href || "";
            if (hreftext.includes("BooleanSliderPCF.css", 0)) {
                thisCSS = styleList[i];
                break;
            }
        }

        //update the stylesheet based on provided properties
        var ruleSearch = thisCSS.cssRules ? thisCSS.cssRules : thisCSS.rules
        for (let i  = 0; i < ruleSearch.length; i++) {
        if (ruleSearch[i].selectorText == typeAndClass) {
            var target = ruleSearch[i]
            break;
            }
        }
        target.style[newRule] = newValue;
    }

	/**
	 * Updates the values to the internal value variable we are storing and also updates the html label that displays the value
	 * @param evt : The "Input Properties" containing the parameters, control metadata and interface functions
	 */
	public refreshData(evt: Event) : void
	{
		console.log("refresh data called");
		//debugger;
		var ToggleVal = (evt.target as HTMLInputElement).checked;
		if(ToggleVal != null){
			if(ToggleVal){
				this._value = true;
			}else{
				this._value = false;
			}
		}
		this._notifyOutputChanged();
	}
	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		this._value = context.parameters.BooleanSlider.raw as boolean;
		this._context = context;
        this._slider.checked = this._value as boolean;
	}


	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {
			BooleanSlider : this._value
		};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
	}
}