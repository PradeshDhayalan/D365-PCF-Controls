import {IInputs, IOutputs} from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { initializeIcons } from 'office-ui-fabric-react'
import {
    DocumentCard,
    DocumentCardActivity,
    DocumentCardTitle,
    DocumentCardDetails,
    DocumentCardImage,
    IDocumentCardStyles,
    IDocumentCardActivityPerson,
    DocumentCardActions
  } from 'office-ui-fabric-react/lib/DocumentCard';
  import { ImageFit } from 'office-ui-fabric-react/lib/Image';
import { string } from "prop-types";
import { DocCardsApp } from './DocumentCard'
import { ID365DocumentCardProps } from './interface'
import { D365CornerDocumentCard,Downloadfile } from './DocumentCard'
/*
  
*/
class CurrentFile implements ComponentFramework.FileObject{
    fileContent: string;
    fileName: string;
    fileSize: number;
    mimeType: string;
}

export class D365DocumentCard implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private _context: ComponentFramework.Context<IInputs>;
    private _container: HTMLDivElement;
    private _groupElements: HTMLElement[] = [];
	private _itemElements: HTMLElement[] = [];
	private _props : ID365DocumentCardProps = {
		documentCards: [],
		_context : this._context
	};
	
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
		// Add control initialization code
        this._context = context;
		this._container = container;

		initializeIcons();
		
		let recordLogicalName = (<any>context).page.entityTypeName;
		let recordId = (<any>context).page.entityId;

		if (recordId!= null) {
			this.getAnnotations(recordId,recordLogicalName).then((docCards: D365CornerDocumentCard[]) => {
				this._props._context = this._context;
				this._props.documentCards = docCards;

				ReactDOM.render(
					React.createElement(DocCardsApp,this._props), this._container
				);
			});
		}
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		// Add code to update control view
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
	}

	private attachmentDownload(id: string): Promise<Downloadfile> {
        debugger;
        return this._context.webAPI.retrieveRecord("annotations", id).then(
            function success(result) {
                let file: Downloadfile = new Downloadfile();
                file.fileContent = result["documentbody"];
                file.fileName = result["filename"];
                file.fileSize = result["filesize"];
                file.mimeType = result["mimetype"];
                return file;               
            });
	}
	
	private getAnnotations(recordId: string, recordLogicalName: string):  Promise<D365CornerDocumentCard[]>
	{
		debugger;

        let query = "?$select=filename,subject,annotationid,filesize,notetext,modifiedon,mimetype&$expand=createdby($select=fullname,entityimage_url)&$filter=filename ne null and _objectid_value eq " + recordId + " and objecttypecode eq '" + recordLogicalName + "' &$orderby=createdon desc";       
        return this._context.webAPI.retrieveMultipleRecords("annotation", query).then(
            function success(result) {
                                    let d365DocCards: D365CornerDocumentCard[] = [];
										for (let i = 0; i < result.entities.length; i++) {
											let ent = result.entities[i];
											let it = new D365CornerDocumentCard(
												ent["subject"] ? ent["subject"].toString(): "",
												ent["annotationid"].toString(),
												ent["description"] ? ent["description"].toString(): "",
												ent["filename"].split('.')[0],
												ent["filename"].split('.')[1],
												ent["filesize"],
												ent["createdby"].fullname,
												ent["modifiedon"],
												ent["createdby"].entityimage_url);

											d365DocCards.push(it);
										}
										return d365DocCards;
                                    }
           , function (error) {
               console.log(error.message);
               let items: D365CornerDocumentCard[] = [];
               return items;
           }
        );

    }
}