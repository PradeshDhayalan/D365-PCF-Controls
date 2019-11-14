import {IInputs, IOutputs} from "./generated/ManifestTypes";
import { D365CornerDocumentCard } from './DocumentCard'

export interface ID365DocumentCardProps {
	documentCards: D365CornerDocumentCard[],
	_context: ComponentFramework.Context<IInputs>,
}