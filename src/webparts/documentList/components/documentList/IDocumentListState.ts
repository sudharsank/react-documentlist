import { IDocumentInfo } from '../../../../services/index';
import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
export interface IDocumentListState {
    loading: boolean;
    documents: IDocumentInfo[];   
    columns: IColumn[]; 
}