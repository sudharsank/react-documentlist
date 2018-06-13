import {IDocumentInfo} from '.';
export interface IDocumentService {
    getAllDocuments(documentURL: string, dateformat: string, showFolder: boolean): Promise<IDocumentInfo[]>;
}