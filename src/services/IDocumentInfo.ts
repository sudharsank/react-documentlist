import {IUserInfo} from '.';
export interface IDocumentInfo {
    Id: string;
    FileName: string;
    FilePath: string;
    Author?: IUserInfo;
    Editor?: IUserInfo;
    Created: string;
    Modified: string;  
    isFile?: boolean;  
}