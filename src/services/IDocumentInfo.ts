import {IUserInfo} from '.';
export interface IDocumentInfo {
    Id: string;
    FileName: string;
    FileExtn?: string;
    FilePath: string;
    Author?: IUserInfo;
    Editor?: IUserInfo;
    Created: string;
    Modified: string;
    isFile?: boolean;
}
