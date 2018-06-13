import { IDocumentInfo, IDocumentService } from '.';
import { sp } from "@pnp/sp";
import { ServiceKey } from '@microsoft/sp-core-library';
import * as moment from 'moment';

export class DocumentService implements IDocumentService {

    public static readonly serviceKey: ServiceKey<IDocumentService> = ServiceKey.create("DS:DocumentService", DocumentService);

    public getAllDocuments(documentURL: string, dateformat: string, showFolder: boolean): Promise<IDocumentInfo[]> {
        return new Promise<IDocumentInfo[]>((resolve: (documents: IDocumentInfo[]) => void, reject: (error: any) => void): void => {
            let documents: IDocumentInfo[] = [];
            sp.web.getList(documentURL).items
                .select("Id", "Title", "FileRef", "FileLeafRef", "FileSystemObjectType", "Created", "Modified", "FSObjType",
                    "Author/Id", "Author/Title", "Author/EMail",
                    "Editor/Id", "Editor/Title", "Editor/EMail")
                .expand("Author", "Editor")
                .get()
                .then((docs: any) => {
                    docs.map((doc, index) => {
                        if (!showFolder){
                            if(doc.FSObjType === 0) {
                                documents.push({
                                    Id: doc.Id,
                                    FileName: doc.FileLeafRef,
                                    FilePath: doc.FileRef,
                                    isFile: doc.FileSystemObjectType === 0,
                                    Created: (dateformat !== "" ? moment(doc.Created).format(dateformat) : moment(doc.Created).format("DD/MM/YYYY")),
                                    Modified: (dateformat !== "" ? moment(doc.Modified).format(dateformat) : moment(doc.Modified).format("DD/MM/YYYY")),
                                    Author: {
                                        Id: doc.Author.Id,
                                        title: doc.Author.Title,
                                        Email: doc.Author.EMail
                                    },
                                    Editor: {
                                        Id: doc.Editor.Id,
                                        title: doc.Editor.Title,
                                        Email: doc.Editor.EMail
                                    }
                                });
                            }
                        } 
                        else {
                            documents.push({
                                Id: doc.Id,
                                FileName: doc.FileLeafRef,
                                FilePath: doc.FileRef,
                                isFile: doc.FileSystemObjectType === 0,
                                Created: (dateformat !== "" ? moment(doc.Created).format(dateformat) : moment(doc.Created).format("DD/MM/YYYY")),
                                Modified: (dateformat !== "" ? moment(doc.Modified).format(dateformat) : moment(doc.Modified).format("DD/MM/YYYY")),
                                Author: {
                                    Id: doc.Author.Id,
                                    title: doc.Author.Title,
                                    Email: doc.Author.EMail
                                },
                                Editor: {
                                    Id: doc.Editor.Id,
                                    title: doc.Editor.Title,
                                    Email: doc.Editor.EMail
                                }
                            });
                        }
                    });
                    resolve(documents);
                });
        });
    }
}