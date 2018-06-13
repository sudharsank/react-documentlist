import * as React from 'react';
import styles from './DocumentCardCL.module.scss';
import { IDocumentCardCLProps } from './IDocumentCardCLProps';
import {
    DocumentCard,
    DocumentCardActivity,
    DocumentCardPreview,
    DocumentCardTitle,
    IDocumentCardPreviewProps,
    DocumentCardType
} from 'office-ui-fabric-react/lib/DocumentCard';

export class DocumentCardCL extends React.Component<IDocumentCardCLProps, {}> {
    constructor(props: IDocumentCardCLProps) {
        super(props);
    }

    public render(): React.ReactElement<IDocumentCardCLProps> {
        const { document } = this.props;
        const previewPropsUsingIcon: IDocumentCardPreviewProps = this.getPreviewIconProps();
        return (
            <div className={styles.documentCardCL}>
                <div className={"ms-Grid-col ms-sm12 ms-md6 ms-lg6 ms-xl6 "}>
                    <DocumentCard type={DocumentCardType.compact} onClickHref={document.FilePath}>
                        <DocumentCardPreview {...previewPropsUsingIcon} />
                        <div className={"ms-DocumentCard-details " + styles.textOverflow}>
                            <DocumentCardTitle title={document.FileName} shouldTruncate={true} />
                            <DocumentCardActivity
                                activity={"Created on " + document.Created}
                                people={[{ name: document.Author.title, profileImageSrc: document.Author.PictureURL }]}
                            />
                        </div>
                    </DocumentCard>
                </div>
            </div>
        );
    }

    private getPreviewIconProps = () => {
        if(this.props.document.isFile){
            return {
                previewImages: [                
                    {   
                        previewIconProps: { iconName: 'OpenFile', className: styles.iconContainer } ,
                        width: 100
                    }
                ]
            };
        }
        else {
            return {
                previewImages: [                
                    {   
                        previewIconProps: { iconName: 'FabricFolderFill', className: styles.iconContainer } ,
                        width: 100
                    }
                ]
            };
        }
    }
}