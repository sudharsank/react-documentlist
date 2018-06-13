import * as React from 'react';
import styles from './DocumentItem.module.scss';
import { IDocumentItemProps } from './IDocumentItemProps';

import { FileTypeIcon, IconType, ImageSize } from "@pnp/spfx-controls-react/lib/FileTypeIcon";
import { Link } from 'office-ui-fabric-react/lib/Link';

export class DocumentItem extends React.Component<IDocumentItemProps, {}> {
    public render(): React.ReactElement<IDocumentItemProps> {
        return (
            <div className={styles.documentItem}>
                <div className={"ms-Grid-col ms-sm12 ms-md6 ms-lg6 ms-xl4 "}>
                    <div className={"ms-Grid-row " + styles.shadowBox}>
                        <div className={"ms-Grid-col ms-sm2 ms-md2 ms-lg2 ms-xl2 " + styles.rightPadCol}>
                            <div className="ms-hiddenMdUp">
                                {!this.props.document.isFile &&
                                    <i className={"ms-Icon ms-Icon--FabricFolderFill " + styles.folderIconSmall} aria-hidden="true" />
                                }
                                {this.props.document.isFile &&
                                    <FileTypeIcon type={IconType.image} size={ImageSize.small}
                                        path={this.props.document.FilePath} />
                                }
                            </div>
                            <div className="ms-hiddenSm">
                                {!this.props.document.isFile &&
                                    <i className={"ms-Icon ms-Icon--FabricFolderFill " + styles.folderIconMedium} aria-hidden="true" />
                                }
                                {this.props.document.isFile &&
                                    <FileTypeIcon type={IconType.image} size={ImageSize.medium}
                                        path={this.props.document.FilePath} />
                                }
                            </div>
                        </div>
                        <div className={"ms-Grid-col ms-sm10 ms-md10 ms-lg10 ms-xl10 " + styles.leftPadCol}>
                            <div className={"ms-Grid-row " + styles.textOverflow}>
                                <Link className={"ms-fontWeight-semibold ms-fontSize-sPlus "}
                                    href={this.props.document.FilePath}>{this.props.document.FileName}</Link>
                            </div>
                            <div className="ms-Grid-row">
                                <span className={"ms-font-s ms-fontWeight-semibold"}>Created By </span>
                                <span className={"ms-font-s"}>{this.props.document.Author.title}</span>
                                <span className={"ms-font-s ms-fontWeight-semibold"}> On </span>
                                <span className={"ms-font-s"}>{this.props.document.Created}</span>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        );
    }
}