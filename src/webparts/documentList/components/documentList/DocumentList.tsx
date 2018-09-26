import * as React from 'react';
import { DisplayMode } from '@microsoft/sp-core-library';
import styles from './DocumentList.module.scss';

import { IDocumentListProps } from './IDocumentListProps';
import { IDocumentListState } from './IDocumentListState';
import { IDocumentService, DocumentService, IDocumentInfo } from '../../../../services/index';
import ConfigContainer from '../../components/configContainer/ConfigContainer';
import { DocumentItem } from '../../components/documentItem/documentItem';
import { DocumentCardCL } from '../../components/documentCardCL/DocumentCardCL';
/** Office UI Fabric Controls */
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { List } from 'office-ui-fabric-react';
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn
} from 'office-ui-fabric-react/lib/DetailsList';
import * as _ from '@microsoft/sp-lodash-subset';
import { FileTypeIcon, IconType, ImageSize } from "@pnp/spfx-controls-react/lib/FileTypeIcon";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { IRectangle } from 'office-ui-fabric-react/lib/Utilities';

import { initializeIcons } from '@uifabric/icons';
import Pagination from 'office-ui-fabric-react-pagination';

const ROWS_PER_PAGE = 3;
const MAX_ROW_HEIGHT = 250;

initializeIcons();

export default class DocumentList extends React.Component<IDocumentListProps, IDocumentListState> {

  private _columnCount: number;
  private _columnWidth: number;
  private _rowHeight: number;
  private _sorted: boolean = false;

  private documentService: IDocumentService;

  constructor(props: IDocumentListProps, state: IDocumentListState) {
    super(props);
    const _columns: IColumn[] = [
      {
        key: "column1",
        name: 'File Type',
        headerClassName: 'DetailsListExample-header--FileIcon',
        className: 'DetailsListExample-cell--FileIcon',
        iconClassName: 'DetailsListExample-Header-FileTypeIcon',
        iconName: 'Page',
        isIconOnly: true,
        fieldName: '',
        minWidth: 16,
        maxWidth: 16,
        onRender: (item: IDocumentInfo) => {
          if (item.isFile) {
            return <FileTypeIcon type={IconType.image} size={ImageSize.small}
              path={item.FilePath} />;
          }
          else {
            return <i className={"ms-Icon ms-Icon--FabricFolderFill " + styles.folderIconSmall} aria-hidden="true" />;
          }
        },
      },
      {
        key: 'column2',
        name: 'File Name',
        fieldName: 'FileName',
        minWidth: 210,
        maxWidth: 350,
        isRowHeader: true,
        isResizable: true,
        //isSorted: true,
        //isSortedDescending: false,
        onColumnClick: this._onColumnClick,
        data: 'string',
        isPadded: true,
        onRender: (item: IDocumentInfo) => {
          return <a className={styles.fileLink} href={item.FilePath}>{item.FileName}</a>;
        }
      },
      {
        key: 'column3',
        name: 'Author',
        fieldName: '',
        minWidth: 150,
        maxWidth: 200,
        isRowHeader: true,
        isResizable: true,
        data: 'string',
        isPadded: true,
        onRender: (item: IDocumentInfo) => {
          return item.Author.title;
        }
      },
      {
        key: 'column4',
        name: 'Created On',
        fieldName: 'Created',
        minWidth: 100,
        maxWidth: 150,
        isRowHeader: true,
        isResizable: true,
        data: 'date',
        isPadded: true,
      },
      {
        key: 'column5',
        name: 'Editor',
        fieldName: '',
        minWidth: 100,
        maxWidth: 150,
        isRowHeader: true,
        isResizable: true,
        data: 'string',
        isPadded: true,
        onRender: (item: IDocumentInfo) => {
          return <span>{item.Editor.title}</span>;
        }
      },
      {
        key: 'column6',
        name: 'Modified On',
        fieldName: 'Modified',
        minWidth: 100,
        maxWidth: 150,
        isRowHeader: true,
        isResizable: true,
        data: 'date',
        isPadded: true,
      },
    ];
    this.state = {
      loading: true,
      documents: [],
      displayDocuments: [],
      currentPage: 1,
      totalPages: 0,
      columns: _columns
    };
    this.documentService = this.props.serviceScope.consume(DocumentService.serviceKey as any) as IDocumentService;

    this._getItemCountForPage = this._getItemCountForPage.bind(this);
    this._getPageHeight = this._getPageHeight.bind(this);
  }

  public render(): React.ReactElement<IDocumentListProps> {
    const { loading, documents, displayDocuments, currentPage, totalPages, columns } = this.state;
    const { displayMode, title, updateProperty, doclibUrl, layoutType } = this.props;

    return (
      <div className={styles.documentList}>

        <div className={"ms-Grid"}>
          <div className={"ms-Grid-row"}>
            <div className={"ms-Grid-col ms-sm2 ms-md1 ms-lg1"}>
              <div className="ms-hiddenMdUp">
                <i className={"ms-Icon ms-Icon--Documentation " + styles.webpartTitleIcon + " " + styles.webpartTitleIconSM} aria-hidden="true"></i>
              </div>
              <div className="ms-hiddenSm">
                <i className={"ms-Icon ms-Icon--Documentation " + styles.webpartTitleIcon} aria-hidden="true"></i>
              </div>
            </div>
            <div className={"ms-Grid-col ms-sm10 ms-md11 ms-lg11 " + styles.noLeftPad}>
              <div className="ms-hiddenMdUp">
                <WebPartTitle displayMode={displayMode}
                  title={title} className={styles.webpartTitle + " " + styles.webpartTitleSM}
                  updateProperty={updateProperty} />
              </div>
              <div className="ms-hiddenSm">
                <WebPartTitle displayMode={displayMode}
                  title={title} className={styles.webpartTitle}
                  updateProperty={updateProperty} />
              </div>
            </div>
          </div>
        </div>

        {!doclibUrl && displayMode === DisplayMode.Edit &&
          <ConfigContainer
            buttonText="Configure"
            currentContext={this.props.currentContext}
            description="Configure the web part properties"
            iconText="Document List properties"
            displayButton={true} />
        }
        {!doclibUrl && displayMode === DisplayMode.Read &&
          <ConfigContainer
            buttonText="Configure"
            currentContext={this.props.currentContext}
            description="Configure the web part properties"
            iconText="Document List properties"
            displayButton={false} />
        }

        {doclibUrl && loading &&
          <Spinner size={SpinnerSize.large} label='Loading Documents...' />}

        {doclibUrl &&
          !loading &&
          documents.length === 0 &&
          <div>Sorry, no documents found</div>
        }

        {/* Box layout */}
        {doclibUrl &&
          !loading &&
          documents.length > 0 &&
          layoutType === "box" &&
          <div className="ms-Grid">
            <div className={"ms-Grid-row " + styles.rowMargin}>
              <Pagination currentPage={currentPage} totalPages={totalPages} onChange={(page) => { this._getPagedItems(page) }} />
              <List items={displayDocuments}
                onRenderCell={this._onRenderCell}
                renderCount={displayDocuments.length}
                getItemCountForPage={this._getItemCountForPage}
                getPageHeight={this._getPageHeight} />
            </div>
          </div>
        }
        {/* List layout */}
        {doclibUrl &&
          !loading &&
          documents.length > 0 &&
          layoutType === "list" &&
          <div className="ms-Grid">
            <div className="ms-Grid-row">
              <Pagination currentPage={currentPage} totalPages={totalPages} onChange={(page) => { this._getPagedItems(page) }} />
              <DetailsList
                items={displayDocuments}
                columns={columns}
                selectionMode={SelectionMode.none}
                selectionPreservedOnEmptyClick={false}
                layoutMode={DetailsListLayoutMode.justified} />
            </div>
          </div>
        }
        {/* Document Card Compact layout */}
        {doclibUrl &&
          !loading &&
          documents.length > 0 &&
          layoutType === "dccl" &&
          <div className="ms-Grid">
            <div className={"ms-Grid-row " + styles.rowMargin}>
              <Pagination currentPage={currentPage} totalPages={totalPages} onChange={(page) => { this._getPagedItems(page) }} />
              <List items={displayDocuments}
                onRenderCell={this._onRenderDCCLCell}
                renderCount={displayDocuments.length}
                getItemCountForPage={this._getItemCountForPage}
                getPageHeight={this._getPageHeight} />
            </div>
          </div>
        }
      </div>
    );
  }

  private _getPagedItems = (page: number): void => {
    let itemsperpage: number = this.props.itemsPerPage;
    let stateDocuments: IDocumentInfo[] = this.state.documents;
    if(this._sorted){
      stateDocuments = stateDocuments.reverse();
    }
    this.setState({
      ...this.state,
      currentPage: page,
      displayDocuments: this.state.documents.slice((page * itemsperpage) - itemsperpage, ((page * itemsperpage) - itemsperpage) + itemsperpage)
    });
  }

  private _onRenderCell = (item: IDocumentInfo, index: number) => {
    return (
      <DocumentItem
        key={item.Id}
        document={item} />
    );
  }

  private _onRenderDCCLCell = (item: IDocumentInfo, index: number) => {
    return (
      <DocumentCardCL
        key={item.Id}
        document={item} />
    );
  }

  public componentDidMount(): void {
    this.bindAllDocuments(this.props.doclibUrl, this.props.dateFormat, this.props.showFolder, this.props.itemsPerPage);
  }

  protected componentShouldUpdate = (newProps: IDocumentListProps) => {
    return (
      this.props.doclibUrl !== newProps.doclibUrl ||
      this.props.itemsPerPage !== newProps.itemsPerPage
    );
  }

  public componentWillReceiveProps(newProps: IDocumentListProps): void {
    if (this.props.doclibUrl !== newProps.doclibUrl ||
      this.props.layoutType !== newProps.layoutType ||
      this.props.dateFormat !== newProps.dateFormat ||
      this.props.showFolder !== newProps.showFolder ||
      this.props.itemsPerPage !== newProps.itemsPerPage) {
      this.setState({
        ...this.state,
        loading: true,
        currentPage: 1
      });
      this.bindAllDocuments(newProps.doclibUrl, newProps.dateFormat, newProps.showFolder, newProps.itemsPerPage);
    }
  }

  /** Get all the documents and store it in the state */
  public bindAllDocuments(docUrl: string, dateformat: string, showFolder: boolean, pagedItems: number) {
    this.documentService.getAllDocuments(docUrl, dateformat, showFolder)
      .then((documents: IDocumentInfo[]): void => {
        this.setState({
          loading: false,
          documents: documents,
          displayDocuments: documents.slice(0, pagedItems),
          totalPages: documents.length / pagedItems
        });
      });
  }

  /** Handle Header column click for the Details list */
  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { columns, documents } = this.state;
    let newItems: IDocumentInfo[] = documents.slice();
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter((currCol: IColumn, idx: number) => {
      return column.key === currCol.key;
    })[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    newItems = this._sortItems(newItems, currColumn.fieldName || '', currColumn.isSortedDescending);
    let page: number = this.state.currentPage;
    let itemsperpage: number = this.props.itemsPerPage;
    this.setState({
      ...this.state,
      columns: newColumns,
      documents: newItems,
      displayDocuments: this.state.documents.slice((page * itemsperpage) - itemsperpage, ((page * itemsperpage) - itemsperpage) + itemsperpage)
    });
    this._sorted = true;
  }

  /** For sorting items on the Details List */
  private _sortItems = (items: IDocumentInfo[], sortBy: string, descending = false): IDocumentInfo[] => {
    if (descending) {
      return items.sort((a: IDocumentInfo, b: IDocumentInfo) => {
        if (a[sortBy] < b[sortBy]) {
          return 1;
        }
        if (a[sortBy] > b[sortBy]) {
          return -1;
        }
        return 0;
      });
    } else {
      return items.sort((a: IDocumentInfo, b: IDocumentInfo) => {
        if (a[sortBy] < b[sortBy]) {
          return -1;
        }
        if (a[sortBy] > b[sortBy]) {
          return 1;
        }
        return 0;
      });
    }
  }

  /** To determine the Item count per page for the List component */
  private _getItemCountForPage(itemIndex: number, surfaceRect: IRectangle): number {
    if (itemIndex === 0) {
      this._columnCount = Math.ceil(surfaceRect.width / MAX_ROW_HEIGHT);
      this._columnWidth = Math.floor(surfaceRect.width / this._columnCount);
      this._rowHeight = this._columnWidth;
    }

    return this._columnCount * ROWS_PER_PAGE;
  }

  /** To determine the Page height for the List component */
  private _getPageHeight(itemIndex: number, surfaceRect: IRectangle): number {
    return this._rowHeight * ROWS_PER_PAGE;
  }
}
