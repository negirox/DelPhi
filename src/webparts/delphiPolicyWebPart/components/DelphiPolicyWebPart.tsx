import * as React from 'react';
import styles from './DelphiPolicyWebPart.module.scss';
import { IDelphiPolicyWebPartProps } from './IDelphiPolicyWebPartProps';
import { getSP } from '../../../pnpjsConfig';
import { SPFI } from '@pnp/sp';
import { IDocument } from '../../../models/IDocument';
import { IDelphiPolicyWebPartState } from './IDelphiPolicyWebPartState';
import { TextField } from '@fluentui/react/lib/TextField';
import { Announced } from '@fluentui/react/lib/Announced';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn } from '@fluentui/react/lib/DetailsList';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { TooltipHost } from '@fluentui/react';
import { Icon } from '@fluentui/react/lib/Icon';
import { ScrollablePane, ScrollbarVisibility } from 'office-ui-fabric-react';
const classNames = mergeStyleSets({
  fileIconHeaderIcon: {
    padding: 0,
    fontSize: '16px',
  },
  fileIconCell: {
    textAlign: 'center',
    selectors: {
      '&:before': {
        content: '.',
        display: 'inline-block',
        verticalAlign: 'middle',
        height: '100%',
        width: '0px',
        visibility: 'hidden',
      },
    },
  },
  fileIconImg: {
    verticalAlign: 'middle',
    maxHeight: '16px',
    maxWidth: '16px',
  },
  controlWrapper: {
    display: 'flex',
    flexWrap: 'wrap',
  },
  exampleToggle: {
    display: 'inline-block',
    marginBottom: '10px',
    marginRight: '30px',
  },
  selectionDetails: {
    marginBottom: '20px',
  },
});
const controlStyles = {
  root: {
    margin: '0 30px 20px 0',
    maxWidth: '300px',
  },
};

export interface IPolicyListDocumentsState {
  columns: IColumn[];
  items: IDocument[];
  selectionDetails?: string;
  announcedMessage?: string;
}

export class PolicyListDocuments extends React.Component<{ items: IDocument[] }, IPolicyListDocumentsState> {
  private _selection: Selection;
  private _allItems: IDocument[];

  constructor(props: { items: IDocument[] }) {
    super(props);

    this._allItems = props.items;

    const columns: IColumn[] = [
      {
        key: 'column1',
        name: 'File Type',
        className: classNames.fileIconCell,
        iconClassName: classNames.fileIconHeaderIcon,
        ariaLabel: 'Column operations for File type, Press to sort on File type',
        iconName: 'Page',
        isIconOnly: true,
        fieldName: 'name',
        minWidth: 16,
        maxWidth: 16,
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => (
          <TooltipHost content={`${item.fileType} file`}>
            <img src={item.iconName} className={classNames.fileIconImg} alt={`${item.fileType} file icon`} />
          </TooltipHost>
        ),
      },
      {
        key: 'column2',
        name: 'Name',
        fieldName: 'name',
        minWidth: 210,
        maxWidth: 350,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClick,
        data: 'string',
        isPadded: true,
      },
      {
        key: 'column3',
        name: 'Date Modified',
        fieldName: 'dateModifiedValue',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        onColumnClick: this._onColumnClick,
        data: 'number',
        onRender: (item: IDocument) => {
          return <span>{item.dateModified}</span>;
        },
        isPadded: true,
      },
      /*       {
              key: 'column4',
              name: 'Modified By',
              fieldName: 'modifiedBy',
              minWidth: 70,
              maxWidth: 90,
              isResizable: true,
              isCollapsible: true,
              data: 'string',
              onColumnClick: this._onColumnClick,
              onRender: (item: IDocument) => {
                return <span>{item.modifiedBy}</span>;
              },
              isPadded: true,
            },
            {
              key: 'column5',
              name: 'File Size',
              fieldName: 'fileSizeRaw',
              minWidth: 70,
              maxWidth: 90,
              isResizable: true,
              isCollapsible: true,
              data: 'number',
              onColumnClick: this._onColumnClick,
              onRender: (item: IDocument) => {
                return <span>{item.fileSize}</span>;
              },
            }, */
    ];

    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({
          selectionDetails: this._getSelectionDetails(),
        });
      },
    });

    this.state = {
      items: this._allItems,
      columns,
      //selectionDetails: this._getSelectionDetails(),
      // announcedMessage: undefined,
    };
  }
  componentWillReceiveProps(props) {
    this.setState({ items: props.items });
  }

  public render() {
    const { columns, items } = this.state;

    return (
      <div>
        <div className={classNames.controlWrapper} style={{ display: 'none' }}>
          <TextField label="Filter by name:" onChange={this._onChangeText} styles={controlStyles} />
          <Announced message={`Number of items after filter applied: ${items.length}.`} />
        </div>
        <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto} style={{ top: '60px', zIndex: 0, minHeight: '195px', maxHeight: '195px' }}>
          {(
            <DetailsList
              items={items}
              compact={false}
              columns={columns}
              selectionMode={SelectionMode.none}
              getKey={this._getKey}
              setKey="none"
              layoutMode={DetailsListLayoutMode.justified}
              isHeaderVisible={items.length > 0}
              onItemInvoked={this._onItemInvoked}

            />
          )}
        </ScrollablePane>
      </div>
    );
  }

  public componentDidUpdate(previousProps: any, previousState: IPolicyListDocumentsState) {

  }

  private _getKey(item: any, index?: number): string {
    return item.key;
  }

  private _onChangeText = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    this.setState({
      items: text ? this._allItems.filter(i => i.name.toLowerCase().indexOf(text) > -1) : this._allItems,
    });
  };

  private _onItemInvoked(item: any): void {
    window.open(item.fileUrl)
  }

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0] as IDocument).name;
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { columns, items } = this.state;
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
        this.setState({
          announcedMessage: `${currColumn.name} is sorted ${currColumn.isSortedDescending ? 'descending' : 'ascending'
            }`,
        });
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = _copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      columns: newColumns,
      items: newItems,
    });
  };
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
  const key = columnKey as keyof T;
  return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}

/* function _generateDocuments() {
  const items: IDocument[] = [];
  for (let i = 0; i < 500; i++) {
    const randomDate = _randomDate(new Date(2012, 0, 1), new Date());
   // const randomFileSize = _randomFileSize();
    const randomFileType =   _randomFileIcon();
    let fileName = _lorem(2);
    fileName = fileName.charAt(0).toUpperCase() + fileName.slice(1).concat(`.${randomFileType.docType}`);
    let userName = _lorem(2);
    userName = userName
      .split(' ')
      .map((name: string) => name.charAt(0).toUpperCase() + name.slice(1))
      .join(' ');
    items.push({
      key: i.toString(),
      name: fileName,
      value: fileName,
      iconName: randomFileType.url,
      fileType: randomFileType.docType,
      //modifiedBy: userName,
      dateModified: randomDate.dateFormatted,
      dateModifiedValue: randomDate.value,
      //fileSize: randomFileSize.value,
     // fileSizeRaw: randomFileSize.rawSize,
    });
  }
  return items;
} */

/* function _randomDate(start: Date, end: Date): { value: number; dateFormatted: string } {
  const date: Date = new Date(start.getTime() + Math.random() * (end.getTime() - start.getTime()));
  return {
    value: date.valueOf(),
    dateFormatted: date.toLocaleDateString(),
  };
}
 */
const FILE_ICONS: { name: string }[] = [
  { name: 'accdb' },
  { name: 'audio' },
  { name: 'code' },
  { name: 'csv' },
  { name: 'docx' },
  { name: 'dotx' },
  { name: 'mpp' },
  { name: 'mpt' },
  { name: 'model' },
  { name: 'one' },
  { name: 'onetoc' },
  { name: 'potx' },
  { name: 'ppsx' },
  { name: 'pdf' },
  { name: 'photo' },
  { name: 'pptx' },
  { name: 'presentation' },
  { name: 'potx' },
  { name: 'pub' },
  { name: 'rtf' },
  { name: 'spreadsheet' },
  { name: 'txt' },
  { name: 'vector' },
  { name: 'vsdx' },
  { name: 'vssx' },
  { name: 'vstx' },
  { name: 'xlsx' },
  { name: 'xltx' },
  { name: 'xsn' },
  { name: 'html' },
];

/* function _randomFileIcon(): { docType: string; url: string } {
  const docType: string = FILE_ICONS[Math.floor(Math.random() * FILE_ICONS.length)].name;
  return {
    docType,
    url: `https://res-1.cdn.office.net/files/fabric-cdn-prod_20221209.001/assets/item-types/16/${docType}.svg`,
  };
} */

/* function _randomFileSize(): { value: string; rawSize: number } {
  const fileSize: number = Math.floor(Math.random() * 100) + 30;
  return {
    value: `${fileSize} KB`,
    rawSize: fileSize,
  };
} */

/* const LOREM_IPSUM = (
  'lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut ' +
  'labore et dolore magna aliqua ut enim ad minim veniam quis nostrud exercitation ullamco laboris nisi ut ' +
  'aliquip ex ea commodo consequat duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore ' +
  'eu fugiat nulla pariatur excepteur sint occaecat cupidatat non proident sunt in culpa qui officia deserunt '
).split(' ');
let loremIndex = 0; */
/* function _lorem(wordCount: number): string {
  const startIndex = loremIndex + wordCount > LOREM_IPSUM.length ? 0 : loremIndex;
  loremIndex = startIndex + wordCount;
  return LOREM_IPSUM.slice(startIndex, loremIndex).join(' ');
} */

export default class DelphiPolicyWebPart extends React.Component<IDelphiPolicyWebPartProps, IDelphiPolicyWebPartState> {
  private _sp: SPFI;
  constructor(props: IDelphiPolicyWebPartProps, state: IDelphiPolicyWebPartState) {
    super(props);
    this.state = {
      policies: new Array<IDocument>(),
      searchText: ''
    }
    this._sp = getSP();
    this.searchPolicy = this.searchPolicy.bind(this);
  }
  private updateInputValue(evt: React.ChangeEvent<HTMLInputElement>): void {
    const val = evt.target.value;
    this.setState({
      searchText: val,
      policies: val === '' ? new Array<IDocument>() : this.state.policies
    });
  }
  private searchPolicy(): void {
    console.log('');
    const results = this._sp.web.lists.getByTitle(this.props.listName).items.
      filter(`substringof('${this.state.searchText}',FileRef)`).expand('FieldValuesAsText,File').
      select('Title,File/Name,FieldValuesAsText/FileRef,FileLeafRef,Modified,Created')();
    results.then(x => {
      console.log(x);
      const items: IDocument[] = [];
      for (let i = 0; i < x.length; i++) {
        const modifiedDate = new Date(x[i].Modified);
        const fileext = x[i].FileLeafRef.split('.')[1];
        let docType = FILE_ICONS.filter(y => { return y.name === fileext.toLowerCase() })[0];
        docType = docType === undefined ? { name: 'photo' } : docType;
        const fileName = x[i].FileLeafRef;
        const file = x[i].FieldValuesAsText;
        const icoUrl = `https://res-1.cdn.office.net/files/fabric-cdn-prod_20230223.001/assets/item-types/32_2x/${docType.name}.png`
        items.push({
          key: i.toString(),
          name: fileName,
          value: fileName,
          iconName: icoUrl,
          fileType: docType.name,
          dateModified: modifiedDate.toLocaleDateString(),
          dateModifiedValue: modifiedDate.valueOf(),
          fileUrl: file.FileRef
        });
      }
      this.setState({ policies: items });
    })
  }
  public render(): React.ReactElement<IDelphiPolicyWebPartProps> {
    return (
      <section className={`${styles.delphiPolicyWebPart}`}>
        <div className={styles.root}>
          <div className={styles.container}>
            <div className={styles.searchBoxRoot}>
              <input type="search" className={styles.inptSearch} id='searchPolicy'
                value={this.state.searchText}
                onChange={evt => this.updateInputValue(evt)}
                onKeyDown ={(ev) => {
                  if (ev.key === 'Enter') {
                    this.searchPolicy();
                    ev.preventDefault();
                  }
                }}
                placeholder="Search For Policies..." />
              <button className={styles.btnSearch} onClick={this.searchPolicy}>Search</button>
              <Icon iconName='ProfileSearch' onClick={this.searchPolicy} className={styles.searchIcon} />
            </div>
            <div style={{ minHeight: '195px' }}>
              <PolicyListDocuments items={this.state.policies} />
            </div>
          </div>
        </div>
      </section>
    );
  }
}
