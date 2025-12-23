import * as React from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { Announced } from '@fluentui/react/lib/Announced';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn } from '@fluentui/react/lib/DetailsList';
import { MarqueeSelection } from '@fluentui/react/lib/MarqueeSelection';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { TooltipHost } from '@fluentui/react';
import { Text } from '@fluentui/react/lib/Text';
import { Link } from '@fluentui/react/lib/Link';

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

export interface IDocLibClassState {
  columns: IColumn[];
  items: IDocument[];
  selectionDetails: string;
  isModalSelection: boolean;
  isCompactMode: boolean,
  announcedMessage?: string;
}

export interface IDocument {
  key: string;
  name: string;
  value: string;
  iconName: string;
  fileType: string;
  modifiedBy: string;
  dateModified: string;
  dateModifiedValue: number;
  fileSize: string;
  fileSizeRaw: number;
}

export const DocLibClass: React.FC = () => {
  const allItems = React.useMemo<IDocument[]>(() => {
    const items: IDocument[] = [];
    for (let i = 0; i < 10; i++) {
      const randomDate = _randomDate(new Date(2012, 0, 1), new Date());
      const randomFileSize = _randomFileSize();
      const randomFileType = _randomFileIcon();
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
        modifiedBy: userName,
        dateModified: randomDate.dateFormatted,
        dateModifiedValue: randomDate.value,
        fileSize: randomFileSize.value,
        fileSizeRaw: randomFileSize.rawSize,
      });
    }
    return items;
  }, []);

  const getKey = React.useCallback((item: IDocument) => item.key, []);

  const [items, setItems] = React.useState<IDocument[]>(allItems);
  const [columns, setColumns] = React.useState<IColumn[]>(() => {
    return [
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
        onColumnClick: onColumnClick,
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
        onColumnClick: onColumnClick,
        data: 'string',
        onRender: (item: IDocument) => (
          // eslint-disable-next-line react/jsx-no-bind
          <Link onClick={() => onItemInvoked(item)} underline>
            {item.name}
          </Link>
        ),
        isPadded: true,
      },
      {
        key: 'column3',
        name: 'Date Modified',
        fieldName: 'dateModifiedValue',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        onColumnClick: onColumnClick,
        data: 'number',
        onRender: (item: IDocument) => <span>{item.dateModified}</span>,
        isPadded: true,
      },
      {
        key: 'column4',
        name: 'Modified By',
        fieldName: 'modifiedBy',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isCollapsible: true,
        data: 'string',
        onColumnClick: onColumnClick,
        onRender: (item: IDocument) => <span>{item.modifiedBy}</span>,
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
        onColumnClick: onColumnClick,
        onRender: (item: IDocument) => <span>{item.fileSize}</span>,
      },
    ];
  });

  const [selectionDetails, setSelectionDetails] = React.useState<string>('No items selected');
  const [isModalSelection, setIsModalSelection] = React.useState<boolean>(false);
  const [isCompactMode, setIsCompactMode] = React.useState<boolean>(false);
  const [announcedMessage, setAnnouncedMessage] = React.useState<string | undefined>(undefined);

  const selectionRef = React.useRef<Selection | null>(null);
  if (!selectionRef.current) {
    selectionRef.current = new Selection({
      onSelectionChanged: () => {
        setSelectionDetails(getSelectionDetails());
      },
      getKey,
    });
  }

  React.useEffect(() => {
    // initialize selection details
    setSelectionDetails(getSelectionDetails());
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  React.useEffect(() => {
    if (!isModalSelection && selectionRef.current) {
      selectionRef.current.setAllSelected(false);
    }
  }, [isModalSelection]);

  function getSelectionDetails(): string {
    const sel = selectionRef.current!;
    const selectionCount = sel.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (sel.getSelection()[0] as IDocument).name;
      default:
        return `${selectionCount} items selected`;
    }
  }

  function onChangeCompactMode(ev: React.MouseEvent<HTMLElement>, checked?: boolean): void {
    setIsCompactMode(checked ?? false);
  }

  function onChangeModalSelection(ev: React.MouseEvent<HTMLElement>, checked?: boolean): void {
    setIsModalSelection(checked ?? false);
  }

  function onChangeText(ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text?: string): void {
    setItems(text ? allItems.filter(i => i.name.toLowerCase().indexOf(text) > -1) : allItems);
  }

  function onItemInvoked(item: IDocument): void {
    alert(`Item invoked: ${item.name}`);
  }

  function onColumnClick(ev: React.MouseEvent<HTMLElement>, column: IColumn): void {
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
        setAnnouncedMessage(`${currColumn.name} is sorted ${currColumn.isSortedDescending ? 'descending' : 'ascending'}`);
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = _copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
    setColumns(newColumns);
    setItems(newItems);
  }

  return (
    <>
      <Text>
        Note: While focusing a row, pressing enter or double clicking will execute onItemInvoked, which in this
        example will show an alert.
      </Text>
      <div className={classNames.controlWrapper}>
        <Toggle
          label="Enable compact mode"
          checked={isCompactMode}
          onChange={onChangeCompactMode}
          onText="Compact"
          offText="Normal"
          styles={controlStyles}
        />
        <Toggle
          label="Enable modal selection"
          checked={isModalSelection}
          onChange={onChangeModalSelection}
          onText="Modal"
          offText="Normal"
          styles={controlStyles}
        />
        <TextField label="Filter by name:" onChange={onChangeText} styles={controlStyles} />
        <Announced message={`Number of items after filter applied: ${items.length}.`} />
      </div>
      <div className={classNames.selectionDetails}>{selectionDetails}</div>
      <Announced message={selectionDetails} />
      {announcedMessage ? <Announced message={announcedMessage} /> : undefined}
      {isModalSelection ? (
        <MarqueeSelection selection={selectionRef.current!}>
          <DetailsList
            items={items}
            compact={isCompactMode}
            columns={columns}
            selectionMode={SelectionMode.multiple}
            setKey="multiple"
            layoutMode={DetailsListLayoutMode.justified}
            isHeaderVisible={true}
            selection={selectionRef.current!}
            selectionPreservedOnEmptyClick={true}
            onItemInvoked={onItemInvoked}
            enterModalSelectionOnTouch={true}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            checkButtonAriaLabel="select row"
          />
        </MarqueeSelection>
      ) : (
        <DetailsList
          items={items}
          compact={isCompactMode}
          columns={columns}
          selectionMode={SelectionMode.none}
          getKey={getKey}
          setKey="none"
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
          onItemInvoked={onItemInvoked}
        />
      )}
    </>
  );
};

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
  const key = columnKey as keyof T;
  return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}

function _randomDate(start: Date, end: Date): { value: number; dateFormatted: string } {
  const date: Date = new Date(start.getTime() + Math.random() * (end.getTime() - start.getTime()));
  return {
    value: date.valueOf(),
    dateFormatted: date.toLocaleDateString(),
  };
}

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
];

function _randomFileIcon(): { docType: string; url: string } {
  const docType: string = FILE_ICONS[Math.floor(Math.random() * FILE_ICONS.length)].name;
  return {
    docType,
    url: `https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/assets/item-types/16/${docType}.svg`,
  };
}

function _randomFileSize(): { value: string; rawSize: number } {
  const fileSize: number = Math.floor(Math.random() * 100) + 30;
  return {
    value: `${fileSize} KB`,
    rawSize: fileSize,
  };
}

const LOREM_IPSUM = (
  'lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut ' +
  'labore et dolore magna aliqua ut enim ad minim veniam quis nostrud exercitation ullamco laboris nisi ut ' +
  'aliquip ex ea commodo consequat duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore ' +
  'eu fugiat nulla pariatur excepteur sint occaecat cupidatat non proident sunt in culpa qui officia deserunt '
).split(' ');
let loremIndex = 0;
function _lorem(wordCount: number): string {
  const startIndex = loremIndex + wordCount > LOREM_IPSUM.length ? 0 : loremIndex;
  loremIndex = startIndex + wordCount;
  return LOREM_IPSUM.slice(startIndex, loremIndex).join(' ');
}
