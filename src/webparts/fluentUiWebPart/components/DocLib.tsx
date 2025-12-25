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
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPService } from '../../../services/SPService';
import { IDocumentDto } from '../../../dtos/IDocumentDto';

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

export interface IDocLibState {
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

interface DocLibProps {
  context: WebPartContext;
  listTitle: string;
}

export const DocLib: React.FC<DocLibProps> = ({ context, listTitle }) => {
  const [allItems, setAllItems] = React.useState<IDocument[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);

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
        fieldName: 'fileType',
        minWidth: 16,
        maxWidth: 16,
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
    const load = async (): Promise<void> => {
      try {
        const service = new SPService(context);
        const dtos = await service.getDocuments(listTitle);

        const mapped = dtos.map(mapDtoToDocument);

        setAllItems(mapped);
        setItems(mapped);
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : "Unknown error";
        setError(message);
      } finally {
        setLoading(false);
      }
    };

    (async () => {
      await load();
    })().catch(() => { });

  }, [context, listTitle]);

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

  function onChangeText(
    ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    text?: string
  ): void {
    const value = text?.toLowerCase() ?? "";
    setItems(value ? allItems.filter(i => i.name.toLowerCase().includes(value)) : allItems);
  }

  function onItemInvoked(item: IDocument): void {
    alert(`Item invoked: ${item.name}`);
  }

  const onColumnClick = React.useCallback(
    (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
      console.log("[onColumnClick] itemsLength", items.length);
      console.log("[onColumnClick] allItemsLength", allItems.length);

      const newColumns = columns.map(col => {
        if (col.key === column.key) {
          return {
            ...col,
            isSorted: true,
            isSortedDescending: !col.isSortedDescending,
          };
        }
        return {
          ...col,
          isSorted: false,
          isSortedDescending: true,
        };
      });

      const currColumn = newColumns.find(col => col.key === column.key)!;

      const newItems = copyAndSort(
        allItems, // always sort the full dataset
        currColumn.fieldName!,
        currColumn.isSortedDescending
      );

      setColumns(newColumns);
      setItems(newItems);
    },
    [items, allItems, columns] // <-- critical
  );


  console.log("RENDER itemsLength", items.length);

  if (loading) {
    return <Text>Loading documents…</Text>;
  }
  if (error) {
    return <Text>Error loading documents: {error}</Text>;
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
            columns={columns.map(col => ({ ...col, onColumnClick, }))}
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
          columns={columns.map(col => ({ ...col, onColumnClick, }))}
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

function mapDtoToDocument(dto: IDocumentDto): IDocument {
  return {
    key: dto.id.toString(),
    name: dto.name,
    value: dto.name,
    fileType: dto.fileType,
    iconName: getFileIconUrl(dto.fileType),
    modifiedBy: dto.modifiedBy ?? "",
    dateModified: dto.modified.toLocaleDateString(),
    dateModifiedValue: dto.modified.getTime(),
    fileSize: `${dto.fileSize} bytes`,
    fileSizeRaw: dto.fileSize,
  };
}

function copyAndSort<T>(
  items: T[],
  columnKey: string,
  isSortedDescending?: boolean
): T[] {
  const key = columnKey as keyof T;

  console.log(
    "[copyAndSort] start",
    { columnKey, isSortedDescending, itemsLength: items.length }
  );

  const result = items.slice().sort((a, b) => {
    const x = a[key];
    const y = b[key];

    console.log("[copyAndSort] compare", {
      key,
      a,
      b,
      x,
      y,
      typeX: typeof x,
      typeY: typeof y,
    });

    // Handle null or undefined
    const xMissing = x === null || x === undefined;
    const yMissing = y === null || y === undefined;

    if (xMissing && yMissing) return 0;
    if (xMissing) return isSortedDescending ? 1 : -1;
    if (yMissing) return isSortedDescending ? -1 : 1;

    const xVal = typeof x === "string" ? x.toLowerCase() : x;
    const yVal = typeof y === "string" ? y.toLowerCase() : y;

    if (xVal === yVal) return 0;

    const res = xVal > yVal ? (isSortedDescending ? -1 : 1) : (isSortedDescending ? 1 : -1);
    console.log("[copyAndSort] result", { xVal, yVal, res });
    return res;
  });

  console.log("[copyAndSort] end", { resultLength: result.length });

  return result;
}

function getFileIconUrl(ext: string): string {
  if (!ext) return "/_layouts/15/images/icgen.png"; // generic icon
  return `/_layouts/15/images/ic${ext.toLowerCase()}.png`;
}

