import * as React from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { Announced } from '@fluentui/react/lib/Announced';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn } from '@fluentui/react/lib/DetailsList';
import { MarqueeSelection } from '@fluentui/react/lib/MarqueeSelection';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { Link, PrimaryButton, TooltipHost } from '@fluentui/react';
import { Text } from '@fluentui/react/lib/Text';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPService } from '../../../services/SPService';
import { IDocumentDto } from '../../../dtos/IDocumentDto';
import { PagedLoader } from '../../../services/paging/PagedLoader';
import { IFieldInfoDto } from '../../../dtos/IFieldInfoDto';
import { CamlHelper } from '../../../services/caml/CamlHelper';
import { set } from '@microsoft/sp-lodash-subset/lib/index';

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

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  fields: Record<string, any>;
}

interface DocLibProps {
  context: WebPartContext;
  listTitle: string;
  listViewName?: string;
}

export const DocLibStreamPaged: React.FC<DocLibProps> = ({ context, listTitle, listViewName }) => {
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);

  const [items, setItems] = React.useState<IDocument[]>([]);
  const [columns, setColumns] = React.useState<IColumn[]>([]);

  const [selectionDetails, setSelectionDetails] = React.useState<string>('No items selected');
  const [isModalSelection, setIsModalSelection] = React.useState<boolean>(false);
  const [isCompactMode, setIsCompactMode] = React.useState<boolean>(false);
  const getKey = React.useCallback((item: IDocument) => item.key, []);

  const selectionRef = React.useRef<Selection | null>(null);
  if (!selectionRef.current) {
    selectionRef.current = new Selection({
      onSelectionChanged: () => {
        setSelectionDetails(getSelectionDetails());
      },
      getKey,
    });
  }

  const [sortField, setSortField] = React.useState<string | undefined>();
  const [sortAscending, setSortAscending] = React.useState<boolean>(true);

  const [filterInput, setFilterInput] = React.useState<string>("");
  const [filterValue, setFilterValue] = React.useState<string>("");

  const service = React.useMemo(() => new SPService(context), [context]);

  const loaderRef = React.useRef<PagedLoader<IDocumentDto> | null>(null);

  React.useEffect(() => {
    const load = async (): Promise<void> => {
      try {
        setLoading(true);

        // 1. Load dynamic fields
        const fields = await service.getListFields(listTitle, listViewName);
        const viewFieldNames = fields.map(f => f.internalName);

        // 2. Build dynamic columns
        const dynamicColumns: IColumn[] = fields
          .filter(f => !f.hidden)
          .map(f => ({
            key: f.internalName,
            name: f.internalName === "DocIcon" ? "" : f.title,
            fieldName: f.internalName,
            minWidth: f.internalName === "DocIcon" ? 16 : 120,
            maxWidth: f.internalName === "DocIcon" ? 16 : 300,
            iconName: f.internalName === "DocIcon" ? "Page" : undefined,
            isResizable: true,
            isSorted: sortField === f.internalName,
            isSortedDescending: !sortAscending,
            onRender: (item: IDocument) => renderDynamicCell(item, f)
          }));

        setColumns(dynamicColumns);

        // 3. Build ViewXml (sorting + filtering)
        const whereXml = filterValue
          ? CamlHelper.buildWhereContainsXml("FileLeafRef", filterValue)
          : undefined;

        const orderByXml = sortField
          ? CamlHelper.buildOrderByXml(sortField, sortAscending)
          : undefined;

        const viewXml = CamlHelper.buildViewXml(
          viewFieldNames,
          whereXml,
          orderByXml,
          5
        );

        // 4. Start stream-based paging
        const iterator = service.getDocumentsStreamPaged(
          listTitle,
          viewXml
        );

        loaderRef.current = new PagedLoader(iterator);

        // 5. Load first page
        const firstPage = await loaderRef.current.loadNextPage();
        const mapped = firstPage.map(mapDtoToDocument);

        setItems(mapped);

      } catch (err) {
        setError(err instanceof Error ? err.message : "Unknown error");
      } finally {
        setLoading(false);
      }
    };

    load().catch(() => { });
  }, [
    listTitle,
    listViewName,
    sortField,
    sortAscending,
    filterValue
  ]);

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
    setFilterInput(text ?? "");
  }

  function onItemInvoked(item: IDocument): void {
    alert(`Item invoked: ${item.name}`);
  }

  const onColumnClick = (ev?: React.MouseEvent<HTMLElement>, column?: IColumn): void => {
    if (!column || !column.fieldName) return;

    const newAsc = column.isSortedDescending; // toggle
    setSortField(column.fieldName);
    setSortAscending(newAsc ?? true);
  };

  function renderDynamicCell(item: IDocument, field: IFieldInfoDto): JSX.Element | null {
    const value = item.fields?.[field.internalName];

    if (field.internalName === 'DocIcon') {
      return <TooltipHost content={`${item.fileType} file`}>
        <img src={item.iconName} className={classNames.fileIconImg} alt={`${item.fileType} file icon`} />
      </TooltipHost>
    } else if (field.internalName === 'LinkFilename') {
      return <Link onClick={() => onItemInvoked(item)} underline>
        {item.fields.FileLeafRef}
      </Link>;
    }

    switch (field.type) {
      case "Text":
      case "Note":
        return <span>{value}</span>;

      case "Number":
      case "Integer":
        return <span>{value}</span>;

      case "DateTime":
        return <span>{value ? new Date(value).toLocaleDateString() : ""}</span>;

      case "User":
        return <span>{value[0]?.title ?? ""}</span>;

      case "UserMulti":
        return (
          <span>
            {Array.isArray(value) ? value.map(v => v.Title).join(", ") : ""}
          </span>
        );

      case "Lookup":
        return <span>{value?.Title ?? ""}</span>;

      case "LookupMulti":
        return (
          <span>
            {Array.isArray(value) ? value.map(v => v.Title).join(", ") : ""}
          </span>
        );

      default:
        return <span>{String(value)}</span>;
    }
  }

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
        <TextField
          label="Filter by name:"
          value={filterInput}
          onChange={onChangeText}
          styles={controlStyles}
        />
        <PrimaryButton
          text="Search"
          onClick={() => setFilterValue(filterInput.toLowerCase())}
          styles={{ root: { marginTop: 8 } }}
        />
        <Announced message={`Number of items after filter applied: ${items.length}.`} />
      </div>
      <div className={classNames.selectionDetails}>{selectionDetails}</div>
      <Announced message={selectionDetails} />
      {loaderRef.current?.hasMore && (
        <PrimaryButton
          text="Load more"
          onClick={async () => {
            const nextPage = await loaderRef.current!.loadNextPage();
            const mapped = nextPage.map(mapDtoToDocument);
            setItems(prev => [...prev, ...mapped]);
          }}
        />
      )}
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
    fields: dto.fields || {},
  };
}

function getFileIconUrl(ext: string): string {
  if (!ext) return "/_layouts/15/images/icgen.png"; // generic icon
  return `/_layouts/15/images/ic${ext.toLowerCase()}.png`;
}