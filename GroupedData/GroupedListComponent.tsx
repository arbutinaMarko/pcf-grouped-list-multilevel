import * as React from "react";
import {
  GroupedList,
  IGroup,
  IGroupedList,
} from "@fluentui/react/lib/GroupedList";
import { IColumn, DetailsRow } from "@fluentui/react/lib/DetailsList";
import {
  Selection,
  SelectionMode,
  SelectionZone,
} from "@fluentui/react/lib/Selection";
import { useConst } from "@fluentui/react-hooks";
import {
  DetailsHeader,
  IDetailsHeaderProps,
  DetailsListLayoutMode,
} from "@fluentui/react/lib/DetailsList";

//#region Types

export type Dataset = ComponentFramework.PropertyTypes.DataSet;
export type IEntityReference = ComponentFramework.EntityReference;
export type ILookupValue = ComponentFramework.LookupValue;

interface IDataSetRecord {
  getFormattedValue(columnName: string): string | undefined;
  getValue(columnName: string): unknown;
  [key: string]: unknown;
}

export type IDynamicItem = Record<
  string,
  | string
  | number
  | number[]
  | boolean
  | IEntityReference
  | IEntityReference[]
  | ILookupValue
  | ILookupValue[]
  | undefined
>;

export interface IGroupedListProps {
  data: Dataset;
  levelCount?: number;
  currentPage: number;
  loadPrevPage: () => void;
  loadNextPage: () => void;
  openRecordForm: (id: string) => void;
}

//#region Utility: Auto-detect first N levels
function getFirstNLevels(dataset: Dataset, levelCount: number): string[] {
  if (!dataset?.columns?.length) return [];
  return dataset.columns.slice(0, levelCount).map((c) => c.name);
}

//#region Groups
export function makeGroupsRecursive(
  dataset: Dataset,
  levels: string[]
): { groups: IGroup[]; items: IDynamicItem[] } {
  const groups: IGroup[] = [];
  const items: IDynamicItem[] = [];

  if (
    !dataset ||
    !Array.isArray(dataset.sortedRecordIds) ||
    !dataset.records ||
    levels.length === 0
  ) {
    return { groups, items };
  }

  const records: IDataSetRecord[] = dataset.sortedRecordIds.map(
    (id) => dataset.records[id] as unknown as IDataSetRecord
  );

  function buildGroups(
    recordsSubset: IDataSetRecord[],
    levelIndex: number,
    parentKey = "",
    currentStart = 0
  ): { groups: IGroup[]; nextIndex: number } {
    if (levelIndex >= levels.length)
      return { groups: [], nextIndex: currentStart };

    const field = levels[levelIndex];
    const uniqueValues = Array.from(
      new Set(
        recordsSubset
          .map((r) => r.getFormattedValue(field))
          .filter((v): v is string => !!v)
      )
    );

    const currentLevelGroups: IGroup[] = [];
    let localIndex = currentStart;

    for (const value of uniqueValues) {
      const filteredRecords = recordsSubset.filter(
        (r) => r.getFormattedValue(field) === value
      );

      const { groups: childGroups, nextIndex } = buildGroups(
        filteredRecords,
        levelIndex + 1,
        `${parentKey}${value}-`,
        localIndex
      );

      const itemCount =
        childGroups.length > 0
          ? childGroups.reduce((sum, g) => sum + g.count, 0)
          : filteredRecords.length;

      if (childGroups.length === 0) {
        for (const record of filteredRecords) {
          const dynamicItem: IDynamicItem = {};
          for (const column of dataset.columns) {
            const key = column.name;
            const formatted =
              record.getFormattedValue(key) ?? (record.getValue(key) as string);
            if (formatted !== undefined) {
              dynamicItem[key] = formatted;
            }
          }
          items.push(dynamicItem);
        }
      }

      currentLevelGroups.push({
        key: `${parentKey}${value}`,
        name: value,
        startIndex: localIndex,
        count: itemCount,
        level: levelIndex,
        isCollapsed: true,
        children: childGroups.length > 0 ? childGroups : undefined,
      });

      localIndex += itemCount;
    }

    return { groups: currentLevelGroups, nextIndex: localIndex };
  }

  // Top-level call
  const { groups: topGroups } = buildGroups(records, 0);
  groups.push(...topGroups);

  return { groups, items };
}

//#region Columns + Items
export function makeColumnsAndItems(
  dataset: Dataset,
  levels: string[]
): { items: IDynamicItem[]; columns: IColumn[] } {
  if (
    !dataset ||
    !Array.isArray(dataset.sortedRecordIds) ||
    !dataset.records ||
    !Array.isArray(dataset.columns)
  ) {
    return { items: [], columns: [] };
  }

  const columns: IColumn[] = dataset.columns
    .filter((col) => !levels.includes(col.name))
    .map((column) => ({
      name: column.displayName,
      fieldName: column.name,
      minWidth: column.visualSizeFactor ?? 100,
      key: column.name,
    }));

  const items: IDynamicItem[] = [];
  const records: IDataSetRecord[] = dataset.sortedRecordIds.map(
    (id) => dataset.records[id] as unknown as IDataSetRecord
  );

  function processGroupHierarchy(
    recordsSubset: IDataSetRecord[],
    levelIndex: number,
    parentValues: Record<string, string>
  ): void {
    if (levelIndex >= levels.length) {
      for (const record of recordsSubset) {
        const dynamicItem: IDynamicItem = {};
        for (const column of dataset.columns) {
          const key = column.name;
          const formatted =
            record.getFormattedValue(key) ?? (record.getValue(key) as string);
          if (formatted !== undefined) {
            dynamicItem[key] = formatted;
          }
        }

        for (const [lvl, val] of Object.entries(parentValues)) {
          dynamicItem[lvl] = val;
        }
        items.push(dynamicItem);
      }
      return;
    }

    const currentField = levels[levelIndex];
    const uniqueValues = Array.from(
      new Set(
        recordsSubset
          .map((r) => r.getFormattedValue(currentField))
          .filter((v): v is string => !!v)
      )
    );

    for (const value of uniqueValues) {
      const filteredRecords = recordsSubset.filter(
        (r) => r.getFormattedValue(currentField) === value
      );

      processGroupHierarchy(filteredRecords, levelIndex + 1, {
        ...parentValues,
        [currentField]: value,
      });
    }
  }

  processGroupHierarchy(records, 0, {});
  return { items, columns };
}

//#region React Component
export const GroupedListComponent = ({
  data,
  levelCount = 2,
  currentPage,
  loadPrevPage,
  loadNextPage,
  openRecordForm
}: IGroupedListProps): JSX.Element => {
  const root = React.useRef<IGroupedList | null>(null);
  const [columns, setColumns] = React.useState<IColumn[]>([]);
  const [items, setItems] = React.useState<IDynamicItem[]>([]);
  const [groups, setGroups] = React.useState<IGroup[]>([]);
  const selection = useConst(() => new Selection());

  const layoutMode = DetailsListLayoutMode.fixedColumns;

  const checkCellStyle = {
    marginRight: `${levelCount * 36}px`,
  };

  React.useEffect(() => {
    if (!data?.sortedRecordIds?.length) return;

    const levels = getFirstNLevels(data, levelCount);

    const { groups, items } = makeGroupsRecursive(data, levels);
    const { columns } = makeColumnsAndItems(data, levels);

    setItems(items);
    setGroups(groups);
    setColumns(columns);
    selection.setItems(items, true);
  }, [data, levelCount]);

  const onRenderCell = (
    nestingDepth?: number,
    item?: IDynamicItem,
    itemIndex?: number,
    group?: IGroup
  ): React.ReactNode =>
    item && typeof itemIndex === "number" && itemIndex > -1 ? (
      <div onDoubleClick={() => openRecordForm(item.id as string)}>
        <DetailsRow
          key={item.id as string}
          columns={columns.map((col) => ({
            ...col,
            onRender: (fieldItem: IDynamicItem) => {
              const value = fieldItem[col.fieldName!];
              if (
                typeof value === "string" &&
                /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}.\d{3}Z$/.test(value)
              ) {
                const date = new Date(value);
                const day = date.getDate().toString().padStart(2, "0");
                const month = (date.getMonth() + 1).toString().padStart(2, "0");
                const year = date.getFullYear();
                return <span>{`${day}.${month}.${year}`}</span>;
              }
              if (value instanceof Date) {
                const day = value.getDate().toString().padStart(2, "0");
                const month = (value.getMonth() + 1)
                  .toString()
                  .padStart(2, "0");
                const year = value.getFullYear();
                return <span>{`${day}.${month}.${year}`}</span>;
              }
              if (
                value &&
                typeof value === "object" &&
                ("id" in value || "entityType" in value)
              ) {
                return (
                  <span
                    style={{
                      cursor: "pointer",
                      textDecoration: "underline",
                      color: "#0078d4",
                    }}
                    onClick={() => openRecordForm(fieldItem.id as string)}
                  >
                    {("name" in value && value.name) ??
                      ("id" in value && value.id) ??
                      ""}
                  </span>
                );
              }
              return <span>{value as string}</span>;
            },
          }))}
          groupNestingDepth={nestingDepth}
          item={item}
          itemIndex={itemIndex}
          selection={selection}
          selectionMode={SelectionMode.multiple}
          group={group}
        />
      </div>
    ) : null;

  return (
    <div className="ms-DetailsList component-wrapper">
      <div className="data-scrollable">
        <DetailsHeader
          columns={columns}
          selectionMode={SelectionMode.multiple}
          selection={selection}
          layoutMode={1}
          ariaLabelForSelectAllCheckbox="Toggle selection"
          ariaLabelForSelectionColumn="Toggle selection"
          onRenderColumnHeaderTooltip={(tooltipHostProps) => (
            <span>{tooltipHostProps?.column?.name}</span>
          )}
          styles={{
            cellIsCheck: checkCellStyle,
          }}
        />
        <SelectionZone
          selection={selection}
          selectionMode={SelectionMode.multiple}
        >
          <GroupedList
            componentRef={root}
            items={items}
            groups={groups}
            groupProps={{
              showEmptyGroups: false,
              isAllGroupsCollapsed: false,
            }}
            onRenderCell={onRenderCell}
          />
        </SelectionZone>
      </div>
      {/* Footer */}
      <div className="groupby-footer-bar">
        <div className="footer-info">
          {(() => {
            return `Page: ${currentPage} | Rows showing: ${data.sortedRecordIds.length}`;
          })()}
        </div>
        <div className="pagination-buttons">
          <button
            className="pagination-btn"
            onClick={loadPrevPage}
            aria-label="Previous page"
            disabled={!data.paging.hasPreviousPage}
          >
            <svg width="20" height="20" viewBox="0 0 20 20" fill="none">
              <path
                d="M13 15L8 10L13 5"
                stroke="#0078d4"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              />
            </svg>
          </button>
          <button
            className="pagination-btn"
            onClick={loadNextPage}
            aria-label="Next page"
            disabled={!data.paging.hasNextPage}
          >
            <svg width="20" height="20" viewBox="0 0 20 20" fill="none">
              <path
                d="M7 5L12 10L7 15"
                stroke="#0078d4"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              />
            </svg>
          </button>
        </div>
      </div>
    </div>
  );
};
