import { useConst, useForceUpdate } from "@fluentui/react-hooks";
import * as React from "react";
import { IObjectWithKey, IRenderFunction, SelectionMode } from "@fluentui/react/lib/Utilities";
import {
  ConstrainMode,
  DetailsList,
  DetailsListLayoutMode,
  DetailsRow,
  IColumn,
  IDetailsHeaderProps,
  IDetailsListProps,
  IDetailsRowStyles,
} from "@fluentui/react/lib/DetailsList";
import { Sticky, StickyPositionType } from "@fluentui/react/lib/Sticky";
import { ContextualMenu, DirectionalHint, IContextualMenuProps } from "@fluentui/react/lib/ContextualMenu";
import { ScrollablePane, ScrollbarVisibility } from "@fluentui/react/lib/ScrollablePane";
import { Stack } from "@fluentui/react/lib/Stack";
import { Overlay } from "@fluentui/react/lib/Overlay";
import { IconButton, DefaultButton, PrimaryButton } from "@fluentui/react/lib/Button";
import { Selection } from "@fluentui/react/lib/Selection";
import { ComboBox, Link } from "@fluentui/react";

type DataSet = ComponentFramework.PropertyHelper.DataSetApi.EntityRecord & IObjectWithKey;

export interface GridProps {
  width?: number;
  height?: number;
  columns: ComponentFramework.PropertyHelper.DataSetApi.Column[];
  records: Record<string, ComponentFramework.PropertyHelper.DataSetApi.EntityRecord>;
  sortedRecordIds: string[];
  hasNextPage: boolean;
  hasPreviousPage: boolean;
  currentPage: number;
  sorting: ComponentFramework.PropertyHelper.DataSetApi.SortStatus[];
  filtering: ComponentFramework.PropertyHelper.DataSetApi.FilterExpression;
  resources: ComponentFramework.Resources;
  itemsLoading: boolean;
  highlightValue: string | null;
  highlightColor: string | null;
  setSelectedRecords: (ids: string[]) => void;
  onNavigate: (item?: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord) => void;
  onSort: (name: string, desc: boolean) => void;
  onFilter: (name: string, condition: "contains" | "equals" | null, values?: string[]) => void;
  onFullScreen: () => void;
  isFullScreen: boolean;
  item?: DataSet;
  onHideColumn: (name: string) => void;
  onPageSizeChange: (newPageSize: number) => void;
  defaultPageSize: number;
  onResetColumns: () => void;
}

const getUniqueColumnValues = (records: Record<string, any>, columnKey: string): string[] => {
  return Array.from(new Set(Object.values(records).map((r) => r?.getFormattedValue(columnKey)).filter(Boolean)));
};

const onRenderDetailsHeader: IRenderFunction<IDetailsHeaderProps> = (props, defaultRender) => {
  if (props && defaultRender) {
    return <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced>{defaultRender({ ...props })}</Sticky>;
  }
  return null;
};

const onRenderItemColumn = (
  item?: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord,
  index?: number,
  column?: IColumn
) => {
  if (column?.fieldName && item) {
    return <>{item?.getFormattedValue(column.fieldName)}</>;
  }
  return <></>;
};

export const Grid = React.memo((props: GridProps) => {
  const {
    records,
    sortedRecordIds,
    columns,
    width,
    height,
    sorting,
    filtering,
    itemsLoading,
    setSelectedRecords,
    onNavigate,
    onSort,
    onFilter,
    resources,
    highlightValue,
    highlightColor,
    onHideColumn,
    onResetColumns,
    onFullScreen,
    isFullScreen
  } = props;

  // Local state for hidden columns (by name)
  const [hiddenColumns, setHiddenColumns] = React.useState<string[]>([]);

  const forceUpdate = useForceUpdate();
  const selection: Selection = useConst(() =>
    new Selection({
      selectionMode: SelectionMode.single,
      onSelectionChanged: () => {
        const items = selection.getItems() as DataSet[];
        const selected = selection.getSelectedIndices().map((i) => items[i]?.getRecordId());
        setSelectedRecords(selected);
        forceUpdate();
      },
    })
  );

  const [contextualMenuProps, setContextualMenuProps] = React.useState<IContextualMenuProps>();
  const [currentPage, setCurrentPage] = React.useState(1);
  const [pageSize, setPageSize] = React.useState<number>(10);
  const [isComponentLoading, setIsLoading] = React.useState(false);
  const [containsFilters, setContainsFilters] = React.useState<Record<string, string[]>>({});
  const [equalsFilters, setEqualsFilters] = React.useState<Record<string, string[]>>({});

  const onContextualMenuDismissed = () => setContextualMenuProps(undefined);

  const getContextualMenuProps = React.useCallback((column: IColumn, ev: React.MouseEvent<HTMLElement>): IContextualMenuProps => {
    const colKey = column.key;
    const uniqueValues = getUniqueColumnValues(records, colKey).sort();

    return {
      items: [
        {
          key: "aToZ",
          name: resources.getString("Label_SortAZ"),
          iconProps: { iconName: "SortUp" },
          onClick: () => {
            onSort(colKey, false);
            setContextualMenuProps(undefined);
          },
        },
        {
          key: "zToA",
          name: resources.getString("Label_SortZA"),
          iconProps: { iconName: "SortDown" },
          onClick: () => {
            onSort(colKey, true);
            setContextualMenuProps(undefined);
          },
        },
        {
          key: "filterContains",
          name: "Filter by Contains",
          iconProps: { iconName: "Filter" },
          subMenuProps: {
            items: [
              {
                key: "containsDropdown",
                name: "Contains",
                onRender: () => {
                  const filterValue = (containsFilters[colKey] && containsFilters[colKey][0]) || "";
                  return (
                    <div style={{ padding: 10, width: 260 }}>
                      <label style={{ display: "block", marginBottom: 6 }}>{`Filter "${column.name}" contains`}</label>
                      <input
                        type="text"
                        // value={containsFilters[colKey]?.[0] || ''}
                        onChange={(e) => {
                          e.stopPropagation();
                          const val = e.target.value;
                          setContainsFilters(prev => {
                            const updated = val ? [val] : [];
                            onFilter(colKey, "contains", updated);
                            return { ...prev, [colKey]: updated };
                          });
                        }}
                        style={{ width: "90%", marginBottom: 8 }}
                        placeholder="Type to filter..."
                      />
                      <DefaultButton
                        text="Clear"
                        onClick={() => {
                          setContainsFilters(prev => {
                            const next = { ...prev };
                            delete next[colKey];
                            onFilter(colKey, null, []);
                            return next;
                          });
                          setContextualMenuProps(undefined);
                        }}
                        styles={{ root: { marginTop: 0 } }}
                      />
                    </div>
                  );
                },
              },
            ],
          },
        },
        {
          key: "filterEquals",
          name: "Filter by Equals",
          iconProps: { iconName: "Equals" },
          subMenuProps: {
            items: [
              {
                key: "equalsDropdown",
                name: "Equals",
                onRender: () => {
                  const selected = equalsFilters[colKey] || [];
                  const allSelected = uniqueValues.length > 0 && selected.length === uniqueValues.length;

                  const options = [
                    {
                      key: "__selectAll__",
                      text: allSelected ? "Unselect All" : "Select All",
                      selected: allSelected,
                    },
                    ...uniqueValues.map((v) => ({
                      key: v,
                      text: v,
                      selected: selected.includes(v),
                    })),
                  ];

                  return (
                    <div style={{ padding: 10, width: 260 }} onMouseDown={e => e.stopPropagation()}>
                      <ComboBox
                        label={`Filter "${column.name}" equals`}
                        multiSelect
                        selectedKey={undefined}
                        options={options}
                        onChange={(_, option) => {
                          if (!option) return;
                          setEqualsFilters(prev => {
                            const current = prev[colKey] || [];

                            let updated: string[] = [];

                            if (option.key === "__selectAll__") {
                              updated = selected.length === uniqueValues.length ? [] : [...uniqueValues];
                              setContextualMenuProps(undefined);
                            } else {
                              updated = option.selected
                                ? [...current, option.key as string]
                                : current.filter((k) => k !== option.key);
                            }

                            onFilter(colKey, "equals", updated);
                            return { ...prev, [colKey]: updated };
                          });
                        }}
                      />
                      <DefaultButton
                        text="Clear"
                        onClick={() => {
                          setEqualsFilters(prev => {
                            const next = { ...prev };
                            delete next[colKey];
                            onFilter(colKey, null, []);
                            return next;
                          });
                          setContextualMenuProps(undefined);
                        }}
                        styles={{ root: { marginTop: 8 } }}
                      />
                    </div>
                  );
                },
              },
            ],
          },
        },
        {
          key: "hideColumn",
          name: resources.getString("Label_HideColumn"),
          iconProps: { iconName: "Hide" },
          onClick: () => {
            setHiddenColumns(prev => [...prev, colKey]);
            if (onHideColumn) onHideColumn(colKey); // still call for parent notification if needed
            setContextualMenuProps(undefined);
          },
        },
      ],
      target: ev.currentTarget as HTMLElement,
      directionalHint: DirectionalHint.bottomLeftEdge,
      isBeakVisible: true,
      onDismiss: onContextualMenuDismissed,
    };
  }, [containsFilters, equalsFilters]);

  const onColumnContextMenu = React.useCallback((column?: IColumn, ev?: React.MouseEvent<HTMLElement>) => {
    if (column && ev) setContextualMenuProps(getContextualMenuProps(column, ev));
  }, [getContextualMenuProps]);

  const onColumnClick = React.useCallback((ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
    if (column && ev) setContextualMenuProps(getContextualMenuProps(column, ev));
  }, [getContextualMenuProps]);

  const items = React.useMemo(() => {
    setIsLoading(false);
    return sortedRecordIds.map((id) => records[id]).filter(Boolean);
  }, [sortedRecordIds, records]);

  const filteredItems = React.useMemo(() => {
    console.log("Filtering with:", { containsFilters, equalsFilters });
    return items.filter(item => {
      const containsMatch = Object.entries(containsFilters).every(([col, vals]) => {
        if (vals.length === 0) return true;
        const itemValue = item.getFormattedValue(col)?.toLowerCase() || "";
        const matches = vals.some(val => itemValue.includes(val.toLowerCase()));
        console.log(`Contains filter for ${col}: item="${itemValue}", filters=[${vals.join(',')}], matches=${matches}`);
        return matches;
      });

      const equalsMatch = Object.entries(equalsFilters).every(([col, vals]) => {
        if (vals.length === 0) return true;
        const itemValue = item.getFormattedValue(col) || "";
        const matches = vals.some(val => itemValue === val);
        console.log(`Equals filter for ${col}: item="${itemValue}", filters=[${vals.join(',')}], matches=${matches}`);
        return matches;
      });

      return containsMatch && equalsMatch;
    });
  }, [items, containsFilters, equalsFilters]);

  const totalPages = Math.ceil(filteredItems.length / pageSize);

  React.useEffect(() => {
    if (currentPage > totalPages) setCurrentPage(totalPages || 1);
  }, [filteredItems, totalPages]);

  const pagedItems = React.useMemo(() => {
    return filteredItems.slice((currentPage - 1) * pageSize, currentPage * pageSize);
  }, [filteredItems, currentPage, pageSize]);

  const gridColumns = React.useMemo(() => {
    return columns
      .filter((col) => !col.isHidden && col.order >= 0 && !hiddenColumns.includes(col.name))
      .sort((a, b) => a.order - b.order)
      .map((col) => {
        const sortOn = sorting?.find((s) => s.name === col.name);
        const isFiltered = containsFilters[col.name]?.length > 0 || equalsFilters[col.name]?.length > 0;
        return {
          key: col.name,
          name: col.displayName,
          fieldName: col.name,
          isSorted: !!sortOn,
          isSortedDescending: sortOn?.sortDirection === 1,
          isFiltered,
          isResizable: true,
          data: col,
          onColumnContextMenu,
          onColumnClick,
        } as IColumn;
      });
  }, [columns, sorting, containsFilters, equalsFilters, hiddenColumns]);

  const onRenderRow: IDetailsListProps["onRenderRow"] = (props) => {
    if (!props) return null;
    const item = props.item as DataSet;
    const customStyles: Partial<IDetailsRowStyles> = {};
    if (highlightColor && highlightValue && item.getValue("HighlightIndicator") == highlightValue) {
      customStyles.root = { backgroundColor: highlightColor };
    }
    return <DetailsRow {...props} styles={customStyles} />;
  };

  console.log("Full Screen Mode:", isFullScreen);

  return (
    <Stack verticalFill grow style={{ width, height }}>
      <Stack.Item>
        <Stack horizontal horizontalAlign="end" verticalAlign="center" style={{ padding: "10px 20px" }}>
          <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="center">
            {!isFullScreen && <Link onClick={onFullScreen}>{resources.getString("Label_ShowFullScreen")}</Link>}
            <PrimaryButton
              text={resources.getString("Label_ResetColumns")}
              iconProps={{ iconName: "RevToggleKey" }}
              onClick={onResetColumns}
            />
          </Stack>
        </Stack>
      </Stack.Item>
      <Stack.Item grow style={{ position: "relative", backgroundColor: "white" }}>
        <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
          <DetailsList
            columns={gridColumns}
            onRenderItemColumn={onRenderItemColumn}
            onRenderDetailsHeader={onRenderDetailsHeader}
            items={pagedItems}
            setKey={`set${currentPage}`}
            layoutMode={DetailsListLayoutMode.fixedColumns}
            constrainMode={ConstrainMode.unconstrained}
            selection={selection}
            onItemInvoked={onNavigate}
            onRenderRow={onRenderRow}
          />
          {contextualMenuProps && <ContextualMenu {...contextualMenuProps} />}
        </ScrollablePane>
        {(itemsLoading || isComponentLoading) && <Overlay />}
      </Stack.Item>
      <Stack.Item>
        <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 10 }} style={{ marginTop: 10 }}>
          <span>Rows per page</span>
          <input
            type="number"
            min={1}
            max={1000}
            value={pageSize}
            onChange={e => {
              let val = Number(e.target.value);
              if (isNaN(val) || val < 1) val = 1;
              if (val > 1000) val = 1000;
              setPageSize(val);
              setCurrentPage(1);
            }}
            style={{ width: 80, marginRight: 8 }}
          />
          <IconButton iconProps={{ iconName: 'Rewind' }} onClick={() => setCurrentPage(1)} disabled={currentPage === 1} />
          <IconButton iconProps={{ iconName: 'Previous' }} onClick={() => setCurrentPage(p => Math.max(1, p - 1))} disabled={currentPage === 1} />
          <span style={{ paddingTop: 5 }}>Page {currentPage} of {totalPages}</span>
          <IconButton iconProps={{ iconName: 'Next' }} onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))} disabled={currentPage === totalPages} />
          <IconButton iconProps={{ iconName: 'FastForward' }} onClick={() => setCurrentPage(totalPages)} disabled={currentPage === totalPages} />
        </Stack>
      </Stack.Item>
    </Stack>
  );
});

Grid.displayName = "Grid";
