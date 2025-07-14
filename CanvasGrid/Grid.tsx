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
  // If no records, return empty array
  if (!records || Object.keys(records).length === 0) {
    return [];
  }

  const allValues = Object.values(records).map((r) => {
    try {
      // First try getFormattedValue
      let value = r?.getFormattedValue(columnKey);
      if (value !== null && value !== undefined && value !== '') {
        return String(value);
      }
      
      // If that doesn't work, try getValue
      value = r?.getValue(columnKey);
      if (value !== null && value !== undefined && value !== '') {
        return String(value);
      }
      
      // Try accessing raw data if available
      if ((r as any).raw && (r as any).raw[columnKey] !== undefined) {
        return String((r as any).raw[columnKey]);
      }
      
      return null;
    } catch (error) {
      // If there's an error accessing the column, return null
      return null;
    }
  }).filter((v): v is string => v !== null && v !== undefined && v !== '');
  
  // Return unique values, sorted
  const uniqueValues = Array.from(new Set(allValues));
  return uniqueValues.sort();
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
    isFullScreen,
    onPageSizeChange
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
  const [pageSizeInput, setPageSizeInput] = React.useState<string>("10");
  const [isComponentLoading, setIsLoading] = React.useState(false);
  const [containsFilters, setContainsFilters] = React.useState<Record<string, string[]>>({});
  const [equalsFilters, setEqualsFilters] = React.useState<Record<string, string[]>>({});

  // Close contextual menu when records change significantly (data refresh)
  // But only if the menu was open before the change
  const previousRecordCount = React.useRef<number>(0);
  React.useEffect(() => {
    const currentRecordCount = Object.keys(records).length;
    
    // Only auto-close if:
    // 1. We had minimal data before (< 5 records)
    // 2. We now have substantial data (>= 5 records) 
    // 3. A contextual menu is currently open
    // 4. This represents a significant increase in data
    if (previousRecordCount.current < 5 && 
        currentRecordCount >= 5 && 
        contextualMenuProps &&
        currentRecordCount > previousRecordCount.current) {
      setContextualMenuProps(undefined);
    }
    
    previousRecordCount.current = currentRecordCount;
  }, [records, contextualMenuProps]);

  const onContextualMenuDismissed = () => setContextualMenuProps(undefined);

  const getContextualMenuProps = React.useCallback((column: IColumn, ev: React.MouseEvent<HTMLElement>): IContextualMenuProps => {
    const colKey = column.key;
    const fieldName = column.fieldName || column.key;
    
    // Try multiple approaches to get unique values
    const uniqueValues1 = getUniqueColumnValues(records, fieldName);
    const uniqueValues2 = fieldName !== colKey ? getUniqueColumnValues(records, colKey) : [];
    
    // Combine results
    const uniqueValues = [...uniqueValues1, ...uniqueValues2];
    const finalUniqueValues = Array.from(new Set(uniqueValues)).sort();

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
          name: containsFilters[colKey]?.length > 0 ? `Filter by Contains ✓` : "Filter by Contains",
          iconProps: { iconName: "Filter" },
          subMenuProps: {
            items: [
              {
                key: "containsDropdown",
                name: "Contains",
                onRender: () => {
                  const currentFilter = containsFilters[colKey]?.[0] || "";
                  
                  return (
                    <div style={{ padding: 10, width: 260 }}>
                      <label style={{ display: "block", marginBottom: 6 }}>
                        {`Filter "${column.name}" contains`}
                        {currentFilter && <span style={{ color: '#0078d4', marginLeft: 5 }}>({currentFilter})</span>}
                      </label>
                      <input
                        key={`${colKey}-${currentFilter}`}
                        type="text"
                        defaultValue={currentFilter}
                        onChange={(e) => {
                          e.stopPropagation();
                          const val = e.target.value;
                          
                          // Update local state immediately
                          setContainsFilters(prev => {
                            const updated = val ? [val] : [];
                            
                            // Also notify parent component (but don't depend on it)
                            try {
                              onFilter(colKey, "contains", updated);
                            } catch (error) {
                              // Parent notification failed, but local filtering will still work
                            }
                            
                            return { ...prev, [colKey]: updated };
                          });
                        }}
                        onKeyDown={(e) => {
                          // Prevent menu from closing on certain keys
                          if (e.key === 'Escape') {
                            setContextualMenuProps(undefined);
                          }
                          e.stopPropagation();
                        }}
                        style={{ width: "90%", marginBottom: 8 }}
                        placeholder="Type to filter..."
                        autoFocus
                      />
                      <DefaultButton
                        text="Clear"
                        onClick={(e) => {
                          // Clear the input field directly
                          const inputElement = (e.target as HTMLElement).closest('div')?.querySelector('input');
                          if (inputElement) {
                            inputElement.value = '';
                          }
                          
                          setContainsFilters(prev => {
                            const next = { ...prev };
                            delete next[colKey];
                            
                            // Also notify parent component
                            try {
                              onFilter(colKey, null, []);
                            } catch (error) {
                              // Parent notification failed, but local filtering will still work
                            }
                            
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
          name: equalsFilters[colKey]?.length > 0 ? `Filter by Equals ✓` : "Filter by Equals",
          iconProps: { iconName: "Filter" }, // Changed from "Equals" to "Filter" since Equals icon isn't registered
          subMenuProps: {
            items: [
              {
                key: "equalsDropdown",
                name: "Equals",
                onRender: () => {
                  const selected = equalsFilters[colKey] || [];
                  
                  // Check if data is still loading by looking at record count and available columns
                  const hasMinimalData = Object.keys(records).length < 5; // Threshold for "still loading"
                  const currentUniqueValues = finalUniqueValues.length > 0 ? finalUniqueValues : [];
                  
                  // If we have no values but data might still be loading, show loading message
                  if (currentUniqueValues.length === 0 && hasMinimalData) {
                    return (
                      <div style={{ padding: 10, width: 260 }}>
                        <div>Loading data... Please try again in a moment.</div>
                      </div>
                    );
                  }
                  
                  // If we have sufficient data but still no values, show no values message
                  if (currentUniqueValues.length === 0) {
                    return (
                      <div style={{ padding: 10, width: 260 }}>
                        <div>No values found for this column</div>
                      </div>
                    );
                  }

                  const allSelected = currentUniqueValues.length > 0 && selected.length === currentUniqueValues.length;

                  const options = [
                    {
                      key: "__selectAll__",
                      text: allSelected ? "Unselect All" : "Select All",
                      selected: allSelected,
                    },
                    ...currentUniqueValues.map((v) => ({
                      key: v,
                      text: v,
                      selected: selected.includes(v)
                    })),
                  ];

                  return (
                    <div style={{ padding: 10, width: 260 }} onMouseDown={e => e.stopPropagation()}>
                      <ComboBox
                        key={`${colKey}-${selected.sort().join(',')}`}
                        placeholder="Select values"
                        label={`Filter "${column.name}" equals`}
                        multiSelect
                        selectedKey={undefined}
                        options={options}
                        styles={{
                          container: { width: 240 },
                          root: { width: 240 }
                        }}
                        onChange={(_, option) => {
                          if (!option) return;
                          setEqualsFilters(prev => {
                            const current = prev[colKey] || [];
                            let updated: string[] = [];

                            if (option.key === "__selectAll__") {
                              updated = current.length === currentUniqueValues.length ? [] : [...currentUniqueValues];
                              
                              // Close the context menu after select all
                              setTimeout(() => {
                                setContextualMenuProps(undefined);
                              }, 100);
                            } else {
                              // Use option.selected to determine action
                              if (option.selected) {
                                updated = current.includes(option.key as string)
                                  ? current
                                  : [...current, option.key as string];
                              } else {
                                updated = current.filter((k) => k !== option.key);
                              }
                            }

                            onFilter(colKey, "equals", updated);
                            return { ...prev, [colKey]: updated };
                          });
                        }}
                        useComboBoxAsMenuWidth={true}
                        allowFreeform={false}
                        autoComplete="on"
                        calloutProps={{
                          calloutMaxHeight: 300,
                          directionalHint: DirectionalHint.rightTopEdge,
                          isBeakVisible: true,
                          gapSpace: 10
                        }}
                        onMouseDown={e => e.stopPropagation()}
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
  }, [containsFilters, equalsFilters, records, columns]); // Added records and columns dependencies

  const onColumnContextMenu = React.useCallback((column?: IColumn, ev?: React.MouseEvent<HTMLElement>) => {
    if (column && ev) setContextualMenuProps(getContextualMenuProps(column, ev));
  }, [getContextualMenuProps]);

  const onColumnClick = React.useCallback((ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
    if (column && ev) setContextualMenuProps(getContextualMenuProps(column, ev));
  }, [getContextualMenuProps]);

  const items = React.useMemo(() => {
    setIsLoading(false);
    const itemsArray = sortedRecordIds.map((id) => records[id]).filter(Boolean);
    return itemsArray;
  }, [sortedRecordIds, records]);

  const filteredItems = React.useMemo(() => {
    if (items.length === 0) {
      return [];
    }

    const filtered = items.filter((item, index) => {
      try {
        // Contains filter logic
        const containsMatch = Object.entries(containsFilters).every(([col, vals]) => {
          if (!vals || vals.length === 0) return true;
          
          let itemValue = "";
          try {
            itemValue = (item.getFormattedValue(col) || "").toLowerCase();
          } catch (error) {
            try {
              itemValue = (item.getValue && item.getValue(col) || "").toString().toLowerCase();
            } catch (fallbackError) {
              return false;
            }
          }
          
          const matches = vals.some(val => {
            if (!val) return true;
            return itemValue.includes(val.toLowerCase());
          });
          
          return matches;
        });

        // Equals filter logic
        const equalsMatch = Object.entries(equalsFilters).every(([col, vals]) => {
          if (!vals || vals.length === 0) return true;
          
          let itemValue = "";
          try {
            itemValue = item.getFormattedValue(col) || "";
          } catch (error) {
            try {
              itemValue = (item.getValue && item.getValue(col) || "").toString();
            } catch (fallbackError) {
              return false;
            }
          }
          
          const matches = vals.some(val => itemValue === val);
          return matches;
        });

        return containsMatch && equalsMatch;
      } catch (error) {
        return false;
      }
    });

    return filtered;
  }, [items, containsFilters, equalsFilters]);

  const totalPages = Math.ceil(filteredItems.length / pageSize);

  React.useEffect(() => {
    if (currentPage > totalPages && totalPages > 0) {
      setCurrentPage(totalPages);
    } else if (totalPages === 0 && currentPage !== 1) {
      setCurrentPage(1);
    }
  }, [filteredItems.length, pageSize, currentPage, totalPages]);

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

  return (
    <Stack verticalFill grow style={{ width, height }}>
      <Stack.Item>
        <Stack horizontal horizontalAlign="end" verticalAlign="center" style={{ padding: "10px 20px" }}>
          <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="center">
            {!isFullScreen && <Link onClick={onFullScreen}>{resources.getString("Label_ShowFullScreen")}</Link>}
            <PrimaryButton
              text={resources.getString("Label_ResetColumns")}
              iconProps={{ iconName: "RevToggleKey" }}
              onClick={() => {
                // Reset all local state
                setHiddenColumns([]);
                setContainsFilters({});
                setEqualsFilters({});
                setCurrentPage(1);
                setPageSize(10);
                setPageSizeInput("10");
                
                // Clear all filters by calling onFilter with null for each filtered column
                Object.keys(containsFilters).forEach(col => onFilter(col, null, []));
                Object.keys(equalsFilters).forEach(col => onFilter(col, null, []));
                
                // Call parent reset function (for sorting and other external state)
                onResetColumns();
              }}
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
            value={pageSizeInput}
            onChange={e => {
              const inputValue = e.target.value;
              setPageSizeInput(inputValue);
              
              // Only update pageSize if it's a valid number
              const val = Number(inputValue);
              if (!isNaN(val) && val >= 1 && val <= 1000) {
                setPageSize(val);
                setCurrentPage(1);
                onPageSizeChange(val);
              }
            }}
            onBlur={e => {
              // On blur, ensure we have a valid value
              const val = Number(e.target.value);
              if (isNaN(val) || val < 1) {
                setPageSizeInput("10");
                setPageSize(10);
                setCurrentPage(1);
                onPageSizeChange(10);
              } else if (val > 1000) {
                setPageSizeInput("1000");
                setPageSize(1000);
                setCurrentPage(1);
                onPageSizeChange(1000);
              }
            }}
            onKeyDown={e => {
              if (e.key === 'Enter') {
                e.currentTarget.blur(); // Trigger validation
              }
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