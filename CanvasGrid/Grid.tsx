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
import { DefaultButton, IconButton } from "@fluentui/react/lib/Button";
import { Selection } from "@fluentui/react/lib/Selection";
import { Link } from "@fluentui/react/lib/Link";
import { on } from "events";
import { Dropdown } from "@fluentui/react";

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
	onFilter: (name: string, filtered: boolean) => void;
	onFullScreen: () => void;
	isFullScreen: boolean;
	item?: DataSet;
	onHideColumn: (name: string) => void;
	onPageSizeChange: (newPageSize: number) => void;
	defaultPageSize: number;
}

const onRenderDetailsHeader: IRenderFunction<IDetailsHeaderProps> = (props, defaultRender) => {
	if (props && defaultRender) {
		return (
			<Sticky stickyPosition={StickyPositionType.Header} isScrollSynced>
				{defaultRender({
					...props,
				})}
			</Sticky>
		);
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
		hasNextPage,
		sorting,
		filtering,
		// currentPage,
		itemsLoading,
		setSelectedRecords,
		onNavigate,
		onSort,
		onFilter,
		resources,
		highlightValue,
		highlightColor,
		onHideColumn	} = props;

	const forceUpdate = useForceUpdate();
	const onSelectionChanged = React.useCallback(() => {
		const items = selection.getItems() as DataSet[];
		const selected = selection.getSelectedIndices().map((index: number) => {
			const item: DataSet | undefined = items[index];
			return item && items[index].getRecordId();
		});

		setSelectedRecords(selected);
		forceUpdate();
	}, [forceUpdate]);

	const selection: Selection = useConst(() => {
		return new Selection({
			selectionMode: SelectionMode.single,
			onSelectionChanged: onSelectionChanged,
		});
	});

	const [isComponentLoading, setIsLoading] = React.useState<boolean>(false);

	const [contextualMenuProps, setContextualMenuProps] = React.useState<IContextualMenuProps>();

	const [currentPage, setCurrentPage] = React.useState(1);
	const [pageSize, setPageSize] = React.useState<number>(10);

	const onContextualMenuDismissed = React.useCallback(() => {
		setContextualMenuProps(undefined);
	}, [setContextualMenuProps]);

	const getContextualMenuProps = React.useCallback(
		(column: IColumn, ev: React.MouseEvent<HTMLElement>): IContextualMenuProps => {
			const menuItems = [
				{
					key: "aToZ",
					name: resources.getString("Label_SortAZ"),
					iconProps: { iconName: "SortUp" },
					canCheck: true,
					checked: column.isSorted && !column.isSortedDescending,
					disable: (column.data as ComponentFramework.PropertyHelper.DataSetApi.Column).disableSorting,
					onClick: () => {
						onSort(column.key, false);
						setContextualMenuProps(undefined);
						setIsLoading(true);
					},
				},
				{
					key: "zToA",
					name: resources.getString("Label_SortZA"),
					iconProps: { iconName: "SortDown" },
					canCheck: true,
					checked: column.isSorted && column.isSortedDescending,
					disable: (column.data as ComponentFramework.PropertyHelper.DataSetApi.Column).disableSorting,
					onClick: () => {
						onSort(column.key, true);
						setContextualMenuProps(undefined);
						setIsLoading(true);
					},
				},
				{
					key: "filter",
					name: resources.getString("Label_DoesNotContainData"),
					iconProps: { iconName: "Filter" },
					canCheck: true,
					checked: column.isFiltered,
					onClick: () => {
						onFilter(column.key, column.isFiltered !== true);
						setContextualMenuProps(undefined);
						setIsLoading(true);
					},
				},
				{
					key: "hideColumn",
					name: resources.getString("Label_HideColumn"),
					iconProps: { iconName: "eye" },
					canCheck: false,
					onClick: () => {
						onHideColumn(column.key);
						setContextualMenuProps(undefined);
						setIsLoading(true);
					},
				},
			];
			return {
				items: menuItems,
				target: ev.currentTarget as HTMLElement,
				directionalHint: DirectionalHint.bottomLeftEdge,
				gapSpace: 10,
				isBeakVisible: true,
				onDismiss: onContextualMenuDismissed,
			};
		},
		[setIsLoading, onFilter, setContextualMenuProps, onHideColumn]
	);

	const onColumnContextMenu = React.useCallback(
		(column?: IColumn, ev?: React.MouseEvent<HTMLElement>) => {
			if (column && ev) {
				setContextualMenuProps(getContextualMenuProps(column, ev));
			}
		},
		[getContextualMenuProps, setContextualMenuProps]
	);

	const onColumnClick = React.useCallback(
		(ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
			if (column && ev) {
				setContextualMenuProps(getContextualMenuProps(column, ev));
			}
		},
		[getContextualMenuProps, setContextualMenuProps]
	);

	const items: (DataSet | undefined)[] = React.useMemo(() => {
		setIsLoading(false);

		const sortedRecords: (DataSet | undefined)[] = sortedRecordIds.map((id) => {
			const record = records[id];
			return record;
		});

		return sortedRecords;
	}, [records, sortedRecordIds, hasNextPage, setIsLoading]);

	const totalPages = Math.ceil(items.length / pageSize);

	React.useEffect(() => {
		if (currentPage > totalPages) {
			setCurrentPage(totalPages);
		}
	}, [items, totalPages, pageSize]);

	const pagedItems = React.useMemo(() => {
		return items.slice((currentPage - 1) * pageSize, currentPage * pageSize);
	}, [items, currentPage, pageSize]);






	const gridColumns = React.useMemo(() => {
		return columns
			.filter((col) => !col.isHidden && col.order >= 0)
			.sort((a, b) => a.order - b.order)
			.map((col) => {
				const sortOn = sorting?.find((s) => s.name === col.name);
				const filtered = filtering?.conditions?.find((f) => f.attributeName == col.name);
				return {
					key: col.name,
					name: col.displayName,
					fieldName: col.name,
					isSorted: sortOn != null,
					isSortedDescending: sortOn?.sortDirection === 1,
					isResizable: true,
					isFiltered: filtered != null,
					data: col,
					onColumnContextMenu: onColumnContextMenu,
					onColumnClick: onColumnClick,
				} as IColumn;
			});
	}, [columns, sorting, onColumnContextMenu, onColumnClick, filtering?.conditions]);

	const rootContainerStyle: React.CSSProperties = React.useMemo(() => {
		return {
			height: height,
			width: width,
		};
	}, [width, height]);

	const onRenderRow: IDetailsListProps["onRenderRow"] = (props) => {
		const customStyles: Partial<IDetailsRowStyles> = {};

		if (props?.item) {
			const item = props.item as DataSet;

			if (highlightColor && highlightValue && item.getValue("HighlightIndicator") == highlightValue) {
				customStyles.root = { backgroundColor: highlightColor };
			}
			return <DetailsRow {...props} styles={customStyles} />;
		}

		return null;
	};


	const pageSizeOptions = [10, 25, 50, 100];


	return (
		<Stack verticalFill grow style={rootContainerStyle}>
			<Stack.Item grow style={{ position: "relative", backgroundColor: "white" }}>
				<ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
					<DetailsList
						columns={gridColumns}
						onRenderItemColumn={onRenderItemColumn}
						onRenderDetailsHeader={onRenderDetailsHeader}
						items={pagedItems}
						setKey={`set${currentPage}`}
						initialFocusedIndex={0}
						checkButtonAriaLabel="select row"
						layoutMode={DetailsListLayoutMode.fixedColumns}
						constrainMode={ConstrainMode.unconstrained}
						selection={selection}
						onItemInvoked={onNavigate}
						onRenderRow={onRenderRow}></DetailsList>
					{contextualMenuProps && <ContextualMenu {...contextualMenuProps} />}
				</ScrollablePane>
				{(itemsLoading || isComponentLoading) && <Overlay />}
			</Stack.Item>
			<Stack.Item>
				<Stack
					horizontal
					horizontalAlign="end"
					tokens={{ childrenGap: 10 }}
					style={{ marginTop: 10 }}
				>
					{/* Page size label and dropdown FIRST */}
					<span >Rows per page</span>
					<Dropdown
						options={pageSizeOptions.map(n => ({ key: n, text: n.toString() }))}
						selectedKey={pageSize}
						onChange={(_, option) => {
							setPageSize(Number(option!.key));
							setCurrentPage(1);
						}}
						styles={{ dropdown: { width: 80 } }}
					/>
					<IconButton
						iconProps={{ iconName: 'Rewind' }}
						text="First"
						onClick={() => setCurrentPage(1)}
						disabled={currentPage === 1}
					/>
					<IconButton
						iconProps={{ iconName: 'Previous' }}
						onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
						disabled={currentPage === 1}
					/>
					<span style={{ paddingTop: 5 }}>
						Page {currentPage} of {totalPages}
					</span>
					<IconButton
						iconProps={{ iconName: 'Next' }}
						onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
						disabled={currentPage === totalPages}
					/>
					<IconButton
						iconProps={{ iconName: 'FastForward' }}
						onClick={() => setCurrentPage(totalPages)}
						disabled={currentPage === totalPages}
					/>
				</Stack>
			</Stack.Item>
		</Stack>
	);
});

Grid.displayName = "Grid";
