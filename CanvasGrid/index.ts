import { initializeIcons } from "@fluentui/react/lib/Icons";
import * as React from "react";
import { createRoot, Root } from 'react-dom/client';
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { Grid } from "./Grid";
import { IColumn } from "@fluentui/react";

// Register icons - but ignore warnings if they have been already registered by Power Apps
initializeIcons(undefined, { disableWarnings: true });

export class CanvasGrid implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	notifyOutputChanged: () => void;
	container: HTMLDivElement;
	context: ComponentFramework.Context<IInputs>;
	sortedRecordsIds: string[] = [];
	resources: ComponentFramework.Resources;
	isTestHarness: boolean;
	records: Record<string, ComponentFramework.PropertyHelper.DataSetApi.EntityRecord>;
	currentPage = 1;
	filteredRecordCount?: number;
	isFullScreen = false;
	private root: Root;
	defaultPageSize = 1; // Default page size, can be adjusted based on requirements


	setSelectedRecords = (ids: string[]): void => {
		this.context.parameters.records.setSelectedRecordIds(ids);
	};

	onNavigate = (item?: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord): void => {
		if (item) {
			this.context.parameters.records.openDatasetItem(item.getNamedReference());
		}
	};

	onSort = (name: string, desc: boolean): void => {
		const sorting = this.context.parameters.records.sorting;
		console.log("Sorting by", name, desc, sorting, this.context.parameters.records);
		while (sorting?.length > 0) {
			sorting.pop();
		}
		if (!sorting || sorting?.length === 0) {
			this.context.parameters.records.sorting = [{
				name: name,
				sortDirection: desc ? 1 : 0,
			}];
		}
		else {
			this.context.parameters.records.sorting.push({
				name: name,
				sortDirection: desc ? 1 : 0,
			});
		}
		this.context.parameters.records.refresh();
	};

	onFilter = (name: string, condition: "contains" | "equals" | null, values?: string[]): void => {
		const filtering = this.context.parameters.records.filtering;
		if (condition) {
			let conditionOperator: number;
			switch (condition) {
				case "contains":
					conditionOperator = 8; // Contains
					break;
				case "equals":
					conditionOperator = 0; // Equals
					break;
				default:
					conditionOperator = 0;
			}
			filtering.setFilter({
				filterOperator: 0, // 0 = And, 1 = Or (choose as appropriate)
				conditions: [
					{
						attributeName: name,
						conditionOperator: conditionOperator,
						value: values && values.length === 1 ? values[0] : values || []
					},
				],
			} as ComponentFramework.PropertyHelper.DataSetApi.FilterExpression);
		} else {
			filtering.clearFilter();
		}
		this.context.parameters.records.refresh();
	};

	loadFirstPage = (): void => {
		this.currentPage = 1;
		this.context.parameters.records.paging.loadExactPage(1);
	};

	loadNextPage = (): void => {
		this.currentPage++;
		this.context.parameters.records.paging.loadExactPage(this.currentPage);
	};

	loadPreviousPage = (): void => {
		this.currentPage--;
		this.context.parameters.records.paging.loadExactPage(this.currentPage);
	};

	laodLastPage = (): void => {
		console.log("Loading last page", this.defaultPageSize);
		this.currentPage = this.defaultPageSize;
		// this.context.parameters.records.paging.loadExactPage(this.currentPage);
	};

	onFullScreen = (): void => {
		this.context.mode.setFullScreen(true);
		console.log("Entering full screen mode");
		this.isFullScreen = true;
		this.notifyOutputChanged();
		this.context.parameters.records.refresh();
	};

	/**
	 * Hides a column by name for the current grid render only (does not mutate global columns array).
	 * This method should be passed to the Grid and used to update a local hidden columns state.
	 */
	onHideColumn = (name: string): void => {
		// This method is intended to be passed to the Grid, which should manage its own local hidden columns state.
		// Do not mutate this.context.parameters.records.columns here.
		// Instead, trigger a re-render or update a local state in the Grid.
		// Optionally, you can notifyOutputChanged or refresh if needed.
		this.notifyOutputChanged();
	}

	onResetColumns = (): void => {
		console.log("Resetting columns");
		const columns = this.context.parameters.records.columns;
		this.context.parameters.records.columns = columns
		this.context.parameters.records.refresh();
		console.log("Columns reset", this.context.parameters.records.columns);
	}

	onPageSizeChange = (newPageSize: number): void => {
		this.defaultPageSize = newPageSize;
		this.context.parameters.records.paging.setPageSize(newPageSize);
		this.context.parameters.records.refresh();
		console.log(`Page size changed to: ${newPageSize} ${this.defaultPageSize}`, this.context.parameters.records.paging);
	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(
		context: ComponentFramework.Context<IInputs>,
		notifyOutputChanged: () => void,
		state: ComponentFramework.Dictionary,
		container: HTMLDivElement
	): void {
		this.notifyOutputChanged = notifyOutputChanged;
		this.container = container;
		this.context = context;
		this.context.mode.trackContainerResize(true);
		this.resources = this.context.resources;
		this.root = createRoot(this.container);
		this.isTestHarness = document.getElementById("control-dimensions") !== null;
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		const dataset = context.parameters.records;
		const datasetChanged = context.updatedProperties.includes("dataset");
		const resetPaging = datasetChanged && !dataset.loading && !dataset.paging.hasPreviousPage && this.currentPage !== 1;

		if (context.updatedProperties.includes("fullscreen_close")) {
			this.isFullScreen = false;
		}
		if (context.updatedProperties.includes("fullscreen_open")) {
			this.isFullScreen = true;
		}

		if (resetPaging) {
			this.currentPage = 1;
		}

		if (resetPaging || datasetChanged || this.isTestHarness || !this.records) {
			this.records = dataset.records;
			this.sortedRecordsIds = dataset.sortedRecordIds;
		}

		const allocatedWidth = parseInt(context.mode.allocatedWidth as unknown as string);
		const allocatedHeight = parseInt(context.mode.allocatedHeight as unknown as string);

		if (this.filteredRecordCount !== this.sortedRecordsIds.length) {
			this.filteredRecordCount = this.sortedRecordsIds.length;
			this.notifyOutputChanged();
		}

		// Render the updated grid
		this.root.render(
			React.createElement(Grid, {
				width: allocatedWidth,
				height: allocatedHeight,
				columns: dataset.columns,
				records: this.records,
				sortedRecordIds: this.sortedRecordsIds,
				sorting: dataset.sorting,
				filtering: dataset.filtering?.getFilter() ?? undefined,
				resources: this.resources,
				highlightValue: this.context.parameters.HighlightValue.raw,
				highlightColor: this.context.parameters.HighlightColor.raw,

				onSort: this.onSort,
				onFilter: this.onFilter,
				onNavigate: this.onNavigate,
				itemsLoading: dataset.loading,
				setSelectedRecords: this.setSelectedRecords,
				onHideColumn: this.onHideColumn,
				onPageSizeChange: this.onPageSizeChange,
				defaultPageSize: this.defaultPageSize ?? 10,
				hasNextPage: dataset.paging.hasNextPage,
				hasPreviousPage: dataset.paging.hasPreviousPage,
				currentPage: this.currentPage ?? 1,
				isFullScreen: this.isFullScreen,
				onFullScreen: this.onFullScreen,
				onResetColumns: this.onResetColumns
			})
		);
	}


	/**
	 * It is called by the framework prior to a control receiving new data.
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {
			FilteredRecordCount: this.filteredRecordCount,
		} as IOutputs;
	}

	/**
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		this.root.unmount();
	}
}