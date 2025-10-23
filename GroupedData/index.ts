import { IInputs, IOutputs } from "./generated/ManifestTypes";
import {
  GroupedListComponent,
  IGroupedListProps,
} from "./GroupedListComponent";
import * as React from "react";
import { PulseLoader } from "react-spinners";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
type DataSet = ComponentFramework.PropertyTypes.DataSet;

export class GroupedData
  implements ComponentFramework.ReactControl<IInputs, IOutputs>
{
  private notifyOutputChanged: () => void;
  private _currentPage: number;
  /**
   * Empty constructor.
   */
  constructor() {
    // Empty
  }

  /**
   * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
   * Data-set values are not initialized here, use updateView.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
   * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
   * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
   */
  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary
  ): void {
    this.notifyOutputChanged = notifyOutputChanged;
    context.parameters.dataset.paging.setPageSize(100); //# If you want to load specific page size, uncomment this 2 lines
    this._currentPage = context.parameters.dataset.paging.firstPageNumber; //# Without this data will be loaded per user settings 50-250 records
    context.parameters.dataset.paging.loadExactPage(this._currentPage); //# This i only way to override that oob setting (without retrieveMultipleRecords)
  }

  /**
   * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
   * @returns ReactElement root react element for the control
   */
  public updateView(
    context: ComponentFramework.Context<IInputs>
  ): React.ReactElement {
    const props: IGroupedListProps = {
      data: context.parameters.dataset,
      currentPage: this._currentPage,
      loadPrevPage: () => this.laodPrevPage(context.parameters.dataset),
      loadNextPage: () => this.loadNextPage(context.parameters.dataset),
      openRecordForm: (id: string) => this.openRecordForm(context, id),
      levelCount: context.parameters.levelCount.raw ?? undefined,
    };
    return React.createElement(GroupedListComponent, props);
  }

  /**
   * It is called by the framework prior to a control receiving new data.
   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
   */
  public getOutputs(): IOutputs {
    return {};
  }

  /**
   * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
   * i.e. cancelling any pending remote calls, removing listeners, etc.
   */
  public destroy(): void {
    // Add code to cleanup control if necessary
  }

  private _createLoader(): React.ReactElement {
    return React.createElement(PulseLoader, {
      color: "#3363ff",
      loading: true,
      cssOverride: {
        display: "flex",
        justifyContent: "center",
        alignItems: "center",
        width: "100%",
        height: "100%",
      },
      size: 15,
      margin: 2,
      speedMultiplier: 0.75,
    });
  }

  private laodPrevPage(dataset: DataSet): void {
    if (dataset.paging.hasPreviousPage) {
      this._currentPage -= 1;
      dataset.paging.loadExactPage(this._currentPage);
    }
  }

  private loadNextPage(dataset: DataSet): void {
    if (dataset.paging.hasNextPage) {
      this._currentPage += 1;
      dataset.paging.loadExactPage(this._currentPage);
    }
  }

  private openRecordForm(context: ComponentFramework.Context<IInputs>, id: string) {
    context.navigation.openForm({
      entityName: context.parameters.dataset.getTargetEntityType(),
      entityId: id
    }).catch((error) => {
      console.error("Could not open record form.")
    })
  }

}
