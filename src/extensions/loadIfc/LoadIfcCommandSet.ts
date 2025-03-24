import { Log } from "@microsoft/sp-core-library";
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs,
} from "@microsoft/sp-listview-extensibility";
import { Dialog } from "@microsoft/sp-dialog";

export interface ILoadIfcCommandSetProperties {
  viewerPageUrl: string;
}

const LOG_SOURCE: string = "LoadIfcCommandSet";

export default class LoadIfcCommandSet extends BaseListViewCommandSet<ILoadIfcCommandSetProperties> {
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized LoadIfcCommandSet");

    // Set initial state of the command's visibility
    const loadIfcCommand: Command = this.tryGetCommand("LOAD_IFC");
    if (loadIfcCommand) {
      loadIfcCommand.visible = false;
    }

    // Add handler for list view state changes
    this.context.listView.listViewStateChangedEvent.add(
      this,
      this._onListViewStateChanged
    );

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case "LOAD_IFC":
        this._handleLoadIfcCommand();
        break;
      default:
        throw new Error("Unknown command");
    }
  }

  private _onListViewStateChanged = (
    args: ListViewStateChangedEventArgs
  ): void => {
    Log.info(LOG_SOURCE, "List view state changed");

    const loadIfcCommand: Command = this.tryGetCommand("LOAD_IFC");
    if (loadIfcCommand) {
      // Show button only when IFC files are selected
      const hasIfcFilesSelected = this.context.listView.selectedRows?.some(
        (row) =>
          row.getValueByName("FileLeafRef").toLowerCase().endsWith(".ifc")
      );
      loadIfcCommand.visible = hasIfcFilesSelected ?? false;
    }

    this.raiseOnChange();
  };

  private _handleLoadIfcCommand(): void {
    if (this.context.listView.selectedRows) {
      const selectedFiles = this.context.listView.selectedRows
        .filter((row) =>
          row.getValueByName("FileLeafRef").toLowerCase().endsWith(".ifc")
        )
        .map((row) => ({
          name: row.getValueByName("FileLeafRef"),
          url: `${this.context.pageContext.web.absoluteUrl}${row.getValueByName(
            "FileRef"
          )}`,
        }));

      if (selectedFiles.length > 0) {
        // Open the viewer page with the first selected IFC file
        const viewerUrl = new URL(this.properties.viewerPageUrl);
        viewerUrl.searchParams.append(
          "fileUrl",
          encodeURIComponent(selectedFiles[0].url)
        );

        // Open in same tab or new tab based on your preference
        window.open(viewerUrl.toString(), "_blank");
      } else {
        Dialog.alert("No IFC files selected").catch(() => {
          /* handle error */
        });
      }
    }
  }
}
