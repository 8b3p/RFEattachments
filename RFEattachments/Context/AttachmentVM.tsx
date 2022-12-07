import * as React from "react";
import {
  IColumn,
  TooltipHost,
  Selection,
  ICommandBarItemProps,
  Link,
} from "@fluentui/react";
import { makeAutoObservable } from "mobx";
import { ServiceProvider } from "pcf-react";
import { axa_requestforexpenditureMetadata } from "../cds-generated/entities/axa_RequestforExpenditure";
import { axa_attachment_axa_attachment_axa_type } from "../cds-generated/enums/axa_attachment_axa_attachment_axa_type";
import CdsService, { cdsServiceName } from "../cdsService/CdsService";
import { IInputs } from "../generated/ManifestTypes";
import { Attachment } from "../types/Attachment";
import styles from "../Components/App.module.css";
import { copyAndSort, fileIconLink } from "../utils/utils";
import { axa_rfestatus } from "../cds-generated/enums/axa_rfestatus";
import { FileToDownload } from "../types/FileToDownload";

export interface IRow {
  key: string;
  type: string;
  fileName: string;
  iconSource: string;
}

export default class AttachmentVM {
  public Attachments: Attachment[] = [];
  public selectedAttachments: Attachment[] = [];
  public listItems: IRow[] = [];
  public listColumns: IColumn[];
  public isControlDisabled: boolean = false;
  public commandBarItems: ICommandBarItemProps[];
  public farCommandBarItems: ICommandBarItemProps[];
  public serviceProvider: ServiceProvider;
  public context: ComponentFramework.Context<IInputs>;
  public isPanelOpen: boolean = false;
  public rfeGuid: string;
  public cdsService: CdsService;
  public isLoading: boolean = false;
  public error: Error;
  public isDeleteDialogOpen: boolean = false;
  public formType: "new" | "edit";
  public selection = new Selection({
    onSelectionChanged: () => {
      this.farCommandBarItems[0].disabled =
        !(this.selection.getSelectedCount() > 0) || this.isControlDisabled;
      this.farCommandBarItems = [...this.farCommandBarItems];
      if (this.selection.getSelectedCount() > 0) {
        this.selectedAttachments = this.selection
          .getSelection()
          .map(item => {
            return this.Attachments.find(attachment => {
              return attachment.attachmentId.id === item.key;
            });
          })
          .filter(item => item !== undefined) as Attachment[];
      } else {
        this.selectedAttachments = [];
      }
    },
  });

  constructor(serviceProvider: ServiceProvider) {
    this.serviceProvider = serviceProvider;
    this.context = serviceProvider.get("context");
    this.cdsService = serviceProvider.get(cdsServiceName);
    this.rfeGuid = Xrm.Page.data.entity.getId();
    this.rfeGuid = this.rfeGuid.substring(1, this.rfeGuid.length - 1); // remove the curly braces, as the RFE ID is stored in the format {guid}
    this.fetchData();
    makeAutoObservable(this);
  }

  public fetchData = async () => {
    this.isLoading = true;
    const [result1] = await Promise.allSettled([
      this.cdsService.retrieveAttachmentByRfeId(this.rfeGuid),
      this.fetchRfeStatus,
    ]);
    if (result1.status === "fulfilled") {
      const result = result1.value;
      if (result instanceof Error) {
        console.error(result.message);
        this.isLoading = false;
        this.error = result;
        return;
      }
      this.Attachments = result;
      this.isLoading = false;
      this.populateUI();
    }
  };

  public async fetchRfeStatus() {
    const result = await this.cdsService.retrieveRfeStatus(this.rfeGuid);
    if (result instanceof Error) {
      console.error(result.message);
      this.isLoading = false;
      this.error = result;
      return;
    }
    this.isControlDisabled = result !== axa_rfestatus.Draft;
    this.commandBarItems[0].disabled = this.isControlDisabled;
    this.farCommandBarItems = [...this.farCommandBarItems];
  }

  public async getFile(attachmentId: string) {
    try {
      const fileToDownload = await this.cdsService.getFile(attachmentId);
      if (fileToDownload instanceof Error) {
        console.error(fileToDownload.message);
        this.isLoading = false;
        this.error = fileToDownload;
        return;
      }
      this.downloadFile(fileToDownload);
    } catch (error) {
      console.error(error);
    }
  }

  public downloadFile(fileToDownload: FileToDownload) {
    const fileURL = this.base64ToUrl(
      fileToDownload.fileContent,
      fileToDownload.mimeType
    );
    console.log(fileURL);
    //download the file
    window.open(fileURL, "_blank");
  }

  public base64ToUrl(base64: string, type: string) {
    const blob = this.base64ToBlob(base64, type);
    const url = URL.createObjectURL(blob);
    return url;
  }

  public base64ToBlob(base64Content: string, contentType: string) {
    const sliceSize = 512;
    const byteCharacters = atob(base64Content);
    const byteArrays = [];

    for (let offset = 0; offset < byteCharacters.length; offset += sliceSize) {
      const slice = byteCharacters.slice(offset, offset + sliceSize);

      const byteNumbers = new Array(slice.length);
      for (let i = 0; i < slice.length; i++) {
        byteNumbers[i] = slice.charCodeAt(i);
      }

      const byteArray = new Uint8Array(byteNumbers);
      byteArrays.push(byteArray);
    }

    const blob = new Blob(byteArrays, { type: contentType });
    return blob;
  }

  public populateUI = () => {
    const _items: IRow[] = [];
    this.Attachments.forEach(attachment => {
      _items.push({
        key: attachment.attachmentId?.id || "",
        type: axa_attachment_axa_attachment_axa_type[attachment.type],
        fileName: this.cdsService.TrimFileExtension(attachment.fileName),
        iconSource: fileIconLink(
          this.cdsService.GetFileExtension(attachment.fileName)
        ).url,
      });
    });
    this.listItems = _items;

    this.listColumns = [
      {
        key: "column1",
        name: "File Type",
        className: styles.fileIconCell,
        iconClassName: styles.fileIconHeaderIcon,
        ariaLabel:
          "Column operations for File type, Press to sort on File type",
        iconName: "Page",
        isIconOnly: true,
        fieldName: "type",
        minWidth: 30,
        maxWidth: 30,
        onRender: (item: IRow) => (
          <TooltipHost content={`${item.fileName} File`}>
            <img
              width={16}
              height={16}
              src={item.iconSource}
              className={styles.fileIconImg}
              alt={`${item.type} file icon`}
            />
          </TooltipHost>
        ),
      },
      {
        key: "column2",
        name: "Type",
        fieldName: "type",
        minWidth: 50,
        maxWidth: 75,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
        onColumnClick: (_ev, column) => {
          this.onColumnClick("type", column);
        },
        data: "string",
        isPadded: true,
      },
      {
        key: "column3",
        name: "File Name",
        fieldName: "fileName",
        minWidth: 210,
        maxWidth: 350,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
        onColumnClick: (_ev, column) => {
          this.onColumnClick("fileName", column);
        },
        onRender: (item: IRow) => (
          <Link href='' onClick={() => this.getFile(item.key)}>
            {item.fileName}
          </Link>
        ),
        data: "string",
        isPadded: true,
      },
    ];
    this.commandBarItems = [
      {
        key: "newItem",
        text: "New",
        cacheKey: "myCacheKey", // changing this key will invalidate this item's cache
        iconProps: { iconName: "Add" },
        disabled: this.isControlDisabled,
        onClick: () => {
          this.formType = "new";
          this.isPanelOpen = true;
        },
      },
    ];
    this.farCommandBarItems = [
      {
        key: "delete",
        text: "Delete",
        // This needs an ariaLabel since it's icon-only
        ariaLabel: "Grid view",
        disabled: true,
        iconOnly: true,
        iconProps: { iconName: "Delete" },
        onClick: () => {
          this.toggleDeleteDialog();
        },
      },
    ];
  };

  public uploadFile = async (
    type: axa_attachment_axa_attachment_axa_type,
    file?: File
  ) => {
    try {
      this.isLoading = true;
      if (this.formType === "new" && file) {
        await this.cdsService.createFiles({
          file,
          type,
          entityId: this.rfeGuid,
          entityLogicalName: axa_requestforexpenditureMetadata.logicalName,
        });
      } else {
        if (!this.isControlDisabled) {
          if (file && type) {
            await this.cdsService.updateFile({
              file,
              type,
              attachmentId: this.selectedAttachments[0].attachmentId.id || "",
            });
          } else {
            await this.cdsService.updateFile({
              type,
              attachmentId: this.selectedAttachments[0].attachmentId.id || "",
            });
          }
        }
      }
      await this.fetchData();
      this.selection.setIndexSelected(0, false, false);
      return;
    } catch (error) {
      console.log(error);
    }
    await this.fetchData();
  };

  public toggleDeleteDialog = () => {
    this.isDeleteDialogOpen = !this.isDeleteDialogOpen;
  };

  public deleteSelectedAttachments = async () => {
    await this.fetchRfeStatus();
    this.isLoading = true;
    await this.cdsService.deleteAttachments(this.selectedAttachments);
    await this.fetchData();
  };

  public onColumnClick = (_columnKey: string, column: IColumn) => {
    const newColumns: IColumn[] = this.listColumns.slice();
    const currColumn: IColumn = newColumns.find(
      currCol => column.key === currCol.key
    ) as IColumn;
    newColumns.forEach((newCol: IColumn) => {
      if (newCol.key === currColumn.key) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = copyAndSort(this.listItems, currColumn);
    this.listColumns = newColumns;
    this.listItems = newItems;
  };

  public getKey = (item: IRow) => {
    return item.key;
  };
}

export const AttachmentVMserviceName = "AttachmentVM";
