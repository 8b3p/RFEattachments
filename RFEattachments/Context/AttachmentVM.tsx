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
import {
  base64ToUrl,
  copyAndSort,
  fileIconLink,
  GetFileExtension,
  TrimFileExtension,
} from "../utils/utils";
import { axa_rfestatus } from "../cds-generated/enums/axa_rfestatus";
import { seperateCamelCaseString } from "../Components/AttachmentPanel";

export interface IRow {
  key: string;
  type: string;
  fileName: string;
  iconSource: string;
  description: string;
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
    if (this.farCommandBarItems && this.farCommandBarItems.length > 0) {
      this.farCommandBarItems[0].disabled = this.isControlDisabled;
      this.farCommandBarItems = [...this.farCommandBarItems];
    }
  }

  public async getFile(attachmentId: string) {
    try {
      const fileToDownload = await this.cdsService.retrieveFile(attachmentId);
      if (fileToDownload instanceof Error) {
        console.error(fileToDownload.message);
        this.isLoading = false;
        this.error = fileToDownload;
        return;
      }
      const fileURL = base64ToUrl(
        fileToDownload.fileContent,
        fileToDownload.mimeType
      );
      //download the file
      window.open(fileURL, "_blank");
    } catch (error) {
      console.error(error);
    }
  }

  public populateUI = () => {
    const _items: IRow[] = [];
    this.Attachments.forEach(attachment => {
      _items.push({
        key: attachment.attachmentId?.id || "",
        type: seperateCamelCaseString(
          axa_attachment_axa_attachment_axa_type[attachment.type]
        ),
        fileName: attachment.isThereFile
          ? TrimFileExtension(attachment.fileName)
          : "",
        iconSource: fileIconLink(GetFileExtension(attachment.fileName)).url,
        description: attachment.description,
      });
    });
    this.listItems = _items;

    this.listColumns = [
      {
        key: "column1",
        name: "File Type",
        className: styles.fileIconCell,
        iconClassName: styles.fileIconHeaderIcon,
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
        minWidth: 100,
        maxWidth: 125,
        isRowHeader: true,
        isResizable: true,
        isSorted: false,
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
        onRender: (item: IRow) => {
          if (item.fileName === "") {
            return <></>;
          }
          return (
            <Link href='' onClick={() => this.getFile(item.key)}>
              {item.fileName}
            </Link>
          );
        },

        data: "string",
        isPadded: true,
      },
      {
        key: "column4",
        name: "Description",
        fieldName: "description",
        minWidth: 300,
        maxWidth: 400,
        isResizable: true,
        isMultiline: true,
        data: "string",
      },
    ];
    this.commandBarItems = [
      {
        key: "newItem",
        text: "New",
        cacheKey: "myCacheKey", // changing this key will invalidate this item's cache
        iconProps: { iconName: "Add" },
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
    description: string,
    file?: File
  ) => {
    try {
      this.isLoading = true;
      if (this.formType === "new" && file) {
        const res = await this.cdsService.createFile({
          file,
          type,
          description,
          entityId: this.rfeGuid,
          entityLogicalName: axa_requestforexpenditureMetadata.logicalName,
        });
        if (res) {
          this.fetchData();
          return res;
        }
      } else {
        if (this.isControlDisabled) {
          this.fetchData();
          return new Error("Control is disabled");
        }
        if (file && type && description) {
          const res = await this.cdsService.updateFile({
            file,
            type,
            description,
            attachmentId: this.selectedAttachments[0].attachmentId.id || "",
          });
          if (res) {
            this.fetchData();
            return res;
          }
        } else {
          const res = await this.cdsService.updateFile({
            type,
            description,
            attachmentId: this.selectedAttachments[0].attachmentId.id || "",
          });
          if (res) {
            this.fetchData();
            return res;
          }
        }
      }
      await this.fetchData();
      this.selection.setIndexSelected(0, false, false);
      return;
    } catch (error) {
      console.error(error);
    }
    await this.fetchData();
  };

  public toggleDeleteDialog = () => {
    this.isDeleteDialogOpen = !this.isDeleteDialogOpen;
  };

  public deleteSelectedAttachments = async () => {
    await this.fetchRfeStatus();
    if (this.isControlDisabled) {
      this.fetchData();
      return new Error("You can't delete attachments on a submitted RFE");
    }
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
