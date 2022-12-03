import * as React from "react";
import {
  IColumn,
  TooltipHost,
  Selection,
  ICommandBarItemProps,
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
  public commandBarItems: ICommandBarItemProps[];
  public farCommandBarItems: ICommandBarItemProps[];
  public serviceProvider: ServiceProvider;
  public context: ComponentFramework.Context<IInputs>;
  public rfeGuid: string;
  public cdsService: CdsService;
  public isLoading: boolean = false;
  public error: Error;
  public isDeleteDialogOpen: boolean = false;
  public selection = new Selection({
    onSelectionChanged: () => {
      this.farCommandBarItems[0].disabled = !(
        this.selection.getSelectedCount() > 0
      );
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
    const result = await this.cdsService.retrieveRecordByRfeId(this.rfeGuid);
    if (result instanceof Error) {
      console.error(result.message);
      this.isLoading = false;
      this.error = result;
      return;
    }
    this.Attachments = result;
    this.isLoading = false;
    this.populateUI();
  };

  public populateUI = () => {
    const _items: IRow[] = [];
    this.Attachments.forEach(attachment => {
      _items.push({
        key: attachment.attachmentId?.id || "",
        type: axa_attachment_axa_attachment_axa_type[attachment.type],
        fileName: this.cdsService.TrimFileExtension(
          attachment.file?.name || ""
        ),
        iconSource: fileIconLink(
          this.cdsService.GetFileExtension(attachment.extension || "")
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
        minWidth: 210,
        maxWidth: 350,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
        onColumnClick: this.onColumnClick,
        data: "string",
        isPadded: true,
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
    file: File,
    type: axa_attachment_axa_attachment_axa_type
  ) => {
    this.isLoading = true;
    await this.cdsService.handleFiles({
      file,
      type,
      entityId: this.rfeGuid,
      entityLogicalName: axa_requestforexpenditureMetadata.logicalName,
    });
    this.fetchData();
  };

  public toggleDeleteDialog = () => {
    this.isDeleteDialogOpen = !this.isDeleteDialogOpen;
  };

  public deleteSelectedAttachments = async () => {
    this.isLoading = true;
    await this.cdsService.deleteAttachments(this.selectedAttachments);
    await this.fetchData();
  };

  public onColumnClick = (
    _ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ) => {
    const newColumns: IColumn[] = this.listColumns.slice();
    const currColumn: IColumn = newColumns.filter(
      currCol => column.key === currCol.key
    )[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
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
