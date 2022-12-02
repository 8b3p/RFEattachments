import * as React from "react";
import { useEffect, useState } from "react";
import { IColumn, TooltipHost } from "@fluentui/react";
import styles from "../Components/App.module.css";
import AttachmentVM from "../Context/AttachmentVM";
import { axa_attachment_axa_attachment_axa_type } from "../cds-generated/enums/axa_attachment_axa_attachment_axa_type";

export interface IRow {
  key: string;
  type: string;
  iconSource: string;
}

export default function useDetailsList({ vm }: { vm: AttachmentVM }) {
  const [columns, setColumns] = useState<IColumn[]>([]);
  const [items, setItems] = useState<IRow[]>([]);
  const [selectionDetails, setSelectionDetails] = useState<string>("");
  const [isModalSelection, setIsModalSelection] = useState<boolean>(false);
  const [isCompactMode, setIsCompactMode] = useState<boolean>(false);
  const [announcedMessage, setAnnouncedMessage] = useState<string | undefined>(
    undefined
  );

  const onColumnClick = (
    _ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ) => {
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(
      currCol => column.key === currCol.key
    )[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
        setAnnouncedMessage(
          `${currColumn.name} is sorted ${
            currColumn.isSortedDescending ? "descending" : "ascending"
          }`
        );
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = copyAndSort(items, currColumn);
    setColumns(newColumns);
    setItems(newItems);
  };

  const fileIconLink = (docType: string): { url: string } => {
    return {
      url: `https://static2.sharepointonline.com/files/fabric/assets/item-types/16/${docType}.svg`,
    };
  };

  const copyAndSort = <T = any,>(items: T[], currCol: IColumn): T[] => {
    const key = currCol.key as keyof T;
    return items
      .slice(0)
      .sort((a: T, b: T) =>
        (currCol.isSortedDescending ? a[key] < b[key] : a[key] > b[key])
          ? 1
          : -1
      );
  };

  useEffect(() => {
    const _items: IRow[] = [];

    vm.Attachments.forEach(attachment => {
      _items.push({
        key: attachment.attachmentId?.id || "",
        type: axa_attachment_axa_attachment_axa_type[attachment.type],
        iconSource: fileIconLink(
          vm.cdsService.GetFileExtension(attachment.extension || "")
        ).url,
      });
    });

    setItems(_items);
    setColumns([
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
        minWidth: 16,
        maxWidth: 16,
        onColumnClick: onColumnClick,
        onRender: (item: IRow) => (
          <TooltipHost content={`${item.type} file`}>
            <img
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
        onColumnClick: onColumnClick,
        data: "string",
        isPadded: true,
      },
    ]);
  }, [vm.Attachments]);

  return {
    columns,
    setColumns,
    items,
    setItems,
    selectionDetails,
    setSelectionDetails,
    isModalSelection,
    setIsModalSelection,
    isCompactMode,
    setIsCompactMode,
    announcedMessage,
    setAnnouncedMessage,
    onColumnClick,
  };
}
