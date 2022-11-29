import * as React from "react";
import { useEffect, useState } from "react";
import { IColumn, TooltipHost } from "@fluentui/react";
import styles from "../Components/App.module.css";
import { filetype } from "../types/FileType";

export interface IDocument {
  key: string;
  name: string;
  type: string;
  fileType: string;
  iconSource: string;
}

export default function useDetailsList() {
  const [columns, setColumns] = useState<IColumn[]>([]);
  const [items, setItems] = useState<IDocument[]>([]);
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
    const _items: IDocument[] = [
      {
        key: "a",
        name: "Document1",
        type: "Word Document",
        fileType: filetype.docx,
        iconSource:
          "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/svg/docx_48x1.svg",
      },
      {
        key: "b",
        name: "Document2",
        type: "PDF Document",
        fileType: filetype.pdf,
        iconSource:
          "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/svg/docx_48x1.svg",
      },
      {
        key: "c",
        name: "Document3",
        type: "Excel Document",
        fileType: filetype.xlsx,
        iconSource:
          "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/svg/docx_48x1.svg",
      },
      {
        key: "e",
        name: "Document5",
        type: "PowerPoint Document",
        fileType: filetype.pptx,
        iconSource:
          "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/svg/docx_48x1.svg",
      },
    ];

    _items.forEach(item => {
      item.iconSource = fileIconLink(item.fileType).url;
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
        fieldName: "name",
        minWidth: 16,
        maxWidth: 16,
        onColumnClick: onColumnClick,
        onRender: (item: IDocument) => (
          <TooltipHost content={`${item.fileType} file`}>
            <img
              src={item.iconSource}
              className={styles.fileIconImg}
              alt={`${item.fileType} file icon`}
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
  }, []);

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
