import { IColumn } from "@fluentui/react";

export const copyAndSort = <T = any>(items: T[], currCol: IColumn): T[] => {
  const key = currCol.key as keyof T;
  return items
    .slice(0)
    .sort((a: T, b: T) =>
      (currCol.isSortedDescending ? a[key] > b[key] : a[key] < b[key]) ? 1 : -1
    );
};

export const fileIconLink = (docType: string): { url: string } => {
  return {
    url: `https://static2.sharepointonline.com/files/fabric/assets/item-types/16/${docType}.svg`,
  };
};
