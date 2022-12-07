import { IColumn } from "@fluentui/react";
import { axa_attachmentMetadata } from "../cds-generated/entities/axa_Attachment";

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

export function makeRequest({
  method,
  fileName,
  url,
  bytes,
  firstRequest,
  offset,
  count,
  fileBytes,
}: {
  method: string;
  fileName: string;
  url: string;
  bytes: Uint8Array | null;
  firstRequest: boolean;
  offset?: number;
  count?: number;
  fileBytes?: Uint8Array;
}) {
  return new Promise(function (resolve, reject) {
    const request = new XMLHttpRequest();
    request.open(method, url);
    if (firstRequest) request.setRequestHeader("x-ms-transfer-mode", "chunked");
    request.setRequestHeader("x-ms-file-name", fileName);
    if (!firstRequest) {
      request.setRequestHeader(
        "Content-Range",
        "bytes " +
          offset +
          "-" +
          ((offset ?? 0) + (count ?? 0) - 1) +
          "/" +
          fileBytes?.length
      );
      request.setRequestHeader("Content-Type", "application/octet-stream");
    }
    request.onload = resolve;
    request.onerror = () => {
      console.dir(request);
      reject(new Error("Error uploading file"));
    };
    if (!firstRequest) request.send(bytes);
    else request.send();
  });
}

export function uploadFile(
  file: File,
  entityId: string,
  entitySetName: string
): Promise<"success" | Error> {
  try {
    const reader = new FileReader();
    const fileName = file.name;
    reader.onload = function () {
      const arrayBuffer = this.result as ArrayBuffer;
      const array = new Uint8Array(arrayBuffer);
      const url =
        parent.Xrm.Utility.getGlobalContext().getClientUrl() +
        "/api/data/v9.1/" +
        entitySetName +
        "(" +
        entityId +
        ")/axa_file";
      // this is the first request. We are passing content as null.
      makeRequest({
        method: "PATCH",
        fileName,
        url,
        bytes: null,
        firstRequest: true,
      })
        .then(function (s) {
          fileChunckUpload({
            response: s,
            fileName: fileName,
            fileBytes: array,
          })
            .then(function () {
              return "success";
            })
            .catch(function (e) {
              return e;
            });
        })
        .catch(function (e) {
          return e;
        });
    };
    reader.readAsArrayBuffer(file);
    return Promise.resolve("success");
  } catch (e) {
    return Promise.reject(e);
  }
}

export async function fileChunckUpload({
  response,
  fileName,
  fileBytes,
}: {
  response: any;
  fileName: string;
  fileBytes: Uint8Array;
}) {
  const req = response.target;
  const url = req.getResponseHeader("location");
  const chunkSize = parseInt(req.getResponseHeader("x-ms-chunk-size"));
  let offset = 0;
  while (offset <= fileBytes.length) {
    const count =
      offset + chunkSize > fileBytes.length
        ? fileBytes.length % chunkSize
        : chunkSize;
    const content = new Uint8Array(count);
    for (let i = 0; i < count; i++) {
      content[i] = fileBytes[offset + i];
    }
    response = await makeRequest({
      method: "PATCH",
      fileName,
      url,
      bytes: content,
      firstRequest: false,
      offset,
      count,
      fileBytes,
    });
    const req = response.target;
    if (req.status === 206) {
      // partial content, so please continue.
      offset += chunkSize;
    } else if (req.status === 204) {
      // request complete.
      break;
    } else {
      // error happened.
      // log error and take necessary action.
      console.log("error happened");
      throw new Error("error happened");
      break;
    }
  }
}
