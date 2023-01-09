import { EntityReference } from "cdsify";
import {
  axa_Attachment,
  axa_attachmentMetadata,
} from "../cds-generated/entities/axa_Attachment";
import { axa_requestforexpenditureMetadata } from "../cds-generated/entities/axa_RequestforExpenditure";
import { axa_attachment_axa_attachment_axa_type } from "../cds-generated/enums/axa_attachment_axa_attachment_axa_type";
import { axa_rfestatus } from "../cds-generated/enums/axa_rfestatus";
import { IInputs } from "../generated/ManifestTypes";
import { Attachment } from "../types/Attachment";
import { FileToDownload } from "../types/FileToDownload";
import { EncodeFile, GetFileExtension } from "../utils/utils";

export default class CdsService {
  context: ComponentFramework.Context<IInputs>;
  constructor(context: ComponentFramework.Context<IInputs>) {
    this.context = context;
  }

  public async retrieveFile(attachmentId: string) {
    try {
      const fileToDownload: FileToDownload = new FileToDownload();
      const fileIdRes = await this.context.webAPI.retrieveRecord(
        axa_attachmentMetadata.logicalName,
        attachmentId,
        "?$select=axa_file"
      );
      const [result, result1] = await Promise.allSettled([
        fetch(
          "/api/data/v9.0/fileattachments?$filter=fileattachmentid eq " +
            fileIdRes.axa_file,
          {
            method: "GET",
          }
        ),
        fetch("/api/data/v9.1/axa_attachments(" + attachmentId + ")/axa_file", {
          method: "GET",
        }),
      ]);
      if (result.status === "fulfilled" && result1.status === "fulfilled") {
        const res = await result.value.json();
        const res1 = await result1.value.json();

        fileToDownload.mimeType = res.value[0].mimetype;
        fileToDownload.fileName = attachmentId;
        fileToDownload.fileContent = res1.value;
        fileToDownload.fileSize = res.value[0].filesizeinbytes;
      }
      return fileToDownload;
    } catch (e: any) {
      console.error(e.message);
      return new Error(e.message);
    }
  }

  public async retrieveRfeStatus(
    rfeId: string
  ): Promise<axa_rfestatus | Error> {
    try {
      const response = await this.context.webAPI.retrieveRecord(
        axa_requestforexpenditureMetadata.logicalName,
        rfeId,
        "?$select=axa_rfestatus"
      );
      if (response) {
        return response.axa_rfestatus as axa_rfestatus;
      } else {
        return new Error("No response from CDS");
      }
    } catch (e: any) {
      console.error(e.message);
      return new Error(e.message);
    }
  }

  public async retrieveAttachmentByRfeId(
    rfeId: string
  ): Promise<Attachment[] | Error> {
    const query = `
      ?fetchXml=
        <fetch>
          <entity name="axa_attachment">
            <attribute name="axa_file" />
            <attribute name="axa_type" />
            <attribute name="axa_rfe" />
            <attribute name="axa_attachmentid" />
            <attribute name="axa_name" />
            <attribute name="axa_file_name" />
            <attribute name="axa_description" />
            <filter>
              <condition attribute="axa_rfe" operator="eq" value="${rfeId}" />
            </filter>
          </entity>
        </fetch>
    `.trim();
    let response;
    try {
      response = await this.context.webAPI.retrieveMultipleRecords(
        axa_attachmentMetadata.logicalName,
        query
      );
    } catch (e: any) {
      console.error(e.message);
    }
    if (response) {
      const attachments = response.entities.map(entity => {
        const attachment = entity as axa_Attachment;
        const newAttachment = new Attachment({
          attachmentId: new EntityReference(
            axa_attachmentMetadata.logicalName,
            attachment.axa_attachmentid
          ),
          type:
            attachment.axa_type ||
            (1 as axa_attachment_axa_attachment_axa_type),
          rfe: new EntityReference(
            attachment.axa_rfe?.entityType,
            attachment.axa_rfe?.id,
            attachment.axa_rfe?.name
          ),
          isThereFile: attachment.axa_file ? true : false,
          fileName: attachment.axa_file_name || "",
          extension: GetFileExtension(attachment.axa_file_name || ""),
          description: attachment.axa_description || "",
        });
        return newAttachment;
      });

      return attachments;
    } else {
      return new Error("No response from CDS");
    }
  }

  public async createFile({
    file,
    type,
    entityId,
    entityLogicalName,
    description,
  }: {
    file: File;
    type: axa_attachment_axa_attachment_axa_type;
    entityId: string;
    entityLogicalName: string;
    description: string;
  }): Promise<void | Error> {
    const attachment = new Attachment({
      attachmentId: new EntityReference(axa_attachmentMetadata.logicalName, ""),
      type: type,
      rfe: new EntityReference(entityLogicalName, entityId),
      extension: GetFileExtension(file.name),
      isThereFile: false,
      fileName: file.name,
      description: description,
    });

    try {
      let data = {
        axa_type: type,
        axa_description: description,
        ["axa_RFE@odata.bind"]: `/${axa_requestforexpenditureMetadata.collectionName}(${entityId})`,
      };
      const response = await this.context.webAPI.createRecord(
        axa_attachmentMetadata.logicalName,
        data
      );
      if (response) {
        attachment.attachmentId.id = response?.id;
      } else {
        return new Error("No response from CDS");
      }
    } catch (e: any) {
      console.error(e.message);
      return new Error(e.message);
    }
    try {
      let response = await this.uploadFile(
        file,
        attachment.attachmentId.id,
        axa_attachmentMetadata.collectionName
      );
      if (response instanceof Error) {
        console.error(response.message);
        return response;
      }
      attachment.isThereFile = true;
    } catch (e: any) {
      console.error(e.message);
      return new Error(e.message);
    }
    return;
  }

  public async updateFile({
    file,
    type,
    description,
    attachmentId,
  }: {
    file?: File;
    type: axa_attachment_axa_attachment_axa_type;
    description: string;
    attachmentId: string;
  }): Promise<void | Error> {
    if (file) {
      const encodedData = await EncodeFile(file);
      try {
        const [response1, response2] = await Promise.allSettled([
          this.uploadFile(
            file,
            attachmentId,
            axa_attachmentMetadata.collectionName
          ),
          this.context.webAPI.updateRecord(
            axa_attachmentMetadata.logicalName,
            attachmentId,
            { axa_type: type, axa_description: description }
          ),
        ]);
        if (
          response1.status === "fulfilled" &&
          response2.status === "fulfilled"
        ) {
          if (response1.value instanceof Error) {
            console.error(response1.value.message);
            return new Error(response1.value.message);
          }
          if (!response2.value) {
            console.error("No response from CDS");
            return new Error("No response from CDS");
          }
        }
      } catch (e: any) {
        console.error(e.message);
        return new Error(e.message);
      }
    } else {
      try {
        const response = await this.context.webAPI.updateRecord(
          axa_attachmentMetadata.logicalName,
          attachmentId,
          { axa_type: type, axa_description: description }
        );
        if (!response) {
          console.error("No response from CDS");
          return new Error("No response from CDS");
        }
      } catch (e: any) {
        console.error(e.message);
        return new Error(e.message);
      }
    }
  }

  private async uploadFile(
    file: File,
    entityId: string,
    entitySetName: string
  ): Promise<"success" | Error> {
    try {
      const fileName = file.name;
      const array = await EncodeFile(file);
      const url =
        parent.Xrm.Utility.getGlobalContext().getClientUrl() +
        "/api/data/v9.1/" +
        entitySetName +
        "(" +
        entityId +
        ")/axa_file";
      let res = await this.makeRequest({
        method: "PATCH",
        fileName,
        url,
        bytes: null,
        firstRequest: true,
      });
      let res1 = await this.fileChunckUpload({
        response: res,
        fileName: fileName,
        fileBytes: array,
      });
      if (res1 instanceof Error || res instanceof Error) {
        return new Error("Error uploading file");
      } else {
        return "success";
      }
    } catch (e: any) {
      return new Error(e.message);
    }
  }

  private async makeRequest({
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
  }): Promise<any> {
    return new Promise(function (resolve, reject) {
      const request = new XMLHttpRequest();
      request.open(method, url);
      if (firstRequest)
        request.setRequestHeader("x-ms-transfer-mode", "chunked");
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
      request.onload = () => {
        if (request.status >= 200 && request.status < 300) {
          resolve(request);
        } else {
          reject(new Error(request.statusText));
        }
      };
      request.onerror = () => {
        reject(new Error(request.statusText));
      };
      if (!firstRequest) request.send(bytes);
      else request.send();
    });
  }

  private async fileChunckUpload({
    response,
    fileName,
    fileBytes,
  }: {
    response: XMLHttpRequest;
    fileName: string;
    fileBytes: Uint8Array;
  }): Promise<"success" | Error> {
    const url = response.getResponseHeader("location") || "";
    const chunkSize = parseInt(
      response.getResponseHeader("x-ms-chunk-size") || ""
    );
    let offset = 0;
    try {
      while (offset <= fileBytes.length) {
        const count =
          offset + chunkSize > fileBytes.length
            ? fileBytes.length % chunkSize
            : chunkSize;
        const content = new Uint8Array(count);
        for (let i = 0; i < count; i++) {
          content[i] = fileBytes[offset + i];
        }
        response = await this.makeRequest({
          method: "PATCH",
          fileName,
          url,
          bytes: content,
          firstRequest: false,
          offset,
          count,
          fileBytes,
        });
        if (response.status === 206) {
          // partial content, so please continue.
          offset += chunkSize;
        } else if (response.status === 204) {
          // request complete.
          return "success";
        } else {
          // error happened.
          // log error and take necessary action.
          console.error("error happened");
          return new Error("error happened" + response.status);
        }
      }
    } catch (e: any) {
      return new Error(e.message);
    }
    return "success";
  }

  public async deleteAttachments(attachments: Attachment[]) {
    let attachmentIds = attachments
      .map(attachment => attachment.attachmentId.id)
      .filter(id => id) as string[];
    try {
      let response = await Promise.allSettled(
        attachmentIds.map(attachmentId => {
          return this.context.webAPI.deleteRecord(
            axa_attachmentMetadata.logicalName,
            attachmentId
          );
        })
      );
      response.forEach((result, _index) => {
        if (result.status !== "fulfilled") {
          console.error(result.reason);
        }
      });
    } catch (e: any) {
      console.error(e.message);
      return new Error(e.message);
    }
  }
}

export const cdsServiceName = "CdsService";
