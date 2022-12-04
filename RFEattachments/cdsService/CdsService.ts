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

export default class CdsService {
  context: ComponentFramework.Context<IInputs>;
  constructor(context: ComponentFramework.Context<IInputs>) {
    this.context = context;
  }

  public async retrieveFileToDownload(attachmentId: string) {
    const result = await this.context.webAPI.retrieveRecord(
      "annotation",
      attachmentId
    );
    let file: FileToDownload = new FileToDownload();
    file.fileContent =
      result.entityType == "annotation" ? result.documentbody : result.body;
    file.fileName = result.filename;
    file.fileSize = result.filesize;
    file.mimeType = result.mimetype;
    return file;
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
            <filter>
              <condition attribute="axa_rfe" operator="eq" value="${rfeId}" />
              <condition attribute="axa_file" operator="not-null" />
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
          file: attachment.axa_file,
          extension: this.GetFileExtension(attachment.axa_file_name || ""),
        });
        return newAttachment;
      });
      console.dir(attachments);

      return attachments;
    } else {
      return new Error("No response from CDS");
    }
  }

  public async handleFiles({
    file,
    type,
    entityId,
    entityLogicalName,
  }: {
    file: File;
    type: axa_attachment_axa_attachment_axa_type;
    entityId: string;
    entityLogicalName: string;
  }): Promise<Attachment | Error> {
    const attachment = new Attachment({
      attachmentId: new EntityReference(axa_attachmentMetadata.logicalName, ""),
      type: type,
      rfe: new EntityReference(entityLogicalName, entityId),
      extension: this.GetFileExtension(file.name),
      file: file,
    });

    const encodedData = await this.EncodeFile(file);
    try {
      let data = {
        axa_type: type,
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

    //upload the file to the File field of the attachment
    if (attachment.attachmentId.id) {
      try {
        const response = await fetch(
          `/api/data/v9.1/axa_attachments(${attachment.attachmentId.id})/axa_file?x-ms-file-name=${file.name}`,
          {
            method: "PATCH",
            headers: {
              Accept: "/",
              "Content-Type": "application/octet-stream",
              Prefer: 'odata.include-annotations="*"',
              "OData-Version": "4.0",
            },
            body: encodedData,
          }
        );
        if (!response.ok) {
          console.error(response.statusText);
          return new Error(response.statusText);
        }
      } catch (e: any) {
        console.error(e.message);
        return new Error(e.message);
      }
    }

    console.log(
      `Attachment: ${file.name} has been uploaded to ${entityLogicalName} with id: ${entityId}`
    );

    return attachment;
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
      response.forEach((result, index) => {
        if (result.status === "fulfilled") {
          console.log(
            `Attachment: ${attachments[index].attachmentId.id} has been deleted`
          );
        } else {
          console.error(result.reason);
        }
      });
    } catch (e: any) {
      console.error(e.message);
      return new Error(e.message);
    }
  }

  public toBase64(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = () => resolve(reader.result as string);
      reader.onerror = error => reject(error);
    });
  }

  public EncodeFile(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = _f => resolve((<string>reader.result).split(",")[1]);
      reader.onerror = error => reject(error);
      reader.readAsDataURL(file);
    });
  }

  public CollectionNameFromLogicalName(entityLogicalName: string): string {
    if (entityLogicalName[entityLogicalName.length - 1] != "s") {
      return `${entityLogicalName}s`;
    } else {
      return `${entityLogicalName}es`;
    }
  }

  public GetFileExtension(fileName: string): string {
    return <string>fileName.split(".").pop();
  }

  public TrimFileExtension(fileName: string): string {
    return <string>fileName.split(".")[0];
  }
}

export const cdsServiceName = "CdsService";
