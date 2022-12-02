import { EntityReference } from "cdsify";
import {
  axa_Attachment,
  axa_attachmentMetadata,
} from "../cds-generated/entities/axa_Attachment";
import { axa_requestforexpenditureMetadata } from "../cds-generated/entities/axa_RequestforExpenditure";
import { axa_attachment_axa_attachment_axa_type } from "../cds-generated/enums/axa_attachment_axa_attachment_axa_type";
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

  public async retrieveRecordByRfeId(
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
      console.dir(response.entities);
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
    });
    let attachmentId;
    try {
      let data = {
        axa_type: type,
        axa_name: "",
        ["axa_RFE@odata.bind"]: `/${axa_requestforexpenditureMetadata.collectionName}(${entityId})`,
      };
      const response = await this.context.webAPI.createRecord(
        axa_attachmentMetadata.logicalName,
        data
      );
      attachmentId = response?.id;
    } catch (e: any) {
      console.error(e.message);
      return new Error(e.message);
    }

    const encodedData = await this.toBase64(file);

    try {
      let result = await this.context.webAPI.createRecord("fileattachment", {
        [`objectid_${axa_attachmentMetadata.logicalName}@odata.bind`]: `/${axa_attachmentMetadata.collectionName}(${attachmentId})`,
        filename: file.name,
        objecttypecode: axa_attachmentMetadata.logicalName,
        regardingfieldname: "axa_file",
        // body: encodedData,
      });
      console.dir(result);
    } catch (e: any) {
      console.error(e.message);
    }

    const response = await this.context.webAPI.updateRecord(
      axa_attachmentMetadata.logicalName,
      attachmentId,
      {
        axa_file: encodedData,
      }
    );

    console.dir(response);
    console.log(file.type);
    console.log(
      `Attachment: ${file.name} has been uploaded to the ${entityLogicalName} with id: ${entityId}`
    );

    attachment.attachmentId = new EntityReference(
      axa_attachmentMetadata.logicalName,
      attachmentId
    );
    attachment.extension = this.GetFileExtension(file.name);
    attachment.type = type;
    attachment.rfe = new EntityReference(
      entityLogicalName,
      entityId,
      entityLogicalName
    );

    return attachment;
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
