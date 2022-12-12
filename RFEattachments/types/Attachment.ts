import { EntityReference } from "cdsify";
import { axa_attachment_axa_attachment_axa_type } from "../cds-generated/enums/axa_attachment_axa_attachment_axa_type";

export class Attachment {
  attachmentId: EntityReference;
  extension?: string;
  type: axa_attachment_axa_attachment_axa_type;
  rfe: EntityReference;
  isThereFile: boolean;
  fileName: string;
  description: string;
  constructor({
    attachmentId,
    extension,
    type,
    rfe,
    isThereFile,
    fileName,
    description,
  }: {
    attachmentId: EntityReference;
    fileType?: string;
    extension?: string;
    type: axa_attachment_axa_attachment_axa_type;
    rfe: EntityReference;
    isThereFile: boolean;
    fileName: string;
    description: string;
  }) {
    this.attachmentId = attachmentId;
    this.extension = extension;
    this.type = type;
    this.rfe = rfe;
    this.isThereFile = isThereFile;
    this.fileName = fileName;
    this.description = description;
  }
}
