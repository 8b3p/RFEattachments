import { EntityReference } from "cdsify";
import { axa_attachment_axa_attachment_axa_type } from "../cds-generated/enums/axa_attachment_axa_attachment_axa_type";

export class Attachment {
  attachmentId: EntityReference;
  extension?: string;
  type: axa_attachment_axa_attachment_axa_type;
  rfe: EntityReference;
  isThereFile: boolean;
  fileName: string;
  constructor({
    attachmentId,
    extension,
    type,
    rfe,
    isThereFile,
    fileName,
  }: {
    attachmentId: EntityReference;
    fileType?: string;
    extension?: string;
    type: axa_attachment_axa_attachment_axa_type;
    rfe: EntityReference;
    isThereFile: boolean;
    fileName: string;
  }) {
    this.attachmentId = attachmentId;
    this.extension = extension;
    this.type = type;
    this.rfe = rfe;
    this.isThereFile = isThereFile;
    this.fileName = fileName;
  }
}
