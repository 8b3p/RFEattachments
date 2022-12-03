import { EntityReference } from "cdsify";
import { axa_attachment_axa_attachment_axa_type } from "../cds-generated/enums/axa_attachment_axa_attachment_axa_type";

export class Attachment {
  attachmentId: EntityReference;
  extension?: string;
  type: axa_attachment_axa_attachment_axa_type;
  rfe: EntityReference;
  file: File;
  constructor({
    attachmentId,
    extension,
    type,
    rfe,
    file,
  }: {
    attachmentId: EntityReference;
    fileType?: string;
    extension?: string;
    type: axa_attachment_axa_attachment_axa_type;
    rfe: EntityReference;
    file: File;
  }) {
    this.attachmentId = attachmentId;
    this.extension = extension;
    this.type = type;
    this.rfe = rfe;
    this.file = file;
  }
}
