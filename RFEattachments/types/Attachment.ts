import { EntityReference } from "cdsify";
import { axa_attachment_axa_attachment_axa_type } from "../cds-generated/enums/axa_attachment_axa_attachment_axa_type";

export class Attachment {
  attachmentId?: EntityReference;
  name?: string;
  extension?: string;
  type?: axa_attachment_axa_attachment_axa_type;
  rfe?: EntityReference;
  file?: File;
  constructor({
    attachmentId,
    name,
    extension,
    type,
    rfe,
    file,
  }: {
    attachmentId?: EntityReference;
    name?: string;
    extension?: string;
    type?: axa_attachment_axa_attachment_axa_type;
    rfe?: EntityReference;
    file?: any;
  }) {
    this.attachmentId = attachmentId;
    this.name = name;
    this.extension = extension;
    this.type = type;
    this.rfe = rfe;
    this.file = file;
  }
}
