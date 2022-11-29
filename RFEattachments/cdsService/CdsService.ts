import { EntityReference } from "cdsify";
import {
  axa_Attachment,
  axa_attachmentMetadata,
} from "../cds-generated/entities/axa_Attachment";
import { IInputs } from "../generated/ManifestTypes";
import { Attachment } from "../types/Attachment";

export default class CdsService {
  context: ComponentFramework.Context<IInputs>;
  constructor(context: ComponentFramework.Context<IInputs>) {
    this.context = context;
  }

  public async retrieveRecordByRfeId(rfeId: string): Promise<Attachment[]> {
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
            </filter>
          </entity>
        </fetch>
    `;

    const response = await this.context.webAPI.retrieveMultipleRecords(
      axa_attachmentMetadata.logicalName,
      query
    );

    const attachments: Attachment[] = response.entities.map(entity => {
      const attachment = entity as axa_Attachment;
      const newAttachment = new Attachment({});
      if (attachment.axa_attachmentid) {
        newAttachment.attachmentId = new EntityReference(
          axa_attachmentMetadata.logicalName,
          attachment.axa_attachmentid
        );
      }
      if (attachment.axa_name) {
        newAttachment.name = attachment.axa_name;
      }
      if (attachment.axa_type) {
        newAttachment.type = attachment.axa_type;
      }
      if (attachment.axa_rfe) {
        newAttachment.rfe = new EntityReference(
          attachment.axa_rfe.entityType,
          attachment.axa_rfe.id,
          attachment.axa_rfe.name
        );
      }
      if (attachment.axa_file) {
        newAttachment.file = attachment.axa_file;
      }
      return newAttachment;
    });
    return attachments;
    // return response.entities as Attachment[];
  }
}

export const cdsServiceName = "CdsService";
