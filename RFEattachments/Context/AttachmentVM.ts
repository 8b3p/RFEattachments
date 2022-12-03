import { IColumn } from "@fluentui/react";
import { makeAutoObservable } from "mobx";
import { ServiceProvider } from "pcf-react";
import { axa_requestforexpenditureMetadata } from "../cds-generated/entities/axa_RequestforExpenditure";
import { axa_attachment_axa_attachment_axa_type } from "../cds-generated/enums/axa_attachment_axa_attachment_axa_type";
import CdsService, { cdsServiceName } from "../cdsService/CdsService";
import { IInputs } from "../generated/ManifestTypes";
import { IRow } from "../Hooks/useDetailsList";
import { Attachment } from "../types/Attachment";

export default class AttachmentVM {
  public Attachments: Attachment[] = [];
  public selectedAttachments: Attachment[] = [];
  public isDeleteEnabled: boolean = false;
  public listItems: IRow[] = [];
  public listColumns: IColumn[];
  public serviceProvider: ServiceProvider;
  public context: ComponentFramework.Context<IInputs>;
  public rfeGuid: string;
  public cdsService: CdsService;
  public isLoading: boolean = false;
  public error: Error;

  constructor(serviceProvider: ServiceProvider) {
    this.serviceProvider = serviceProvider;
    this.context = serviceProvider.get("context");
    this.cdsService = serviceProvider.get(cdsServiceName);
    this.rfeGuid = Xrm.Page.data.entity.getId();
    this.rfeGuid = this.rfeGuid.substring(1, this.rfeGuid.length - 1); // remove the curly braces, as the RFE ID is stored in the format {guid}
    this.fetchData();
    makeAutoObservable(this);
  }
  fetchData = async () => {
    this.isLoading = true;
    const result = await this.cdsService.retrieveRecordByRfeId(this.rfeGuid);
    if (result instanceof Error) {
      console.error(result.message);
      this.isLoading = false;
      this.error = result;
      return;
    }
    this.Attachments = result;
    this.isLoading = false;
  };

  uploadFile = async (
    file: File,
    type: axa_attachment_axa_attachment_axa_type
  ) => {
    this.isLoading = true;
    await this.cdsService.handleFiles({
      file,
      type,
      entityId: this.rfeGuid,
      entityLogicalName: axa_requestforexpenditureMetadata.logicalName,
    });
    this.fetchData();
  };

  public deleteSelectedAttachments = async () => {
    this.isLoading = true;
    await this.cdsService.deleteAttachments(this.selectedAttachments);
    this.fetchData();
  };
}

export const AttachmentVMserviceName = "AttachmentVM";
