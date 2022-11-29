import { makeAutoObservable } from "mobx";
import { ServiceProvider } from "pcf-react";
import { cdsServiceName } from "../cdsService/CdsService";
import { IInputs } from "../generated/ManifestTypes";
import { Attachment } from "../types/Attachment";

export default class AttachmentVM {
  public Attachments: Attachment[] = [];
  public serviceProvider: ServiceProvider;
  public context: ComponentFramework.Context<IInputs>;
  public rfeGuid: string;
  cdsService: any;

  constructor(serviceProvider: ServiceProvider) {
    this.serviceProvider = serviceProvider;
    this.context = serviceProvider.get("context");
    this.cdsService = serviceProvider.get(cdsServiceName);
    this.rfeGuid = this.rfeGuid = Xrm.Page.data.entity.getId();
    this.Attachments = this.cdsService.retrieveRecordByRfeId(this.rfeGuid);
    console.dir(this.Attachments);
    makeAutoObservable(this);
  }
}

export const AttachmentVMserviceName = "AttachmentVM";
