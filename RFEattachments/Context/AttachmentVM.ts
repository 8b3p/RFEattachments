import { makeAutoObservable } from "mobx";
import { ServiceProvider } from "pcf-react";
import { Attachment } from "../types/Attachment";

export class AttachmentVM {
  public AttachmentList: Attachment[] = [];
  public serviceProvider: ServiceProvider;

  constructor(serviceProvider: ServiceProvider) {
    this.serviceProvider = serviceProvider;
    makeAutoObservable(this);
  }
}

export const AttachmentVMserviceName = "AttachmentVM";
