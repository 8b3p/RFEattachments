import { IInputs } from "../generated/ManifestTypes";

export default class CdsService {
  context: ComponentFramework.Context<IInputs>;
  constructor(context: ComponentFramework.Context<IInputs>) {
    this.context = context;
  }
}

export const cdsServiceName = "CdsService";
