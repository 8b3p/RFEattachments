import { ServiceProvider, StandardControlReact } from "pcf-react";
import React = require("react");
import ReactDOM = require("react-dom");
import CdsService, { cdsServiceName } from "./cdsService/CdsService";
import App from "./Components/App";
import Temp from "./Components/temp";
import AttachmentVM, { AttachmentVMserviceName } from "./Context/AttachmentVM";
import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class RFEattachments extends StandardControlReact<IInputs, IOutputs> {
  constructor() {
    super();
    this.renderOnParametersChanged = false;
    this.initServiceProvider = (serviceProvider: ServiceProvider) => {
      serviceProvider.register("context", this.context);
      serviceProvider.register(cdsServiceName, new CdsService(this.context));
      serviceProvider.register(
        AttachmentVMserviceName,
        new AttachmentVM(serviceProvider)
      );
    };
    this.reactCreateElement = (
      container: HTMLDivElement,
      _width: number | undefined,
      _height: number | undefined,
      serviceProvider: ServiceProvider
    ) => {
      ReactDOM.render(
        React.createElement(App, {
          // context: this.context,
          // width: width,
          // height: height,
          serviceProvider: serviceProvider,
        }),
        container
      );
    };
  }
}
