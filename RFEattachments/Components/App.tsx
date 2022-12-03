import * as React from "react";
import AttachmentVMProvider from "../Context/context";
import { observer } from "mobx-react-lite";
import { ServiceProvider } from "pcf-react";
import AttachmentVM, { AttachmentVMserviceName } from "../Context/AttachmentVM";
import AttachmentsList from "./AttachmentsList";
import { IInputs } from "../generated/ManifestTypes";

interface props {
  context: ComponentFramework.Context<IInputs>;
  width?: number;
  height?: number;
  serviceProvider: ServiceProvider;
}

const App = ({ serviceProvider }: props) => {
  const vm = serviceProvider.get<AttachmentVM>(AttachmentVMserviceName);

  return (
    <AttachmentVMProvider value={vm}>
      <AttachmentsList />
    </AttachmentVMProvider>
  );
};

export default observer(App);
