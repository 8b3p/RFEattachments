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

// const lightTheme: PartialTheme = {
//   palette: {
//     themePrimary: "#0078d4",
//     themeLighterAlt: "#eff6fc",
//     themeLighter: "#deecf9",
//     themeLight: "#c7e0f4",
//     themeTertiary: "#71afe5",
//     themeSecondary: "#2b88d8",
//     themeDarkAlt: "#106ebe",
//     themeDark: "#005a9e",
//     themeDarker: "#004578",
//     neutralLighterAlt: "#f8f8f8",
//     neutralLighter: "#f4f4f4",
//     neutralLight: "#eaeaea",
//     neutralQuaternaryAlt: "#dadada",
//     neutralQuaternary: "#d0d0d0",
//     neutralTertiaryAlt: "#c8c8c8",
//     neutralTertiary: "#a6a6a6",
//     neutralSecondary: "#858585",
//     neutralPrimaryAlt: "#4b4b4b",
//     neutralPrimary: "#333333",
//     neutralDark: "#272727",
//     black: "#1d1d1d",
//     white: "#ffffff",
//   },
// };
// const darkTheme: PartialTheme = {
//   palette: {
//     themePrimary: "#0078d4",
//     themeLighterAlt: "#eff6fc",
//     themeLighter: "#deecf9",
//     themeLight: "#c7e0f4",
//     themeTertiary: "#71afe5",
//     themeSecondary: "#2b88d8",
//     themeDarkAlt: "#106ebe",
//     themeDark: "#005a9e",
//     themeDarker: "#004578",
//     neutralLighterAlt: "#1b1b1b",
//     neutralLighter: "#252525",
//     neutralLight: "#303030",
//     neutralQuaternaryAlt: "#3b3b3b",
//     neutralQuaternary: "#595959",
//     neutralTertiaryAlt: "#a6a6a6",
//     neutralTertiary: "#c8c8c8",
//     neutralSecondary: "#d0d0d0",
//     neutralPrimaryAlt: "#dadada",
//     neutralPrimary: "#f4f4f4",
//     neutralDark: "#f8f8f8",
//     black: "#ffffff",
//     white: "#333333",
//   },
// };

const App = ({ serviceProvider }: props) => {
  const vm = serviceProvider.get<AttachmentVM>(AttachmentVMserviceName);

  return (
    <AttachmentVMProvider value={vm}>
        <AttachmentsList />
    </AttachmentVMProvider>
  );
};

export default observer(App);
