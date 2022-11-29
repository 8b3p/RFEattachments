import * as React from "react";
import { Label } from "@fluentui/react";
import { observer } from "mobx-react-lite";

const App = () => {
  return <Label>Hello World</Label>;
};

export default observer(App);
