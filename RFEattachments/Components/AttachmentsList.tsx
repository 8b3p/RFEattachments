import * as React from "react";
import {
  DetailsList,
  DetailsListLayoutMode,
  MarqueeSelection,
} from "@fluentui/react";
import styles from "./App.module.css";
import { useAttachmentVM } from "../Context/context";
import { observer } from "mobx-react-lite";
import Actions from "./CommandBar";

const AttachmentsList = () => {
  const vm = useAttachmentVM();

  return (
    <div className={styles.container}>
      <Actions />

      <MarqueeSelection selection={vm.selection}>
        <DetailsList
          items={vm.listItems}
          columns={vm.listColumns}
          getKey={vm.getKey}
          setKey='multiple'
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
          selection={vm.selection}
          selectionPreservedOnEmptyClick={true}
          onItemInvoked={() => {
            vm.formType = "edit";
            vm.isPanelOpen = true;
          }}
          enterModalSelectionOnTouch={true}
          ariaLabelForSelectionColumn='Toggle selection'
          ariaLabelForSelectAllCheckbox='Toggle selection for all items'
          checkButtonAriaLabel='select row'
        />
      </MarqueeSelection>
    </div>
  );
};

export default observer(AttachmentsList);
