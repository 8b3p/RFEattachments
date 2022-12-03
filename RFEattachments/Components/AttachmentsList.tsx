import * as React from "react";
import {
  DetailsList,
  DetailsListLayoutMode,
  MarqueeSelection,
  Selection,
} from "@fluentui/react";
import styles from "./App.module.css";
import { useAttachmentVM } from "../Context/context";
import useDetailsList from "../Hooks/useDetailsList";
import { observer } from "mobx-react-lite";
import Actions from "./CommandBar";
import { Attachment } from "../types/Attachment";

const AttachmentsList = () => {
  const vm = useAttachmentVM();
  const { typeOprtions: typeOptions, isCompactMode } = useDetailsList({
    vm,
  });

  const selection = new Selection({
    onSelectionChanged: () => {
      vm.isDeleteEnabled = selection.getSelectedCount() > 0;
      console.log(selection.getSelectedCount() > 0);
      if (selection.getSelectedCount() > 0) {
        console.log(true);
        vm.selectedAttachments = selection
          .getSelection()
          .map(item => {
            return vm.Attachments.find(attachment => {
              return attachment.attachmentId.id === item.key;
            });
          })
          .filter(item => item !== undefined) as Attachment[];
      } else {
        console.log(false);
        vm.selectedAttachments = [];
      }
    },
  });

  const getKey = (item: any) => {
    return item.key;
  };

  return (
    <div className={styles.container}>
      <Actions typeOptions={typeOptions} />

      <MarqueeSelection selection={selection}>
        <DetailsList
          items={vm.listItems}
          compact={isCompactMode}
          columns={vm.listColumns}
          getKey={getKey}
          setKey='multiple'
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
          selection={selection}
          selectionPreservedOnEmptyClick={true}
          onItemInvoked={item => alert(`Item invoked: ${item.name}`)}
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
