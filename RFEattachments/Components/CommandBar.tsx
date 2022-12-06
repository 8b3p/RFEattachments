import {
  CommandBar,
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  FocusTrapZone,
  Layer,
  Overlay,
  PrimaryButton,
  Spinner,
  SpinnerSize,
} from "@fluentui/react";
import { observer } from "mobx-react-lite";
import { useBoolean } from "@fluentui/react-hooks";
import * as React from "react";
import { useAttachmentVM } from "../Context/context";
import AttachmentPanel from "./AttachmentPanel";
import styles from "./App.module.css";

interface props {}

const dialogContentProps = {
  type: DialogType.normal,
  title: "Delete Attachments",
  closeButtonAriaLabel: "Close",
  subText: "Are you sure you want to delete the selected attachments?",
};

const Actions = ({}: props) => {
  const vm = useAttachmentVM();
  const [isLoading, { setTrue: startLoading, setFalse: stopLoading }] =
    useBoolean(false);

  const modalProps = { isBlocking: false };

  return (
    <>
      <CommandBar items={vm.commandBarItems} farItems={vm.farCommandBarItems} />
      <AttachmentPanel />
      <Dialog
        hidden={!vm.isDeleteDialogOpen}
        onDismiss={() => {
          vm.toggleDeleteDialog();
        }}
        dialogContentProps={dialogContentProps}
        modalProps={modalProps}
      >
        {isLoading && (
          <FocusTrapZone disabled={!isLoading}>
            <Layer>
              <Overlay isDarkThemed={true}>
                <Spinner className={styles.spinner} size={SpinnerSize.large} />
              </Overlay>
            </Layer>
          </FocusTrapZone>
        )}
        <DialogFooter>
          <PrimaryButton
            onClick={async () => {
              startLoading();
              await vm.deleteSelectedAttachments();
              stopLoading();
              vm.toggleDeleteDialog();
            }}
            text='Delete'
          />
          <DefaultButton
            onClick={() => {
              vm.toggleDeleteDialog();
            }}
            text='Cancel'
          />
        </DialogFooter>
      </Dialog>
    </>
  );
};

export default observer(Actions);
