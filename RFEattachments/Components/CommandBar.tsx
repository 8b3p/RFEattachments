import {
  CommandBar,
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  FocusTrapZone,
  Layer,
  MessageBar,
  MessageBarType,
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

const dialogContentProps = {
  type: DialogType.normal,
  title: "Delete Attachments",
  closeButtonAriaLabel: "Close",
  subText: "Are you sure you want to delete the selected attachments?",
};

const Actions = () => {
  const vm = useAttachmentVM();
  const [isLoading, { setTrue: startLoading, setFalse: stopLoading }] =
    useBoolean(false);
  const [error, setError] = React.useState<string | undefined>(undefined);

  const modalProps = { isBlocking: false };

  return (
    <>
      <CommandBar items={vm.commandBarItems} farItems={vm.farCommandBarItems} />
      {error && (
        <MessageBar
          messageBarType={MessageBarType.error}
          onDismiss={() => setError(undefined)}
          dismissButtonAriaLabel='Close'
        >
          {error}
        </MessageBar>
      )}
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
              let response = await vm.deleteSelectedAttachments();
              vm.toggleDeleteDialog();
              stopLoading();
              if (response instanceof Error) {
                setError(response.message);
              }
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
