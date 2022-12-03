import {
  CommandBar,
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  FocusTrapZone,
  ICommandBarItemProps,
  Layer,
  Overlay,
  PrimaryButton,
  Spinner,
  SpinnerSize,
} from "@fluentui/react";
import { observer } from "mobx-react-lite";
import { useBoolean } from "@fluentui/react-hooks";
import * as React from "react";
import { useEffect, useState } from "react";
import { useAttachmentVM } from "../Context/context";
import NewAttachmentPanel from "./NewAttachmentPanel";
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
  const [isOpen, { setTrue: openPanel, setFalse: dismissPanel }] =
    useBoolean(false);
  const [commandBarItems, setCommandBarItems] = useState<
    ICommandBarItemProps[]
  >([]);

  const modalProps = { isBlocking: false };

  useEffect(() => {
    setCommandBarItems([
      {
        key: "newItem",
        text: "New",
        cacheKey: "myCacheKey", // changing this key will invalidate this item's cache
        iconProps: { iconName: "Add" },
        onClick: () => {
          openPanel();
        },
      },
    ]);
  }, [openPanel, vm.Attachments]);

  return (
    <>
      <CommandBar items={commandBarItems} farItems={vm.farCommandBarItems} />
      <NewAttachmentPanel isOpen={isOpen} dismissPanel={dismissPanel} />
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
