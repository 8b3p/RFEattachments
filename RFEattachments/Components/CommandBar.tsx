import {
  CommandBar,
  ICommandBarItemProps,
  IDropdownOption,
  setVirtualParent,
} from "@fluentui/react";
import { observer } from "mobx-react-lite";
import * as React from "react";
import { useEffect, useState } from "react";
import { axa_attachment_axa_attachment_axa_type } from "../cds-generated/enums/axa_attachment_axa_attachment_axa_type";
import { useAttachmentVM } from "../Context/context";

interface props {
  typeOptions: IDropdownOption[];
}

const Actions = ({}: props) => {
  const vm = useAttachmentVM();
  const [file, setFile] = useState<File>();
  const [typeInput, setTypeInput] =
    useState<axa_attachment_axa_attachment_axa_type>();
  const [commandBarItems, setCommandBarItems] = useState<
    ICommandBarItemProps[]
  >([]);
  const [farCommandBarItems, setFarCommandBarItems] =
    useState<ICommandBarItemProps[]>();

  const onTypeSelectHandler = (
    option: axa_attachment_axa_attachment_axa_type
  ): void => {
    if (option) {
      console.log("theres options");
      setTypeInput(option);
    }
  };

  const fileInputChangeHandler = (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    if (event.target.files) {
      console.log("theres files");
      setFile(event.target.files[0]);
    } else {
      setFile(undefined);
    }
  };

  const buttonClickedHandler = () => {
    if (file && typeInput) {
      vm.uploadFile(file, typeInput);
      setFile(undefined);
      setTypeInput(undefined);
    }
  };

  useEffect(() => {
    console.log("useEffect ran");
    setCommandBarItems([
      {
        key: "newItem",
        text: "New",
        disabled: file && typeInput ? false : true,
        cacheKey: "myCacheKey", // changing this key will invalidate this item's cache
        iconProps: { iconName: "Add" },
        onClick: () => {
          buttonClickedHandler();
        },
      },
      {
        key: "file",
        text: file ? vm.cdsService.TrimFileExtension(file.name) : "Select File",
        iconProps: { iconName: "Upload" },
        preferMenuTargetAsEventTarget: true,
        onClick: (
          ev?:
            | React.MouseEvent<HTMLElement, MouseEvent>
            | React.KeyboardEvent<HTMLElement>
            | undefined
        ) => {
          ev?.persist();

          Promise.resolve().then(() => {
            const inputElement = document.createElement("input");
            inputElement.style.visibility = "hidden";
            inputElement.setAttribute("type", "file");

            document.body.appendChild(inputElement);

            const target = ev?.target as HTMLElement | undefined;

            if (target) {
              setVirtualParent(inputElement, target);
            }

            inputElement.onchange = (ev: any) => {
              fileInputChangeHandler(ev);
            };

            inputElement.click();

            if (target) {
              setVirtualParent(inputElement, null);
            }

            setTimeout(() => {
              inputElement.remove();
            }, 30000);
          });
        },
      },
      {
        key: "type",
        text: typeInput
          ? axa_attachment_axa_attachment_axa_type[typeInput]
          : "Choose Type",
        cacheKey: "myCacheKey", // changing this key will invalidate this item's cache
        iconProps: { iconName: "NumberedListText" },
        subMenuProps: {
          items: [
            {
              key: "dfa",
              text: "DFA",
              onClick: () => {
                onTypeSelectHandler(axa_attachment_axa_attachment_axa_type.DFA);
              },
            },
            {
              key: "bidwaiver",
              text: "Bid Waiver",
              onClick: () => {
                onTypeSelectHandler(
                  axa_attachment_axa_attachment_axa_type.BidWaiver
                );
              },
            },
            {
              key: "miscdocs",
              text: "Misc Docs",
              onClick: () => {
                onTypeSelectHandler(
                  axa_attachment_axa_attachment_axa_type.MiscDocs
                );
              },
            },
            {
              key: "quote",
              text: "Quote",
              onClick: () => {
                onTypeSelectHandler(
                  axa_attachment_axa_attachment_axa_type.Quote
                );
              },
            },
          ],
        },
      },
    ]);
    setFarCommandBarItems([
      {
        key: "delete",
        text: "Delete",
        // This needs an ariaLabel since it's icon-only
        ariaLabel: "Grid view",
        disabled: vm.isDeleteEnabled ? false : true,
        iconOnly: true,
        iconProps: { iconName: "Delete" },
        onClick: () => {
          vm.deleteSelectedAttachments();
        },
      },
      {
        key: "gotorecord",
        text: "Go To Record",
        // This needs an ariaLabel since it's icon-only
        ariaLabel: "Go To Record",
        iconOnly: true,
        iconProps: { iconName: "FileSymlink" },
        onClick: () => console.log("Going To Record!"),
      },
    ]);
  }, [file, typeInput, vm.isDeleteEnabled]);

  
  return <CommandBar items={commandBarItems} farItems={farCommandBarItems} />;
};

export default observer(Actions);
