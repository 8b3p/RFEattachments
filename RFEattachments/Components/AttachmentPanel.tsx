import {
  DefaultButton,
  Dropdown,
  FocusTrapZone,
  IDropdownOption,
  Layer,
  MessageBar,
  MessageBarType,
  Overlay,
  Panel,
  PrimaryButton,
  setVirtualParent,
  Spinner,
  SpinnerSize,
  Stack,
  Text,
  TextField,
} from "@fluentui/react";
import { observer } from "mobx-react-lite";
import * as React from "react";
import { useAttachmentVM } from "../Context/context";
import styles from "./App.module.css";
import { useBoolean } from "@fluentui/react-hooks";
import { axa_attachment_axa_attachment_axa_type } from "../cds-generated/enums/axa_attachment_axa_attachment_axa_type";

export function seperateCamelCaseString(str: string) {
  return str.replace(/([a-z0-9])([A-Z])/g, "$1 $2");
}

const typeOptions: IDropdownOption[] = [
  {
    key: axa_attachment_axa_attachment_axa_type.DFA,
    text: seperateCamelCaseString(
      axa_attachment_axa_attachment_axa_type[
        axa_attachment_axa_attachment_axa_type.DFA
      ]
    ),
  },
  {
    key: axa_attachment_axa_attachment_axa_type.Quote,
    text: seperateCamelCaseString(
      axa_attachment_axa_attachment_axa_type[
        axa_attachment_axa_attachment_axa_type.Quote
      ]
    ),
  },
  {
    key: axa_attachment_axa_attachment_axa_type.MiscDocs,
    text: seperateCamelCaseString(
      axa_attachment_axa_attachment_axa_type[
        axa_attachment_axa_attachment_axa_type.MiscDocs
      ]
    ),
  },
  {
    key: axa_attachment_axa_attachment_axa_type.BidWaiver,
    text: seperateCamelCaseString(
      axa_attachment_axa_attachment_axa_type[
        axa_attachment_axa_attachment_axa_type.BidWaiver
      ]
    ),
  },
  {
    key: axa_attachment_axa_attachment_axa_type.PaybackWorksheet,
    text: seperateCamelCaseString(
      axa_attachment_axa_attachment_axa_type[
        axa_attachment_axa_attachment_axa_type.PaybackWorksheet
      ]
    ),
  },
];

const AttachmentPanel = () => {
  const vm = useAttachmentVM();
  const [errorMessage, setErrorMessage] = React.useState<string>();
  const [file, setFile] = React.useState<File>();
  const [typeInput, setTypeInput] =
    React.useState<axa_attachment_axa_attachment_axa_type>();
  const [description, setDescription] = React.useState<string>();
  const [isLoading, { setTrue: startLoading, setFalse: stopLoading }] =
    useBoolean(false);
  const [fileName, setFileName] = React.useState<string>();

  const onTypeSelectHandler = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption<any> | undefined
  ): void => {
    if (option) {
      setTypeInput(option.key as axa_attachment_axa_attachment_axa_type);
    }
  };

  const resetState = () => {
    setFile(undefined);
    setFileName(undefined);
    setTypeInput(undefined);
    setDescription("");
    setErrorMessage(undefined);
  };

  const onFileInputChangeHandler = (event: any) => {
    Promise.resolve().then(() => {
      const inputElement = document.createElement("input");
      inputElement.style.visibility = "hidden";
      inputElement.setAttribute("type", "file");

      document.body.appendChild(inputElement);

      const target = event?.target as HTMLElement | undefined;

      if (target) {
        setVirtualParent(inputElement, target);
      }

      inputElement.onchange = (ev: any) => {
        if (ev.target.files) {
          setFile(ev.target.files[0]);
          setFileName(ev.target.files[0].name);
        } else {
          setFile(undefined);
          setFileName(undefined);
        }
      };

      inputElement.click();

      if (target) {
        setVirtualParent(inputElement, null);
      }

      setTimeout(() => {
        inputElement.remove();
      }, 30000);
    });
  };

  const onSaveButtonClickHandler = async () => {
    if (vm.formType === "new") {
      if (file && typeInput && description) {
        startLoading();
        await vm.uploadFile(typeInput, description, file);
        resetState();
        stopLoading();
        vm.isPanelOpen = false;
      }
      setErrorMessage("Please fill out all fields");
    } else if (vm.formType === "edit") {
      if (file && typeInput && description) {
        startLoading();
        await vm.uploadFile(typeInput, description, file);
        resetState();
        stopLoading();
        vm.isPanelOpen = false;
      } else if (typeInput && description) {
        startLoading();
        const res = await vm.uploadFile(typeInput, description);
        resetState();
        stopLoading();
        vm.isPanelOpen = false;
      }
      setErrorMessage("Please fill out all fields");
    }
  };

  React.useEffect(() => {
    if (vm.formType === "edit" && vm.isPanelOpen) {
      fetchData();
    }
  }, [vm.isPanelOpen]);

  const fetchData = async () => {
    let selected = vm.selectedAttachments[0];
    setDescription(selected.description);
    setTypeInput(selected.type);
    setFileName(selected.fileName);
  };

  const onRenderFooterContent = () => (
    <div>
      <PrimaryButton
        onClick={onSaveButtonClickHandler}
        className={styles["panel-button-styles"]}
      >
        Save
      </PrimaryButton>
      <DefaultButton
        onClick={() => {
          resetState();

          vm.isPanelOpen = false;
        }}
      >
        Cancel
      </DefaultButton>
    </div>
  );

  return (
    <div>
      <Panel
        headerText={
          vm.formType === "new" ? "New Attachment" : "Edit Attachment"
        }
        isOpen={vm.isPanelOpen}
        onOpen={() => {
          resetState();
        }}
        headerTextProps={{ style: { marginBottom: "2em" } }}
        onDismiss={() => {
          resetState();

          vm.isPanelOpen = false;
        }}
        onRenderFooterContent={onRenderFooterContent}
        // You MUST provide this prop! Otherwise screen readers will just say "button" with no label.
        closeButtonAriaLabel='Close'
        isFooterAtBottom={true}
      >
        <Stack className={styles["dropdown-container"]} horizontal>
          {vm.formType === "edit" ? (
            <Dropdown
              placeholder='Choose Type'
              selectedKey={typeInput}
              options={typeOptions}
              onChange={(e, option?) => {
                setErrorMessage(undefined);
                onTypeSelectHandler(e, option);
              }}
            />
          ) : (
            <Dropdown
              placeholder='Choose Type'
              options={typeOptions}
              onChange={(e, option?) => {
                setErrorMessage(undefined);
                onTypeSelectHandler(e, option);
              }}
            />
          )}
        </Stack>
        <Stack className={styles["dropdown-container"]} horizontal>
          <TextField
            placeholder='Description'
            multiline
            autoAdjustHeight
            maxLength={300}
            value={description}
            onChange={(_e, value) => {
              setErrorMessage(undefined);
              setDescription(value || "");
            }}
          />
        </Stack>
        <Stack className={styles["dropdown-container"]} horizontal>
          <DefaultButton
            onClick={e => {
              setErrorMessage(undefined);
              onFileInputChangeHandler(e);
            }}
          >
            <Text nowrap>{fileName ? fileName : "Select File"}</Text>
          </DefaultButton>
        </Stack>
        {errorMessage && (
          <MessageBar
            messageBarType={MessageBarType.blocked}
            isMultiline={false}
            onDismiss={() => setErrorMessage(undefined)}
            dismissButtonAriaLabel='Close'
          >
            {errorMessage}
          </MessageBar>
        )}
      </Panel>
      {isLoading && (
        <FocusTrapZone disabled={!isLoading}>
          <Layer>
            <Overlay isDarkThemed={true}>
              <Spinner className={styles.spinner} size={SpinnerSize.large} />
            </Overlay>
          </Layer>
        </FocusTrapZone>
      )}
    </div>
  );
};

export default observer(AttachmentPanel);
