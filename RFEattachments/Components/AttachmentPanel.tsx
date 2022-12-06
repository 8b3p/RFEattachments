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
} from "@fluentui/react";
import { observer } from "mobx-react-lite";
import * as React from "react";
import { axa_attachment_axa_attachment_axa_type } from "../cds-generated/enums/axa_attachment_axa_attachment_axa_type";
import { useAttachmentVM } from "../Context/context";
import styles from "./App.module.css";
import { useBoolean } from "@fluentui/react-hooks";
import { Attachment } from "../types/Attachment";

const typeOptions: IDropdownOption[] = [
  {
    key: axa_attachment_axa_attachment_axa_type.BidWaiver,
    text: axa_attachment_axa_attachment_axa_type[
      axa_attachment_axa_attachment_axa_type.BidWaiver
    ],
  },
  {
    key: axa_attachment_axa_attachment_axa_type.DFA,
    text: axa_attachment_axa_attachment_axa_type[
      axa_attachment_axa_attachment_axa_type.DFA
    ],
  },
  {
    key: axa_attachment_axa_attachment_axa_type.MiscDocs,
    text: axa_attachment_axa_attachment_axa_type[
      axa_attachment_axa_attachment_axa_type.MiscDocs
    ],
  },
  {
    key: axa_attachment_axa_attachment_axa_type.Quote,
    text: axa_attachment_axa_attachment_axa_type[
      axa_attachment_axa_attachment_axa_type.Quote
    ],
  },
];

const AttachmentPanel = () => {
  const vm = useAttachmentVM();
  const [errorMessage, setErrorMessage] = React.useState<string>();
  const [file, setFile] = React.useState<File>();
  const [typeInput, setTypeInput] =
    React.useState<axa_attachment_axa_attachment_axa_type>();
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
          console.log("theres files");
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
      if (file && typeInput) {
        startLoading();
        await vm.uploadFile(typeInput, file);
        setFile(undefined);
        setFileName(undefined);
        setTypeInput(undefined);
        setErrorMessage(undefined);
        stopLoading();
        vm.isPanelOpen = false;
      }
      setErrorMessage("Please select a file and type");
    } else if (vm.formType === "edit") {
      if (file && typeInput) {
        startLoading();
        await vm.uploadFile(typeInput, file);
        setFile(undefined);
        setFileName(undefined);
        setTypeInput(undefined);
        setErrorMessage(undefined);
        stopLoading();
        vm.isPanelOpen = false;
      } else if (typeInput) {
        startLoading();
        await vm.uploadFile(typeInput);
        setFile(undefined);
        setFileName(undefined);
        setTypeInput(undefined);
        setErrorMessage(undefined);
        stopLoading();
        vm.isPanelOpen = false;
      }
      setErrorMessage("Please select a file and type");
    }
  };

  React.useEffect(() => {
    console.log("useEffect ran");
    if (vm.formType === "edit") {
      console.log("vm.formType", vm.formType);
      fetchData();
    }
  }, [vm.isPanelOpen]);

  const fetchData = async () => {
    let selected = vm.selectedAttachments[0];
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
          setFile(undefined);
          setFileName(undefined);
          setTypeInput(undefined);
          setErrorMessage(undefined);
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
          setFile(undefined);
          setFileName(undefined);
          setTypeInput(undefined);
          setErrorMessage(undefined);
        }}
        headerTextProps={{ style: { marginBottom: "2em" } }}
        onDismiss={() => {
          setFile(undefined);
          setFileName(undefined);
          setTypeInput(undefined);
          setErrorMessage(undefined);
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