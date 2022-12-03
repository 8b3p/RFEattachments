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
} from "@fluentui/react";
import { observer } from "mobx-react-lite";
import * as React from "react";
import { axa_attachment_axa_attachment_axa_type } from "../cds-generated/enums/axa_attachment_axa_attachment_axa_type";
import { useAttachmentVM } from "../Context/context";
import styles from "./App.module.css";
import { useBoolean } from "@fluentui/react-hooks";

interface props {
  dismissPanel: () => void;
  isOpen: boolean;
}

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

const NewAttachmentPanel = ({ isOpen, dismissPanel }: props) => {
  const vm = useAttachmentVM();
  const [errorMessage, setErrorMessage] = React.useState<string>();
  const [file, setFile] = React.useState<File>();
  const [typeInput, setTypeInput] =
    React.useState<axa_attachment_axa_attachment_axa_type>();
  const [isLoading, { setTrue: startLoading, setFalse: stopLoading }] =
    useBoolean(false);

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
        } else {
          setFile(undefined);
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
    console.dir(file);
    console.dir(typeInput);
    console.log(file && typeInput);
    if (file && typeInput) {
      startLoading();
      await vm.uploadFile(file, typeInput);
      setFile(undefined);
      setTypeInput(undefined);
      setErrorMessage(undefined);
      stopLoading();
      dismissPanel();
    }
    setErrorMessage("Please select a file and type");
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
          setTypeInput(undefined);
          setErrorMessage(undefined);
          dismissPanel();
        }}
      >
        Cancel
      </DefaultButton>
    </div>
  );

  return (
    <div>
      <Panel
        headerText='Create New Attachment'
        isOpen={isOpen}
        onOpen={() => {
          setFile(undefined);
          setTypeInput(undefined);
          setErrorMessage(undefined);
        }}
        onDismiss={() => {
          setFile(undefined);
          setTypeInput(undefined);
          setErrorMessage(undefined);
          dismissPanel();
        }}
        onRenderFooterContent={onRenderFooterContent}
        // You MUST provide this prop! Otherwise screen readers will just say "button" with no label.
        closeButtonAriaLabel='Close'
        isFooterAtBottom={true}
      >
        <Stack horizontal>
          <Stack.Item grow={1}>
            <Dropdown
              placeholder='Choose Type'
              options={typeOptions}
              onChange={(e, option?) => {
                setErrorMessage(undefined);
                onTypeSelectHandler(e, option);
              }}
            />
          </Stack.Item>
          <Stack.Item grow={1}>
            <DefaultButton
              text={file ? file.name : "Select File"}
              onClick={e => {
                setErrorMessage(undefined);
                onFileInputChangeHandler(e);
              }}
            />
          </Stack.Item>
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

export default observer(NewAttachmentPanel);
