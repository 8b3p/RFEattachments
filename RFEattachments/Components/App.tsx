import * as React from "react";
import {
  DetailsList,
  DetailsListLayoutMode,
  MarqueeSelection,
  Selection,
  Announced,
  PrimaryButton,
  Dropdown,
  IDropdownOption,
} from "@fluentui/react";
import styles from "./App.module.css";
import AttachmentVMProvider from "../Context/context";
import useDetailsList from "../Hooks/useDetailsList";
import { axa_attachment_axa_attachment_axa_type } from "../cds-generated/enums/axa_attachment_axa_attachment_axa_type";
import { observer } from "mobx-react-lite";
import { ServiceProvider } from "pcf-react";
import AttachmentVM, { AttachmentVMserviceName } from "../Context/AttachmentVM";

interface props {
  serviceProvider: ServiceProvider;
}

const App = ({ serviceProvider }: props) => {
  const vm = serviceProvider.get<AttachmentVM>(AttachmentVMserviceName);
  const {
    items,
    columns,
    isCompactMode,
    announcedMessage,
    setSelectionDetails,
  } = useDetailsList({vm});
  const [files, setFiles] = React.useState<File | null>(null);
  const [typeInput, setTypeInput] =
    React.useState<axa_attachment_axa_attachment_axa_type>();

  const selection = new Selection({
    onSelectionChanged: () => {
      setSelectionDetails(getSelectionDetails());
    },
  });

  const getSelectionDetails = () => {
    const selectionCount = selection.getSelectedCount();
    switch (selectionCount) {
      case 0:
        return "No items selected";
      case 1:
        return "1 item selected: " + (selection.getSelection()[0] as any).name;

      default:
        return `${selectionCount} items selected`;
    }
  };

  const getKey = (item: any) => {
    return item.key;
  };

  const onTypeSelectHandler = (
    _ev: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (option) {
      setTypeInput(option.key as axa_attachment_axa_attachment_axa_type);
    }
  };

  const fileInputChangeHandler = (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    setFiles(event.target.files?.item(0) || null);
  };

  const buttonClickedHandler = () => {
    if (files && typeInput) {
      vm.uploadFile(files, typeInput);
    }
  };

  const options: IDropdownOption[] = [
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
  return (
    <AttachmentVMProvider value={vm}>
      <div className={styles.container}>
        <div className={styles.controlWrapper}>
          <PrimaryButton text='New' onClick={buttonClickedHandler} />
          <Dropdown
            options={options}
            placeholder='Select Type'
            onChange={onTypeSelectHandler}
          />
          <Announced
            message={`Number of items after filter applied: ${items.length}.`}
          />
          <input type='file' onChange={fileInputChangeHandler} />
        </div>
        {announcedMessage && <Announced message={announcedMessage} />}
        {/* {selectionDetails && (
        <Label className={styles.selectionDetails}>{selectionDetails}</Label>
      )} */}

        <MarqueeSelection selection={selection}>
          <DetailsList
            items={items}
            compact={isCompactMode}
            columns={columns}
            // selectionMode={SelectionMode.multiple}
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
    </AttachmentVMProvider>
  );
};

export default observer(App);

// setItems(prev => {
//   return text
//     ? prev.filter(i => i.name.toLowerCase().indexOf(text) > -1)
//     : prev;
// });

//* this is for filtering the items from the search box
