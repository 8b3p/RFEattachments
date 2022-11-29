import * as React from "react";
import {
  DetailsList,
  DetailsListLayoutMode,
  Label,
  MarqueeSelection,
  SelectionMode,
  Selection,
  Announced,
  TextField,
  Toggle,
} from "@fluentui/react";
import useDetailsList from "../Hooks/useDetailsList";
import styles from "./App.module.css";
import { useEffect } from "react";

const App = () => {
  const {
    items,
    setItems,
    columns,
    isCompactMode,
    announcedMessage,
    isModalSelection,
    selectionDetails,
    setSelectionDetails,
  } = useDetailsList();

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

  useEffect(() => {
    console.dir(items);
    console.dir(columns);
    console.log(announcedMessage);
  }, [items, columns, announcedMessage]);

  const onTextChange = (
    _ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    text?: string
  ): void => {
    setItems(prev => {
      return text
        ? prev.filter(i => i.name.toLowerCase().indexOf(text) > -1)
        : prev;
    });
  };

  return (
    <div className={styles.container}>
      <div className={styles.controlWrapper}>
        <TextField
          label='Filter by name:'
          onChange={onTextChange}
          styles={styles["control-styles"]}
        />
        <Announced
          message={`Number of items after filter applied: ${items.length}.`}
        />
      </div>
      {announcedMessage && <Announced message={announcedMessage} />}
      <Label className={styles.selectionDetails}>{selectionDetails}</Label>

      <MarqueeSelection selection={selection}>
        <DetailsList
          items={items}
          compact={isCompactMode}
          columns={columns}
          selectionMode={SelectionMode.multiple}
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

export default App;
