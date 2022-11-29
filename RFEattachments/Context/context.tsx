import React = require("react");
import AttachmentVM from "./AttachmentVM";

export const AttachmentVMContext = React.createContext<AttachmentVM>(
  {} as AttachmentVM
);

interface props {
  value: AttachmentVM;
  children: React.ReactNode;
}

const AttachmentVMProvider = ({ value, children }: props) => {
  return (
    <AttachmentVMContext.Provider value={value}>
      {children}
    </AttachmentVMContext.Provider>
  );
};

export const useAttachmentVM = () => React.useContext(AttachmentVMContext);

export default AttachmentVMProvider;
