import { MSGraphClientV3 } from "@microsoft/sp-http";

export interface IAdUserFormProps {
  graphClient: MSGraphClientV3;
  onClose?: () => void;
}
