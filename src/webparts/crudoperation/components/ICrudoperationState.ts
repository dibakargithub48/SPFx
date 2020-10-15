import { IListItem } from "./IListItem";

export interface ICrudoperationState {
  status: string;
  items: IListItem[];
  textValue: string;
}
