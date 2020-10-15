import * as React from "react";
import styles from "./Crudoperation.module.scss";
import { ICrudoperationProps } from "./ICrudoperationProps";
import { escape } from "@microsoft/sp-lodash-subset";

import { ICrudoperationState } from "./ICrudoperationState";
import { IListItem } from "./IListItem";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IItemAddResult } from "@pnp/sp/items";

import {
  DefaultButton,
  PrimaryButton,
  Stack,
  IStackTokens,
} from "office-ui-fabric-react";

export default class Crudoperation extends React.Component<
  ICrudoperationProps,
  ICrudoperationState
> {
  constructor(props: ICrudoperationProps, state: ICrudoperationState) {
    super(props);

    this.state = {
      status: "Ready",
      items: [],
      textValue: "",
    };
  }

  public render(): React.ReactElement<ICrudoperationProps> {
    const items: JSX.Element[] = this.state.items.map(
      (item: IListItem, i: number): JSX.Element => {
        return (
          <li>
            ({item.Id}) - {item.Title}{" "}
          </li>
        );
      }
    );

    return (
      <div className={styles.crudoperation}>
        <div
          className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}
        >
          <div className="ms-Grid" dir="ltr">
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg2">
                <PrimaryButton text="Get All Item" onClick={this.GetAllItems} />
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg2">
                <PrimaryButton text="Create Item" onClick={this.CreateItem} />
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg2">
                <PrimaryButton text="Update Item" onClick={this.UpdateItem} />
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg2">
                <PrimaryButton text="Delete Item" onClick={this.DeleteItem} />
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg2"></div>
            </div>
            <div style={{ marginTop: 10 }} className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg4">
                <input
                  type="text"
                  value={this.state.textValue}
                  onChange={this.handleChange}
                />
              </div>
            </div>
          </div>
          <div style={{ marginTop: 20 }} className="ms-Grid" dir="ltr">
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md12 ms-lg12">
                Current Status : {this.state.status}
                <br />
                Textbox Value : {this.state.textValue}
                <br />
                <ul>{items}</ul>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private CreateItem = async () => {
    this.setState({
      status: "Creating item...",
      items: [],
    });

    // add an item to the list
    const iar: IItemAddResult = await sp.web.lists
      .getByTitle(this.props.listName)
      .items.add({
        Title: `Item ${new Date()}`,
      });

    console.log(iar);

    this.GetAllItems();
  };

  private GetAllItems = async () => {
    // get all the items from a list
    this.setState({
      status: "Retrieved items...",
      items: [],
    });

    const allItems: any[] = await sp.web.lists
      .getByTitle(this.props.listName)
      .items.get();
    console.log(allItems);

    this.setState({
      status: "Retrieved items...",
      items: allItems,
    });
  };

  private UpdateItem = async () => {
    let list = sp.web.lists.getByTitle(this.props.listName);
    let itemToUpdateId = parseInt(this.state.textValue);

    const i = await list.items.getById(itemToUpdateId).update({
      Title: `Item Updated ${new Date()}`,
    });
    console.log(i);

    this.GetAllItems();
  };

  private DeleteItem = async () => {
    let list = sp.web.lists.getByTitle(this.props.listName);
    let itemToDeleteId = parseInt(this.state.textValue);
    await list.items.getById(itemToDeleteId).delete();

    this.GetAllItems();
  };

  private handleChange = (event) => {
    this.setState({ textValue: event.target.value });
  };
}
