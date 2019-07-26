import * as React from "react";
import styles from "./AgarbList.module.scss";
import { IAgarbListProps } from "./IAgarbListProps";
import { escape } from "@microsoft/sp-lodash-subset";
import SimpleTable from "./SimpleTable";

export default class AgarbList extends React.Component<IAgarbListProps, {}> {
  constructor(props: IAgarbListProps) {
    super(props);
  }

  public render(): JSX.Element {
    return (
      <div className={styles.agarbList}>
        <div className={styles.container}>
          <SimpleTable items={this.props.items} />
        </div>
      </div>
    );
  }
}
