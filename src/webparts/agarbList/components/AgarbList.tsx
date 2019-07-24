import * as React from 'react';
import styles from './AgarbList.module.scss';
import { IAgarbListProps } from './IAgarbListProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class AgarbList extends React.Component<IAgarbListProps, {}> {
  public render(): React.ReactElement<IAgarbListProps> {
    return (
      <div className={ styles.agarbList }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
