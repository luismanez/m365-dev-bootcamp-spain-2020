import * as React from 'react';
import styles from './GraphGroups.module.scss';

import { IGraphGroup } from "../../../models/IGraphGroup";
import GroupDetail from './GroupDetail';
import { IMicrosoftTeams } from '@microsoft/sp-webpart-base';

export interface IGroupListProps {
  groups: IGraphGroup[];
  isTeamsMessagingExtension?: boolean;
  teamsContext?: IMicrosoftTeams;
}

export interface IGroupListState {}

export default class GroupList extends React.Component<IGroupListProps, IGroupListState> {
  public render(): React.ReactElement<IGroupListProps> {

    const groups = this.props.groups.map(group => {
      return <GroupDetail
        group={group}
        isTeamsMessagingExtension={this.props.isTeamsMessagingExtension}
        teamsContext={this.props.teamsContext} />;
    });

    return (
      <div className={styles.graphGroups}>
        <div className={styles.cards}>
          {groups}
        </div>
      </div>
    );
  }
}
