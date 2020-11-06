import * as React from 'react';
import { IGraphGroupsProps } from './IGraphGroupsProps';

import { graph } from "@pnp/graph/presets/all";
import { IGraphGroupsState } from './IGraphGroupsState';
import { Shimmer } from 'office-ui-fabric-react/lib/components/Shimmer/Shimmer';
import { ShimmerElementType } from 'office-ui-fabric-react/lib/components/Shimmer';
import GroupList from './GroupList';

export default class GraphGroups extends React.Component<IGraphGroupsProps, IGraphGroupsState> {

  constructor(props: IGraphGroupsProps) {
    super(props);
    this.state = {
      groups: []
    };
  }

  public componentDidMount(): void {
    graph.groups
      .setEndpoint("beta")
      .top(20)
      .select("id, displayName, description")
      .filter("resourceProvisioningOptions/Any(x:x eq 'Team')") // only Teams: https://docs.microsoft.com/en-us/graph/teams-list-all-teams#get-a-list-of-groups-using-beta-apis
      .get()
      .then(groups => {
        this.setState({
          groups: groups.map((g, index) => {
            index += 10;
            return {
              displayName: g.displayName,
              id: g.id,
              description: g.description,
              thumbnailUrl: `https://picsum.photos/id/${index}/200/100`
            };
          })
        });
      });
  }

  public render(): React.ReactElement<IGraphGroupsProps> {
    if(this.state.groups.length <= 0) {
      return(
        <Shimmer
          shimmerElements={[
            { type: ShimmerElementType.line, width: 246, height: 246 },
            { type: ShimmerElementType.gap, width: '2%' },
            { type: ShimmerElementType.line, width: 246, height: 246 },
            { type: ShimmerElementType.gap, width: '2%' },
            { type: ShimmerElementType.line, width: 246, height: 246 },
            { type: ShimmerElementType.gap, width: '2%' },
            { type: ShimmerElementType.line, width: '100%', height: 246 }
          ]}
        />
      );
    }

    return (
      <GroupList
        groups={this.state.groups}
        isTeamsMessagingExtension={this.props.isTeamsMessagingExtension}
        teamsContext={this.props.teamsContext} />
    );
  }
}
