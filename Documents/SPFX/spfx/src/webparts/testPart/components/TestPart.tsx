import * as React from 'react';
import type { ITestPartProps } from './ITestPartProps';
import { Form } from "./Day04/Form";

export default class TestPart extends React.Component<ITestPartProps, {}> {
  public render(): React.ReactElement<ITestPartProps> {
    return (
      <div>
        {this.props.context && <Form context={this.props.context} />}
      </div>
    );
  }
}
