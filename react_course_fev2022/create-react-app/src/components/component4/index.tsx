import React from "react";

type ComponentProps = {
    message: string;
}

type ComponentState = {}

export class Component4 extends React.Component<ComponentProps, ComponentState>
{
    render() {
        return <h1>{this.props.message}</h1>;
      }
}