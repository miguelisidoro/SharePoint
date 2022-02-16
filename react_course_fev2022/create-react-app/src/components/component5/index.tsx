import React from "react";

type ComponentProps = {}

type ComponentState = {
    message: string;
}

export class Component5 extends React.Component<ComponentProps, ComponentState>
{  
    constructor(props: ComponentProps) {
        super(props);

        this.state = { message: 'This is a default message'};
    }

    onChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        console.log(e.target.value);
        this.setState({ message: e.target.value });
    }

    render() {
        return (
        <>
            <h1>{this.state.message}</h1>
            <input type="text" id="messageText" placeholder="changeMessage" onChange={ this.onChange } />
        </>)
      }
}