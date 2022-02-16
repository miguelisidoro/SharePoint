import React from "react";

type ChildComponentProps = {
    message: string;
}

type ParentComponentProps = {
    message: string;
}

type ParentComponentState = {
    parentMessage: string;
}

class ChildComponent extends React.Component<ChildComponentProps>
{
    render() {
        return (
            <>
                <h1>Child Component</h1>
            </>)
    }
}

export default class ParentComponent extends React.Component<ParentComponentProps, ParentComponentState>
{
    constructor(props: ParentComponentProps) {
        super(props);

        this.state = {
            parentMessage: "Default message"
        }
    }

    onChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        console.log(e.target.value);
        this.setState({ parentMessage: e.target.value });
    }

    render() {
        return (
            <>
                <h1>Parent Component</h1>
                <input type="text" onChange={this.onChange}></input>
                <ChildComponent message={this.state.parentMessage}></ChildComponent>
            </>)
    }
}