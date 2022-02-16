import React from "react";

type ChildComponentProps = {
    onMessageChange: (message: string) => void
}

type ParentComponentProps = {
    message: string;
}

type ParentComponentState = {
    message: string;
}

class ChildComponent extends React.Component<ChildComponentProps>
{
    render() {
        return (
            <>
                <h1>Child Component</h1>
                <input type="text" onChange={e => this.props.onMessageChange(e.target.value)} />
            </>)
    }
}

export default class ParentComponent extends React.Component<ParentComponentProps, ParentComponentState>
{
    constructor(props: ParentComponentProps) {
        super(props);

        this.state = {
            message: "Default message"
        }
    }

    onChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        console.log(e.target.value);
        this.setState({ message: e.target.value });
    }

    changeMessage(message: string) {
        this.setState({ message });
    }

    render() {
        return (
            <>
                <h1>Parent Component</h1>
                <ChildComponent onMessageChange={this.changeMessage}></ChildComponent>
            </>)
    }
}