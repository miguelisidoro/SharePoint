import React from "react";
import {useParams} from "react-router-dom";

type ComponentProps = {
    message: string;
}

// Old class component definition
// type ComponentState = {}

// export class Component4 extends React.Component<ComponentProps, ComponentState>
// {
//     render() {
//         return <h1>{this.props.message}</h1>;
//     }
// }

// Functional component
export const Component8 = (props: ComponentProps) => {

    const {defaultMessage} = useParams()

    return <h1>{defaultMessage}</h1>;
}