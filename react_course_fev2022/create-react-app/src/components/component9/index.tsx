import React, { useState, useEffect } from "react";

type ComponentProps = {
    message: string;
}

// Functional component
export const Component9 = (props: ComponentProps) => {
    const [message, setMessage] = useState<string>(props.message)
    
    //componentDidMount - useEffect permite manipular lifecycle do componente
    //componentDidMount = []
    useEffect(() => {
        console.log("Component did mount")
    }, []);

    //componentDidUpdate = [message] - semelhante a:
    useEffect(() => {
        console.log("Fetch data from backend - search as you type")
    }, [message]);

    console.log("Render funcional component");

    return (
        <>
            <input type="text" onChange={e => setMessage(e.target.value)} />
            <h1>{message}</h1>
        </>
    );
}