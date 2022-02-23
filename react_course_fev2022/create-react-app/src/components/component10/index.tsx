import React, { useState, useEffect } from "react";

type ComponentProps = {
    message: string;
}

const useMessage = (props: ComponentProps) => {
    const [message, setMessage] = useState<string>(props.message)
    
    //componentDidMount - useEffect permite manipular lifecycle do componente
    //componentDidMount = []
    useEffect(() => {
        console.log("Component did mount")
    }, []);

    //componentDidUpdate = [message] - semelhante a:
    useEffect(() => {
        console.log("Fetch data from backend - executed if message variable is changed")
    }, [message]);

    return {message, setMessage};
}

// Functional component
export const Component10 = (props: ComponentProps) => {
    
    console.log("Render funcional component");

    const {message, setMessage} = useMessage(props);

    return (
        <>
            <input type="text" onChange={e => setMessage(e.target.value)} />
            <h1>{message}</h1>
        </>
    );
}