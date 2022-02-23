import React, { useState } from "react";

type ComponentProps = {
    message: string;
}

// Functional component
export const Component9 = (props: ComponentProps) => {
    const [message, setMessage] = useState<string>(props.message)
    
    return (
        <>
            <input type="text" onChange={e => setMessage(e.target.value)} />
            <h1>{props.message}</h1>
        </>
    );
}