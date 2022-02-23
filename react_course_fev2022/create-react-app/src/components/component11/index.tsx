import React from "react";
import { useSelector } from 'react-redux'
import { RootState } from "../../store";

// Functional component
export const Component11 = () => {

    const message = useSelector((state: RootState) => {
      return state.example.message;
     })

    return <h1>{message}</h1>;
}