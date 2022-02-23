import React from "react";
import { useSelector, useDispatch } from 'react-redux'
import { RootState } from "../../store";
import * as example from "../../store/example.duck"

// Functional component
export const Component11 = () => {

    const dispatch = useDispatch()

    const message = useSelector((state: RootState) => {
      return state.example.message;
     })

    const onChange = (e: React.ChangeEvent<HTMLInputElement>) => {
      const payload = {
        message: e.target.value
      }

      dispatch(example.actions.setMessage(payload));
    }

    return <>
      <input type="text" onChange={onChange} />
      <h1>{message}</h1>
    </>;
}