import {createAction, createReducer, PayloadAction} from '@reduxjs/toolkit'
import { Reducer } from 'redux'

export type ExampleState = {
    message?: string;
}

const initialState: ExampleState = {
    message: 'This is a default message'
}

type SetMessageActionType = {
    message: string;
}

export const actions = {
    setMessage: createAction<SetMessageActionType>('SET_MESSAGE')
}

export const reducer: Reducer<ExampleState> = createReducer(initialState, {
    [actions.setMessage.type]: (state: ExampleState, action: PayloadAction<SetMessageActionType>) => {
        return {...state, message: action.payload.message}
    }
})
