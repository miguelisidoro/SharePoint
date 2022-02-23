import {configureStore} from '@reduxjs/toolkit'
import * as example from './example.duck'

const store = configureStore({
    reducer: {
        example: example.reducer
    }
});

export type RootState = ReturnType<typeof store.getState>

export default store

