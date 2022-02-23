import {configureStore} from '@reduxjs/toolkit'
import * as example from './example.duck'
import * as collections from './collections.duck'

const store = configureStore({
    reducer: {
        example: example.reducer,
        collections: collections.reducer
    }
});

export type RootState = ReturnType<typeof store.getState>

export default store

