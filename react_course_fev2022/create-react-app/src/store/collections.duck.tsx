import {createAction, createReducer, PayloadAction} from '@reduxjs/toolkit'
import { Reducer } from 'redux'
import { Collection } from '../models/Collection'

export type CollectionState = {
    listCollections?: Array<Collection>,
    selectedCollection?: Collection,
    search?: string
}

const initialState: CollectionState = {
    listCollections: [],
    selectedCollection: undefined,
    search: undefined
}

type SetSelectedCollectionActionType = {
    selectedCollection: Collection;
}

type SetCollectionsActionType = {
    collections?: Array<Collection>;
}

export const actions = {
    setSelectedCollection: createAction<SetSelectedCollectionActionType>('SET_SELECTED_COLLECTION'),
    setCollectionList: createAction<SetCollectionsActionType>('SET_COLLECTIONS')
}

export const reducer: Reducer<CollectionState> = createReducer(initialState, {
    [actions.setSelectedCollection.type]: (state: CollectionState, action: PayloadAction<SetSelectedCollectionActionType>) => {
        return {...state, selectedCollection: action.payload.selectedCollection}
    },
    [actions.setCollectionList.type]: (state: CollectionState, action: PayloadAction<SetCollectionsActionType>) => {
        return {...state, listCollections: action.payload.collections}
    }
})
