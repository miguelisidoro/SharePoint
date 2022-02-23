import React, { useEffect } from 'react';
import { NTFCollectionList } from '../components/NTFCollectionList';
import { Collection } from '../models/Collection';
import axios from 'axios';
import { useDispatch, useSelector } from 'react-redux';
import * as collectionsStore from '../store/collections.duck';
import { RootState } from '../store/'
import { NTFCollectionItem } from '../components/NTFCollectionItem';

const collectionList: Array<Collection> = []

const useCollections = () => {
    const dispatch = useDispatch();

    const collections = useSelector((state: RootState) => state.collections.listCollections);

    useEffect(() => {
        axios.get('https://api.opensea.io/api/v1/collections?offset=0&limit=10')
            .then(function (response) {
                const payload = {
                    collections: response.data.collections
                }

                dispatch(collectionsStore.actions.setCollectionList(payload))
            })
            .catch(function (error) {
                // handle error
                console.log(error);
            })
            .then(function () {
                // always executed
            })
    }, []);

    return collections;
}

export const NTFCollectionsListPage = () => {

    const collectionList = useCollections();

    return <>
        <NTFCollectionList listCollections={collectionList} />
    </>
}