import React, { useEffect } from 'react';
import axios from 'axios';
import { useParams } from 'react-router-dom';
import { useDispatch, useSelector } from 'react-redux';
import * as collectionsStore from '../store/collections.duck';
import { RootState } from '../store/'
import { NTFCollectionItem } from '../components/NTFCollectionItem';

const useSelectedCollection = () => {
    const { collectionId } = useParams()
    const dispatch = useDispatch();

    const selectedCollection = useSelector((state: RootState) => state.collections.selectedCollection);

    useEffect(() => {
        axios.get('https://api.opensea.io/api/v1/collection/' + collectionId)
            .then(function (response) {
                const payload = {
                    selectedCollection: response.data.collection
                }

                dispatch(collectionsStore.actions.setSelectedCollection(payload))
            })
            .catch(function (error) {
                // handle error
                console.log(error);
            })
            .then(function () {
                // always executed
            })
    }, [collectionId]);

    return selectedCollection;
}

export const NTFCollectionsDetailPage = () => {
    const selectedCollection = useSelectedCollection();
    
    return <>
        <NTFCollectionItem collection={selectedCollection} />
    </>
}