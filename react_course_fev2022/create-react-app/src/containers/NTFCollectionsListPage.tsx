import React from 'react';
import NTFCollectionList from '../components/NTFCollectionList';
import { Collection } from '../models/Collection';

const collectionList: Array<Collection> = [
    new Collection(
        { name: 'Name 1', description: 'description 1', imageUrl: '/assets/1.jpg' },
    ),
    new Collection(
        { name: 'Name 2', description: 'description 2', imageUrl: '/assets/2.jpg' },
    ),
    new Collection(
        { name: 'Name 3', description: 'description 3', imageUrl: '/assets/3.jpg' }
    )
];

export class NTFCollectionsListPage extends React.Component {
    render() {
        return (
            <NTFCollectionList listCollections={collectionList} />
        )
    }
}