import React from "react";
import { Collection } from "../../models/Collection";
import { NTFCollectionItem } from "../NTFCollectionItem";

type NTFCollectionListProps = {
    listCollections: Array<Collection>
}

type NTFCollectionListState = {

}

export default class NTFCollectionListOld extends React.Component<NTFCollectionListProps, NTFCollectionListState>
{
    render() {
        return (
            <>
                {
                    this.props.listCollections.map(
                        collection => <NTFCollectionItem collection={collection} />
                    )
                }
            </>
        )
    }
}