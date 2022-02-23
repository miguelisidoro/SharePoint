import React from "react";
import { Collection } from "../../models/Collection";
import { NTFCollectionItem } from "../NTFCollectionItem";

type NTFCollectionListProps = {
  listCollections: Array<Collection>
}

export const NTFCollectionList = (props: NTFCollectionListProps) => {
  return (
    <>
      {
        props.listCollections.map(
          collection => <NTFCollectionItem collection={collection} />
        )
      }
    </>
  )
}