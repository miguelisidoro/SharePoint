import React from "react";
import { Link } from "react-router-dom";
import { Collection } from "../../models/Collection";
import { NTFCollectionItem } from "../NTFCollectionItem";

type NTFCollectionListProps = {
  listCollections?: Array<Collection>
}

export const NTFCollectionList = (props: NTFCollectionListProps) => {
  return (
    <>
      {
        props.listCollections?.map(
          collection => 
          <Link to={`/collections/${collection.slug}`}>
            <NTFCollectionItem collection={collection} />
          </Link>
        )
      }
    </>
  )
}