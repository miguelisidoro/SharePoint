import React from "react";
import { Collection } from "../../models/Collection";
import ListItem from '@mui/material/ListItem';
import ListItemText from '@mui/material/ListItemText';
import ListItemAvatar from '@mui/material/ListItemAvatar';
import Avatar from '@mui/material/Avatar';
import Typography from '@mui/material/Typography';

type NTFCollectionItemProps = {
  collection: Collection
}

export const NTFCollectionItem = (props: NTFCollectionItemProps) => {
  return (
    <>
      <ListItem alignItems="flex-start">
        <ListItemAvatar>
          <Avatar alt="Remy Sharp" src={props.collection.imageUrl} />
        </ListItemAvatar>
        <ListItemText
          primary={props.collection.name}
          secondary={
            <React.Fragment>
              <Typography
                sx={{ display: 'inline' }}
                component="span"
                variant="body2"
                color="text.primary"
              >
                {props.collection.description}
              </Typography>
            </React.Fragment>
          }
        />
      </ListItem>
    </>)
}