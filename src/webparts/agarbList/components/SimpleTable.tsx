import * as React from "react";
import { createStyles, Theme, makeStyles } from "@material-ui/core/styles";
import Table from "@material-ui/core/Table";
import TableBody from "@material-ui/core/TableBody";
import TableCell from "@material-ui/core/TableCell";
import TableHead from "@material-ui/core/TableHead";
import TableRow from "@material-ui/core/TableRow";
import Paper from "@material-ui/core/Paper";
import IlistItem from "./IListItem";

const useStyles = makeStyles((theme: Theme) =>
  createStyles({
    root: {
      width: "100%",
      marginTop: theme.spacing(3),
      overflowX: "auto"
    },
    table: {
      minWidth: 650
    }
  })
);

export default function SimpleTable(props) {
  const classes = useStyles(props);

  console.log(props.items);

  return (
    <Paper className={classes.root}>
      <Table className={classes.table}>
        <TableHead>
          <TableRow>
            <TableCell>ID</TableCell>
            <TableCell align="right">Title</TableCell>
            <TableCell align="right">Modified</TableCell>
            <TableCell align="right">Modified by</TableCell>
          </TableRow>
        </TableHead>
        <TableBody>
          {props.items.map((item: IlistItem) => (
            <TableRow key={item.ID}>
              <TableCell component="th" scope="row">
                {item.ID}
              </TableCell>
              <TableCell align="right">{item.Title}</TableCell>
              <TableCell align="right">{item.Modified}</TableCell>
              <TableCell align="right">{item.ModifiedBy}</TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>
    </Paper>
  );
}
