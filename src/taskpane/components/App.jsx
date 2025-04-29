import * as React from "react";
import PropTypes from "prop-types";
import Header from "./Header";
import ParagraphManager from "./ParagraphManager";
import { makeStyles, tokens } from "@fluentui/react-components";
import { 
  processParagraphs, 
  updateParagraphById,
  addCommentToParagraph,
  highlightParagraph,
  scrollToParagraph
} from "../taskpane";

const useStyles = makeStyles({
  root: {
    height: "100vh",
    backgroundColor: tokens.colorNeutralBackground1,
    display: "flex",
    flexDirection: "column",
    overflow: "hidden",
  },
  content: {
    flex: 1,
    padding: "4px 8px 8px",
    overflowY: "auto",
    display: "flex",
    flexDirection: "column",
  }
});

const App = () => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <Header message="Document Analysis Tool" />
      <div className={styles.content}>
        <ParagraphManager 
          processParagraphs={processParagraphs} 
          updateParagraphById={updateParagraphById}
          addCommentToParagraph={addCommentToParagraph}
          highlightParagraph={highlightParagraph}
          scrollToParagraph={scrollToParagraph}
        />
      </div>
    </div>
  );
};

export default App;
