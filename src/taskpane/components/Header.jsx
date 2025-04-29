import * as React from "react";
import PropTypes from "prop-types";
import { Text, makeStyles, tokens } from "@fluentui/react-components";

const useStyles = makeStyles({
  header: {
    display: "flex",
    flexDirection: "row",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "6px 10px",
    background: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundInverted,
    boxShadow: "0 1px 2px rgba(0, 0, 0, 0.1)",
    minHeight: "auto",
  },
  titleContainer: {
    display: "flex",
    flexDirection: "row",
    alignItems: "center",
    gap: "6px",
  },
  title: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    lineHeight: tokens.lineHeightBase100,
  },
  subtitle: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightRegular,
    opacity: 0.9,
  }
});

const Header = (props) => {
  const { message } = props;
  const styles = useStyles();

  return (
    <header className={styles.header}>
      <div className={styles.titleContainer}>
        <Text className={styles.title}>Pramata</Text>
        <Text className={styles.subtitle}>{message}</Text>
      </div>
    </header>
  );
};

Header.propTypes = {
  message: PropTypes.string,
};

export default Header;
