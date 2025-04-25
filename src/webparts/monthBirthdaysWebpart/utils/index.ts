import { FontWeights } from "@fluentui/react";
import { mergeStyleSets } from "@fluentui/react";
import { getTheme } from "@fluentui/react";
import { CSSProperties } from "react";

// Themed styles for the example.
const theme = getTheme();
export const mainStyles = mergeStyleSets({
  buttonArea: {
    float: "right",
    display: "inline-block",
  },
  callout: {
    maxWidth: 300,
  },
  header: {
    padding: "18px 24px 12px",
  },
  title: [
    theme.fonts.xLarge,
    {
      margin: 0,
      color: theme.palette.neutralPrimary,
      fontWeight: FontWeights.semilight,
    },
  ],
  inner: {
    height: "100%",
    padding: "0 24px 20px",
  },
  actions: {
    position: "relative",
    marginTop: 20,
    width: "100%",
    whiteSpace: "nowrap",
  },
  link: [
    theme.fonts.medium,
    {
      color: theme.palette.neutralPrimary,
    },
  ],
  subtext: {
    margin: 0,
    fontSize: "12px",
    color: "#323130",
    fontWeight: 300,
  },
});

export const webpartHeaderStyles = {
  fontSize: "24px",
  fontWeight: "100",
  color: "rgb(51, 51, 51)",
  height: "36px",
};

export const webpartHeaderEditStyles: CSSProperties = {
  backgroundColor: "transparent",
  border: "none",
  boxSizing: "border-box",
  display: "block",
  margin: 0,
  outline: 0,
  overflow: "hidden",
  resize: "none",
  whiteSpace: "pre",
  width: "100%",
  fontFamily: "inherit",
  fontSize: "inherit",
  fontWeight: "inherit",
  lineHeight: "inherit",
  textAlign: "inherit",
  color: "inherit",
  height: "inherit",
};

export const webpartHeaderViewStyles = {
  color: "inherit",
};
