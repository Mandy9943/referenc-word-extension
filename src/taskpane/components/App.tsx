import { Button, makeStyles, Text, tokens } from "@fluentui/react-components";
import * as React from "react";
import { analyzeDocument, removeReferences } from "../word";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    backgroundColor: tokens.colorNeutralBackground1,
    padding: "20px",
  },
  buttonContainer: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    gap: "20px",
    maxWidth: "600px",
    margin: "0 auto",
    padding: "24px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: "8px",
    boxShadow: tokens.shadow4,
  },
  title: {
    fontSize: "24px",
    fontWeight: "600",
    marginBottom: "16px",
    color: tokens.colorNeutralForeground1,
    textAlign: "center",
  },
  button: {
    width: "100%",
    maxWidth: "300px",
    height: "40px",
  },
  status: {
    padding: "12px",
    borderRadius: "4px",
    width: "100%",
    maxWidth: "300px",
    textAlign: "center",
    transition: "all 0.3s ease",
  },
  statusLoading: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
  },
  statusSuccess: {
    backgroundColor: tokens.colorPaletteGreenBackground1,
    color: tokens.colorPaletteGreenForeground1,
  },
  statusError: {
    backgroundColor: tokens.colorPaletteRedBackground1,
    color: tokens.colorPaletteRedForeground1,
  },
  statusIdle: {
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground3,
  },
});

type Status = "idle" | "loading" | "success" | "error";

const App: React.FC = () => {
  const styles = useStyles();
  const [status, setStatus] = React.useState<Status>("idle");

  const handleAnalyzeDocument = async () => {
    setStatus("loading");
    try {
      await analyzeDocument();
      setStatus("success");
    } catch (error) {
      setStatus("error");
    }
  };

  const handleRemoveReferences = async () => {
    setStatus("loading");
    try {
      await removeReferences();
      setStatus("success");
    } catch (error) {
      setStatus("error");
    }
  };

  const getStatusDisplay = () => {
    const baseClassName = `${styles.status} `;
    switch (status) {
      case "loading":
        return (
          <div className={baseClassName + styles.statusLoading}>
            <Text>Processing...</Text>
          </div>
        );
      case "success":
        return (
          <div className={baseClassName + styles.statusSuccess}>
            <Text>Operation completed successfully!</Text>
          </div>
        );
      case "error":
        return (
          <div className={baseClassName + styles.statusError}>
            <Text>An error occurred. Please try again.</Text>
          </div>
        );
      default:
        return (
          <div className={baseClassName + styles.statusIdle}>
            <Text>Ready to process</Text>
          </div>
        );
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.buttonContainer}>
        <Text className={styles.title}>Reference Manager</Text>
        <Button
          appearance="secondary"
          onClick={handleRemoveReferences}
          disabled={status === "loading"}
          className={styles.button}
        >
          Remove References
        </Button>
        <Button
          appearance="primary"
          onClick={handleAnalyzeDocument}
          disabled={status === "loading"}
          className={styles.button}
        >
          Add References
        </Button>
        gemini v1
        {getStatusDisplay()}
      </div>
    </div>
  );
};

export default App;
