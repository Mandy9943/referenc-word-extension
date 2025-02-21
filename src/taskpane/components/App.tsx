import { Button, makeStyles, Text, tokens } from "@fluentui/react-components";
import * as React from "react";
import {
  analyzeDocument,
  humanizeDocument,
  humanizeSelectedText,
  removeReferences,
  stopHumanizeProcess,
} from "../taskpane";

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
  buttonGreen: {
    backgroundColor: "#008000",
    color: tokens.colorNeutralForegroundOnBrand,
    "&:hover": {
      backgroundColor: tokens.colorPaletteGreenBackground1,
    },
  },
  buttonBlue: {
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
    "&:hover": {
      backgroundColor: tokens.colorBrandBackgroundHover,
    },
  },
  buttonYellow: {
    backgroundColor: tokens.colorPaletteYellowBackground2,
    color: tokens.colorNeutralForeground1,
    "&:hover": {
      backgroundColor: tokens.colorPaletteYellowBackground1,
    },
  },
  buttonRed: {
    backgroundColor: tokens.colorPaletteRedBackground2,
    color: tokens.colorNeutralForegroundOnBrand,
    "&:hover": {
      backgroundColor: tokens.colorPaletteRedBackground1,
    },
  },
});

type Status = "idle" | "loading" | "success" | "error";

const App: React.FC = () => {
  const styles = useStyles();
  const [status, setStatus] = React.useState<Status>("idle");
  const [isValidHost, setIsValidHost] = React.useState(false);
  const [isHumanizing, setIsHumanizing] = React.useState(false);

  React.useEffect(() => {
    Office.onReady((info) => {
      setIsValidHost(info.host === Office.HostType.Word || info.host === Office.HostType.PowerPoint);
    });
  }, []);

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

  const handleHumanizeDocument = async () => {
    setStatus("loading");
    setIsHumanizing(true);
    try {
      await humanizeDocument();
      setStatus("success");
    } catch (error) {
      setStatus("error");
    } finally {
      setIsHumanizing(false);
    }
  };

  const handleHumanizeSelectedText = async () => {
    setStatus("loading");
    setIsHumanizing(true);
    try {
      await humanizeSelectedText();
      setStatus("success");
    } catch (error) {
      setStatus("error");
    } finally {
      setIsHumanizing(false);
    }
  };

  const handleStopHumanize = () => {
    stopHumanizeProcess();
    setIsHumanizing(false);
    setStatus("idle");
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
          <div>
            <Text>Ready to process</Text>
          </div>
        );
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.buttonContainer}>
        <Text className={styles.title}>Essay Manager</Text>
        {isValidHost ? (
          <>
            <Button
              appearance="secondary"
              onClick={handleRemoveReferences}
              disabled={status === "loading"}
              className={styles.button}
            >
              Remove References
            </Button>
            <Button
              onClick={handleAnalyzeDocument}
              disabled={status === "loading"}
              className={`${styles.button}`}
              style={{ backgroundColor: "rgb(26 167 26)", color: "#fff" }}
            >
              Add References
            </Button>
            <Button
              appearance="primary"
              onClick={handleHumanizeDocument}
              disabled={status === "loading"}
              className={`${styles.button} ${styles.buttonBlue}`}
            >
              Humanize All Text
            </Button>
            <Button
              appearance="primary"
              onClick={handleHumanizeSelectedText}
              disabled={status === "loading"}
              className={`${styles.button}`}
              style={{ backgroundColor: "rgb(155 163 7)", color: "#fff" }}
            >
              Humanize Selected Text
            </Button>
            <Button
              appearance="primary"
              onClick={handleStopHumanize}
              disabled={!isHumanizing}
              className={`${styles.button}`}
              style={{ backgroundColor: "rgb(255 10 10)", color: "#fff" }}
            >
              Stop Humanize Process
            </Button>
          </>
        ) : (
          <Text>
            This add-in is optimized for Word and PowerPoint. Some features may not be available in other applications.
          </Text>
        )}
        v1.5
        {getStatusDisplay()}
      </div>
    </div>
  );
};

export default App;
