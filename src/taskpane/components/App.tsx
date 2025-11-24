/* global Office setInterval clearInterval */
import { Button, Checkbox, makeStyles, Text, tokens } from "@fluentui/react-components";
import { Timer24Regular } from "@fluentui/react-icons";
import * as React from "react";
import {
  analyzeDocument,
  paraphraseSelectedText,
  removeLinks,
  removeReferences,
  removeWeirdNumbers,
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
  const [insertEveryOther, setInsertEveryOther] = React.useState(false);
  const [paraphraseTime, setParaphraseTime] = React.useState<number | null>(null);
  const timerRef = React.useRef<any>(null);

  React.useEffect(() => {
    Office.onReady((info) => {
      setIsValidHost(info.host === Office.HostType.Word || info.host === Office.HostType.PowerPoint);
    });
  }, []);

  const handleAnalyzeDocument = async () => {
    setStatus("loading");
    try {
      // @ts-ignore
      await analyzeDocument(insertEveryOther);
      setStatus("success");
    } catch (error) {
      setStatus("error");
    }
  };

  const handleClean = async () => {
    setStatus("loading");
    try {
      await removeReferences();
      await removeLinks(false);
      await removeWeirdNumbers();
      setStatus("success");
    } catch (error) {
      setStatus("error");
    }
  };

  const handleParaphraseText = async () => {
    setStatus("loading");
    setParaphraseTime(0);
    const startTime = Date.now();

    if (timerRef.current) clearInterval(timerRef.current);

    timerRef.current = setInterval(() => {
      setParaphraseTime((Date.now() - startTime) / 1000);
    }, 100);

    try {
      await paraphraseSelectedText();
      setStatus("success");
    } catch (error) {
      setStatus("error");
    } finally {
      if (timerRef.current) {
        clearInterval(timerRef.current);
        timerRef.current = null;
      }
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
              onClick={handleClean}
              disabled={status === "loading"}
              className={styles.button}
              style={{ backgroundColor: "#cf6760ff" }}
            >
              Clean
            </Button>
            <Button
              appearance="secondary"
              onClick={handleParaphraseText}
              disabled={status === "loading"}
              className={styles.button}
              style={{ backgroundColor: "#7d7fd6ff" }}
            >
              Paraphrase
            </Button>

            <div
              style={{ display: "flex", alignItems: "center", marginBottom: "10px", width: "100%", maxWidth: "300px" }}
            >
              <Checkbox
                checked={insertEveryOther}
                onChange={(_e, data) => setInsertEveryOther(data.checked === true)}
                label="Insert references every other paragraph"
              />
            </div>
            <Button
              onClick={handleAnalyzeDocument}
              disabled={status === "loading"}
              className={`${styles.button}`}
              style={{ backgroundColor: "rgb(26 167 26)", color: "#fff" }}
            >
              Add References
            </Button>
          </>
        ) : (
          <Text>
            This add-in is optimized for Word and PowerPoint. Some features may not be available in other applications.
          </Text>
        )}
        v2.9
        {getStatusDisplay()}
        {paraphraseTime !== null && (
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              gap: "8px",
              marginBottom: "10px",
            }}
          >
            <Timer24Regular />
            <Text>{paraphraseTime.toFixed(1)}s</Text>
          </div>
        )}
      </div>
    </div>
  );
};

export default App;
