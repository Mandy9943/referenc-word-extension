/* global Office setInterval clearInterval */
import { Button, Checkbox, makeStyles, Text, tokens } from "@fluentui/react-components";
import { Timer24Regular } from "@fluentui/react-icons";
import * as React from "react";
import {
  analyzeDocument,
  normalizeBoldText,
  paraphraseSelectedTextStandard as paraphraseDocumentStandard,
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

interface FailedService {
  url: string;
  instanceNumber: number;
}

const App: React.FC = () => {
  const styles = useStyles();
  const [status, setStatus] = React.useState<Status>("idle");
  const [isValidHost, setIsValidHost] = React.useState(false);
  const [insertEveryOther, setInsertEveryOther] = React.useState(false);
  const [paraphraseTime, setParaphraseTime] = React.useState<number | null>(null);
  const [failedService, setFailedService] = React.useState<FailedService | null>(null);
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
      await normalizeBoldText();
      setStatus("success");
    } catch (error) {
      setStatus("error");
    }
  };

  const handleParaphraseText = async () => {
    setStatus("loading");
    setParaphraseTime(0);
    setFailedService(null);
    const startTime = Date.now();

    if (timerRef.current) clearInterval(timerRef.current);

    timerRef.current = setInterval(() => {
      setParaphraseTime((Date.now() - startTime) / 1000);
    }, 100);

    try {
      await paraphraseSelectedText();
      setStatus("success");
      setFailedService(null);
    } catch (error) {
      setStatus("error");
      // Extract failed service info from error message if available
      if (error.message && error.message.includes("failed with status")) {
        const match = error.message.match(/API request to (https:\/\/[^\s]+) failed/);
        if (match) {
          const url = match[1];
          const instanceNumber = url.includes("v3") ? 3 : url.includes("v2") ? 2 : 1;
          setFailedService({ url, instanceNumber });
        }
      }
    } finally {
      if (timerRef.current) {
        clearInterval(timerRef.current);
        timerRef.current = null;
      }
    }
  };

  const handleParaphraseTextStandard = async () => {
    setStatus("loading");
    setParaphraseTime(0);
    setFailedService(null);
    const startTime = Date.now();

    if (timerRef.current) clearInterval(timerRef.current);

    timerRef.current = setInterval(() => {
      setParaphraseTime((Date.now() - startTime) / 1000);
    }, 100);

    try {
      await paraphraseDocumentStandard();
      setStatus("success");
      setFailedService(null);
    } catch (error) {
      setStatus("error");
      // Extract failed service info from error message if available
      if (error.message && error.message.includes("failed with status")) {
        const match = error.message.match(/API request to (https:\/\/[^\s]+) failed/);
        if (match) {
          const url = match[1];
          const instanceNumber = url.includes("v3") ? 3 : url.includes("v2") ? 2 : 1;
          setFailedService({ url, instanceNumber });
        }
      }
    } finally {
      if (timerRef.current) {
        clearInterval(timerRef.current);
        timerRef.current = null;
      }
    }
  };

  const handleRestartService = async () => {
    if (!failedService) return;

    try {
      setStatus("loading");
      const restartUrl = `${failedService.url}/restart`;
      console.log(`Restarting service: ${restartUrl}`);

      const response = await fetch(restartUrl, {
        method: "POST",
      });

      if (response.ok) {
        setStatus("success");
        setFailedService(null);
      } else {
        setStatus("error");
      }
    } catch (error) {
      console.error("Error restarting service:", error);
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
            {failedService && (
              <Button appearance="primary" onClick={handleRestartService} style={{ marginTop: "10px" }}>
                Restart Paraphrase {failedService.instanceNumber}
              </Button>
            )}
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
              SIMPLE + SHORT
            </Button>
            <Button
              appearance="secondary"
              onClick={handleParaphraseTextStandard}
              disabled={status === "loading"}
              className={styles.button}
              style={{ backgroundColor: "#5a9bd6ff" }}
            >
              STANDARD
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
        v2.14
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
