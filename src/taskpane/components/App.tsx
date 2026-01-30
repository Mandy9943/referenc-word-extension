/* global Office setInterval clearInterval */
import { Button, Checkbox, makeStyles, Text, tokens } from "@fluentui/react-components";
import { Timer24Regular } from "@fluentui/react-icons";
import * as React from "react";
import {
  analyzeDocument,
  ChangeMetrics,
  normalizeBoldText,
  ParaphraseResult,
  paraphraseSelectedText,
  paraphraseSelectedTextStandard,
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
  statusWarning: {
    backgroundColor: tokens.colorPaletteYellowBackground1,
    color: tokens.colorPaletteYellowForeground1,
    marginTop: "8px",
    fontSize: "12px",
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
  metricsContainer: {
    width: "100%",
    maxWidth: "300px",
    padding: "12px",
    borderRadius: "4px",
    backgroundColor: tokens.colorNeutralBackground3,
    marginTop: "8px",
  },
  metricsRow: {
    display: "flex",
    justifyContent: "space-between",
    marginBottom: "4px",
  },
  previewText: {
    fontSize: "11px",
    color: tokens.colorNeutralForeground3,
    fontStyle: "italic",
    marginTop: "4px",
    wordBreak: "break-word" as const,
  },
  changeGood: {
    color: tokens.colorPaletteGreenForeground1,
    fontWeight: "600",
  },
  changeWarning: {
    color: tokens.colorPaletteYellowForeground1,
    fontWeight: "600",
  },
  changeDanger: {
    color: tokens.colorPaletteRedForeground1,
    fontWeight: "600",
  },
});

type Status = "idle" | "loading" | "success" | "error";

interface FailedAccount {
  accountId: string;
}

const App: React.FC = () => {
  const styles = useStyles();
  const [status, setStatus] = React.useState<Status>("idle");
  const [isValidHost, setIsValidHost] = React.useState(false);
  const [insertEveryOther, setInsertEveryOther] = React.useState(false);
  const [paraphraseTime, setParaphraseTime] = React.useState<number | null>(null);
  const [failedAccount, setFailedAccount] = React.useState<FailedAccount | null>(null);
  const [warnings, setWarnings] = React.useState<string[]>([]);
  const [errorMessage, setErrorMessage] = React.useState<string | null>(null);
  const [changeMetrics, setChangeMetrics] = React.useState<ChangeMetrics | null>(null);
  const timerRef = React.useRef<any>(null);

  React.useEffect(() => {
    Office.onReady((info) => {
      setIsValidHost(info.host === Office.HostType.Word || info.host === Office.HostType.PowerPoint);
    });
  }, []);

  const handleAnalyzeDocument = async () => {
    setStatus("loading");
    setWarnings([]);
    setErrorMessage(null);
    try {
      // @ts-ignore
      await analyzeDocument(insertEveryOther);
      setStatus("success");
    } catch (error) {
      setStatus("error");
      setErrorMessage(error.message || "An error occurred");
    }
  };

  const handleClean = async () => {
    setStatus("loading");
    setWarnings([]);
    setErrorMessage(null);
    try {
      await removeReferences();
      await removeLinks(false);
      await removeWeirdNumbers();
      await normalizeBoldText();
      setStatus("success");
    } catch (error) {
      setStatus("error");
      setErrorMessage(error.message || "An error occurred");
    }
  };

  const handleParaphraseText = async () => {
    setStatus("loading");
    setParaphraseTime(0);
    setFailedAccount(null);
    setWarnings([]);
    setErrorMessage(null);
    setChangeMetrics(null);
    const startTime = Date.now();

    if (timerRef.current) clearInterval(timerRef.current);

    timerRef.current = setInterval(() => {
      setParaphraseTime((Date.now() - startTime) / 1000);
    }, 100);

    try {
      const result: ParaphraseResult = await paraphraseSelectedText();
      setStatus("success");
      setFailedAccount(null);
      if (result.warnings && result.warnings.length > 0) {
        setWarnings(result.warnings);
      }
      if (result.metrics) {
        setChangeMetrics(result.metrics);
      }
    } catch (error) {
      setStatus("error");
      setErrorMessage(error.message || "An error occurred");
      // Extract failed account info from error message if available
      if (error.message) {
        const match = error.message.match(/Account (acc[123]) failed/);
        if (match) {
          setFailedAccount({ accountId: match[1] });
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
    setFailedAccount(null);
    setWarnings([]);
    setErrorMessage(null);
    setChangeMetrics(null);
    const startTime = Date.now();

    if (timerRef.current) clearInterval(timerRef.current);

    timerRef.current = setInterval(() => {
      setParaphraseTime((Date.now() - startTime) / 1000);
    }, 100);

    try {
      const result: ParaphraseResult = await paraphraseSelectedTextStandard();
      setStatus("success");
      setFailedAccount(null);
      if (result.warnings && result.warnings.length > 0) {
        setWarnings(result.warnings);
      }
      if (result.metrics) {
        setChangeMetrics(result.metrics);
      }
    } catch (error) {
      setStatus("error");
      setErrorMessage(error.message || "An error occurred");
      // Extract failed account info from error message if available
      if (error.message) {
        const match = error.message.match(/Account (acc[123]) failed/);
        if (match) {
          setFailedAccount({ accountId: match[1] });
        }
      }
    } finally {
      if (timerRef.current) {
        clearInterval(timerRef.current);
        timerRef.current = null;
      }
    }
  };

  const handleRestartAccount = async () => {
    if (!failedAccount) return;

    try {
      setStatus("loading");
      const restartUrl = `https://analizeai.com/restart/${failedAccount.accountId}`;
      console.log(`Restarting account: ${restartUrl}`);

      const response = await fetch(restartUrl, {
        method: "POST",
      });

      if (response.ok) {
        setStatus("success");
        setFailedAccount(null);
        setErrorMessage(null);
      } else {
        setStatus("error");
        setErrorMessage(`Failed to restart account: ${response.status}`);
      }
    } catch (error) {
      console.error("Error restarting account:", error);
      setStatus("error");
      setErrorMessage(error.message || "Failed to restart account");
    }
  };

  // Helper to get change level styling
  const getChangeClass = (percent: number) => {
    if (percent >= 60) return styles.changeGood;
    if (percent >= 40) return styles.changeWarning;
    return styles.changeDanger;
  };

  const getChangeWarning = (percent: number) => {
    const reusedPercent = 100 - percent;
    if (percent < 40) return `⚠️ Reusing ~${reusedPercent}% of original words`;
    if (percent < 60) return `⚡ Reusing ~${reusedPercent}% of original words`;
    return `✅ Only ~${reusedPercent}% of original words reused`;
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
          <>
            <div className={baseClassName + styles.statusSuccess}>
              <Text>Operation completed successfully!</Text>
            </div>
            {changeMetrics && (
              <div className={styles.metricsContainer}>
                <Text weight="semibold" size={300}>
                  📊 Change Metrics
                </Text>
                <div className={styles.metricsRow} style={{ marginTop: "8px" }}>
                  <Text size={200}>Words Before:</Text>
                  <Text size={200} weight="semibold">
                    {changeMetrics.originalWordCount}
                  </Text>
                </div>
                <div className={styles.metricsRow}>
                  <Text size={200}>Words After:</Text>
                  <Text size={200} weight="semibold">
                    {changeMetrics.newWordCount}
                  </Text>
                </div>
                <div className={styles.metricsRow}>
                  <Text size={200}>Words Changed:</Text>
                  <Text size={200} weight="semibold" className={getChangeClass(changeMetrics.wordChangePercent)}>
                    {changeMetrics.wordsChanged} ({changeMetrics.wordChangePercent}%)
                  </Text>
                </div>
                <div style={{ marginTop: "8px", textAlign: "center" }}>
                  <Text size={200} className={getChangeClass(changeMetrics.wordChangePercent)}>
                    {getChangeWarning(changeMetrics.wordChangePercent)}
                  </Text>
                </div>
                <div style={{ marginTop: "8px" }}>
                  <Text size={200} weight="semibold">
                    Before:
                  </Text>
                  <div className={styles.previewText}>"{changeMetrics.originalPreview}"</div>
                </div>
                <div style={{ marginTop: "4px" }}>
                  <Text size={200} weight="semibold">
                    After:
                  </Text>
                  <div className={styles.previewText}>"{changeMetrics.newPreview}"</div>
                </div>
              </div>
            )}
            {warnings.length > 0 && (
              <div className={baseClassName + styles.statusWarning}>
                <Text weight="semibold">Warnings:</Text>
                {warnings.map((warning, idx) => (
                  <div key={idx} style={{ marginTop: "4px" }}>
                    <Text size={200}>• {warning}</Text>
                  </div>
                ))}
              </div>
            )}
          </>
        );
      case "error":
        return (
          <div className={baseClassName + styles.statusError}>
            <Text>{errorMessage || "An error occurred. Please try again."}</Text>
            {failedAccount && (
              <Button appearance="primary" onClick={handleRestartAccount} style={{ marginTop: "10px" }}>
                Restart Account {failedAccount.accountId}
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
        v3.1
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
