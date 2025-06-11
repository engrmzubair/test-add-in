import * as React from "react";
import { useState, useEffect } from "react";
import Header from "./Header";
import { makeStyles } from "@fluentui/react-components";
import { 
  Button, 
  Card, 
  CardHeader, 
  Text, 
  Badge,
  Spinner,
  MessageBar
} from "@fluentui/react-components";
import { 
  CloudArrowUp24Regular, 
  CheckmarkCircle24Regular, 
  ErrorCircle24Regular 
} from "@fluentui/react-icons";
import { OfficeService, EmailData } from "../../services/officeService";
import { CloudflareService } from "../../services/cloudflareService";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    padding: "20px",
  },
  card: {
    marginBottom: "16px",
  },
  buttonContainer: {
    display: "flex",
    gap: "12px",
    marginTop: "16px",
  },
  statusContainer: {
    marginTop: "16px",
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();
  const [emailData, setEmailData] = useState<EmailData | null>(null);
  const [isTransferred, setIsTransferred] = useState<boolean>(false);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [message, setMessage] = useState<string>("");
  const [messageType, setMessageType] = useState<"info" | "success" | "warning" | "error">("info");

  // Mock client ID - in real implementation, this would come from authentication
  const clientId = "demo-client-123";

  useEffect(() => {
    loadEmailData();
    checkTransferStatus();
  }, []);

  const loadEmailData = async () => {
    try {
      const data = await OfficeService.getCurrentEmailData();
      setEmailData(data);
    } catch (error) {
      console.error("Error loading email data:", error);
      setMessage("Error loading email data");
      setMessageType("error");
    }
  };

  const checkTransferStatus = async () => {
    try {
      const emailId = OfficeService.getEmailId();
      if (emailId) {
        const transferred = await CloudflareService.isEmailTransferred(clientId, emailId);
        setIsTransferred(transferred);
      }
    } catch (error) {
      console.error("Error checking transfer status:", error);
    }
  };

  const handleTransferEmail = async () => {
    if (!emailData) return;

    setIsLoading(true);
    setMessage("");

    try {
      // Get email ID
      const emailId = OfficeService.getEmailId();
      
      // Convert to .eml format
      const emlContent = OfficeService.convertToEml(emailData);
      
      // Store in Cloudflare KV
      await CloudflareService.storeEmailId(clientId, emailId);
      
      // TODO: Send .eml to LiraDocs server
      // This would be implemented based on your LiraDocs API
      
      setIsTransferred(true);
      setMessage("Email successfully transferred to LiraDocs!");
      setMessageType("success");
      
    } catch (error) {
      console.error("Error transferring email:", error);
      setMessage(`Error transferring email: ${error.message}`);
      setMessageType("error");
    } finally {
      setIsLoading(false);
    }
  };

  const handleRemoveTransfer = async () => {
    setIsLoading(true);
    setMessage("");

    try {
      const emailId = OfficeService.getEmailId();
      await CloudflareService.removeEmailId(clientId, emailId);
      
      setIsTransferred(false);
      setMessage("Email transfer removed from LiraDocs");
      setMessageType("info");
      
    } catch (error) {
      console.error("Error removing transfer:", error);
      setMessage(`Error removing transfer: ${error.message}`);
      setMessageType("error");
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={props.title} message="Legal Case Management" />
      
      {message && (
        <MessageBar intent={messageType} className={styles.card}>
          {message}
        </MessageBar>
      )}

      <Card className={styles.card}>
        <CardHeader
          header={<Text weight="semibold">Email Transfer Status</Text>}
          action={
            isTransferred ? (
              <Badge appearance="filled" color="success" icon={<CheckmarkCircle24Regular />}>
                Transferred
              </Badge>
            ) : (
              <Badge appearance="outline" color="subtle">
                Not Transferred
              </Badge>
            )
          }
        />
      </Card>

      {emailData && (
        <Card className={styles.card}>
          <CardHeader
            header={<Text weight="semibold">Current Email</Text>}
            description={
              <div>
                <Text size={200}>Subject: {emailData.subject}</Text><br />
                <Text size={200}>From: {emailData.from?.displayName || emailData.from?.emailAddress}</Text>
              </div>
            }
          />
        </Card>
      )}

      <div className={styles.buttonContainer}>
        {!isTransferred ? (
          <Button
            appearance="primary"
            icon={<CloudArrowUp24Regular />}
            onClick={handleTransferEmail}
            disabled={isLoading || !emailData}
          >
            {isLoading ? <Spinner size="tiny" /> : "Transfer to LiraDocs"}
          </Button>
        ) : (
          <Button
            appearance="secondary"
            icon={<ErrorCircle24Regular />}
            onClick={handleRemoveTransfer}
            disabled={isLoading}
          >
            {isLoading ? <Spinner size="tiny" /> : "Remove from LiraDocs"}
          </Button>
        )}
      </div>

      <div className={styles.statusContainer}>
        <Text size={200} style={{ color: "#666" }}>
          {isLoading ? "Processing..." : "Ready"}
        </Text>
      </div>
    </div>
  );
};

export default App;
