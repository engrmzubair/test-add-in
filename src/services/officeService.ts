/**
 * Service for interacting with Office.js Mailbox APIs
 * Handles email content extraction, .eml conversion, and email operations
 */

/* global Office */

export interface EmailData {
  subject: string;
  body: string;
  from: Office.EmailAddressDetails;
  to: Office.EmailAddressDetails[];
  cc?: Office.EmailAddressDetails[];
  bcc?: Office.EmailAddressDetails[];
  attachments: Office.AttachmentDetails[];
  internetMessageId?: string;
  dateTimeCreated?: Date;
}

export class OfficeService {
  /**
   * Get the current email item data
   */
  static async getCurrentEmailData(): Promise<EmailData> {
    return new Promise((resolve, reject) => {
      Office.onReady(() => {
        const item = Office.context.mailbox.item;
        
        if (!item) {
          reject(new Error("No email item available"));
          return;
        }

        // Get email body
        item.body.getAsync(Office.CoercionType.Html, (bodyResult) => {
          if (bodyResult.status === Office.AsyncResultStatus.Failed) {
            reject(new Error(`Failed to get email body: ${bodyResult.error.message}`));
            return;
          }

          // Helper function to convert Recipients to EmailAddressDetails[]
          const getEmailAddresses = (recipients: any): Office.EmailAddressDetails[] => {
            if (!recipients) return [];
            if (Array.isArray(recipients)) return recipients;
            // For Recipients object, we'll need to get the addresses asynchronously
            // For now, return empty array and handle this in a separate method
            return [];
          };

          const emailData: EmailData = {
            subject: item.subject || "",
            body: bodyResult.value || "",
            from: item.from,
            to: getEmailAddresses(item.to),
            cc: getEmailAddresses(item.cc),
            bcc: getEmailAddresses(item.bcc),
            attachments: item.attachments || [],
            internetMessageId: item.internetMessageId,
            dateTimeCreated: item.dateTimeCreated
          };

          resolve(emailData);
        });
      });
    });
  }

  /**
   * Get a unique identifier for the current email
   */
  static getEmailId(): string {
    const item = Office.context.mailbox.item;
    
    if (!item) {
      throw new Error("No email item available");
    }

    // Use internetMessageId if available, otherwise use itemId
    return item.internetMessageId || item.itemId || "";
  }

  /**
   * Convert email data to .eml format
   */
  static convertToEml(emailData: EmailData): string {
    const formatEmailAddress = (addr: Office.EmailAddressDetails) => 
      `"${addr.displayName}" <${addr.emailAddress}>`;

    const formatEmailAddresses = (addresses: Office.EmailAddressDetails[]) =>
      addresses.map(formatEmailAddress).join(", ");

    let emlContent = "";

    // Email headers
    emlContent += `Subject: ${emailData.subject}\r\n`;
    emlContent += `From: ${formatEmailAddress(emailData.from)}\r\n`;
    
    if (emailData.to && emailData.to.length > 0) {
      emlContent += `To: ${formatEmailAddresses(emailData.to)}\r\n`;
    }
    
    if (emailData.cc && emailData.cc.length > 0) {
      emlContent += `Cc: ${formatEmailAddresses(emailData.cc)}\r\n`;
    }
    
    if (emailData.bcc && emailData.bcc.length > 0) {
      emlContent += `Bcc: ${formatEmailAddresses(emailData.bcc)}\r\n`;
    }

    if (emailData.dateTimeCreated) {
      emlContent += `Date: ${emailData.dateTimeCreated.toUTCString()}\r\n`;
    }

    if (emailData.internetMessageId) {
      emlContent += `Message-ID: ${emailData.internetMessageId}\r\n`;
    }

    emlContent += `MIME-Version: 1.0\r\n`;
    emlContent += `Content-Type: text/html; charset=utf-8\r\n`;
    emlContent += `Content-Transfer-Encoding: quoted-printable\r\n`;
    emlContent += `\r\n`;

    // Email body
    emlContent += emailData.body;

    return emlContent;
  }

  /**
   * Insert text at the current cursor position (for compose mode)
   */
  static async insertText(text: string): Promise<void> {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item?.body.setSelectedDataAsync(
        text,
        { coercionType: Office.CoercionType.Text },
        (asyncResult: Office.AsyncResult<void>) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            reject(new Error(asyncResult.error.message));
          } else {
            resolve();
          }
        }
      );
    });
  }

  /**
   * Check if the current context is compose mode
   */
  static isComposeMode(): boolean {
    return Office.context.mailbox.item?.itemType === Office.MailboxEnums.ItemType.Message &&
           Office.context.mailbox.item?.itemClass === "IPM.Note";
  }

  /**
   * Check if the current context is read mode
   */
  static isReadMode(): boolean {
    return Office.context.mailbox.item?.itemType === Office.MailboxEnums.ItemType.Message &&
           Office.context.mailbox.item?.itemClass !== "IPM.Note";
  }

  /**
   * Get current user's email address
   */
  static getCurrentUserEmail(): string {
    return Office.context.mailbox.userProfile.emailAddress;
  }

  /**
   * Get current user's display name
   */
  static getCurrentUserDisplayName(): string {
    return Office.context.mailbox.userProfile.displayName;
  }
} 