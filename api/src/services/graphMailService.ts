import { getGraphAccessToken } from './graphAuthService';
import { fetchWithRetry } from '../utils/graphFetch';

const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

/**
 * GRAPH PERMISSION NOTE
 * ----------------------
 * Sending mail via Microsoft Graph (POST /users/{upn}/sendMail) requires
 * the APPLICATION (not delegated) permission `Mail.Send`, granted with
 * ADMIN CONSENT on the same Entra ID app registration already used for
 * `Reports.Read.All` / `Directory.Read.All` / `User.Read.All` /
 * `Sites.Read.All` (see graphAuthService.ts and api/README.md).
 *
 * `Mail.Send` is a SENSITIVE permission: with application permissions it
 * allows the app to send mail AS ANY USER IN THE TENANT, not just one
 * mailbox. This backend mitigates that by only ever calling sendMail for
 * ONE specific, configured mailbox - the value of the `REMINDER_SENDER_UPN`
 * application setting (intended to be `SPAdmin01@stewart.com`, the
 * SharePoint administration service account) - never a mailbox chosen by
 * the caller/request body. There is no code path in this service that
 * accepts an arbitrary "from" address.
 */

export interface ISendMailRequest {
  toEmail: string;
  subject: string;
  htmlBody: string;
}

function getSenderUpn(): string {
  const upn = process.env.REMINDER_SENDER_UPN;
  if (!upn) {
    throw new Error(
      'Missing required application setting "REMINDER_SENDER_UPN". Configure it in ' +
      'api/local.settings.json for local development, or in Azure Function App ' +
      'Settings for production (expected value: SPAdmin01@stewart.com).'
    );
  }
  return upn;
}

/**
 * Sends an HTML email via Microsoft Graph from the configured
 * REMINDER_SENDER_UPN mailbox. Callers must never log the access token
 * or the full raw Graph response - only high-level status/outcome.
 */
export async function sendMail(request: ISendMailRequest): Promise<void> {
  const senderUpn = getSenderUpn();
  const token = await getGraphAccessToken();

  const url = `${GRAPH_BASE_URL}/users/${encodeURIComponent(senderUpn)}/sendMail`;

  const payload = {
    message: {
      subject: request.subject,
      body: {
        contentType: 'HTML',
        content: request.htmlBody
      },
      toRecipients: [
        { emailAddress: { address: request.toEmail } }
      ]
    },
    saveToSentItems: true
  };

  const response = await fetchWithRetry(
    url,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    },
    `sendMail from ${senderUpn}`
  );

  // Microsoft Graph sendMail returns 202 Accepted on success, with no
  // response body.
  if (response.status !== 202) {
    const errorBody = await response.text().catch(() => '');
    throw new Error(
      `Microsoft Graph sendMail failed with status ${response.status} ${response.statusText}. ${errorBody}`.trim()
    );
  }
}
