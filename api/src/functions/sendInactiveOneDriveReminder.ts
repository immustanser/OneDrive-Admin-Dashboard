import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { sendMail } from '../services/graphMailService';

/**
 * POST /api/send-inactive-onedrive-reminder
 *
 * Sends a governance "inactivity reminder" email to a SINGLE inactive
 * OneDrive owner, using data already present on the Inactive OneDrives
 * tab's row (no additional Microsoft Graph reads are performed here -
 * see PERFORMANCE note below).
 *
 * AUTHENTICATION: like every other endpoint in this Function App, this
 * route is registered with `authLevel: 'anonymous'` at the Azure
 * Functions runtime level, but the Function App itself has Microsoft
 * Entra ID Authentication (App Service "Easy Auth") enabled with
 * anonymous access DISABLED at the platform level - unauthenticated
 * requests never reach this handler. The SPFx client calls this route
 * through AadHttpClient (see GraphService.ts), exactly like the
 * existing /api/onedrive-dashboard and /api/user-profile routes. This
 * endpoint introduces no new authentication mechanism and does not
 * change that setup.
 *
 * PERFORMANCE: this handler makes exactly ONE Microsoft Graph call
 * (POST /users/{REMINDER_SENDER_UPN}/sendMail) per request - it never
 * re-fetches the dashboard, user profile, or usage reports. It relies
 * entirely on the request body supplied by the already-loaded,
 * already-cached Inactive OneDrives grid data.
 */

interface ISendReminderRequestBody {
  userDisplayName?: string;
  userEmail?: string;
  managerName?: string;
  department?: string;
  storageUsed?: string;
  storageUsedGB?: number;
  daysInactive?: number;
  lastActivityDate?: string;
  oneDriveUrl?: string;
}

const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

const GENERIC_FAILURE_MESSAGE =
  'Unable to send reminder email. Please try again or contact the SharePoint Administration team.';

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function buildReminderEmailHtml(body: ISendReminderRequestBody): string {
  const displayName = escapeHtml(body.userDisplayName || '');
  const lastActivityDate = escapeHtml(body.lastActivityDate || 'N/A');
  const daysInactive = escapeHtml(String(body.daysInactive));
  const storageUsed = escapeHtml(body.storageUsed || 'N/A');
  const department = escapeHtml(body.department || 'N/A');
  const managerName = escapeHtml(body.managerName || 'N/A');

  return `
    <p>Hello ${displayName},</p>
    <p>Our records indicate that your OneDrive has not had recent activity.</p>
    <p><strong>OneDrive Details:</strong></p>
    <ul>
      <li>Last Activity Date: ${lastActivityDate}</li>
      <li>Days Inactive: ${daysInactive}</li>
      <li>Storage Used: ${storageUsed}</li>
      <li>Department: ${department}</li>
      <li>Manager: ${managerName}</li>
    </ul>
    <p>If this OneDrive is still required for business purposes, no action is needed.</p>
    <p>If the data is no longer needed, please review and remove unnecessary content where appropriate.</p>
    <p>If you believe this message was received in error, please contact the SharePoint Administration team.</p>
    <p>Thank you,<br/>SharePoint Administration Team</p>
  `.trim();
}

export async function sendInactiveOneDriveReminder(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  let body: ISendReminderRequestBody;
  try {
    body = (await request.json()) as ISendReminderRequestBody;
  } catch {
    return {
      status: 400,
      jsonBody: { success: false, message: 'Request body must be valid JSON.' }
    };
  }

  context.log('sendInactiveOneDriveReminder: reminder request received.');

  if (!body || !body.userDisplayName || !body.userEmail || body.daysInactive === undefined || body.daysInactive === null) {
    return {
      status: 400,
      jsonBody: {
        success: false,
        message: 'Missing required fields. "userDisplayName", "userEmail" and "daysInactive" are required.'
      }
    };
  }

  if (!EMAIL_REGEX.test(body.userEmail)) {
    return {
      status: 400,
      jsonBody: { success: false, message: 'The provided userEmail is not a valid email address.' }
    };
  }

  context.log(`sendInactiveOneDriveReminder: recipient=${body.userEmail}`);

  try {
    context.log('sendInactiveOneDriveReminder: reminder email send started.');

    await sendMail({
      toEmail: body.userEmail,
      subject: 'OneDrive Activity Reminder',
      htmlBody: buildReminderEmailHtml(body)
    });

    context.log('sendInactiveOneDriveReminder: reminder email send succeeded.');

    return {
      status: 200,
      jsonBody: { success: true, message: 'Reminder email sent successfully.' }
    };
  } catch (error) {
    // Never log the error object's raw Graph response/token details -
    // only the resulting error message.
    context.error('sendInactiveOneDriveReminder: reminder email send failed.', error instanceof Error ? error.message : 'Unknown error');

    return {
      status: 502,
      jsonBody: { success: false, message: GENERIC_FAILURE_MESSAGE }
    };
  }
}

app.http('sendInactiveOneDriveReminder', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'send-inactive-onedrive-reminder',
  handler: sendInactiveOneDriveReminder
});
