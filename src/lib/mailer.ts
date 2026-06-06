// serverside/src/lib/mailer.ts
import nodemailer from "nodemailer";
import config from "../config/config";

const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: { user: config.emailUser, pass: config.emailPass },
});

export async function sendEmail({
  to,
  subject,
  html,
  attachments,
}: {
  to: string;
  subject: string;
  html: string;
  attachments?: Array<{
    filename: string;
    path?: string;
    content?: string | Buffer;
    cid?: string;
  }>;
}) {
  await transporter.sendMail({
    from: `"${config.appName}" <${config.emailUser}>`,
    to,
    subject,
    html,
    attachments,
  });
}
