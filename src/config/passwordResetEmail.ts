// serverside/src/config/passwordResetEmail.ts
import nodemailer from "nodemailer";
import config from "./config";

const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: config.emailUser,
    pass: config.emailPass,
  },
});

export const sendRecoveryEmail = async (
  to: string,
  token: string,
  name: string,
) => {
  const link = `${config.frontendUrl}/reset-password/${token}`;

  await transporter.sendMail({
    from: `"${config.appName} Security" <${config.emailUser}>`,
    to,
    subject: `SECURITY PROTOCOL: Password Recovery - ${config.appName}`,
    html: `
    <div style="font-family: 'Segoe UI', Arial, sans-serif; padding: 20px; background: #f4f4f4;">
      <div style="max-width: 600px; margin: auto; background: #ffffff; border-radius: 8px; overflow: hidden; border: 1px solid #e2e8f0;">
        <div style="background: #cfba29; padding: 20px; text-align: center;">
          <h1 style="margin: 0; color: #064e3b; font-size: 20px; text-transform: uppercase; letter-spacing: 2px;">Identity Recovery</h1>
        </div>
        <div style="padding: 30px; color: #334155; line-height: 1.6;">
          <p>Hello <strong>${name}</strong>,</p>
          <p>A request has been initiated to reset the Security Key for your institutional account at <strong>${config.instName}</strong>.</p>
          <p>To establish a new connection and update your credentials, please use the secure link below:</p>
          <div style="text-align: center; margin: 30px 0;">
            <a href="${link}" style="padding: 12px 25px; background: #156504; color: #ffffff; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">Update Security Key</a>
          </div>
          <p style="font-size: 12px; color: #94a3b8;">This link is single-use and expires in 60 minutes. If you did not request this, please contact your administrator immediately.</p>
        </div>
        <div style="background: #f8fafc; padding: 15px; text-align: center; font-size: 11px; color: #94a3b8;">
          Protected by ${config.appName} Infrastructure
        </div>
      </div>
    </div>
  `,
  });
};
