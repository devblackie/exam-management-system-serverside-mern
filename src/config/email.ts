import nodemailer from "nodemailer";
import config from "./config";

const transporter = nodemailer.createTransport({
  service: "gmail", // or use SMTP provider
  auth: {
    user: config.emailUser,
    pass: config.emailPass,
  },
});

// Send invite email with token link
export const sendInviteEmail = async (
  to: string,
  token: string,
  name: string
) => {
  const link = `${config.frontendUrl}/register/${token}`;

  await transporter.sendMail({
    from: `"${config.appName}" <${config.emailUser}>`,
    to,
    subject: `Account Invitation - ${config.appName}`,
    html: `
     <div style="font-family: Arial, sans-serif; padding: 7px;background:#cfba29; border:1px solid #ddd; border-radius:5px;">
     <p>Hello <strong>${name}</strong> ,</p>   
     <p>Welcome to <strong>${config.appName}</strong></p>
        <p>You have been invited to join the examination system for <strong>${config.instName}</strong>.</p>
        <p>Please set your password using the link below (expires in 24 hours):</p>
        <a href="${link}" style="padding:5px 7px;background:#156504;color:#fff;text-decoration:none;border-radius:5px;">Register</a>
        <p>If you did not request this, please ignore this email.</p>
      </div>
    `,
  });
};
