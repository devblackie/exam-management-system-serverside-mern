// serverside/src/services/twoFactorService.ts
import nodemailer from "nodemailer";
import config from "../config/config";

const getTransporter = () =>
  nodemailer.createTransport({
    host: config.emailHot || "smtp.gmail.com",
    port: Number(config.emailPot) || 587,
    secure: false,
    auth: {user: config.emailUser, pass: config.emailPass},
  });

export const sendOTPEmail = async (
  email: string,
  name: string,
  otp: string,
  purpose: "login" | "register",
): Promise<void> => {
  const transporter = getTransporter();

  const subject =
    purpose === "login" ? "Your login verification code" : "Verify your registration";

  const fromName = config.appName || "EMS Academic System";

  await transporter.sendMail({
    from: `"${fromName}" <${config.emailUser}>`,
    to: email,
    subject,
    html: `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
</head>
<body style="margin:0;padding:0;background:#f5f4f0;font-family:'Georgia',serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="padding:40px 0;">
    <tr>
      <td align="center">
        <table width="480" cellpadding="0" cellspacing="0"
               style="background:#1a3a2a;border-radius:16px;overflow:hidden;">
          <!-- Header -->
          <tr>
            <td style="padding:36px 40px 24px;text-align:center;">
              <div style="display:inline-block;width:48px;height:48px;background:#c9a227;
                          border-radius:12px;line-height:48px;font-size:24px;">
                
              </div>
              <h1 style="color:#c9a227;font-size:13px;letter-spacing:4px;
                         text-transform:uppercase;margin:16px 0 0;font-weight:400;">
                ${purpose === "login" ? "Login Verification" : "Registration Verification"}
              </h1>
            </td>
          </tr>
          <!-- Body -->
          <tr>
            <td style="padding:0 40px 36px;">
              <p style="color:#a8c4b0;font-size:15px;line-height:1.6;margin:0 0 24px;">
                Hi ${name},
              </p>
              <p style="color:#a8c4b0;font-size:15px;line-height:1.6;margin:0 0 32px;">
                ${
                  purpose === "login"
                    ? "Enter the code below to complete your login. It expires in <strong style='color:#c9a227'>10 minutes</strong>."
                    : "Enter the code below to verify your registration. It expires in <strong style='color:#c9a227'>10 minutes</strong>."
                }
              </p>
              <!-- OTP Code Box -->
              <div style="background:#0d2218;border-radius:12px;padding:32px;
                          text-align:center;margin-bottom:32px;">
                <div style="font-family:'Courier New',monospace;font-size:48px;
                            font-weight:700;letter-spacing:16px;color:#c9a227;
                            text-shadow:0 0 30px rgba(201,162,39,0.3);">
                  ${otp}
                </div>
              </div>
              <p style="color:#5a7a65;font-size:13px;line-height:1.6;margin:0;">
                If you did not request this, you can safely ignore this email.
                Your account remains secure.
              </p>
            </td>
          </tr>
          <!-- Footer -->
          <tr>
            <td style="padding:20px 40px;border-top:1px solid #2a4a3a;
                       text-align:center;">
              <p style="color:#3a5a4a;font-size:11px;letter-spacing:2px;
                        text-transform:uppercase;margin:0;">
                Exams Management System · Secured
              </p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>
    `,
  });
};
