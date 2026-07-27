"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.sendEmail = sendEmail;
// serverside/src/lib/mailer.ts
const nodemailer_1 = __importDefault(require("nodemailer"));
const config_1 = __importDefault(require("../config/config"));
const transporter = nodemailer_1.default.createTransport({
    service: "gmail",
    auth: { user: config_1.default.emailUser, pass: config_1.default.emailPass },
});
async function sendEmail({ to, subject, html, attachments, }) {
    await transporter.sendMail({
        from: `"${config_1.default.appName}" <${config_1.default.emailUser}>`,
        to,
        subject,
        html,
        attachments,
    });
}
