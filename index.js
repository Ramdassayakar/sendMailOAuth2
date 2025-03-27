require("dotenv").config();
const express = require("express");
const axios = require("axios");
const { PublicClientApplication } = require("@azure/msal-node");
 
// Dynamic import for "open" (fixes the require() error)
const open = (...args) => import("open").then(({ default: open }) => open(...args));
 
const app = express();
const port = 3000;
 
let accessToken = null; // Stores the access token
let refreshToken = null; // Stores the refresh token
 
const msalConfig = {
    auth: {
        clientId: process.env.CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
        clientSecret: process.env.CLIENT_SECRET,
    },
};
 
const pca = new PublicClientApplication(msalConfig);
 
// Redirect to Microsoft Login
app.get("/login", (req, res) => {
    const authUrl = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/authorize?client_id=${process.env.CLIENT_ID}&response_type=code&redirect_uri=${process.env.REDIRECT_URI}&response_mode=query&scope=https://graph.microsoft.com/.default&state=12345`;
    res.redirect(authUrl);
});
 
// Callback from Microsoft OAuth
app.get("/redirect", async (req, res) => {
    const code = req.query.code;
    if (!code) return res.send("No authorization code received.");
 
    try {
        const tokenResponse = await axios.post(
            `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
            new URLSearchParams({
                client_id: process.env.CLIENT_ID,
                client_secret: process.env.CLIENT_SECRET,
                grant_type: "authorization_code",
                code: code,
                redirect_uri: process.env.REDIRECT_URI,
                scope: "https://graph.microsoft.com/Mail.Send offline_access",
            }).toString(),
            { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
        );
 
        accessToken = tokenResponse.data.access_token;
        refreshToken = tokenResponse.data.refresh_token; // Save refresh token
 
        res.send(`<h2>Authentication successful!</h2> <p>Email will be sent every 1 minute.</p>`);
       
        // Start automatic email sending
        startEmailScheduler();
    } catch (error) {
        console.error(error.response?.data || error.message);
        res.send("Error obtaining access token.");
    }
});
 
// Function to refresh Access Token using Refresh Token
async function refreshAccessToken() {
    if (!refreshToken) {
        console.error("No refresh token available.");
        return null;
    }
 
    try {
        const tokenResponse = await axios.post(
            `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
            new URLSearchParams({
                client_id: process.env.CLIENT_ID,
                client_secret: process.env.CLIENT_SECRET,
                grant_type: "refresh_token",
                refresh_token: refreshToken,
                scope: "https://graph.microsoft.com/Mail.Send offline_access",
            }).toString(),
            { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
        );
 
        accessToken = tokenResponse.data.access_token;
        refreshToken = tokenResponse.data.refresh_token; // Update refresh token
        console.log("Access token refreshed successfully!");
    } catch (error) {
        console.error("Error refreshing access token:", error.response?.data || error.message);
    }
}
 
// Function to send email
async function sendEmail() {
    if (!accessToken) {
        console.error("No valid access token, attempting refresh...");
        await refreshAccessToken();
    }
 
    if (!accessToken) {
        console.error("Failed to refresh access token. Skipping email.");
        return;
    }
 
    try {
        const emailData = {
            message: {
                subject: "Test Email from Node.js",
                body: { contentType: "Text", content: "Hello! This is a test email using OAuth2." },
                toRecipients: [{ emailAddress: { address: "chandan.yadav@sorigin.co" } }],
            },
            saveToSentItems: "true",
        };
 
        await axios.post("https://graph.microsoft.com/v1.0/me/sendMail", emailData, {
            headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
        });
 
        console.log("Email sent successfully at", new Date().toLocaleTimeString());
    } catch (error) {
        console.error("Error sending email:", error.response?.data || error.message);
    }
}
 
// Function to start email scheduler (every 1 minute)
function startEmailScheduler() {
    console.log("Email scheduler started. Sending an email every 1 minute...");
    sendEmail(); // Send the first email immediately
    setInterval(sendEmail, 60 * 1000); // Repeat every 1 minute
}
 
app.listen(port, async () => {
    console.log(`App running on http://localhost:${port}`);
    await open(`http://localhost:${port}/login`); // Ensures compatibility
});
 
 
 
// require("dotenv").config();
// const express = require("express");
// const axios = require("axios");
 
// const app = express();
// const port = 3000;
 
// let accessToken = null; // Store the access token
 
// // Function to get Access Token using Client Credentials Flow
// async function getAccessToken() {
//     try {
//         const response = await axios.post(
//             `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
//             new URLSearchParams({
//                 client_id: process.env.CLIENT_ID,
//                 client_secret: process.env.CLIENT_SECRET,
//                 grant_type: "client_credentials",
//                 scope: "https://graph.microsoft.com/.default"
//             }).toString(),
//             { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
//         );
 
//         accessToken = response.data.access_token;
//         console.log("✅ Access Token Obtained!");
//     } catch (error) {
//         console.error("❌ Error obtaining access token:", error.response?.data || error.message);
//     }
// }
 
// // Function to send email
// async function sendEmail() {
//     if (!accessToken) {
//         console.error("⚠️ No valid access token, attempting to fetch a new one...");
//         await getAccessToken();
//     }
 
//     if (!accessToken) {
//         console.error("❌ Failed to obtain access token. Skipping email.");
//         return;
//     }
 
//     try {
//         const emailData = {
//             message: {
//                 subject: "Automated Email from Node.js",
//                 body: { contentType: "Text", content: "Hello! This is an automated email sent using OAuth2 Client Credentials Flow." },
//                 toRecipients: [{ emailAddress: { address: "chandan.yadav@sorigin.co" } }],
//             },
//             saveToSentItems: "true",
//         };
 
//         await axios.post("https://graph.microsoft.com/v1.0/users/web@sorigin.co/sendMail", emailData, {
//             headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
//         });
 
//         console.log("✅ Email sent successfully at", new Date().toLocaleTimeString());
//     } catch (error) {
//         console.error("❌ Error sending email:", error.response?.data || error.message);
//     }
// }
 
// // Function to start email scheduler (every 1 minute)
// function startEmailScheduler() {
//     console.log("📧 Email scheduler started. Sending an email every 1 minute...");
//     sendEmail(); // Send the first email immediately
//     setInterval(sendEmail, 60 * 1000); // Repeat every 1 minute
// }
 
// // Start the app
// app.listen(port, async () => {
//     console.log(`🚀 Server running on http://localhost:${port}`);
//     await getAccessToken(); // Fetch Access Token on startup
//     startEmailScheduler(); // Start sending emails every 1 minute
// });
 
