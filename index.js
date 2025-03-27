require("dotenv").config();
const express = require("express");
const axios = require("axios");
const { PublicClientApplication } = require("@azure/msal-node");

// Dynamic import for "open" (fixes the require() error)
const open = (...args) => import("open").then(({ default: open }) => open(...args));

const app = express();
const port = 3000;

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
                scope: "https://graph.microsoft.com/Mail.Send offline_access"

                // scope: "https://graph.microsoft.com/.default",
            }).toString(),
            { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
        );

        const accessToken = tokenResponse.data.access_token;
        res.send(`<h2>Authentication successful!</h2> <a href="/send-mail?token=${accessToken}">Send Email</a>`);
    } catch (error) {
        console.error(error.response?.data || error.message);
        res.send("Error obtaining access token.");
    }
});

// Send email using Microsoft Graph API
app.get("/send-mail", async (req, res) => {
    const accessToken = req.query.token;
    if (!accessToken) return res.send("Missing access token.");

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

        res.send("Email sent successfully!");
    } catch (error) {
        console.error(error.response?.data || error.message);
        res.send("Error sending email.");
    }
});

app.listen(port, async () => {
    console.log(`App running on http://localhost:${port}`);
    await open(`http://localhost:${port}/login`); // Ensures compatibility
});
