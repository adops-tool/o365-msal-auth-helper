# O365 MSAL Authentication Helper

## Overview

This utility performs an interactive Microsoft identity platform authorization flow and creates an MSAL token cache for use with the [O365 Python library](https://github.com/O365/python-o365). It is intended for applications that need delegated access to Microsoft 365 mail but cannot rely on O365's built-in interactive authentication flow.

The script opens an authorization-code flow through MSAL, exchanges the returned one-time code for tokens, and serializes the complete cache to `o365_token.txt`. Saving the MSAL cache instead of a standalone access token preserves the account and refresh-token data that an O365 application needs for subsequent sessions.

## Prerequisites

Before using the utility, ensure that you have:

- Python 3.9 or later.
- A Microsoft Entra ID tenant (formerly Azure AD).
- An app registration configured as a **public client**. In the Azure portal, enable **Allow public client flows** under **Authentication**.
- The native-client redirect URI `https://login.microsoftonline.com/common/oauth2/nativeclient` configured for the app registration.
- The following delegated Microsoft Graph permissions configured for the app:
  - `Mail.ReadWrite`
  - `Mail.Send`
  - `User.Read`
- User or administrator consent for those permissions, as required by your tenant's consent policies.

Do not create or configure a client secret for this script. Public client applications cannot safely protect a secret on the user's device.

## Installation

1. Clone the repository and enter its directory:

   ```bash
   git clone <repository-url>
   cd o365-msal-token-cache-generator
   ```

2. Create a virtual environment:

   ```bash
   python -m venv .venv
   ```

3. Activate the virtual environment:

   On Linux or macOS:

   ```bash
   source .venv/bin/activate
   ```

   On Windows PowerShell:

   ```powershell
   .venv\Scripts\Activate.ps1
   ```

4. Install the dependencies:

   ```bash
   python -m pip install -r requirements.txt
   ```

5. Create a `.env` file in the repository root and provide the identifiers from your app registration:

   ```dotenv
   AZURE_CLIENT_ID=00000000-0000-0000-0000-000000000000
   AZURE_TENANT_ID=00000000-0000-0000-0000-000000000000
   ```

## Usage

1. Activate the virtual environment if it is not already active.
2. Run the authentication helper:

   ```bash
   python auth.py
   ```

3. Copy the authorization URL printed by the utility and open it in a browser.
4. Sign in with the Microsoft 365 account that the O365 application will use and grant the requested permissions.
5. After authentication, the browser redirects to the native-client URL. The resulting page may be blank; this is expected.
6. Copy the **complete URL from the browser address bar**. It must include the `code` and `state` query parameters.
7. Paste the URL into the terminal prompt and press Enter.
8. After a successful token exchange, confirm that `o365_token.txt` was created in the current directory. Configure your O365 application to use this serialized token cache.

The authorization code is short-lived and can be used only once. If the exchange fails, restart the script to create a new authorization flow and URL.

## Security Note

The `.env` file contains identifiers used to configure authentication, and `o365_token.txt` can contain access tokens, refresh tokens, and account metadata. Exposing the token cache may allow unauthorized access to the permitted Microsoft 365 resources.

Add both files to `.gitignore` before running the utility or committing changes:

```gitignore
.env
o365_token.txt
```

Never commit, share, log, or upload these files. If a token cache is exposed, revoke the affected user's sessions or application consent as appropriate, remove the exposed file from repository history, and authenticate again to generate a new cache.
