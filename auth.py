"""Create an MSAL token cache that can be reused by the O365 library.

This command-line utility performs an interactive Azure AD authorization-code
flow for a public client application. It directs the user to Microsoft's sign-in
page, accepts the final redirect URL, exchanges the authorization response for
tokens, and stores the resulting MSAL cache in ``o365_token.txt``.

The O365 library can reuse this file because ``SerializableTokenCache`` writes
MSAL's complete cache representation rather than an isolated access token. The
serialized data may include access tokens, refresh tokens, account metadata, and
expiration information, allowing MSAL-backed clients to renew access without
requiring the user to sign in for every run. Treat the generated file as a
credential and never commit it to source control.
"""

import os
import urllib.parse

import msal
from dotenv import load_dotenv
from msal.exceptions import MsalServiceError


# Load local configuration before resolving the Azure application identifiers.
load_dotenv()

CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
TENANT_ID = os.getenv("AZURE_TENANT_ID")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# Request only the delegated Microsoft Graph permissions required by the client.
SCOPES = ["Mail.ReadWrite", "Mail.Send", "User.Read"]
REDIRECT_URI = "https://login.microsoftonline.com/common/oauth2/nativeclient"
TOKEN_CACHE_PATH = "o365_token.txt"


def main() -> None:
    """Run the interactive authorization flow and persist the MSAL token cache."""
    print("--- MSAL TOKEN CACHE GENERATOR ---")

    # Fail early with a clear configuration error instead of contacting an
    # invalid tenant authority or constructing an application without a client ID.
    if not CLIENT_ID or not TENANT_ID:
        print(
            "\nConfiguration error: AZURE_CLIENT_ID and AZURE_TENANT_ID "
            "must be set in the .env file."
        )
        return

    # SerializableTokenCache is required because a regular in-memory MSAL cache
    # cannot be written to disk. Its serialized schema is the format expected by
    # O365's MSAL-based token backend, including refresh and account metadata.
    cache = msal.SerializableTokenCache()

    # A public client has no client secret because authentication occurs on the
    # user's device. The cache receives all token updates made by this MSAL app.
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache,
    )

    # Start the authorization-code flow directly through MSAL. The returned flow
    # contains both the sign-in URL and transient state used to validate the reply.
    flow = app.initiate_auth_code_flow(
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
    )

    print("\nOpen this authorization URL in your browser:")
    print(flow["auth_uri"])

    # The native-client redirect may display a blank page. Its address bar still
    # contains the authorization response that MSAL needs to finish the flow.
    final_url = input(
        "\nPaste the complete URL from the final browser page here: "
    ).strip()

    if "code=" not in final_url:
        print("\nAuthorization error: the URL does not contain a 'code' parameter.")
        return

    try:
        # Convert the redirect query into the mapping expected by MSAL while
        # retaining fields such as code, state, and session_state.
        parsed_url = urllib.parse.urlparse(final_url)
        query_params = {
            key: values[0]
            for key, values in urllib.parse.parse_qs(parsed_url.query).items()
        }

        # MSAL validates the returned state against the original flow and sends
        # the one-time authorization code to Azure AD's token endpoint. Azure AD
        # then returns the tokens, which MSAL also records in the attached cache.
        result = app.acquire_token_by_auth_code_flow(
            flow,
            auth_response=query_params,
        )

        if "access_token" not in result:
            error_message = result.get(
                "error_description",
                "Azure AD returned an unspecified error.",
            )
            print(f"\nToken acquisition failed: {error_message}")
            return

        # Persist the entire cache, not merely the short-lived access token, so
        # O365 can deserialize it and use its refresh data in later sessions.
        with open(TOKEN_CACHE_PATH, "w", encoding="utf-8") as cache_file:
            cache_file.write(cache.serialize())

        print(f"\nSuccess: the MSAL token cache was saved to {TOKEN_CACHE_PATH}.")
        print("The cache is now ready for use by your O365 application.")
    except (KeyError, ValueError, MsalServiceError) as error:
        # Report malformed responses and MSAL service failures without exposing
        # the token cache or other credential material.
        print(f"\nAuthorization failed: {error}")


if __name__ == "__main__":
    # Keep imports side-effect free while retaining direct command-line execution.
    main()
