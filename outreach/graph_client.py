from __future__ import annotations

import os

import config
from outreach.models import QueueEmail


GRAPH_SENDMAIL_URL = "https://graph.microsoft.com/v1.0/me/sendMail"


class GraphClient:
    def __init__(self) -> None:
        self.client_id = os.getenv(config.GRAPH_CLIENT_ID_ENV, "").strip()
        self.tenant_id = os.getenv(config.GRAPH_TENANT_ID_ENV, config.GRAPH_DEFAULT_TENANT).strip()
        if not self.client_id:
            raise RuntimeError(
                f"{config.GRAPH_CLIENT_ID_ENV} is not set. "
                "Create or request an Entra app registration with delegated Mail.Send."
            )

    def send_email(self, email: QueueEmail) -> str:
        import msal
        import requests

        app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
        )
        accounts = app.get_accounts()
        token = None
        if accounts:
            token = app.acquire_token_silent(config.GRAPH_SCOPES, account=accounts[0])

        if not token:
            flow = app.initiate_device_flow(scopes=config.GRAPH_SCOPES)
            if "user_code" not in flow:
                raise RuntimeError(f"Could not create Microsoft device-code flow: {flow}")
            print(flow["message"])
            token = app.acquire_token_by_device_flow(flow)

        if "access_token" not in token:
            raise RuntimeError(f"Microsoft authentication failed: {token}")

        payload = {
            "message": {
                "subject": email.subject,
                "body": {
                    "contentType": "Text",
                    "content": email.body,
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": email.recipient_email,
                        }
                    }
                ],
            },
            "saveToSentItems": True,
        }
        try:
            response = requests.post(
                GRAPH_SENDMAIL_URL,
                headers={
                    "Authorization": f"Bearer {token['access_token']}",
                    "Content-Type": "application/json",
                },
                json=payload,
                timeout=30,
            )
        except requests.exceptions.Timeout as exc:
            raise TimeoutError("Microsoft Graph sendMail timed out; send status is unknown.") from exc
        if response.status_code == 202:
            return "202 Accepted"
        response.raise_for_status()
        return f"{response.status_code} {response.text}"
