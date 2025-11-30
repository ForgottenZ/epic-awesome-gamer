# -*- coding: utf-8 -*-
"""
Microsoft Graph mailer for sending free game digests via OAuth.
"""
from __future__ import annotations

from typing import Iterable

import httpx
import msal
from loguru import logger

from models import PromotionGame
from settings import settings


class MsGraphMailer:
    def __init__(
        self,
        client_id: str | None,
        client_secret: str | None,
        tenant_id: str | None,
        sender: str | None,
        recipient: str | None,
    ):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.sender = sender
        self.recipient = recipient

        authority = f"https://login.microsoftonline.com/{self.tenant_id}" if self.tenant_id else None
        self._confidential_client = (
            msal.ConfidentialClientApplication(
                client_id=self.client_id,
                client_credential=self.client_secret,
                authority=authority,
            )
            if all([self.client_id, self.client_secret, authority])
            else None
        )

    def is_configured(self) -> bool:
        if not all([self.client_id, self.client_secret, self.tenant_id, self.sender, self.recipient]):
            logger.warning(
                "Microsoft Graph mailer is not fully configured; skip sending digest.",
                missing_configs=[
                    name
                    for name, value in [
                        ("MS_CLIENT_ID", self.client_id),
                        ("MS_CLIENT_SECRET", self.client_secret),
                        ("MS_TENANT_ID", self.tenant_id),
                        ("MS_SENDER_ADDRESS", self.sender),
                        ("MS_RECIPIENT_ADDRESS", self.recipient),
                    ]
                    if not value
                ],
            )
            return False

        if not self._confidential_client:
            logger.error("Failed to initialize MSAL confidential client; check configuration.")
            return False

        return True

    def _acquire_token(self) -> str | None:
        assert self._confidential_client
        scopes = ["https://graph.microsoft.com/.default"]

        result = self._confidential_client.acquire_token_silent(scopes=scopes, account=None)
        if not result:
            result = self._confidential_client.acquire_token_for_client(scopes=scopes)

        if not result or "access_token" not in result:
            logger.error(
                "Failed to acquire Microsoft Graph access token.",
                error=result.get("error") if result else None,
                description=result.get("error_description") if result else None,
            )
            return None

        return result["access_token"]

    @staticmethod
    def _build_body(promotions: Iterable[PromotionGame]) -> str:
        items = list(promotions)
        if not items:
            return "<p>No free games detected from Epic Store at this time.</p>"

        game_items = "".join(
            f"<li><strong>{game.title}</strong> - <a href=\"{game.url}\">{game.url}</a></li>"
            for game in items
        )
        return f"<p>Here are the current free games on Epic Store:</p><ul>{game_items}</ul>"

    async def send_free_game_digest(self, promotions: Iterable[PromotionGame]):
        if not self.is_configured():
            return

        token = self._acquire_token()
        if not token:
            return

        payload = {
            "message": {
                "subject": "Epic Store - Weekly Free Games Update",
                "body": {"contentType": "HTML", "content": self._build_body(promotions)},
                "toRecipients": [{"emailAddress": {"address": self.recipient}}],
                "from": {"emailAddress": {"address": self.sender}},
            },
            "saveToSentItems": "true",
        }

        headers = {"Authorization": f"Bearer {token}"}
        url = f"https://graph.microsoft.com/v1.0/users/{self.sender}/sendMail"

        async with httpx.AsyncClient(timeout=30) as client:
            response = await client.post(url, json=payload, headers=headers)

        if response.is_success:
            logger.success("Free game digest email sent via Microsoft Graph.")
        else:
            logger.error(
                "Failed to send free game digest email.",
                status_code=response.status_code,
                content=response.text,
            )


ms_mailer = MsGraphMailer(
    client_id=settings.MS_CLIENT_ID,
    client_secret=settings.MS_CLIENT_SECRET.get_secret_value()
    if settings.MS_CLIENT_SECRET
    else None,
    tenant_id=settings.MS_TENANT_ID,
    sender=settings.MS_SENDER_ADDRESS,
    recipient=settings.MS_RECIPIENT_ADDRESS,
)
