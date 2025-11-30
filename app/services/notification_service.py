# -*- coding: utf-8 -*-
"""
Utilities for sending free game summaries through Microsoft email (OAuth).
"""

from typing import Iterable

import httpx
import msal
from loguru import logger

from models import PromotionGame
from settings import settings


class MicrosoftMailClient:
    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        sender: str,
        recipients: list[str],
    ):
        self.authority = f"https://login.microsoftonline.com/{tenant_id}"
        self.client_id = client_id
        self.client_secret = client_secret
        self.sender = sender
        self.recipients = recipients

    @classmethod
    def from_settings(cls, config=settings) -> "MicrosoftMailClient":
        return cls(
            tenant_id=config.MS_TENANT_ID,
            client_id=config.MS_CLIENT_ID,
            client_secret=config.MS_CLIENT_SECRET.get_secret_value(),
            sender=config.MS_SENDER,
            recipients=config.MS_RECIPIENTS,
        )

    def _acquire_token(self) -> str:
        app = msal.ConfidentialClientApplication(
            self.client_id,
            authority=self.authority,
            client_credential=self.client_secret,
        )
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" not in result:
            raise RuntimeError(
                f"Failed to acquire Microsoft Graph token: {result.get('error_description')}"
            )
        return result["access_token"]

    async def send_promotions(self, promotions: Iterable[PromotionGame]):
        if not promotions:
            logger.debug("No promotions to send, skipping email notification")
            return
        if not self.recipients:
            logger.warning(
                "Microsoft email recipients are not configured; skipping free game notification"
            )
            return

        token = self._acquire_token()

        items = []
        for promotion in promotions:
            items.append(
                f"<li><a href=\"{promotion.url}\">{promotion.title}</a> — {promotion.description}</li>"
            )
        html_body = "".join(
            [
                "<p>以下是本周可领取的 Epic 免费游戏：</p>",
                "<ul>",
                "".join(items),
                "</ul>",
                "<p>邮件由 epic-awesome-gamer 自动发送。</p>",
            ]
        )

        message = {
            "message": {
                "subject": "Epic 免费游戏更新",
                "body": {"contentType": "HTML", "content": html_body},
                "toRecipients": [{"emailAddress": {"address": r}} for r in self.recipients],
            }
        }

        url = f"https://graph.microsoft.com/v1.0/users/{self.sender}/sendMail"
        headers = {"Authorization": f"Bearer {token}"}

        async with httpx.AsyncClient(timeout=30) as client:
            resp = await client.post(url, json=message, headers=headers)
            resp.raise_for_status()

        logger.success("Free game notification sent via Microsoft email")

