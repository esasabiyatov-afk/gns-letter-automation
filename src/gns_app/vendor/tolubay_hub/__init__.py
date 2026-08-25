"""Проверенный read-only HTTP-клиент Tolubay ABS."""

__version__ = "0.5.1+gns.1"

from .tolubay import (
    AccountRecord,
    AuthenticationError,
    CustomerQuestionnaire,
    CustomerSummary,
    ProtocolError,
    TolubayClient,
    TolubayConfig,
    TolubayError,
)

__all__ = [
    "AccountRecord",
    "AuthenticationError",
    "CustomerQuestionnaire",
    "CustomerSummary",
    "ProtocolError",
    "TolubayClient",
    "TolubayConfig",
    "TolubayError",
]
