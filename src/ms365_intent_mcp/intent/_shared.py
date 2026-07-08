"""Shared nested types used across intent payloads and responses."""

from __future__ import annotations

from typing import Annotated, Literal

from pydantic import BaseModel, ConfigDict, EmailStr, Field


class Recipient(BaseModel):
    """One email recipient. Used in email draft payloads."""

    model_config = ConfigDict(extra="forbid")
    email: Annotated[EmailStr, Field(description="RFC 5322 email address")]
    name: Annotated[str | None, Field(default=None, min_length=1)]


class Attendee(BaseModel):
    """One event attendee — email + optional role."""

    model_config = ConfigDict(extra="forbid")
    email: EmailStr
    name: str | None = None
    type: Literal["required", "optional", "resource"] = "required"


class BaseResponse(BaseModel):
    """Every tool response inherits from this. Subclasses pin ``type`` to a Literal."""

    model_config = ConfigDict(extra="forbid")
    type: str


class ErrorResponse(BaseResponse):
    """Uniform error envelope. Any tool can return this on failure."""

    type: Literal["error"] = "error"
    code: Literal[
        "not_found",
        "insufficient_scope",
        "invalid_id",
        "graph_api_error",
        "validation_error",
        "rate_limited",
    ]
    message: str
    retryable: bool = False
