"""Typed error codes used by the pipeline. Stable contract for the API layer."""
from enum import Enum

class ErrorCode(str, Enum):
    EXCEL_INVALID = "EXCEL_INVALID"
    EXCEL_EMPTY = "EXCEL_EMPTY"
    EXCEL_INSUFFICIENT_DATA = "EXCEL_INSUFFICIENT_DATA"
    AI_SATURATED = "AI_SATURATED"
    AI_RESPONSE_INVALID = "AI_RESPONSE_INVALID"
    PLANNER_REJECTED_PROMPT = "PLANNER_REJECTED_PROMPT"
    PYTHON_RUNTIME_ERROR = "PYTHON_RUNTIME_ERROR"
    TIMEOUT = "TIMEOUT"

class PipelineError(Exception):
    def __init__(self, code: ErrorCode, message: str, details: str = "",
                 user_action: str = "report_bug", retry_after_seconds: int = 0):
        self.code = code
        self.message = message
        self.details = details
        self.user_action = user_action
        self.retry_after_seconds = retry_after_seconds
        super().__init__(f"[{code.value}] {message}")

    def to_dict(self) -> dict:
        return {
            "code": self.code.value,
            "message": self.message,
            "details": self.details,
            "user_action": self.user_action,
            "retry_after_seconds": self.retry_after_seconds,
        }
