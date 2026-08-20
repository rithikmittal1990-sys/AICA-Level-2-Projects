"""Errors raised by the PDF extraction layer."""

from __future__ import annotations


class PDFExtractionError(Exception):
    """Base error for upload and extraction failures."""

    status_code = 400
    error_code = "pdf_extraction_error"

    def __init__(self, message: str) -> None:
        super().__init__(message)
        self.message = message

    def to_dict(self) -> dict[str, str]:
        return {"error": self.error_code, "message": self.message}


class InvalidPDFError(PDFExtractionError):
    error_code = "invalid_pdf"


class CorruptedPDFError(PDFExtractionError):
    error_code = "corrupted_pdf"


class EncryptedPDFError(PDFExtractionError):
    error_code = "password_protected_pdf"
    status_code = 400


class EmptyPDFError(PDFExtractionError):
    error_code = "empty_pdf"


class PDFTooLargeError(PDFExtractionError):
    error_code = "pdf_too_large"
    status_code = 413


class TrialBalanceExtractionError(PDFExtractionError):
    """Trial balance upload and parsing failures."""

    error_code = "trial_balance_error"


class InvalidTrialBalanceError(TrialBalanceExtractionError):
    error_code = "invalid_trial_balance"


class TrialBalanceTooLargeError(TrialBalanceExtractionError):
    error_code = "trial_balance_too_large"
    status_code = 413
