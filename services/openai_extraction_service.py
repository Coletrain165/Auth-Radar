"""
Auth Radar - OpenAI Extraction Service

Sends PDF page images to OpenAI's vision-capable model and extracts
structured authorization data using Structured Outputs (JSON schema).
"""

import base64
import json
import logging
import os

from openai import OpenAI

from config import OPENAI_API_KEY, OPENAI_MODEL

logger = logging.getLogger(__name__)

# --- The JSON schema that OpenAI Structured Outputs will enforce ---
AUTH_EXTRACTION_SCHEMA = {
    "type": "json_schema",
    "json_schema": {
        "name": "auth_extraction",
        "strict": True,
        "schema": {
            "type": "object",
            "properties": {
                "auth_number": {
                    "anyOf": [{"type": "string"}, {"type": "null"}],
                    "description": "Authorization number exactly as shown on the form.",
                },
                "date_approved": {
                    "anyOf": [{"type": "string"}, {"type": "null"}],
                    "description": "Date approved in YYYY-MM-DD format.",
                },
                "date_auth_expired": {
                    "anyOf": [{"type": "string"}, {"type": "null"}],
                    "description": "Date authorization expired in YYYY-MM-DD format.",
                },
                "participant_name": {
                    "anyOf": [{"type": "string"}, {"type": "null"}],
                    "description": "Participant name as shown, preferably LAST, FIRST.",
                },
                "participant_id": {
                    "anyOf": [{"type": "string"}, {"type": "null"}],
                    "description": "Participant ID exactly as shown.",
                },
                "source_page": {
                    "type": "integer",
                    "description": "1-indexed page number used for extraction.",
                },
                "confidence": {
                    "type": "object",
                    "properties": {
                        "auth_number": {"type": "number"},
                        "date_approved": {"type": "number"},
                        "date_auth_expired": {"type": "number"},
                        "participant_name": {"type": "number"},
                        "participant_id": {"type": "number"},
                    },
                    "required": [
                        "auth_number",
                        "date_approved",
                        "date_auth_expired",
                        "participant_name",
                        "participant_id",
                    ],
                    "additionalProperties": False,
                },
                "notes": {
                    "type": "string",
                    "description": "Notes about uncertainty, low confidence, or multiple possible values.",
                },
            },
            "required": [
                "auth_number",
                "date_approved",
                "date_auth_expired",
                "participant_name",
                "participant_id",
                "source_page",
                "confidence",
                "notes",
            ],
            "additionalProperties": False,
        },
    },
}

# --- System prompt for OpenAI ---
SYSTEM_PROMPT = """You are extracting authorization data from a scanned healthcare authorization form.

Return only valid JSON matching the provided schema.

Extract these fields:
- Auth #
- Date Approved
- Date Auth. Expired
- Participant's Name
- Participant ID

Rules:
- Do not guess.
- If a value is missing, unreadable, or uncertain, return null.
- Normalize dates to YYYY-MM-DD.
- Preserve IDs and auth numbers exactly as shown.
- Use the Treatment Authorization Form page when available.
- The participant fields may appear under a section labeled MEMBER (Participant).
- The date fields usually appear near the top of the Treatment Authorization Form.
- Include confidence scores from 0 to 1 for each field.
- In the notes field, explain any uncertainty or if multiple possible values were found."""


class OpenAIExtractionService:
    """
    Sends page images to OpenAI and extracts structured auth data.

    Uses the OpenAI Responses API with vision-capable models and
    Structured Outputs to guarantee consistent JSON responses.
    """

    def __init__(self, api_key: str = "", model: str = ""):
        self.api_key = api_key or OPENAI_API_KEY
        self.model = model or OPENAI_MODEL

        if not self.api_key:
            raise ValueError(
                "OPENAI_API_KEY is not configured. "
                "Set it in your .env file or environment variables."
            )

        self.client = OpenAI(api_key=self.api_key)

    def extract_from_image(self, image_path: str, page_num: int) -> dict:
        """
        Extract authorization data from a single page image.

        Args:
            image_path: Path to the PNG/JPG image of the PDF page.
            page_num: 1-indexed page number (for metadata).

        Returns:
            Dict matching the auth extraction schema.

        Raises:
            RuntimeError: If the API call fails or response is unparseable.
        """
        if not os.path.isfile(image_path):
            raise FileNotFoundError(f"Image not found: {image_path}")

        # Encode image to base64
        with open(image_path, "rb") as f:
            image_data = base64.b64encode(f.read()).decode("utf-8")

        # Determine MIME type
        ext = os.path.splitext(image_path)[1].lower()
        mime_type = "image/png" if ext == ".png" else "image/jpeg"

        logger.info("Sending page %d image to OpenAI (%s)...", page_num, self.model)

        # --- OpenAI API call using chat completions with vision ---
        # This uses the standard chat completions endpoint which supports
        # vision inputs and structured outputs (response_format).
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "text",
                                "text": (
                                    f"Extract the authorization data from this page image. "
                                    f"This is page {page_num} of the PDF document."
                                ),
                            },
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:{mime_type};base64,{image_data}",
                                    "detail": "high",
                                },
                            },
                        ],
                    },
                ],
                response_format=AUTH_EXTRACTION_SCHEMA,
                temperature=0.0,
                max_tokens=1000,
            )
        except Exception as e:
            raise RuntimeError(f"OpenAI API call failed: {e}") from e

        # Parse the structured response
        raw_content = response.choices[0].message.content
        logger.info("OpenAI extraction completed for page %d", page_num)
        logger.info("Raw response: %s", raw_content[:500])

        try:
            result = json.loads(raw_content)
        except json.JSONDecodeError as e:
            raise RuntimeError(
                f"OpenAI returned invalid JSON: {e}\nRaw: {raw_content[:500]}"
            ) from e

        # Log what we got
        fields_found = [f for f in ["auth_number", "date_approved", "date_auth_expired",
                                     "participant_name", "participant_id"] if result.get(f)]
        logger.info("Fields extracted: %s", fields_found if fields_found else "NONE")

        # Ensure source_page is set correctly
        result["source_page"] = page_num

        return result

    def extract_from_images(self, image_pages: list[tuple[int, str]]) -> dict:
        """
        Extract from multiple page images. Tries the first page with the best
        result; falls back to subsequent pages if extraction quality is low.

        Args:
            image_pages: List of (page_num_0indexed, image_path) tuples.

        Returns:
            Best extraction result dict.
        """
        best_result = None
        best_score = -1

        for page_num_0, image_path in image_pages:
            page_num_1 = page_num_0 + 1  # Convert to 1-indexed for display

            try:
                result = self.extract_from_image(image_path, page_num_1)
            except Exception as e:
                logger.warning("Extraction failed for page %d: %s", page_num_1, e)
                continue

            # Score based on how many fields were extracted
            score = self._score_result(result)
            logger.info("Page %d extraction score: %d/5", page_num_1, score)

            if score > best_score:
                best_score = score
                best_result = result

            # If we got all 5 fields with good confidence, no need to try more pages
            if score >= 5 and self._min_confidence(result) >= 0.7:
                break

        if best_result is None:
            return {
                "auth_number": None,
                "date_approved": None,
                "date_auth_expired": None,
                "participant_name": None,
                "participant_id": None,
                "source_page": image_pages[0][0] + 1 if image_pages else 0,
                "confidence": {
                    "auth_number": 0.0,
                    "date_approved": 0.0,
                    "date_auth_expired": 0.0,
                    "participant_name": 0.0,
                    "participant_id": 0.0,
                },
                "notes": "OpenAI extraction failed for all candidate pages.",
            }

        return best_result

    def _score_result(self, result: dict) -> int:
        """Count non-null extracted fields."""
        fields = [
            "auth_number",
            "date_approved",
            "date_auth_expired",
            "participant_name",
            "participant_id",
        ]
        return sum(1 for f in fields if result.get(f) is not None)

    def _min_confidence(self, result: dict) -> float:
        """Get the minimum confidence among non-null fields."""
        confidence = result.get("confidence", {})
        fields = [
            "auth_number",
            "date_approved",
            "date_auth_expired",
            "participant_name",
            "participant_id",
        ]
        scores = [
            confidence.get(f, 0.0)
            for f in fields
            if result.get(f) is not None
        ]
        return min(scores) if scores else 0.0
