"""Prepare report screenshots on the server, then send email through the local VPN."""

from __future__ import annotations

import argparse
import json
import mimetypes
import os
import sys
import tempfile
import uuid
from pathlib import Path
from urllib import request
from urllib.error import HTTPError, URLError


DEFAULT_SERVER_URL = "https://auto-report-secm.lintasarta.dev"
DEFAULT_SMTP_URL = "http://10.34.144.197/secm-portal/smtp/api_send_email"
DEFAULT_RECIPIENT = "ilhaam.akmal@lintasarta.co.id"

REPORT_DEFAULTS = {
    1: (
        "Daily Tracking Report 1",
        '<html><body><h3>Daily Tracking Report 1</h3>'
        '<p><img src="cid:target-complete-table" style="max-width:100%;"></p>'
        '<p><img src="cid:target-complete-legend" style="max-width:100%;"></p>'
        '<p><img src="cid:target-after-table" style="max-width:100%;"></p>'
        '<p><img src="cid:target-after-legend" style="max-width:100%;"></p></body></html>',
    ),
    2: (
        "Daily Tracking Report 2 Ongoing",
        '<html><body><h3>Daily Tracking Report 2 Ongoing</h3>'
        '<p><img src="cid:report-2-ongoing-table" style="max-width:100%;"></p>'
        '<p><img src="cid:report-2-ongoing-legend" style="max-width:100%;"></p>'
        '<p><img src="cid:report-2-after-table" style="max-width:100%;"></p>'
        '<p><img src="cid:report-2-after-legend" style="max-width:100%;"></p>'
        '<p><img src="cid:report-2-aging-table" style="max-width:100%;"></p></body></html>',
    ),
    3: (
        "Daily Tracking Report 3 iPhone",
        '<html><body><h3>Daily Tracking Report 3 iPhone</h3>'
        '<p><img src="cid:report-3-pivot" style="max-width:100%;"></p></body></html>',
    ),
    4: (
        "Daily Tracking Report 4 IDE",
        '<html><body><h3>Daily Tracking Report 4 IDE</h3>'
        '<p><img src="cid:report-4-complete-table" style="max-width:100%;"></p>'
        '<p><img src="cid:report-4-complete-legend" style="max-width:100%;"></p>'
        '<p><img src="cid:report-4-after-table" style="max-width:100%;"></p>'
        '<p><img src="cid:report-4-after-legend" style="max-width:100%;"></p></body></html>',
    ),
}


def encode_multipart(fields: dict[str, str], attachment: Path, file_field: str) -> tuple[bytes, str]:
    boundary = f"----LocalReportEmail{uuid.uuid4().hex}"
    chunks: list[bytes] = []
    for name, value in fields.items():
        chunks.extend([
            f"--{boundary}\r\n".encode(),
            f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode(),
            value.encode("utf-8"),
            b"\r\n",
        ])

    content_type = mimetypes.guess_type(attachment.name)[0] or "application/octet-stream"
    safe_filename = attachment.name.replace('"', "")
    chunks.extend([
        f"--{boundary}\r\n".encode(),
        (
            f'Content-Disposition: form-data; name="{file_field}"; '
            f'filename="{safe_filename}"\r\nContent-Type: {content_type}\r\n\r\n'
        ).encode(),
        attachment.read_bytes(),
        b"\r\n",
        f"--{boundary}--\r\n".encode(),
    ])
    return b"".join(chunks), f"multipart/form-data; boundary={boundary}"


def encode_multipart_attachments(fields: dict[str, str], attachments: list[Path]) -> tuple[bytes, str]:
    boundary = f"----LocalReportEmail{uuid.uuid4().hex}"
    chunks: list[bytes] = []
    for name, value in fields.items():
        chunks.extend([
            f"--{boundary}\r\n".encode(),
            f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode(),
            value.encode("utf-8"),
            b"\r\n",
        ])
    for attachment in attachments:
        content_type = mimetypes.guess_type(attachment.name)[0] or "application/octet-stream"
        chunks.extend([
            f"--{boundary}\r\n".encode(),
            (
                'Content-Disposition: form-data; name="attachment[]"; '
                f'filename="{attachment.name.replace(chr(34), "")}"\r\n'
                f"Content-Type: {content_type}\r\n\r\n"
            ).encode(),
            attachment.read_bytes(),
            b"\r\n",
        ])
    chunks.append(f"--{boundary}--\r\n".encode())
    return b"".join(chunks), f"multipart/form-data; boundary={boundary}"


def post_multipart(
    url: str,
    fields: dict[str, str],
    attachment: Path,
    timeout: int,
    cookie: str | None = None,
    file_field: str = "attachment[]",
    attachments: list[Path] | None = None,
) -> tuple[int, str]:
    if attachments:
        body, content_type = encode_multipart_attachments(fields, [attachment, *attachments])
    else:
        body, content_type = encode_multipart(fields, attachment, file_field)
    headers = {"Content-Type": content_type}
    if cookie:
        headers["Cookie"] = cookie
    api_request = request.Request(url, data=body, method="POST", headers=headers)
    try:
        with request.urlopen(api_request, timeout=timeout) as response:
            return response.status, response.read().decode("utf-8", errors="replace")
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {exc.code} from {url}: {detail}") from exc
    except URLError as exc:
        raise RuntimeError(f"Could not connect to {url}: {exc.reason}") from exc


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", type=int, choices=REPORT_DEFAULTS, required=True)
    parser.add_argument("--workbook", type=Path, required=True)
    parser.add_argument("--to", default=DEFAULT_RECIPIENT)
    parser.add_argument("--subject")
    parser.add_argument("--message-file", type=Path, help="Optional UTF-8 HTML template containing the report CID placeholders")
    parser.add_argument("--server-url", default=os.getenv("REPORT_SERVER_URL", DEFAULT_SERVER_URL))
    parser.add_argument("--smtp-url", default=os.getenv("SMTP_API_URL", DEFAULT_SMTP_URL))
    parser.add_argument("--smtp-cookie", default=os.getenv("SMTP_API_COOKIE"))
    parser.add_argument("--prepare-only", action="store_true", help="Publish screenshots without calling the VPN SMTP API")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    workbook = args.workbook.resolve()
    if not workbook.is_file() or workbook.suffix.lower() not in {".xlsx", ".xlsm"}:
        print(f"Workbook must be an existing .xlsx or .xlsm file: {workbook}", file=sys.stderr)
        return 2

    default_subject, default_message = REPORT_DEFAULTS[args.report]
    subject = args.subject or default_subject
    message = args.message_file.read_text(encoding="utf-8") if args.message_file else default_message
    prepare_url = f"{args.server_url.rstrip('/')}/email-report/prepare"

    try:
        print(f"Preparing Report {args.report} screenshots on {args.server_url} ...")
        _, response_body = post_multipart(
            prepare_url,
            {
                "report_number": str(args.report),
                "recipient": args.to,
                "subject": subject,
                "message": message,
            },
            workbook,
            timeout=30 * 60,
            file_field="workbook",
        )
        prepared = json.loads(response_body)
        email_message = str(prepared["email_message"])
        public_images = prepared.get("public_images", [])
        print("Published screenshots:")
        for public_image in public_images:
            print(f"  {public_image}")

        if args.prepare_only:
            print("Preparation complete. SMTP delivery skipped.")
            return 0

        print(f"Sending email through VPN SMTP API to {args.to} ...")
        image_paths: list[Path] = []
        with tempfile.TemporaryDirectory(prefix="report-email-images-") as temp_dir:
            for index, public_image in enumerate(public_images, start=1):
                image_path = Path(temp_dir) / f"report-{args.report}-image-{index}.jpg"
                with request.urlopen(str(public_image), timeout=60) as image_response:
                    image_path.write_bytes(image_response.read())
                image_paths.append(image_path)
            status, smtp_body = post_multipart(
                args.smtp_url,
                {
                    "to": args.to,
                    "subject": subject,
                    "message": email_message,
                    # The SMTP gateway must render this field as HTML.
                    "is_html": "true",
                },
                workbook,
                timeout=5 * 60,
                cookie=args.smtp_cookie,
                attachments=image_paths,
            )
        print(f"SMTP API responded {status}: {smtp_body}")
        return 0
    except (KeyError, ValueError, RuntimeError, json.JSONDecodeError) as exc:
        print(f"Email workflow failed: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
