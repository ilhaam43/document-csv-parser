FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY csv_to_excel.py csv_to_excel_api.py csv_to_excel_on_going.py generate_iphone_tracking.py generate_ide_tracking.py app.py send_report_1_email.py ./
COPY templates ./templates
COPY static ./static
RUN mkdir -p input-today output-today vlookup-yesterday

EXPOSE 8000

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
