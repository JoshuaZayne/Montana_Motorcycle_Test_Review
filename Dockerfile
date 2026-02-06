FROM python:3.11-slim

LABEL maintainer="JoshuaZayne"
LABEL description="Montana Motorcycle Test - Flashcard Generator"

WORKDIR /app

# Install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy source
COPY generate_flashcards.py .

# Output directory inside the container
RUN mkdir -p /app/output

# Override the output directory so files land in /app/output
ENV FLASHCARD_OUTPUT_DIR=/app/output

CMD ["python", "generate_flashcards.py"]
