FROM python:3.11-slim-bookworm
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONIOENCODING=utf-8 \
    OPENBLAS_NUM_THREADS=1 \
    OMP_NUM_THREADS=1 \
    MKL_NUM_THREADS=1
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt && pip check
COPY main.py excelqa.py ./
COPY scripts/__init__.py scripts/validate_submission.py ./scripts/
# The grader can mount a root-owned output directory without granting other UIDs write access.
USER 0:0
ENTRYPOINT ["python", "main.py"]
