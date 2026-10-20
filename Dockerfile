# Use a Debian-based Python slim image
# Note: pyproject.toml requires Python >= 3.12 (.python-version specifies 3.12)
FROM python:3.12-slim-bookworm

# Set environment variables
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    UV_COMPILE_BYTECODE=1 \
    UV_LINK_MODE=copy \
    PATH="/app/.venv/bin:$PATH"

# Install system dependencies (pandoc for pypandoc conversion)
RUN apt-get update && apt-get install -y --no-install-recommends \
    pandoc \
    curl \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# Install uv from official binary image
COPY --from=ghcr.io/astral-sh/uv:latest /uv /uvx /bin/

# Set working directory
WORKDIR /app

# Copy dependency definition files first for optimal layer caching
COPY pyproject.toml uv.lock ./

# Install project dependencies into virtual environment
RUN uv sync --frozen --no-install-project --no-dev

# Copy source code and project metadata
COPY README.md ./
COPY src/ ./src/

# Install the project itself
RUN uv sync --frozen --no-dev

# Command to run the MCP server
CMD ["python", "-m", "mcp_qpgenerator.server"]
