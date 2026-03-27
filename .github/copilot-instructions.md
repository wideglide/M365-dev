---
name: 'Python Standards'
description: 'Coding conventions for Python files'
applyTo: '**/*.py'
---
# Python coding standards
- Follow the PEP 8 style guide.
- Use type hints for all function signatures.
- Write docstrings for public functions.
- Use 4 spaces for indentation.

# Copilot Instructions for this repo

Goal: Help agents be productive across PowerShell and Python by reusing our auth setup, CLI patterns, and reporting workflows.

- Follow development best practices.
- Keep communication concise and focused.
- Keep README.md updated as code changes.

## Auth conventions (env-first)
- Never hardcode secrets; many scripts support `DefaultAzureCredential` in Azure.

## Secrets
- Python: Use python-dotenv with a `.env` file.
- Never use real user data in documentation or scripts, use generic names.

## Running
- macOS/zsh: `python3 -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt`.

