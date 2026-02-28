# CLAUDE.md — AI Assistant Guide

## Project Overview

**Repository:** `Kedros64/telegram-channel-land-water`

This project is a Telegram channel tool or bot focused on land and water topics. As the codebase evolves, this file will be updated to reflect the actual implementation details, conventions, and workflows.

> **Note:** This repository is in its initial state. This CLAUDE.md establishes baseline conventions for AI assistants contributing to the project. Update this file as the project grows.

---

## Repository Structure

```
telegram-channel-land-water/
├── CLAUDE.md          # This file — AI assistant guide
└── (source files to be added)
```

Update this section as directories and files are added.

---

## Development Workflow

### Branch Naming

- Feature branches: `feature/<short-description>`
- Bug fixes: `fix/<short-description>`
- Claude-managed branches: `claude/<task-id>` (managed automatically)

### Commit Messages

Use clear, imperative commit messages:

```
Add Telegram bot initialization
Fix land-parcel data parsing
Update water quality alert threshold
```

- Keep the subject line under 72 characters
- Use the body to explain *why*, not *what*

### Git Workflow

```bash
# Start a new feature
git checkout -b feature/<description>

# Push to remote
git push -u origin <branch-name>

# Never force-push to main/master
```

---

## Code Conventions

### General

- Prefer clarity over cleverness
- Keep functions small and single-purpose
- Document non-obvious logic with inline comments
- Avoid over-engineering — implement what is needed now

### Environment Variables

- Never hardcode secrets, tokens, or API keys in source files
- Use a `.env` file for local development (add `.env` to `.gitignore`)
- Document all required environment variables in a `.env.example` file

Expected variables (update as the project grows):

```
TELEGRAM_BOT_TOKEN=       # Bot API token from @BotFather
TELEGRAM_CHANNEL_ID=      # Target channel ID or @username
```

### Error Handling

- Handle errors at system boundaries (external APIs, user input, file I/O)
- Log errors with enough context to diagnose issues
- Do not silently swallow exceptions

---

## Testing

Document test commands here as a test suite is established. Example:

```bash
# Run all tests
pytest

# Run a specific test file
pytest tests/test_bot.py
```

- Write tests for business logic and data-processing functions
- Do not test framework internals or third-party library behavior

---

## Key Decisions & Context

| Decision | Rationale |
|---|---|
| (none yet) | (add decisions here as they are made) |

---

## Common Tasks for AI Assistants

1. **Adding a new feature:** Read existing related code before writing anything. Follow the patterns already established.
2. **Fixing a bug:** Understand the root cause before changing code. Do not add workarounds that mask the underlying issue.
3. **Updating dependencies:** Check for breaking changes in changelogs. Run tests after updating.
4. **Adding environment variables:** Add them to `.env.example` with a comment explaining their purpose.
5. **Updating this file:** Keep CLAUDE.md current whenever the project structure, conventions, or workflows change significantly.

---

## What to Avoid

- Do not commit `.env` files or any file containing real credentials
- Do not add unnecessary abstractions or generalize prematurely
- Do not introduce dependencies without a clear reason
- Do not push directly to `main` or `master`
- Do not amend commits that have already been pushed

---

*Last updated: 2026-02-28 — Initial scaffold (empty repository)*
