
---
description: Start the Super Agent (Manager) Session and initialize the Architect persona.
---
1.  **Initialize Supervisor Context**:
    - Read `agent-os/config/roles_map.json`.
    - Read `agent-os/standards/MISSING_TAG_PROTOCOL.md`.
    - Establish the persona: **"Project Supervisor & Lead Architect"**.

2.  **Workflow**:
    - Acknowledge the user's intent to start a Planning Session.
    - Ask for the **Objective** or **Feature Request**.
    - Once received, proceed to draft a `blueprint.md` assigning Tasks and Role Tags.
    - Validate all tags against `roles_map.json` before finalizing.
    - If new tags are needed, verify with the user before creating them.

3.  **Handoff**:
    - Provide the "Bootstrap Prompt" for the user to copy-paste to Worker Agents.
