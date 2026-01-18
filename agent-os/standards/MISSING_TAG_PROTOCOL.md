
# Protocol: Missing Configuration Handling

> **CRITICAL INSTRUCTION**: This protocol overrides all other instructions regarding task execution.

## 1. Validation Phase (Before Action)
Upon receiving a list of Role Tags (e.g., from a Blueprint or User input), the Agent **MUST** perform the following atomic check:

1.  **Load Map**: Read `agent-os/config/roles_map.json`.
2.  **Verify Tags**: For *every* assigned tag, verify if a corresponding key exists in the `roles` object.
3.  **Verify Content**: If the key exists, verify the file path is non-empty.

## 2. Failure Handling (If ANY tag is missing)
If *any* assigned tag is not found in the map:

1.  **HALT IMMEDIATELY**. Do not attempt to guess or hallucinate the role's purpose.
2.  **REPORT**: Output a specific error block:
    ```markdown
    > [!CAUTION]
    > **CONFIGURATION MISSING**
    > The following tags are assigned but undefined in `roles_map.json`:
    > - [Missing Tag Name 1]
    > - [Missing Tag Name 2]
    ```
3.  **PROPOSE**: Ask the user:
    "Shall I create a new rule file for `[Missing Tag]`? If yes, please provide the rules, or authorize me to draft them based on our context."

## 3. Recovery Phase
Only proceed to task execution AFTER:
1.  The user provides the rule content.
2.  The new file is created (e.g., `agent-os/standards/[tag_name].md`).
3.  The `agent-os/config/roles_map.json` is updated with the new mapping.
4.  The Agent re-initializes its context.
