# Session Rolling Spec

Last updated: 2026-01-26
Source session: ses_416619046ffefwzYQcRGEIKGLC

## Context
- Project: rewrite legacy VB.NET production management system with SAP B1 integration into a modern stack (jet-platform).
- Long-term goal: decouple from .NET/Visual Studio stack, move to future-proof, AI-collaboration-friendly architecture.

## Phase 1 Goal (Must Be 1:1 Parity)
- Functional parity with the legacy system for the user-owned features.
- No new behavior changes, no optimizations, minimal change.
- Keep `jID` as the global unique transaction/document ID (name unchanged).

## Confirmed Scope (Phase 1)
- Login
- HOME page (no change)
- Feature list
- Purchase Request
- Expense Claim
- Document Search (replicate current behavior)
- Account Settings + Password Change (1:1 behavior from existing WebForms)

## Key Decisions
- `jID` is the only global unique transaction/document ID; permissions are fully decoupled from `jID`.
- Old permission flags (`AP_App`, `PU_App`, `Approver`, `Admin`) will be migration sources only, then fully retired.
- New permission system should be RBAC (roles/permissions), ABAC only if needed later.
- Account settings and password change should keep current behavior (no hashing, no new security logic) for Phase 1.

## Permission Model (Target Design)
- Core tables: `users`, `roles`, `permissions`, `user_roles`, `role_permissions`.
- Suggested permissions:
  - `expense_claim:create|read|update|approve|read_all`
  - `purchase_request:create|read|update|approve|read_all`
  - `document:read|read_all`
  - `system:maintenance_bypass`
- Migration mapping:
  - `AP_App=1` -> `ap-approver`
  - `PU_App=1` -> `pu-approver`
  - `Approver=1` -> `expense-approver`
  - `Admin=1` -> `admin`

## Account Settings + Password Change (Phase 1)
- Entry points:
  - Keep existing Home page entry.
  - Add a new entry in the feature list pointing to the same page.
- API (minimal):
  - `GET /api/v1/users/me`
  - `PUT /api/v1/users/me`
  - `POST /api/v1/users/me/password`
- Fields (match legacy behavior): `name`, `email`, `expDept`, `empSeries`.

## Document Search (Phase 1)
- Copy current query conditions and permission checks from legacy code.
- No new fields, no UI changes.

## SQL DDL (RBAC Tables)
```sql
-- Core users table (new system). Use existing users if already defined.
CREATE TABLE users (
  id           VARCHAR(50) PRIMARY KEY,
  name         NVARCHAR(100) NOT NULL,
  email        VARCHAR(200),
  exp_dept     NVARCHAR(50),
  emp_series   NVARCHAR(50)
);

CREATE TABLE roles (
  id           UUID PRIMARY KEY,
  code         VARCHAR(100) NOT NULL UNIQUE,
  name         NVARCHAR(100) NOT NULL,
  description  NVARCHAR(255)
);

CREATE TABLE permissions (
  id           UUID PRIMARY KEY,
  code         VARCHAR(100) NOT NULL UNIQUE,
  description  NVARCHAR(255)
);

CREATE TABLE user_roles (
  user_id      VARCHAR(50) NOT NULL,
  role_id      UUID NOT NULL,
  PRIMARY KEY (user_id, role_id),
  FOREIGN KEY (user_id) REFERENCES users(id),
  FOREIGN KEY (role_id) REFERENCES roles(id)
);

CREATE TABLE role_permissions (
  role_id       UUID NOT NULL,
  permission_id UUID NOT NULL,
  PRIMARY KEY (role_id, permission_id),
  FOREIGN KEY (role_id) REFERENCES roles(id),
  FOREIGN KEY (permission_id) REFERENCES permissions(id)
);

-- Optional ABAC extension (only if needed later).
CREATE TABLE policy_rules (
  id              UUID PRIMARY KEY,
  resource        VARCHAR(50) NOT NULL,
  action          VARCHAR(50) NOT NULL,
  conditions_json JSONB NOT NULL,
  enabled         BOOLEAN NOT NULL DEFAULT TRUE
);

CREATE INDEX idx_user_roles_user ON user_roles(user_id);
CREATE INDEX idx_user_roles_role ON user_roles(role_id);
CREATE INDEX idx_role_permissions_role ON role_permissions(role_id);
CREATE INDEX idx_role_permissions_perm ON role_permissions(permission_id);
```

## API Spec (Phase 1)

### Response Envelope
Success:
```json
{
  "data": {},
  "traceId": "..."
}
```

Error:
```json
{
  "error": {
    "code": "STRING",
    "message": "STRING",
    "details": {}
  },
  "traceId": "..."
}
```

### Auth / Me
GET /api/v1/users/me
- Response data:
```json
{
  "id": "u123",
  "name": "Alice",
  "email": "a@example.com",
  "expDept": "FD",
  "empSeries": "A",
  "roles": ["ap-approver"],
  "permissions": ["expense_claim:approve", "document:read_all"]
}
```

### Account Settings
PUT /api/v1/users/me
- Request:
```json
{
  "name": "Alice",
  "email": "a@example.com",
  "expDept": "FD",
  "empSeries": "A"
}
```
- Response data: same shape as GET /api/v1/users/me

POST /api/v1/users/me/password
- Request:
```json
{
  "currentPassword": "old",
  "newPassword": "new",
  "confirmPassword": "new"
}
```
- Response data:
```json
{
  "ok": true
}
```
- Behavior: verify current password matches legacy value; update to new password; no hashing changes in Phase 1.

### Roles
GET /api/v1/admin/roles
- Response data:
```json
[
  {"id":"uuid","code":"ap-approver","name":"AP Approver","description":"..."}
]
```

POST /api/v1/admin/roles
- Request:
```json
{
  "code": "ap-approver",
  "name": "AP Approver",
  "description": "..."
}
```
- Response data: role object

PUT /api/v1/admin/roles/{id}
- Request: same as POST
- Response data: role object

DELETE /api/v1/admin/roles/{id}
- Response data:
```json
{
  "ok": true
}
```

### Permissions
GET /api/v1/admin/permissions
- Response data:
```json
[
  {"id":"uuid","code":"expense_claim:approve","description":"..."}
]
```

POST /api/v1/admin/permissions (optional)
- Request:
```json
{
  "code": "document:read_all",
  "description": "..."
}
```
- Response data: permission object

### Role Permissions
GET /api/v1/admin/roles/{id}/permissions
- Response data:
```json
[
  "expense_claim:approve",
  "document:read_all"
]
```

PUT /api/v1/admin/roles/{id}/permissions
- Request:
```json
{
  "permissionCodes": ["expense_claim:approve", "document:read_all"]
}
```
- Response data:
```json
{
  "ok": true
}
```

### User Roles
GET /api/v1/admin/users/{id}/roles
- Response data:
```json
[
  "ap-approver",
  "admin"
]
```

PUT /api/v1/admin/users/{id}/roles
- Request:
```json
{
  "roleCodes": ["ap-approver", "admin"]
}
```
- Response data:
```json
{
  "ok": true
}
```

## Open Items / Next Actions
- Enable file edit/write permissions in the environment (if still disabled).
- Keep this rolling spec updated after each confirmed decision.
- If requested, expand into:
  - Full API schema (requests/responses/errors)
  - UI flow details and page state specs
