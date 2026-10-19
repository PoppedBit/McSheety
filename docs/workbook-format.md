# Workbook format (formatVersion 1)

This page is for people preparing a SharePoint workbook for McSheety. The fixture `fixtures/equipment/workbook.json` (built with `pnpm fixtures:build`) is a complete example. Check any workbook with `mcsheety validate <file.xlsx>`.

## Data tables

- One worksheet per logical table, holding **one formal Excel Table** (Insert → Table) named `tbl_<WorksheetName>`, e.g. sheet `Equipment` with table `tbl_Equipment`.
- Header names:
  - are case-sensitive, unique and non-blank
  - match `^[A-Za-z][A-Za-z0-9_]*$` (camelCase recommended, no spaces)
  - must not start with `_`
- Every table has an `id` column that is:
  - type `string` or `uuid`
  - `required` and `unique`
  - default `generated`
  - blank `writeRoles`

  Records are always located by `id`, never by row position.

- Headers without a schema row are ignored: never returned, and left blank on create.
- Prefer plain values. McSheety refuses to overwrite cells that contain formulas.

## Metadata tables

Worksheet names start with `_`; table names start with `mc_`.

### `_mcsheety` → `mc_config` (optional)

| key           | value                                |
| ------------- | ------------------------------------ |
| formatVersion | `1` (anything else disables the API) |
| schemaVersion | free text                            |
| policyVersion | free text                            |
| applicationId | free text                            |

### `_schema` → `mc_schema`

Columns: `table | column | type | required | default | unique | options`

- **type:** `string`, `number`, `boolean`, `date` (`YYYY-MM-DD`), `time` (`HH:mm:ss`), `datetime` (ISO 8601 with `Z` or offset, stored as UTC text), `uuid`, `enum`.
- **required / unique:** `TRUE` / `FALSE` (Excel booleans or the text `true`/`false`; blank means false).
- **default:**

  | Value               | Meaning                                                   |
  | ------------------- | --------------------------------------------------------- |
  | blank               | No default                                                |
  | `generated`         | New UUID (string/uuid columns)                            |
  | `now`               | Current UTC time (datetime)                               |
  | `claim:<attribute>` | The caller's trusted attribute, e.g. `claim:departmentId` |
  | anything else       | A literal valid for the type                              |

- **options:**

  | Type   | Syntax                                 |
  | ------ | -------------------------------------- |
  | enum   | `A\|B\|C` (required)                   |
  | number | `min=0`, `max=100`, or `min=0;max=100` |
  | string | `maxLength=200`                        |

### `_permissions` → `mc_permissions`

Columns: `table | role | read | create | update | delete` (booleans).

Default deny. Roles are Entra **app roles** (or roles mapped from immutable group IDs) and are case-sensitive.

### `_columns` → `mc_columns`

Columns: `table | column | readRoles | writeRoles`

- `*` means any role that has the matching table action.
- Otherwise list roles separated by commas.
- Blank means nobody.
- A column with no row here is hidden and read-only (with a warning).

### `_rowPolicies` → `mc_rowPolicies`

Columns: `table | action | role | column | operator | claim`

| operator   | meaning                                                                                              |
| ---------- | ---------------------------------------------------------------------------------------------------- |
| `equals`   | The record's `column` must equal the caller's trusted attribute `claim`. Fails if either is missing. |
| `allowAll` | No row restriction for this role/action. `column` and `claim` must be blank.                         |

Semantics:

- No rule for a role/action means no row restriction beyond table permission.
- Several `equals` rules for the same role/action must **all** pass.
- A user is allowed if **any one** of their roles allows the action on that row:
  - For create, the role must allow the new row.
  - For update, it must allow both the old and the new row.
  - For delete, it must allow the existing row.

## Values

- Blank cells are `null`. Writing `null` or `""` clears the cell.
- Excel date serials are converted for `date`/`time` columns. In `datetime` columns they are flagged, because the timezone is unknown.
- Invalid cells (e.g. text in a number column after a manual edit) are returned as `null`, with an `issues` entry for columns the caller can read.
- User text starting with `=`, `+`, `-`, `@`, tab or a newline is rejected, to prevent formula injection.

## Fail-closed rules

- A malformed metadata header or a missing metadata table disables all tables.
- Any error in a table's metadata (bad type, unknown operator, policy for an unknown column, invalid default, missing Excel table or header) disables **that table** until fixed.
