## Purpose
Add `Dictionary.ParseStringToDictProcedure` (`Dictionary.cls`) as a new public procedure method that parses a JSON-like string into an already-instanced `Dictionary` object (empty or pre-populated).

## Background
The `Dictionary` class in ExcelSteps provides a cross-platform dictionary. This change adds a procedure method to parse JSON-like text into a `Dictionary` (newly initialized or already populated), for example:
```
{"key1":"value1", "key2":"value2"}   value1 and value2 are Strings
{"key1":True, "key2":False}          values are Boolean
{"key1":23.1, "key2":100}           values are numeric
{"name":"Alice", "age":25, "active":True}  mixed types
```

Reference documents:
- [[VBA Project Architecture]] - Standard VBA project structure
- [[Projects/VBA_Development/VBA Project Changes]] - Planning mode process
- Dictionary.cls source code: `Excel_Steps/src/Dictionary.cls`
- Existing tests: `Excel_Steps/tests/tests_Dictionary.bas`

## Data I/O Descriptions

**Input:**
- `dict` (Object) - Dictionary instance (can be empty or pre-populated)
- `jsonStr` (String) - JSON-like string with format: `{"key":"value", "key2":value2}`

**Output:**
- Returns Boolean (`True` success, `False` validation/parsing failure)
- Populates `dict` with parsed key/value pairs
- Preserves existing entries except where keys are overwritten by parsed input

## Project Architecture

**Module Structure:**
- **Dictionary.cls** - Add one new public procedure method plus related public methods in the same class
- No new classes required
- No changes to `Validation.bas` (`New_Dictionary` already exists)

**Procedure Pattern:**
`Dictionary.cls` exposes one public procedure method: `dict.ParseStringToDictProcedure(jsonStr)`.
The procedure method calls public methods in sequence (validate, split, parse, type-detect, add), and requires `dict` to be initialized before the call (empty or pre-populated).

## Test Architecture

**Test Module:** `tests_Dictionary.bas` (existing module)
**Procedure Group:** `procs.dictionary` (existing Procedure attribute in Procedures.cls)

**New Tests to Add:**
- `test_ParseStringToDictProcedure` as the integration test for the full procedure method
- One or more unit tests for each sub-method in the procedure

## Discussion: JSON Syntax Flexibility

**Design Decision: Quote Requirements**
- **Keys**: Support both quoted (`"key"`) and unquoted (`key`) for flexibility
- **String values**: Require quotes to distinguish from Boolean/Numeric
- **Boolean/Numeric**: Unquoted (standard JSON behavior)

**Rationale:** Simplifies user experience while maintaining type safety. Unquoted keys are easier to type; quoted string values prevent ambiguity.

**String Parsing Strategy:**
1. Validate outer braces
2. Strip braces and split by comma while respecting quoted strings
3. Parse each key:value pair and trim whitespace
4. Detect value type (quoted String, Boolean, numeric) or fail validation

## Testing Considerations

**Test Module Structure:** Add tests to existing `tests_Dictionary.bas` module under `procs.dictionary` procedure group

## Procedure Outline

**ParseStringToDictProcedure** - Public Dictionary procedure method to parse JSON-like string and populate Dictionary
* **`ParseStringToDictProcedure`**: [[procPlan_ParseStringToDictProcedure]]
* `ValidateAndStripBraces` - Validate outer braces and remove them; return inner string
* `SplitIntoPairs` - Split string into array of key:value pair strings (handle commas in quoted strings)
* `ParsePair` - Parse single "key:value" string into key and value; detect value type
* `DetectValueType` - Determine if value is String, Boolean, Double or invalid
* `AddParsedValue` - Add parsed key-value to dictionary using dict.Add with IsObject check
