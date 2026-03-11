SetErrs Behavior

**Use Case Matrix Describing SetErrs utility actions**
"overridden" means in SetErrs calling routine after SetErrs returns
SetErrs can be called with its CallingFunction argument as:
* "driver" lower case string --for calls from a driver sub routine to trigger initializing errs and reporting messages from driver sub or nested sub-functions
* "non-bool" lower case string --for calls from a non-boolean sub or function in the flow of a nested procedure originally called by a driver sub or by a unit test
* CallingFunction Boolean --for calls from a Boolean function containing error handling SetErrs/RecordErr code structure

| Use Case                     | First `SetErrs` Call Type                                                              |    `errs` Before Call | Default `IsTesting` |                                    Default `IsShowMsgs` | Expected Message Behavior                              |
| ---------------------------- | -------------------------------------------------------------------------------------- | --------------------: | ------------------: | ------------------------------------------------------: | ------------------------------------------------------ |
| Driver normal                | `"driver"`                                                                             |             `Nothing` |             `False` |                                                  `True` | Show messages from called Boolean functions            |
| Driver silent                | `"driver"` then manually set `.IsShowMsgs=False`                                       |             `Nothing` |             `False` | `True` then overridden to `False` after calling SetErrs | Suppress message display from called Boolean functions |
| Default test flow (Boolean)  | Boolean function calls `SetErrs`                                                       |             `Nothing` |              `True` |                                                 `False` | No message display                                     |
| Demo test flow (Boolean)     | Pre-init call (`"non-bool"` or helper), then set `.IsShowMsgs=True`, then Boolean call | `Nothing` on pre-init |              `True` | `False` then overridden to `True` after calling SetErrs | Show messages intentionally                            |
| Default test flow (non-bool) | `"non-bool"`                                                                           |             `Nothing` |              `True` |                                                 `False` | No message display                                     |
| Demo test flow (non-bool)    | `"non-bool"` pre-init then `.IsShowMsgs=True`                                          |             `Nothing` |              `True` | `False` then overridden to `True` after calling SetErrs | Show messages intentionally                            |

**Key principle**

- If `errs` is already initialized, `SetErrs` should not reset `IsShowMsgs`/`IsTesting` unless you explicitly want a hard reset (typically only for a new driver run).
- That preserves your demo workflow where you pre-initialize and then intentionally set `.IsShowMsgs=True`.

**Potential `SetErrs` Logic Flow (no code edits yet)**

1. Determine call type:
    
2. `IsDriverCall = (CallingFunction = "driver")`
    
3. `IsNonBoolCall = (CallingFunction = "non-bool")`
    
4. `IsBoolCall = Not IsDriverCall And Not IsNonBoolCall`
    
5. Initialize Boolean return value:
    
6. If `IsBoolCall`, set `CallingFunction = True`.
    
7. Decide whether to initialize `errs`:
    
8. If `errs Is Nothing`, initialize.
    
9. Optionally: if `IsDriverCall`, force reinitialize for a clean driver context.
    
10. On initialization, set defaults by call type:
    
11. Driver init defaults:
    
12. `IsTesting = False`
    
13. `IsShowMsgs = True` (can be turned off explicitly by driver code)
    
14. Direct call init defaults (Boolean or non-bool):
    
15. `IsTesting = True`
    
16. `IsShowMsgs = False`
    
17. If `errs` already exists and you are not forcing reset:
    
18. Leave `IsTesting` and `IsShowMsgs` unchanged.
    
19. This preserves explicit toggles set by test/demo setup.
    
20. Record/report behavior should follow:
    
21. Driver path: allowed to report, but actual UI display controlled by `IsShowMsgs`.
    
22. Testing path: suppressed by default; only displayed if explicitly forced via `IsShowMsgs=True`.
    

**Why this matches your intent**

- `IsTesting` becomes the mode flag for direct/non-driver invocation.
- `IsShowMsgs` becomes the explicit override knob for both driver and testing/demo scenarios.
- Pre-initialization workflows behave predictably because later Boolean calls do not clobber the flags.

If you want, next I can draft the exact `SetErrs` pseudocode in VBA style (still no file changes) matching this flow one-to-one.

| A (0) | B (1)        | C (2)           | D (3)                                                                | E (4) | F (5)       | G (6)      |
| ----- | ------------ | --------------- | -------------------------------------------------------------------- | ----- | ----------- | ---------- |
| iCode | Module       | Routine         | sMsg_String                                                          | sVal  | iMsgDevUser | VBAProject |
| 100   | modInterface | RefreshRCDriver | Base                                                                 |       | FALSE       | ExcelSteps |
| 104   | modInterface | RefreshRCDriver | VBA error handling is off.  Set Const IsErrorHandle = True to enable |       | TRUE        | ExcelSteps |
| 110   | modInterface | RefreshAPI      | Base                                                                 |       | FALSE       | ExcelSteps |
| 120   | modInterface | Auto_Open       | Base                                                                 |       | FALSE       | ExcelSteps |
| 130   | modInterface | RefreshDriver   | Base                                                                 |       | FALSE       | ExcelSteps |