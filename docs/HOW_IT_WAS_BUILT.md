# How this tool was built

This document records how the tool grew from a manual routine into a small
automation tool, and what went wrong on the way. It is meant as a development log
and as a list of lessons for anyone automating CUCM with AXL.

**Process note.** The tool was built with AI assistance (Claude). The author defined
the task, the CUCM rules and the defaults, ran every version against a live CUCM 15,
reported each error message, and decided how the tool should behave. The assistant wrote and corrected
the code. Each fix below started with a real error from the real system.

## 1. The starting point: a manual routine

For every new employee the author had to:

1. find out whether a phone is needed and its model (7945 or 7975),
2. open an existing user's phone in *Device → Phone*, use **Copy**, change the description,
3. set line 1: directory number, route partition, voice mail profile *No Voice Mail*, CSS *your standard CSS*,
4. subscribe the phone to the Login/Logout phone service,
5. do the same for a **device profile** (Extension Mobility),
6. create the **end user** with User ID = extension, a PIN, and the profile as controlled profile,
7. write the name next to the extension in an Excel list.

Goal: do all of it from one form, with the same defaults every time.

## 2. The idea: AXL

CUCM's **AXL API** can do everything the web pages do. The first version was a small
script that copied a template phone and created a line, then grew step by step.

## 3. Problems met and how each was solved

| # | Symptom (real message) | Cause | Fix |
|---|---|---|---|
| 1 | Template phone could not be named `TEMPLATE-7945` | a phone name must be `SEP` + hex MAC | templates use a fake MAC, e.g. `SEP000000007945`; the description holds the readable name |
| 2 | `Missing element usage (addLine.line.usage)` | CUCM 15 requires `usage` on a new line | send `usage = Device` |
| 3 | `unexpected keyword argument 'certificateStatus'` | `getPhone` returns read-only fields that `addPhone` rejects | read the allowed field names of the `XPhone` type from the WSDL and copy only those |
| 4 | `Invalid value for confidentialAccessMode` | blank/empty values returned by `get` are rejected by `add` | drop `confidentialAccess` and strip empty values recursively |
| 5 | `Service has no operation 'listIpPhoneService'` | AXL operation names are plural | use `listIpPhoneServices` |
| 6 | Service "not found" although it existed | lookup used the wrong request; names are case-sensitive | list all services once and match by name, ignoring case; the error lists the real names |
| 7 | Phone showed **Login unavailable (22)** | *Enable Extension Mobility* was off on the copied phone | the tool sets it on every new phone |
| 8 | `unhashable type: OrderedDict` when copying a real user's phone | service names come back as small objects, not text | normalise the service name before comparing |
| 9 | `unexpected keyword argument 'phoneServiceCategory'` | subscriptions of a real phone carry read-only extras | keep only the six fields a subscription accepts |
| 10 | `Missing element presenceGroupName` | an end user needs a presence group | send the default presence group |
| 11 | A failed run left a line or a profile behind | each step succeeded before a later one failed | a **rollback** list: every created object is removed again, newest first |
| 12 | Hard-coded values in the script | passwords and site names must not live in code | moved to `config.ini`; later the site defaults too (`[defaults]`) |

## 4. Design decisions

- **Check first, create second.** All duplicate checks run before the first object is created.
- **Copy templates, do not invent settings.** A ready-made phone and profile per model carry the device settings; the tool only adds the line and the user data.
- **Fixed template per model**, set in `config.ini`, instead of choosing a real user's phone each time. `make_templates.py` builds clean templates (no lines, speed dials or owner).
- **Never guess when deleting.** An existing phone with the same MAC is replaced only if it is not registered *and* older than 60 days *and* the operator confirms. If the status or date cannot be read, nothing is deleted.
- **Never overwrite a used number** in the Excel list, and always back it up first.
- **A failed optional step must not undo a good result.** If the auto login or the Excel update fails, the user stays created and the log says what to do by hand.
- **Protocol consistency** (SIP / SCCP) is checked before creating anything, because Extension Mobility needs the profile and the phone to match.

## 5. Features added over time

1. Phone + line from a template (the original idea)
2. Device profile, then the end user with its controlled profile
3. Extension Mobility service subscription, without duplicates
4. A GUI (Tkinter) with an activity log and progress bar
5. Per-model templates in `config.ini`; the optional `[protocols]` check (for a SIP-only model)
6. Extension Mobility **auto login** through the EM API (with retries while the phone registers)
7. Handling of an **existing phone** with the same MAC (Real-Time status + last registration date)
8. Automatic **first / last name** from the user name
9. **Excel phone list** update with backups
10. Preparation for publishing: all site-specific values moved out of the code

## 6. Testing approach

- Every version was run against a real CUCM with a spare phone and a spare extension.
- Each error message was pasted back and fixed in the next version.
- Pure logic (name splitting, date parsing, the existing-phone decision table, the protocol check, the Excel update and the rollback order) was also tested against simulated data.
- Not yet covered by automated tests: the GUI itself and calls to a real CUCM. Behaviour that depends on your CUCM version (for example which database column holds the "last registered" date) should be checked on the first run.

## 7. Lessons learned

- AXL `get...` results cannot simply be sent back to `add...`; whitelist the fields from the WSDL.
- Read the exact error text: CUCM usually names the field that is wrong.
- Build the rollback early; it makes every later experiment safe.
- Keep secrets and site names out of the code from the start.
- Ask for confirmation before any delete, and prefer "do nothing" when information is missing.

## 8. Ideas for the future

Batch creation from a CSV with preview and dry-run, reading onboarding emails and proposing users for review,
changing a user's CSS, and picking free extensions automatically.
