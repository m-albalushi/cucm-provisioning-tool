# CUCM Provisioning Tool

A small desktop tool (Python + Tkinter) that creates a complete new user on
**Cisco Unified Communications Manager (CUCM 15)** from one form, using the
**AXL API**. It replaces about ten manual screens in CUCM Administration.

> Built with AI assistance (Claude). The requirements, the CUCM rules, the
> testing on a live system and every design decision are the author's; the code
> was written together with the assistant and corrected after each real error.
> See [docs/HOW_IT_WAS_BUILT.md](docs/HOW_IT_WAS_BUILT.md).

## What it does

For one user it creates, in this order:

1. **Phone line** (directory number) with partition, calling search space and voice mail profile
2. **Phone**: a copy of a ready-made *template phone* of the chosen model
3. **Device profile line** + **device profile**: a copy of a ready-made *template profile*
4. **End user** with the device profile as controlled profile (Extension Mobility)
5. **Extension Mobility auto login** on the new phone (optional)
6. **Excel phone list**: writes the user's name next to the extension (optional)

If any step fails, everything created so far is removed again, so a retry
always starts from a clean state.

## Features

- One ready-made template **phone and profile per model** (7945, 7975, 8865, ...), set in `config.ini`
- **Checks before creating anything**: phone MAC, directory number (both partitions), profile name, end user
- **Existing phone with the same MAC**: left alone if it is *Registered*, or unregistered for less than 60 days; offered for replacement (with a confirmation) if unregistered for longer
- **Protocol check** (SIP / SCCP) of the templates, and phone/profile protocol consistency (needed by Extension Mobility)
- Copies templates safely: only fields the AXL `add` operations accept, so read-only fields cannot break the request
- Subscribes new phones and profiles to the Extension Mobility phone service, without duplicates
- **First / last name** filled in automatically from the user name (editable)
- **Phone list update** with backup before every change, never overwriting a used number
- Colour-coded activity log and progress bar

## Requirements

- Cisco Unified CM 15 with the AXL web service activated
- Python 3.10+ (on Linux also `python3-tk`)
- `python -m pip install -r requirements.txt`
- The Cisco **AXL toolkit** files `AXLAPI.wsdl`, `AXLSoap.xsd`, `AXLEnums.xsd`
  (CUCM Administration → *Application → Plugins → Cisco AXL Toolkit*; use the files from `schema/current`). They are not included in this repository.

## Project files

| File | Purpose |
|---|---|
| `cucm_phone_creator.py` | the tool (engine + GUI) |
| `make_templates.py` | helper that builds the template phone and profile of a model |
| `config.example.ini` | example settings; copy to `config.ini` |
| `requirements.txt` | Python packages |
| `LICENSE` | MIT license |
| `docs/HOW_IT_WAS_BUILT.md` | development story, problems met and how they were solved |

Local files that are **not** in the repository (see `.gitignore`): `config.ini`,
`phone-list.xlsx`, `phone-list-backups/`, the three AXL toolkit files.

## What you need to provide (and where to get it)

No site-specific value is stored in this repository. Before the first run, collect these
and enter them in your own `config.ini` (never commit that file):

| Setting in `config.ini` | What it is | Where to get it |
|---|---|---|
| `ip` | Address of your CUCM publisher | Your telephony/network team, or the address you use to open CUCM Administration |
| `axl_user`, `axl_password` | Application user with AXL access | Create it yourself in CUCM (see *Setup, step 1*). Choose your own strong password |
| `[extension_mobility] user/password` (optional) | Separate user for the auto login | Same place; leave out to re-use the AXL user |
| `wsdl_file` + `AXLSoap.xsd` + `AXLEnums.xsd` | Cisco AXL toolkit files | Download from your own CUCM: *Application -> Plugins -> Cisco AXL Toolkit* (use `schema/current`) |
| `phone_line_partition`, `profile_line_partition` | Route partitions for the phone line and the device profile line | *Call Routing -> Class of Control -> Partition* |
| `line_css` | Calling search space of the lines | *Call Routing -> Class of Control -> Calling Search Space* |
| `voicemail_profile` | Voice mail profile for new lines | *Advanced Features -> Voice Mail -> Voice Mail Profile* |
| `external_phone_mask` | External number mask of the profile line (optional) | Your dial plan; leave empty if you do not use one |
| `pin` | PIN given to new end users | Your own choice; it must pass your credential policy (*User Management -> Credential Policy*) |
| `presence_group` | Presence group of new users | *System -> Presence Group* (the default is usually `Standard Presence group`) |
| `services` | Name of the Extension Mobility phone service | *Device -> Device Settings -> Phone Services* (exact name) |
| `[phones]` / `[profiles]` | Template phone and device profile of each model | Create them with `make_templates.py` or by hand (see *Setup, step 2*) |
| `phone-list.xlsx` (optional) | Your own Excel list of extensions and names | Your own file; keep it outside the repository |

The tool refuses to start while any `REPLACE_...` example value is still in `config.ini`.

## Setup

### 1. CUCM side (once)

1. *Cisco Unified Serviceability → Tools → Service Activation*: activate **Cisco AXL Web Service** (and keep **Cisco Extension Mobility** running).
2. Create an **access control group** that carries the role **Standard AXL API Access**
   (*User Management → User Settings → Access Control Group → Add New*, then *Related Links → Assign Role to Access Control Group*).
3. Create an **application user** (*User Management → Application User*) and add these groups:
   - the group from step 2 (AXL)
   - `Standard RealtimeAndTraceCollection` (registration status of phones)
   - `Standard EM Authentication Proxy Rights` (Extension Mobility auto login)
4. Under *Edit Credential* tick **Does Not Expire** so the password does not expire silently.

### 2. Templates (once per phone model)

The tool copies one phone and one device profile per model. Create them by hand
(*Device → Phone* and *User Management → User Settings → Device Profile*) or let
`make_templates.py` copy them from an existing phone and profile:

```
python make_templates.py
```

Notes:
- A phone name must be `SEP` + 12 hex characters, so templates use a fake MAC such as `SEP000000007945`.
- Templates carry **no lines**: the tool creates a fresh line for every user.
- The **Device Protocol** (SIP or SCCP) of a template cannot be changed later, and the profile must use the same protocol as the phone.

### 3. Settings

Copy `config.example.ini` to `config.ini` and fill in your values (CUCM IP, AXL user and
password, site defaults, template names). Keep `config.ini` private.

### 4. Run

```
python cucm_phone_creator.py
```

## Using the tool

| Field / option | Meaning |
|---|---|
| Phone model | model from `config.ini`; the "Copies from" line shows which template pair is used |
| MAC address | with or without `:` or `-` |
| User name | becomes the phone description, the profile name and the display name |
| First / last name | filled automatically: first word = first name, the rest = last name (a hyphen keeps one word); editable |
| Directory number | extension; also becomes the end user's User ID |
| Device profile / End user / Log in automatically | what to create; End user needs the profile, auto login needs the end user |
| Update phone-list.xlsx | only runs when phone, profile and end user were all created |

### Existing phone with the same MAC

| Status in CUCM | Result |
|---|---|
| does not exist | continue |
| Registered | stop, nothing is created |
| not registered, last registered within 60 days | stop, nothing is created |
| not registered, older than 60 days | ask for confirmation, delete the old phone, create the new one |
| not registered, date unknown | stop (never deletes when unsure) |

### Phone list

`phone-list.xlsx` (same folder) has the extension in column 1 and the name in column 2 (columns are found by their headers).
Before every update a backup is saved in `phone-list-backups/`. A name is only written into an **empty** cell; if the file is open in Excel,
the extension is missing or duplicated, or the cell is already used, nothing is written and the log tells you what to type by hand.

## How it works

The tool talks only to your CUCM publisher, over HTTPS (port 8443):

- **AXL** (`/axl/`): `addLine`, `getPhone`, `addPhone`, `getDeviceProfile`, `addDeviceProfile`, `addUser`, `listIpPhoneServices`, `resetPhone`, `executeSQLQuery`, and the matching `remove...` operations for the rollback
- **Real-Time Information Service** (`/realtimeservice2/`): registration status of an existing phone
- **Extension Mobility API** (`/emservice/EMServiceServlet`): logs the user into the new phone

Template copies keep only the fields that the WSDL lists for the `add` operation, so
read-only values returned by `get...` never reach `add...`.

## Security notes

- `config.ini` holds the AXL password **in plain text**. Never commit it, and use a dedicated application user with only the groups listed above.
- TLS certificate checking is switched off (`verify=False`) because CUCM normally uses a self-signed certificate. Use it only on a trusted network, or export your CUCM certificate and enable verification.
- The tool never sends data anywhere except to your CUCM.

## Troubleshooting

| Message / symptom | Cause and fix |
|---|---|
| `Unauthorized` / HTTP 401 at the start | wrong user/password, or the user is not in a group that carries *Standard AXL API Access* |
| "not allowed to read phone registration status (RIS)" | add the user to `Standard RealtimeAndTraceCollection` |
| Phone shows "Login unavailable (22)" | *Enable Extension Mobility* is off on the phone (the tool turns it on for new phones) |
| `Phone service "..." was not found` | the service name in `services` must match *Device → Device Settings → Phone Services* exactly; the error lists the real names |
| Auto login fails (everything else was created) | the phone must be plugged in and registered; the user needs `Standard EM Authentication Proxy Rights`; log in manually with User ID + PIN |
| `Missing element presenceGroupName` | set `presence_group` in `[defaults]` |
| PIN rejected | your credential policy blocks simple PINs; change the policy or the `pin` value |
| "template ... uses protocol ..." | the template phone/profile has the wrong protocol; create it again with the right one |
| "Directory number ... already exists" | the number is in use in that partition; delete the leftover or choose another number |
| "Phone list NOT updated: ... open in Excel" | close `phone-list.xlsx` and type the shown value by hand |

## Roadmap (ideas)

- Batch mode from a list (CSV), with preview and dry-run
- Reading onboarding emails and proposing users for review
- Change a user's line CSS
- Choosing a free extension automatically from the phone list and CUCM

## Disclaimer

This tool changes a production telephony system. Test it first with a spare phone and a spare
extension, keep backups, and use it at your own risk. Not affiliated with or endorsed by Cisco.

## Author

Musallam Al Balushi, network and IT infrastructure engineer (Muscat, Oman).
LinkedIn: linkedin.com/in/musallam1albalushi

## License

MIT. See the `LICENSE` file.
