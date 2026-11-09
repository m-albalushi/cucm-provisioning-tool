"""
CUCM Provisioning Tool (GUI)
----------------------------
Creates a complete new user on Cisco Unified Communications Manager (CUCM 15)
through the AXL API, from one small form:

  PHONE           : directory number (line) + copy of a ready-made template phone
  DEVICE PROFILE  : directory number (line) + copy of a ready-made template profile
  END USER        : with the device profile as controlled profile (Extension Mobility)
  AUTO LOGIN      : logs the user into the new phone (Extension Mobility API)
  PHONE LIST      : writes the user's name next to the extension in an Excel list

If anything fails, everything created so far is removed again, so a retry
starts clean. An existing phone with the same MAC is replaced only when it is
not registered and has been unused for a long time.

Install once:   python -m pip install -r requirements.txt
Run:            python cucm_phone_creator.py
Files needed in the same folder: config.ini (copy of config.example.ini),
AXLAPI.wsdl, AXLSoap.xsd, AXLEnums.xsd (from the Cisco AXL toolkit).

Connection details, site defaults and the ready-made phone/profile of each
model are NOT in this script. They are read from config.ini, which sits in the
same folder. Keep config.ini private (it contains the AXL password).
"""

import os
import re
import sys
import time
import shutil
import configparser
from datetime import datetime
from xml.sax.saxutils import escape
import tkinter as tk
from tkinter import ttk, messagebox
import urllib3
from requests import Session
from requests.auth import HTTPBasicAuth
from zeep import Client, Settings
from zeep.transports import Transport
from zeep.helpers import serialize_object

urllib3.disable_warnings()

# ------------------------------------------------------------------
# CONNECTION + MODELS: set in config.ini (same folder), not here.
# Filled in by load_config() when the tool starts.
# ------------------------------------------------------------------
CUCM_IP = ""
AXL_USER = ""
AXL_PASS = ""
WSDL_FILE = ""
EM_APP_USER = ""
EM_APP_PASS = ""
TEMPLATES = {}            # model -> ready-made phone (from [phones])
PROFILE_TEMPLATES = {}    # model -> ready-made device profile (from [profiles])
PROTOCOLS = {}            # model -> required protocol, SIP or SCCP (from [protocols], optional)

# ------------------------------------------------------------------
# SITE DEFAULTS: set in config.ini [defaults]. Filled in by load_config().
# ------------------------------------------------------------------
DEFAULT_PARTITION = ""            # route partition of the PHONE line
DEFAULT_PROFILE_PARTITION = ""    # route partition of the DEVICE PROFILE line
DEFAULT_CSS = ""                  # calling search space of the lines
DEFAULT_VOICEMAIL = ""            # voice mail profile of the lines
E164_MASK = ""                    # external phone number mask of the profile line (optional)
DEFAULT_PIN = ""                  # PIN of every new end user
DEFAULT_PRESENCE_GROUP = ""       # presence group of every new end user
DEFAULT_SERVICES = ""             # phone services to subscribe to, comma separated
ENABLE_EXTENSION_MOBILITY = True  # tick "Enable Extension Mobility" on new phones

# ------------------------------------------------------------------
# BEHAVIOUR SETTINGS (safe to edit here)
# ------------------------------------------------------------------

# Existing phone with the same MAC address
OLD_PHONE_DAYS = 60        # not registered for longer than this (about 2 months) = old
CONFIRM_DELETE = True      # ask "delete the old phone?" before deleting (False = no question)

# Phone list (Excel): the user's name is written next to the extension
PHONE_LIST_FILE = "phone-list.xlsx"   # in the same folder as this script
PHONE_LIST_SHEET = ""                 # empty = the first sheet
CLEAR_FILL_WHEN_USED = False          # True = also remove the cell colour (e.g. blue = free)
PHONE_LIST_BACKUPS_KEEP = 50          # newest backups to keep in "phone-list-backups"

# Automatic login (Extension Mobility API). The application user must be allowed
# to use the Extension Mobility API (e.g. group "Standard EM Authentication Proxy
# Rights"). It re-uses the AXL user unless [extension_mobility] is set in config.ini.
EM_LOGIN_TRIES = 6          # how many times to try (the phone must be registered)
EM_LOGIN_WAIT = 15          # seconds between tries

# ----------------------------- config -----------------------------

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "config.ini")

CONFIG_TEMPLATE = """; CUCM Provisioning Tool - settings
; Copy this file to config.ini and replace every value marked REPLACE.
; Keep config.ini private: it contains the AXL password. Never commit it to git.
; (Where to find each value: see "What you need to provide" in README.md)

[cucm]
; REPLACE: address of your CUCM publisher (IP or hostname)
ip = 192.0.2.10
; REPLACE: the AXL application user you create in CUCM
axl_user = REPLACE_AXL_USER
; REPLACE: that user's password (never share or commit it)
axl_password = REPLACE_AXL_PASSWORD
; Cisco AXL toolkit file, downloaded from your own CUCM
wsdl_file = AXLAPI.wsdl

; Optional: a separate application user for the automatic login.
; Leave these two lines out to re-use the AXL user above.
[extension_mobility]
; user = REPLACE_EM_USER
; password = REPLACE_EM_PASSWORD

; Site defaults applied to every new user (names must match CUCM exactly)
[defaults]
; REPLACE: route partition for the phone line
phone_line_partition = REPLACE_PHONE_LINE_PARTITION
; REPLACE: route partition for the device profile line
profile_line_partition = REPLACE_PROFILE_LINE_PARTITION
; REPLACE: calling search space for the lines
line_css = REPLACE_CSS_NAME
; REPLACE: voice mail profile name (use the profile that means "no voicemail" if you want none)
voicemail_profile = REPLACE_VOICEMAIL_PROFILE
; External Phone Number Mask of the profile line, e.g. 55501XXX. Leave empty for none.
external_phone_mask =
; REPLACE: the PIN given to every new end user. It must pass your credential policy.
pin = REPLACE_PIN
; Name of an existing presence group
presence_group = Standard Presence group
; Name of the Extension Mobility phone service (exact name from CUCM)
services = REPLACE_EM_SERVICE_NAME
enable_extension_mobility = true

; ONE ready-made PHONE per model (device name = SEP + 12 hex characters)
; To add a model: add a line here AND a line with the same model in [profiles].
[phones]
7945 = SEP000000007945
7975 = SEP000000007975

; ONE ready-made DEVICE PROFILE per model (exact profile name)
[profiles]
7945 = TEMPLATE-7945
7975 = TEMPLATE-7975

; Optional: the protocol (SIP or SCCP) that each model's templates must use.
; The tool stops BEFORE creating anything if a template does not match.
; Example:  8865 = SIP
[protocols]
"""


class ConfigError(Exception):
    pass


def load_config():
    """Read config.ini (created as a template on first run)."""
    global CUCM_IP, AXL_USER, AXL_PASS, WSDL_FILE, EM_APP_USER, EM_APP_PASS
    global DEFAULT_PARTITION, DEFAULT_PROFILE_PARTITION, DEFAULT_CSS, DEFAULT_VOICEMAIL
    global E164_MASK, DEFAULT_PIN, DEFAULT_PRESENCE_GROUP, DEFAULT_SERVICES
    global ENABLE_EXTENSION_MOBILITY

    if not os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            f.write(CONFIG_TEMPLATE)
        raise ConfigError(
            "config.ini was not found, so a template was created here:\n\n"
            f"{CONFIG_FILE}\n\nOpen it, fill in your values, save it, "
            "and start the tool again.")

    cp = configparser.ConfigParser(interpolation=None)
    cp.optionxform = str            # keep capital letters in names
    try:
        cp.read(CONFIG_FILE, encoding="utf-8")
    except configparser.Error as e:
        raise ConfigError(f"config.ini could not be read:\n{e}")

    def need(section, key):
        v = cp.get(section, key, fallback="").strip()
        if not v:
            raise ConfigError(f'config.ini: "{key}" is missing in [{section}].')
        return v

    CUCM_IP = need("cucm", "ip")
    AXL_USER = need("cucm", "axl_user")
    AXL_PASS = need("cucm", "axl_password")
    wsdl = cp.get("cucm", "wsdl_file", fallback="").strip() or "AXLAPI.wsdl"
    WSDL_FILE = wsdl if os.path.isabs(wsdl) else os.path.join(BASE_DIR, wsdl)

    EM_APP_USER = cp.get("extension_mobility", "user", fallback="").strip() or AXL_USER
    EM_APP_PASS = cp.get("extension_mobility", "password", fallback="").strip() or AXL_PASS

    phones = dict(cp.items("phones")) if cp.has_section("phones") else {}
    profiles = dict(cp.items("profiles")) if cp.has_section("profiles") else {}
    if not phones:
        raise ConfigError("config.ini: the [phones] section is empty.")
    for m in phones:
        if m not in profiles:
            raise ConfigError(f'config.ini: model {m} is in [phones] but has no '
                              f'line in [profiles].')
    for m in profiles:
        if m not in phones:
            raise ConfigError(f'config.ini: model {m} is in [profiles] but has no '
                              f'line in [phones].')
    if not cp.has_section("defaults"):
        raise ConfigError(
            "config.ini has no [defaults] section. Copy it from config.example.ini and "
            "set your values: phone_line_partition, profile_line_partition, line_css, "
            "voicemail_profile, external_phone_mask, pin, presence_group, services.")

    def dflt(key, required=True):
        v = cp.get("defaults", key, fallback="").strip()
        if required and not v:
            raise ConfigError(f'config.ini: "{key}" is missing in [defaults].')
        return v

    DEFAULT_PARTITION = dflt("phone_line_partition")
    DEFAULT_PROFILE_PARTITION = dflt("profile_line_partition")
    DEFAULT_CSS = dflt("line_css")
    DEFAULT_VOICEMAIL = dflt("voicemail_profile")
    E164_MASK = dflt("external_phone_mask", required=False)
    DEFAULT_PIN = dflt("pin")
    DEFAULT_PRESENCE_GROUP = dflt("presence_group")
    DEFAULT_SERVICES = dflt("services")
    ENABLE_EXTENSION_MOBILITY = cp.getboolean("defaults", "enable_extension_mobility",
                                              fallback=True)

    # Stop early if example placeholders were left in config.ini
    left = [name for name, val in (
        ("ip / axl_user / axl_password", CUCM_IP + AXL_USER + AXL_PASS),
        ("[defaults]", DEFAULT_PARTITION + DEFAULT_PROFILE_PARTITION + DEFAULT_CSS
         + DEFAULT_VOICEMAIL + DEFAULT_PIN + DEFAULT_SERVICES)) if "REPLACE_" in val]
    if left:
        raise ConfigError("config.ini still contains example values (REPLACE_...) in: "
                          + ", ".join(left) + ".\nOpen it and fill in your own values.")
    if CUCM_IP == "192.0.2.10":
        raise ConfigError("config.ini still has the example CUCM address (192.0.2.10). "
                          "Enter the address of your CUCM publisher.")

    protocols = dict(cp.items("protocols")) if cp.has_section("protocols") else {}
    for m, proto in protocols.items():
        if m not in phones:
            raise ConfigError(f'config.ini: model {m} is in [protocols] but not in [phones].')
        if proto.strip().upper() not in ("SIP", "SCCP"):
            raise ConfigError(f'config.ini: protocol for model {m} must be SIP or SCCP, '
                              f'not "{proto}".')
    PROTOCOLS.clear()
    PROTOCOLS.update({m: v.strip().upper() for m, v in protocols.items()})
    TEMPLATES.clear()
    TEMPLATES.update(phones)
    PROFILE_TEMPLATES.clear()
    PROFILE_TEMPLATES.update(profiles)


CLIENT = None
_SERVICE_CATALOG = None


def get_axl():
    global CLIENT, _SERVICE_CATALOG
    _SERVICE_CATALOG = None
    session = Session()
    session.verify = False
    session.auth = HTTPBasicAuth(AXL_USER, AXL_PASS)
    CLIENT = Client(
        WSDL_FILE,
        settings=Settings(strict=False, xml_huge_tree=True),
        transport=Transport(session=session, timeout=30),
    )
    return CLIENT.create_service(
        "{http://www.cisco.com/AXLAPIService/}AXLAPIBinding",
        f"https://{CUCM_IP}:8443/axl/",
    )


# ---------------------------- helpers -----------------------------

def clean_mac(mac):
    mac = mac.replace(":", "").replace("-", "").replace(".", "").upper().strip()
    if len(mac) != 12 or any(c not in "0123456789ABCDEF" for c in mac):
        raise ValueError("MAC address must be 12 hex characters, e.g. 001122334455")
    return mac


def strip_empty(value):
    """Recursively remove None and empty-string values (CUCM rejects some blanks)."""
    if isinstance(value, dict):
        out = {}
        for k, v in value.items():
            v = strip_empty(v)
            if v is None or v == "":
                continue
            out[k] = v
        return out
    if isinstance(value, list):
        return [strip_empty(v) for v in value]
    return value


def allowed_fields(typename):
    """Field names the add operation accepts for a type (XPhone, XDeviceProfile)."""
    for t in CLIENT.wsdl.types.types:
        if t.qname and t.qname.localname == typename:
            return {name for name, _ in t.elements}
    raise RuntimeError(f"Could not find {typename} type in the WSDL.")


def copy_from_template(template_obj, typename):
    """Keep only fields the add operation accepts (drops read-only ones)."""
    allowed = allowed_fields(typename)
    obj = {k: v for k, v in template_obj.items() if k in allowed}
    for bad in ("versionStamp", "ctiid", "uuid", "confidentialAccess",
                "ownerUserName", "mobilityUserIdName"):
        obj.pop(bad, None)
    return strip_empty(obj)


def service_catalog(axl):
    global _SERVICE_CATALOG
    if _SERVICE_CATALOG is None:
        try:
            r = serialize_object(axl.listIpPhoneServices(
                searchCriteria={"serviceName": "%"},
                returnedTags={"serviceName": "", "asciiServiceName": "",
                              "serviceUrl": ""}))
        except Exception as e:
            raise RuntimeError(f"Could not read phone services from CUCM: {e}")
        ret = r.get("return") or {}
        _SERVICE_CATALOG = ret.get("ipPhoneServices") or ret.get("ipPhoneService") or []
    return _SERVICE_CATALOG


def build_service_entry(axl, svc_name):
    items = service_catalog(axl)
    found = None
    for s in items:
        if (s.get("serviceName") or "").strip().lower() == svc_name.strip().lower():
            found = s
            break
    if not found:
        names = ", ".join(i.get("serviceName") or "?" for i in items)
        raise RuntimeError(
            f'Phone service "{svc_name}" was not found. '
            f'Services in CUCM: {names or "(none returned)"}')
    name = found["serviceName"]
    return strip_empty({
        "telecasterServiceName": name,
        "name": name,
        "url": found.get("serviceUrl"),
        "urlButtonIndex": 0,
        "urlLabel": name,
        "serviceNameAscii": found.get("asciiServiceName") or name,
    })


def svc_key(entry):
    """Service name of a subscription entry (CUCM may return it as a dict)."""
    v = entry.get("telecasterServiceName")
    if isinstance(v, dict):
        v = v.get("_value_1") or v.get("name") or v.get("uuid")
    return str(v).strip().lower() if v is not None else ""


SERVICE_FIELDS = ("telecasterServiceName", "name", "url", "urlButtonIndex",
                  "urlLabel", "serviceNameAscii")


def clean_service(entry):
    """Keep only the fields a subscription accepts (a copied real phone also
    carries read-only extras such as phoneServiceCategory)."""
    out = {k: entry[k] for k in SERVICE_FIELDS if k in entry}
    v = out.get("telecasterServiceName")
    if isinstance(v, dict):                     # plain name instead of the dict
        out["telecasterServiceName"] = v.get("_value_1") or v.get("name")
    return strip_empty(out)


def missing_service_entries(axl, template_obj, services, log, what):
    have = {svc_key(x)
            for x in (template_obj.get("services") or {}).get("service") or []}
    entries = []
    for svc in services:
        if svc.strip().lower() in have:
            log(f"Service {svc} already comes from the {what} template.")
            continue
        log(f"Checking phone service {svc} ...")
        entries.append(build_service_entry(axl, svc))
    return entries


def merge_services(obj, entries):
    """Existing + new subscriptions, one per service (no duplicates)."""
    current, seen = [], set()
    for e in list((obj.get("services") or {}).get("service") or []) + entries:
        key = svc_key(e)
        if key in seen:
            continue
        seen.add(key)
        current.append(clean_service(e))
    if current:
        obj["services"] = {"service": current}


def phone_device_name(text):
    """Accept SEPxxxxxxxxxxxx or a plain MAC; return the CUCM device name."""
    t = text.strip()
    if t.upper().startswith("SEP"):
        return t.upper()
    try:
        return "SEP" + clean_mac(t)
    except ValueError:
        return t   # some other device name, use as typed


def split_name(full):
    """First word = first name, everything after it = last name.
    'john doe' -> ('john', 'doe')
    'abdul-rahman al rahbi' -> ('abdul-rahman', 'al rahbi')   (a hyphen keeps one word)
    'John Al Smith' -> ('John', 'Al Smith')"""
    parts = full.strip().split()
    if len(parts) <= 1:
        return "", full.strip()
    return parts[0], " ".join(parts[1:])


def user_exists(axl, userid):
    try:
        axl.getUser(userid=userid)
        return True
    except Exception:
        return False


def phone_exists(axl, name):
    try:
        axl.getPhone(name=name)
        return True
    except Exception:
        return False


def profile_exists(axl, name):
    try:
        axl.getDeviceProfile(name=name)
        return True
    except Exception:
        return False


def line_exists(axl, dn, partition):
    r = serialize_object(axl.listLine(
        searchCriteria={"pattern": dn, "routePartitionName": partition},
        returnedTags={"pattern": ""}))
    return bool(r["return"])


def add_line(axl, dn, partition, description, css, voicemail):
    axl.addLine(line={
        "pattern": dn,
        "usage": "Device",
        "routePartitionName": partition,
        "description": description,
        "alertingName": description,
        "asciiAlertingName": description,
        "shareLineAppearanceCssName": css,
        "voiceMailProfileName": voicemail,
    })


def rollback(axl, done, log):
    """Remove everything created so far, newest first."""
    for kind, *args in reversed(done):
        try:
            if kind == "phone":
                axl.removePhone(name=args[0])
            elif kind == "profile":
                axl.removeDeviceProfile(name=args[0])
            elif kind == "user":
                axl.removeUser(userid=args[0])
            elif kind == "line":
                axl.removeLine(pattern=args[0], routePartitionName=args[1])
            log(f"Rolled back {kind} {' / '.join(args)}")
        except Exception as e:
            log(f"Could not remove {kind} {' / '.join(args)} (delete manually): {e}")


def wait(seconds, pump):
    """Sleep without freezing the window."""
    end = time.time() + seconds
    while time.time() < end:
        if pump:
            pump()
        time.sleep(0.2)


def em_login(device, userid, profile, log, pump):
    """Log a user into a phone with the Extension Mobility API."""
    url = f"https://{CUCM_IP}:8443/emservice/EMServiceServlet"
    xml = (
        "<request><appInfo>"
        f"<appID>{escape(EM_APP_USER)}</appID>"
        f"<appCertificate>{escape(EM_APP_PASS)}</appCertificate>"
        "</appInfo><login>"
        f"<deviceName>{escape(device)}</deviceName>"
        f"<userID>{escape(userid)}</userID>"
        f"<deviceProfile>{escape(profile)}</deviceProfile>"
        "</login></request>"
    )
    last = ""
    for attempt in range(1, EM_LOGIN_TRIES + 1):
        try:
            r = Session()
            r.verify = False
            resp = r.post(url, data={"xml": xml}, timeout=30)
            text = resp.text.strip()
            if "<success" in text.lower():
                return True, "ok"
            last = text
        except Exception as e:
            last = str(e)
        log(f"Auto login try {attempt}/{EM_LOGIN_TRIES} not ready yet ...")
        if attempt < EM_LOGIN_TRIES:
            wait(EM_LOGIN_WAIT, pump)
    return False, last


# ------------------- existing phone with the same MAC -------------------

def ris_status(device):
    """Registration status(es) of a phone from the Real-Time service (RIS)."""
    url = f"https://{CUCM_IP}:8443/realtimeservice2/services/RISService70"
    body = (
        '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" '
        'xmlns:soap="http://schemas.cisco.com/ast/soap">'
        "<soapenv:Header/><soapenv:Body><soap:selectCmDevice>"
        "<soap:StateInfo></soap:StateInfo><soap:CmSelectionCriteria>"
        "<soap:MaxReturnedDevices>10</soap:MaxReturnedDevices>"
        "<soap:DeviceClass>Phone</soap:DeviceClass>"
        "<soap:Model>255</soap:Model>"
        "<soap:Status>Any</soap:Status>"
        "<soap:NodeName></soap:NodeName>"
        "<soap:SelectBy>Name</soap:SelectBy>"
        f"<soap:SelectItems><soap:item><soap:Item>{escape(device)}</soap:Item>"
        "</soap:item></soap:SelectItems>"
        "<soap:Protocol>Any</soap:Protocol>"
        "<soap:DownloadStatus>Any</soap:DownloadStatus>"
        "</soap:CmSelectionCriteria></soap:selectCmDevice></soapenv:Body>"
        "</soapenv:Envelope>"
    )
    headers = {
        "Content-Type": "text/xml; charset=utf-8",
        "SOAPAction": '"http://schemas.cisco.com/ast/soap/action/#RisPort#SelectCmDevice"',
    }
    sess = Session()
    sess.verify = False
    sess.auth = HTTPBasicAuth(AXL_USER, AXL_PASS)
    try:
        resp = sess.post(url, data=body.encode("utf-8"), headers=headers, timeout=30)
    except Exception as e:
        raise RuntimeError(
            f"Could not reach the CUCM Real-Time service to check if {device} is "
            f"registered. The old phone was left alone. ({e})")
    if resp.status_code in (401, 403):
        raise RuntimeError(
            "The CUCM user is not allowed to read phone registration status (RIS). "
            "Add it to the user group 'Standard RealtimeAndTraceCollection'. "
            "The old phone was left alone.")
    if resp.status_code != 200 or "faultstring" in resp.text.lower():
        raise RuntimeError(
            f"The Real-Time service answered with an error (HTTP {resp.status_code}). "
            f"The old phone was left alone. {resp.text[:300]}")
    return re.findall(r"<(?:\w+:)?Status>\s*([^<]+?)\s*</", resp.text)


def parse_when(value):
    """Turn a CUCM date value (epoch seconds/ms or a date text) into a datetime."""
    if value is None:
        return None
    t = str(value).strip()
    if not t or t.lower() in ("none", "null"):
        return None
    try:
        n = float(t)
        if n <= 0:
            return None
        if n > 1e11:                 # milliseconds
            n /= 1000.0
        return datetime.fromtimestamp(n)
    except ValueError:
        pass
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%b %d, %Y %I:%M:%S %p"):
        try:
            return datetime.strptime(t, fmt)
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(t)
    except ValueError:
        return None


def row_to_dict(row):
    if isinstance(row, dict):
        return {str(k).lower(): v for k, v in row.items()}
    out = {}
    for child in row:                # lxml element from executeSQLQuery
        out[str(child.tag).split("}")[-1].lower()] = child.text
    return out


def last_registered(axl, phone_name):
    """(datetime or None, text with the raw values) for a phone, read from CUCM."""
    sql = ("select * from registrationdynamic where fkdevice = "
           f"(select pkid from device where name = '{phone_name}')")
    try:
        r = axl.executeSQLQuery(sql=sql)
        ret = r["return"]
        rows = (ret["row"] if ret is not None else None) or []
    except Exception as e:
        return None, f"could not read the date from CUCM: {e}"
    if not rows:
        return None, "CUCM has no registration record for this phone"
    d = row_to_dict(rows[0])
    found = {k: d.get(k) for k in ("lastseen", "lastactive", "datetimestamp")
             if k in d}
    dates = [w for w in (parse_when(v) for v in found.values()) if w]
    raw = ", ".join(f"{k}={v}" for k, v in found.items()) or "no date columns"
    # the most recent date is the safe one (a phone is only deleted when ALL are old)
    return (max(dates) if dates else None), raw


def check_existing_phone(axl, phone_name, log):
    """None  = no phone with this MAC, carry on.
    dict  = old unregistered phone that may be deleted.
    Raises = phone must be left alone (registered, recent, or unknown)."""
    if not phone_exists(axl, phone_name):
        return None
    statuses = ris_status(phone_name)
    log(f"Phone {phone_name} already exists. Status: "
        f"{', '.join(statuses) if statuses else 'None (not registered)'}")
    if any(x in ("Registered", "PartiallyRegistered") for x in statuses):
        raise RuntimeError(
            f"Phone {phone_name} already exists and is Registered Now, so it was "
            "left alone. Nothing was created.")
    when, raw = last_registered(axl, phone_name)
    log(f"Registration dates from CUCM: {raw}")
    if when is None:
        raise RuntimeError(
            f"Phone {phone_name} already exists and is not registered, but its Last "
            f"Registered date could not be read ({raw}), so it was left alone. "
            "Nothing was created.")
    age = (datetime.now() - when).days
    log(f"Last registered {when:%b %d, %Y %I:%M %p} ({age} days ago).")
    if age <= OLD_PHONE_DAYS:
        raise RuntimeError(
            f"Phone {phone_name} already exists, is not registered, and was last "
            f"registered {age} days ago (less than {OLD_PHONE_DAYS} days), so it was "
            "left alone. Nothing was created.")
    return {"when": when, "age": age}


def check_protocols(model, src_phone, phone_tmpl, src_profile, prof_tmpl, log):
    """Make sure the templates use the right protocol (SIP / SCCP).
    Raises before anything is created if not."""
    phone_proto = str(phone_tmpl.get("protocol") or "").strip().upper()
    expected = PROTOCOLS.get(model, "")
    if expected and phone_proto != expected:
        raise RuntimeError(
            f"The template phone {src_phone} for model {model} uses protocol "
            f"{phone_proto or 'unknown'}, but config.ini requires {expected}. "
            "Nothing was created.")
    if prof_tmpl is not None:
        prof_proto = str(prof_tmpl.get("protocol") or "").strip().upper()
        if prof_proto and phone_proto and prof_proto != phone_proto:
            raise RuntimeError(
                f'The template profile "{src_profile}" uses {prof_proto} but the '
                f"template phone {src_phone} uses {phone_proto}. Extension Mobility "
                "needs the same protocol on both. Nothing was created.")
    log(f"Protocol of the templates: {phone_proto or 'unknown'}"
        + (f" (required {expected})" if expected else ""))


# ---------------------- phone list (Excel) update ----------------------

def _dn_text(value):
    """Cell value -> extension text (3933, 3933.0 and ' 3933 ' all become '3933')."""
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return str(int(value))
    return str(value).strip()


def _find_list_columns(ws):
    """(header row, extension column, name column), found by header words."""
    for r in range(1, 6):
        dn_col = name_col = None
        for c in range(1, ws.max_column + 1):
            v = str(ws.cell(r, c).value or "").lower()
            if dn_col is None and any(k in v for k in ("directory", "extension")):
                dn_col = c
            elif name_col is None and any(k in v for k in ("profile", "name", "description")):
                name_col = c
        if dn_col and name_col:
            return r, dn_col, name_col
    return 1, 1, 2        # fallback: columns A and B, data from row 2


def update_phone_list(dn, name, log):
    """Type the user's name next to the extension in phone-list.xlsx.
    Backs the file up first, never overwrites a name, never raises.
    Returns a short note for the final message."""
    path = PHONE_LIST_FILE if os.path.isabs(PHONE_LIST_FILE) \
        else os.path.join(BASE_DIR, PHONE_LIST_FILE)
    todo = f"Please type it by hand:  {dn}  ->  {name}"
    try:
        import openpyxl
    except ImportError:
        log("Phone list NOT updated: the 'openpyxl' package is missing "
            "(run: python -m pip install openpyxl). " + todo)
        return " Phone list NOT updated."
    if not os.path.exists(path):
        log(f"Phone list NOT updated: {path} was not found. " + todo)
        return " Phone list NOT updated."

    # 1) backup first - without a backup nothing is written
    try:
        bdir = os.path.join(os.path.dirname(path), "phone-list-backups")
        os.makedirs(bdir, exist_ok=True)
        base, ext = os.path.splitext(os.path.basename(path))
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]      # to the millisecond
        backup = os.path.join(bdir, f"{base}_{stamp}{ext}")
        shutil.copy2(path, backup)
        old = sorted(f for f in os.listdir(bdir)
                     if f.startswith(base + "_") and f.endswith(ext))
        for f in old[:-PHONE_LIST_BACKUPS_KEEP]:
            os.remove(os.path.join(bdir, f))
    except Exception as e:
        log(f"Phone list NOT updated: could not make a backup first ({e}). " + todo)
        return " Phone list NOT updated."

    # 2) write the name
    try:
        wb = openpyxl.load_workbook(path)
        ws = wb[PHONE_LIST_SHEET] if PHONE_LIST_SHEET else wb.worksheets[0]
        head, dn_col, name_col = _find_list_columns(ws)
        rows = [r for r in range(head + 1, ws.max_row + 1)
                if _dn_text(ws.cell(r, dn_col).value) == str(dn).strip()]
        if not rows:
            log(f"Phone list NOT updated: extension {dn} is not in the list. " + todo)
            return " Phone list NOT updated."
        if len(rows) > 1:
            log(f"Phone list NOT updated: extension {dn} appears {len(rows)} times "
                "in the list. " + todo)
            return " Phone list NOT updated."
        cell = ws.cell(rows[0], name_col)
        current = str(cell.value).strip() if cell.value is not None else ""
        if current:
            if current.lower() == name.strip().lower():
                log(f"Phone list: {dn} already has this name. Nothing changed.")
                return " Phone list already had the name."
            log(f'Phone list NOT updated: {dn} already belongs to "{current}" in the '
                "list, so it was not overwritten. " + todo)
            return " Phone list NOT updated."
        cell.value = name
        if CLEAR_FILL_WHEN_USED:
            from openpyxl.styles import PatternFill
            cell.fill = PatternFill(fill_type=None)
        wb.save(path)
    except PermissionError:
        log("Phone list NOT updated: the file is open in Excel or locked. Close it "
            "and type the name by hand (or run again). " + todo)
        return " Phone list NOT updated (file is open)."
    except Exception as e:
        log(f"Phone list NOT updated: {e}. " + todo)
        return " Phone list NOT updated."

    log(f"Phone list updated: {dn} -> {name}  (backup saved in phone-list-backups)")
    return " Phone list updated."


# ----------------------------- main -------------------------------

def create_user(axl, model, mac, description, dn, partition, css, voicemail,
                services, make_profile, profile_partition, log,
                make_user=False, auto_login=False, pump=None, confirm=None,
                first_name=None, last_name=None, update_list=False):
    phone_name = "SEP" + clean_mac(mac)
    # Each model has its own fixed phone and profile to copy from
    src_phone = TEMPLATES[model]
    src_profile = PROFILE_TEMPLATES[model]

    # ---- checks first: nothing is created if any of these fail ----
    stale = check_existing_phone(axl, phone_name, log)   # None, or old phone to replace
    if line_exists(axl, dn, partition):
        raise RuntimeError(f"Directory number {dn} already exists in {partition}.")
    if auto_login and not make_user:
        raise RuntimeError("Auto login needs 'Create end user' turned on.")
    if make_user and not make_profile:
        raise RuntimeError("Creating the end user needs 'Create device profile' turned on.")
    if make_user and user_exists(axl, dn):
        raise RuntimeError(f"End user {dn} already exists.")
    if make_profile:
        if profile_exists(axl, description):
            raise RuntimeError(f'Device profile "{description}" already exists.')
        if line_exists(axl, dn, profile_partition):
            raise RuntimeError(
                f"Directory number {dn} already exists in {profile_partition}.")

    log(f"Reading phone to copy: {src_phone} ...")
    phone_tmpl = serialize_object(
        axl.getPhone(name=src_phone)["return"]["phone"])
    phone_svc = missing_service_entries(axl, phone_tmpl, services, log, "phone")

    if make_profile:
        log(f"Reading profile to copy: {src_profile} ...")
        prof_tmpl = serialize_object(
            axl.getDeviceProfile(name=src_profile)["return"]["deviceProfile"])
        prof_svc = missing_service_entries(axl, prof_tmpl, services, log, "profile")

    check_protocols(model, src_phone, phone_tmpl, src_profile,
                    prof_tmpl if make_profile else None, log)

    # ---- old unregistered phone with the same MAC: replace it ----
    if stale:
        question = (f"Phone {phone_name} already exists, is NOT registered, and was "
                    f"last registered {stale['when']:%b %d, %Y} ({stale['age']} days "
                    "ago).\n\nDelete the old phone and create the new one?")
        if CONFIRM_DELETE and confirm and not confirm(question):
            raise RuntimeError("Cancelled: the old phone was kept. Nothing was created.")
        log(f"Deleting old phone {phone_name} ...")
        axl.removePhone(name=phone_name)

    # ---- create, with rollback on any error ----
    done = []
    try:
        # PHONE
        log(f"Creating phone line {dn} in {partition} ...")
        add_line(axl, dn, partition, description, css, voicemail)
        done.append(("line", dn, partition))

        phone = copy_from_template(phone_tmpl, "XPhone")
        phone["name"] = phone_name
        phone["description"] = description
        phone["enableExtensionMobility"] = ENABLE_EXTENSION_MOBILITY
        phone["lines"] = {"line": [{
            "index": 1,
            "label": description,
            "display": description,
            "displayAscii": description,
            "dirn": {"pattern": dn, "routePartitionName": partition},
        }]}
        merge_services(phone, phone_svc)
        log(f"Creating phone {phone_name} ...")
        axl.addPhone(phone=phone)
        done.append(("phone", phone_name))

        # DEVICE PROFILE
        if make_profile:
            log(f"Creating profile line {dn} in {profile_partition} ...")
            add_line(axl, dn, profile_partition, description, css, voicemail)
            done.append(("line", dn, profile_partition))

            prof = copy_from_template(prof_tmpl, "XDeviceProfile")
            prof["name"] = description
            prof["description"] = description
            prof_line = {
                "index": 1,
                "label": description,          # Line Text Label
                "display": description,        # Display (Caller ID)
                "displayAscii": description,   # ASCII Display (Caller ID)
                "dirn": {"pattern": dn, "routePartitionName": profile_partition},
            }
            if E164_MASK:
                prof_line["e164Mask"] = E164_MASK   # External Phone Number Mask
            prof["lines"] = {"line": [prof_line]}
            merge_services(prof, prof_svc)
            log(f'Creating device profile "{description}" ...')
            axl.addDeviceProfile(deviceProfile=prof)
            done.append(("profile", description))

        # END USER (controls the device profile for Extension Mobility)
        if make_user:
            first, last = split_name(description)
            if last_name and last_name.strip():          # names typed/edited in the form
                first, last = (first_name or "").strip(), last_name.strip()
            user = {
                "userid": dn,                  # User ID = directory number
                "selfService": dn,             # Self-Service User ID
                "pin": DEFAULT_PIN,
                "firstName": first,
                "lastName": last,
                "displayName": description,
                "telephoneNumber": dn,
                "presenceGroupName": DEFAULT_PRESENCE_GROUP,
                "phoneProfiles": {"profileName": [description]},  # controlled profile
            }
            ok = allowed_fields("XUser")
            dropped = [k for k in user if k not in ok]
            if dropped:
                log(f"Note: CUCM does not accept these user fields, skipped: {dropped}")
            user = strip_empty({k: v for k, v in user.items() if k in ok})
            log(f"Creating end user {dn} ({description}) ...")
            axl.addUser(user=user)
            done.append(("user", dn))
    except Exception:
        log("Failed, removing what was created ...")
        rollback(axl, done, log)
        raise

    # everything is created: write the name into the Excel phone list
    list_note = ""
    if update_list:
        if make_profile and make_user:
            list_note = update_phone_list(dn, description, log)
        else:
            log("Phone list not updated (it is only updated when the phone, the "
                "device profile and the end user were all created).")
            list_note = " Phone list not updated."

    login_note = ""
    if auto_login:
        try:
            axl.resetPhone(name=phone_name)
            log("Phone reset sent, waiting for it to register ...")
        except Exception as e:
            log(f"Could not reset the phone (reset it manually): {e}")
        ok, detail = em_login(phone_name, dn, description, log, pump)
        if ok:
            log(f"Logged {dn} into {phone_name}.")
            login_note = " Logged in on the phone."
        else:
            log(f"AUTO LOGIN FAILED (everything else was created): {detail}")
            login_note = " Auto login FAILED - see the log."

    log(f"DONE: {phone_name} | {dn} | {description}"
        + (" | + device profile" if make_profile else "")
        + (" | + end user" if make_user else ""))
    return list_note + login_note


# ------------------------------ GUI -------------------------------

BG, CARD, NAVY = "#EEF1F6", "#FFFFFF", "#0F1B3D"
ACCENT, ACCENT_DARK = "#2563EB", "#1D4ED8"
TEXT, MUTED, BORDER = "#1F2937", "#6B7280", "#D5DBE6"
C_OK, C_WARN, C_BAD, C_INFO = "#4ADE80", "#FBBF24", "#F87171", "#E2E8F0"
FONT = "Segoe UI"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CUCM Provisioning Tool")
        self.configure(bg=BG)
        self.geometry("680x900")
        self.minsize(660, 720)

        # ---- form variables ----
        self.model = tk.StringVar(value=next(iter(TEMPLATES)))
        self.mac = tk.StringVar()
        self.name = tk.StringVar()
        self.dn = tk.StringVar()
        self.partition = tk.StringVar(value=DEFAULT_PARTITION)
        self.css = tk.StringVar(value=DEFAULT_CSS)
        self.vm = tk.StringVar(value=DEFAULT_VOICEMAIL)
        self.services = tk.StringVar(value=DEFAULT_SERVICES)
        self.make_profile = tk.BooleanVar(value=True)
        self.make_user = tk.BooleanVar(value=True)
        self.auto_login = tk.BooleanVar(value=True)
        self.update_list = tk.BooleanVar(value=True)
        self.profile_partition = tk.StringVar(value=DEFAULT_PROFILE_PARTITION)
        self.first_name = tk.StringVar()
        self.last_name = tk.StringVar()
        self._first_edited = False
        self._last_edited = False
        self.copy_info = tk.StringVar()
        self.status = tk.StringVar(value="Ready")
        self.adv_open = False

        self._styles()
        self._build()
        self.show_sources()
        self.name.trace_add("write", self._autofill_names)
        self.make_profile.trace_add("write", self._deps)
        self.make_user.trace_add("write", self._deps)
        self.bind("<Return>", lambda e: self.run())

    # ---------------------------- styling ----------------------------
    def _styles(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        st.configure(".", font=(FONT, 10), background=CARD, foreground=TEXT)
        st.configure("Card.TLabel", background=CARD, foreground=TEXT)
        st.configure("Field.TLabel", background=CARD, foreground=MUTED,
                     font=(FONT, 9, "bold"))
        st.configure("Muted.TLabel", background=CARD, foreground=MUTED,
                     font=(FONT, 9))
        st.configure("Status.TLabel", background=BG, foreground=MUTED,
                     font=(FONT, 10, "bold"))
        st.configure("Modern.TEntry", padding=8, fieldbackground="#FFFFFF",
                     bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER)
        st.map("Modern.TEntry",
               bordercolor=[("focus", ACCENT)],
               lightcolor=[("focus", ACCENT)],
               darkcolor=[("focus", ACCENT)])
        st.configure("Modern.TCombobox", padding=7, arrowsize=14,
                     fieldbackground="#FFFFFF", bordercolor=BORDER,
                     lightcolor=BORDER, darkcolor=BORDER)
        st.map("Modern.TCombobox",
               fieldbackground=[("readonly", "#FFFFFF")],
               bordercolor=[("focus", ACCENT)])
        st.configure("Card.TCheckbutton", background=CARD, foreground=TEXT,
                     font=(FONT, 10, "bold"))
        st.map("Card.TCheckbutton", background=[("active", CARD)])
        st.configure("Accent.TButton", background=ACCENT, foreground="#FFFFFF",
                     font=(FONT, 12, "bold"), padding=(12, 12), borderwidth=0,
                     focusthickness=0, focuscolor=ACCENT)
        st.map("Accent.TButton",
               background=[("active", ACCENT_DARK), ("disabled", "#93B4F5")],
               foreground=[("disabled", "#FFFFFF")])
        st.configure("Link.TButton", background=CARD, foreground=ACCENT,
                     font=(FONT, 9, "bold"), padding=2, borderwidth=0,
                     focusthickness=0)
        st.map("Link.TButton", background=[("active", CARD)],
               foreground=[("active", ACCENT_DARK)])
        st.configure("Slim.Horizontal.TProgressbar", troughcolor="#DCE2EE",
                     background=ACCENT, bordercolor=BG, lightcolor=ACCENT,
                     darkcolor=ACCENT, thickness=6)

    def _card(self, parent, title, row):
        outer = tk.Frame(parent, bg=CARD, highlightbackground=BORDER,
                         highlightthickness=1)
        outer.grid(row=row, column=0, sticky="ew", pady=(0, 12))
        tk.Label(outer, text=title.upper(), bg=CARD, fg=ACCENT,
                 font=(FONT, 9, "bold")).pack(anchor="w", padx=16, pady=(12, 0))
        inner = tk.Frame(outer, bg=CARD)
        inner.pack(fill="x", padx=16, pady=(6, 14))
        inner.columnconfigure(0, weight=1, uniform="c")
        inner.columnconfigure(1, weight=1, uniform="c")
        return outer, inner

    def _field(self, parent, label, widget, row, col, span=1):
        ttk.Label(parent, text=label, style="Field.TLabel").grid(
            row=row * 2, column=col, columnspan=span, sticky="w",
            padx=(0 if col == 0 else 8, 0), pady=(8, 3))
        widget.grid(row=row * 2 + 1, column=col, columnspan=span, sticky="ew",
                    padx=(0, 8) if col == 0 and span == 1 else (0, 0))

    def _entry(self, parent, var):
        return ttk.Entry(parent, textvariable=var, style="Modern.TEntry")

    # ----------------------------- layout ----------------------------
    def _build(self):
        header = tk.Frame(self, bg=NAVY)
        header.pack(fill="x")
        tk.Label(header, text="CUCM Provisioning Tool", bg=NAVY, fg="white",
                 font=(FONT, 20, "bold")).pack(anchor="w", padx=24, pady=(18, 0))
        tk.Label(header,
                 text="Phone   \u2022   Device profile   \u2022   End user   \u2022   Auto login",
                 bg=NAVY, fg="#9FB0D6", font=(FONT, 10)).pack(
                     anchor="w", padx=24, pady=(2, 16))
        tk.Frame(self, bg=ACCENT, height=3).pack(fill="x")

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=18, pady=(14, 10))
        body.columnconfigure(0, weight=1)
        body.rowconfigure(5, weight=1)

        # --- card 1: the user ---
        _, c1 = self._card(body, "New user", 0)
        self.model_box = ttk.Combobox(
            c1, textvariable=self.model, state="readonly",
            style="Modern.TCombobox",
            values=[m for m in TEMPLATES if m in PROFILE_TEMPLATES])
        self.model_box.bind("<<ComboboxSelected>>", lambda e: self.show_sources())
        self._field(c1, "PHONE MODEL", self.model_box, 0, 0)
        self._field(c1, "MAC ADDRESS", self._entry(c1, self.mac), 0, 1)
        self._field(c1, "USER NAME", self._entry(c1, self.name), 1, 0)
        self._field(c1, "DIRECTORY NUMBER", self._entry(c1, self.dn), 1, 1)
        first_entry = self._entry(c1, self.first_name)
        last_entry = self._entry(c1, self.last_name)
        first_entry.bind("<KeyPress>", lambda e: self._mark_edited("_first_edited", e))
        last_entry.bind("<KeyPress>", lambda e: self._mark_edited("_last_edited", e))
        self._field(c1, "FIRST NAME (AUTO)", first_entry, 2, 0)
        self._field(c1, "LAST NAME (AUTO)", last_entry, 2, 1)
        ttk.Label(c1, textvariable=self.copy_info, style="Muted.TLabel",
                  wraplength=560).grid(row=6, column=0, columnspan=2,
                                       sticky="w", pady=(10, 0))

        # --- card 2: what to create ---
        _, c2 = self._card(body, "What to create", 1)
        row = tk.Frame(c2, bg=CARD)
        row.grid(row=0, column=0, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Checkbutton(row, text="Device profile", variable=self.make_profile,
                        style="Card.TCheckbutton").pack(side="left", padx=(0, 26))
        ttk.Checkbutton(row, text="End user", variable=self.make_user,
                        style="Card.TCheckbutton").pack(side="left", padx=(0, 26))
        ttk.Checkbutton(row, text="Log in automatically", variable=self.auto_login,
                        style="Card.TCheckbutton").pack(side="left")
        row2 = tk.Frame(c2, bg=CARD)
        row2.grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Checkbutton(row2, text="Update phone-list.xlsx", variable=self.update_list,
                        style="Card.TCheckbutton").pack(side="left")
        ttk.Label(c2, text="The phone and its line are always created. The phone "
                  "list is updated only when phone, profile and end user are all created.",
                  style="Muted.TLabel", wraplength=560).grid(
                      row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))
        self.adv_btn = ttk.Button(c2, text="\u25B8  Advanced defaults",
                                  style="Link.TButton", command=self.toggle_adv)
        self.adv_btn.grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))

        # --- card 3: advanced defaults (hidden until opened) ---
        self.adv_card, c3 = self._card(body, "Advanced defaults", 2)
        self._field(c3, "PHONE LINE PARTITION", self._entry(c3, self.partition), 0, 0)
        self._field(c3, "PROFILE LINE PARTITION",
                    self._entry(c3, self.profile_partition), 0, 1)
        self._field(c3, "LINE CSS", self._entry(c3, self.css), 1, 0)
        self._field(c3, "VOICE MAIL PROFILE", self._entry(c3, self.vm), 1, 1)
        self._field(c3, "SERVICES (COMMA SEPARATED)",
                    self._entry(c3, self.services), 2, 0, span=2)
        self.adv_card.grid_remove()

        # --- action ---
        self.btn = ttk.Button(body, text="Create User", style="Accent.TButton",
                              command=self.run, cursor="hand2")
        self.btn.grid(row=3, column=0, sticky="ew", pady=(2, 8))

        stat = tk.Frame(body, bg=BG)
        stat.grid(row=4, column=0, sticky="ew", pady=(0, 8))
        stat.columnconfigure(1, weight=1)
        self.status_lbl = ttk.Label(stat, textvariable=self.status,
                                    style="Status.TLabel")
        self.status_lbl.grid(row=0, column=0, sticky="w")
        self.bar = ttk.Progressbar(stat, mode="indeterminate", length=240,
                                   style="Slim.Horizontal.TProgressbar")
        self.bar.grid(row=0, column=1, sticky="e")

        # --- log ---
        logwrap = tk.Frame(body, bg="#0B1220", highlightbackground="#1E293B",
                           highlightthickness=1)
        logwrap.grid(row=5, column=0, sticky="nsew")
        top = tk.Frame(logwrap, bg="#0B1220")
        top.pack(fill="x", padx=12, pady=(8, 0))
        tk.Label(top, text="ACTIVITY LOG", bg="#0B1220", fg="#64748B",
                 font=(FONT, 8, "bold")).pack(side="left")
        clear = tk.Label(top, text="clear", bg="#0B1220", fg="#60A5FA",
                         font=(FONT, 8, "bold"), cursor="hand2")
        clear.pack(side="right")
        clear.bind("<Button-1>", lambda e: self.clear_log())
        inner = tk.Frame(logwrap, bg="#0B1220")
        inner.pack(fill="both", expand=True, padx=(12, 4), pady=(4, 8))
        self.out = tk.Text(inner, height=9, state="disabled", wrap="word",
                           bg="#0B1220", fg=C_INFO, insertbackground=C_INFO,
                           font=("Consolas", 10), relief="flat", borderwidth=0,
                           padx=2, pady=2)
        sb = ttk.Scrollbar(inner, orient="vertical", command=self.out.yview)
        self.out.configure(yscrollcommand=sb.set)
        self.out.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.out.tag_configure("time", foreground="#64748B")
        self.out.tag_configure("info", foreground=C_INFO)
        self.out.tag_configure("ok", foreground=C_OK)
        self.out.tag_configure("warn", foreground=C_WARN)
        self.out.tag_configure("bad", foreground=C_BAD)

    # ---------------------------- behaviour --------------------------
    def toggle_adv(self):
        self.adv_open = not self.adv_open
        if self.adv_open:
            self.adv_card.grid()
            self.adv_btn.configure(text="\u25BE  Advanced defaults")
        else:
            self.adv_card.grid_remove()
            self.adv_btn.configure(text="\u25B8  Advanced defaults")

    def _deps(self, *_):
        if not self.make_profile.get():
            self.make_user.set(False)
        if not self.make_user.get():
            self.auto_login.set(False)

    def _mark_edited(self, flag, event):
        """The user typed in first/last name: stop auto-filling that field."""
        if event.char or event.keysym in ("BackSpace", "Delete"):
            setattr(self, flag, True)

    def _autofill_names(self, *_):
        text = self.name.get()
        if not text.strip():                  # name cleared: start fresh
            self._first_edited = self._last_edited = False
            self.first_name.set("")
            self.last_name.set("")
            return
        first, last = split_name(text)
        if not self._first_edited:
            self.first_name.set(first)
        if not self._last_edited:
            self.last_name.set(last)

    def show_sources(self):
        m = self.model.get()
        self.copy_info.set(
            f"Copies from  \u2192  phone {TEMPLATES.get(m, '?')}   |   "
            f"profile {PROFILE_TEMPLATES.get(m, '?')}")

    def set_status(self, text, color=MUTED):
        self.status.set(text)
        ttk.Style(self).configure("Status.TLabel", foreground=color)

    def _confirm(self, question):
        return messagebox.askyesno("Delete old phone?", question, icon="warning")

    def clear_log(self):
        self.out.configure(state="normal")
        self.out.delete("1.0", "end")
        self.out.configure(state="disabled")

    def log(self, msg):
        low = msg.lower()
        if msg.startswith("ERROR") or "failed" in low:
            tag = "bad"
        elif msg.startswith(("DONE", "Logged")):
            tag = "ok"
        elif msg.startswith(("Rolled", "Note", "Could not")) or "auto login try" in low:
            tag = "warn"
        else:
            tag = "info"
        self.out.configure(state="normal")
        self.out.insert("end", time.strftime("%H:%M:%S  "), "time")
        self.out.insert("end", msg + "\n", tag)
        self.out.see("end")
        self.out.configure(state="disabled")
        self.update_idletasks()

    def run(self):
        if str(self.btn.cget("state")) == "disabled":
            return
        model = self.model.get()
        mac = self.mac.get()
        name = self.name.get().strip()
        dn = self.dn.get().strip()
        part = self.partition.get().strip()
        css = self.css.get().strip()
        vm = self.vm.get().strip()
        services = [x.strip() for x in self.services.get().split(",") if x.strip()]
        make_profile = self.make_profile.get()
        pprofile = self.profile_partition.get().strip()
        make_user = self.make_user.get()
        auto_login = self.auto_login.get()
        first = self.first_name.get().strip()
        last = self.last_name.get().strip()

        if not (mac.strip() and name and dn and part and css and vm):
            messagebox.showwarning("Missing data", "Please fill in the model, "
                                   "MAC address, user name and directory number.")
            return
        if make_user and not last:
            messagebox.showwarning("Missing data", "The end user needs a last name.")
            return
        if make_profile and not pprofile:
            messagebox.showwarning("Missing data", "Enter the profile line partition "
                                   "(Advanced defaults).")
            return

        self.btn.configure(state="disabled", text="Working ...")
        self.set_status(f"Creating {name} ...", ACCENT)
        self.bar.start(12)
        try:
            axl = get_axl()
            note = create_user(axl, model, mac, name, dn, part, css, vm, services,
                               make_profile, pprofile, self.log,
                               make_user, auto_login, self.update,
                               confirm=self._confirm,
                               first_name=first, last_name=last,
                               update_list=self.update_list.get())
            extra = (" + device profile" if make_profile else "") \
                + (" + end user" if make_user else "")
            self.set_status(f"\u2714  Created {name} ({dn})", "#16A34A")
            messagebox.showinfo("Success",
                                f"Phone{extra} created for {name} ({dn}).{note}")
            self.mac.set("")
            self.name.set("")
            self.dn.set("")
        except Exception as e:
            self.log(f"ERROR: {e}")
            self.set_status("\u2716  Failed - nothing was left behind in CUCM",
                            "#DC2626")
            messagebox.showerror("Error", str(e))
        finally:
            self.bar.stop()
            self.btn.configure(state="normal", text="Create User")


if __name__ == "__main__":
    try:
        load_config()
    except ConfigError as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Settings needed", str(e))
        sys.exit(1)
    App().mainloop()
