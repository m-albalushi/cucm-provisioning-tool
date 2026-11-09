"""
Template Maker for the CUCM Provisioning Tool
---------------------------------------------
Creates, for ONE phone model, a ready-made TEMPLATE phone and a ready-made
TEMPLATE device profile by copying the settings of an existing phone/profile,
without lines, speed dials or user data. After this, new users are copied from
the templates and no real user's phone or profile is used any more.

Run (same folder as cucm_phone_creator.py and config.ini):
    python make_templates.py

It only READS the phone and profile you name and CREATES two new objects.
Nothing is deleted (only the new phone is removed again if the profile fails).
"""

import cucm_phone_creator as c

PERSONAL_FIELDS = ("lines", "speeddials", "busyLampFields", "blfDirectedCallParks",
                   "primaryPhoneName", "addOnModules")


def ask(prompt):
    while True:
        value = input(prompt).strip()
        if value:
            return value
        print("  (cannot be empty)")


def template_mac(model):
    """Fake but valid MAC for the template phone, e.g. 7945 -> 000000007945."""
    if model.isdigit() and len(model) <= 12:
        return model.zfill(12)
    return c.clean_mac(ask("Model is not only digits. Fake MAC for the template "
                           "phone (12 hex characters, e.g. 000000000001): "))


def make_phone_template(axl, src_phone, name, model, services, log):
    log(f"Reading phone {src_phone} ...")
    src = c.serialize_object(axl.getPhone(name=src_phone)["return"]["phone"])
    svc = c.missing_service_entries(axl, src, services, log, "phone")
    phone = c.copy_from_template(src, "XPhone")
    for k in PERSONAL_FIELDS:
        phone.pop(k, None)
    phone["name"] = name
    phone["description"] = f"TEMPLATE-{model} - do not use"
    phone["enableExtensionMobility"] = c.ENABLE_EXTENSION_MOBILITY
    c.merge_services(phone, svc)
    log(f"Creating template phone {name} ...")
    axl.addPhone(phone=phone)


def make_profile_template(axl, src_profile, name, model, services, log):
    log(f"Reading device profile {src_profile} ...")
    src = c.serialize_object(
        axl.getDeviceProfile(name=src_profile)["return"]["deviceProfile"])
    svc = c.missing_service_entries(axl, src, services, log, "profile")
    prof = c.copy_from_template(src, "XDeviceProfile")
    for k in PERSONAL_FIELDS:
        prof.pop(k, None)
    prof["name"] = name
    prof["description"] = f"TEMPLATE-{model} - do not use"
    c.merge_services(prof, svc)
    log(f'Creating template device profile "{name}" ...')
    axl.addDeviceProfile(deviceProfile=prof)


def main():
    print("\n=== CUCM Template Maker ===\n")
    try:
        c.load_config()
    except c.ConfigError as e:
        print(e)
        return

    model = ask("Phone model (for example 7945): ")
    src_phone = c.phone_device_name(ask("Existing phone to copy the settings from (SEP...): "))
    src_profile = ask("Existing device profile to copy the settings from (exact name): ")

    phone_name = "SEP" + template_mac(model)
    profile_name = f"TEMPLATE-{model}"
    services = [s.strip() for s in c.DEFAULT_SERVICES.split(",") if s.strip()]

    print(f"\nWill create:\n  template phone   : {phone_name}\n"
          f"  template profile : {profile_name}\n"
          f"  copied from      : {src_phone}  and  \"{src_profile}\"")
    if input("\nContinue? (y/n): ").strip().lower() != "y":
        print("Cancelled.")
        return

    try:
        axl = c.get_axl()
        if c.phone_exists(axl, phone_name):
            raise RuntimeError(f"Template phone {phone_name} already exists. "
                               "Delete it in CUCM first if you want to rebuild it.")
        if c.profile_exists(axl, profile_name):
            raise RuntimeError(f'Template profile "{profile_name}" already exists. '
                               "Delete it in CUCM first if you want to rebuild it.")

        make_phone_template(axl, src_phone, phone_name, model, services, print)
        try:
            make_profile_template(axl, src_profile, profile_name, model, services, print)
        except Exception:
            print("Profile failed, removing the new template phone again ...")
            try:
                axl.removePhone(name=phone_name)
            except Exception as e:
                print(f"Could not remove {phone_name} (delete it manually): {e}")
            raise
    except Exception as e:
        print(f"\nERROR: {e}")
        return

    print(f"""
DONE. In CUCM, open both new templates and check them:
  - Device > Phone > {phone_name}
  - User Management > User Settings > Device Profile > {profile_name}
Remove anything personal that you do not want every new user to get.

Then put these lines in config.ini (replace the old line of model {model}):

  [phones]
  {model} = {phone_name}

  [profiles]
  {model} = {profile_name}
""")


if __name__ == "__main__":
    main()
