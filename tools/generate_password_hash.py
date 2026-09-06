"""Generate the salt and hash that gate the Advanced Settings tab.

The app never stores the password itself: it stores a random salt and the
SHA-256 of salt + password, and compares against those at unlock time. Run
this when the password needs to change, then paste the two printed values
into ``SALT`` and ``HASH`` on ``PasswordPopup`` in ``src/gui/main_window.py``.

    python tools/generate_password_hash.py

The password is read from a hidden prompt so it never lands in this file,
in your shell history, or in the repository.
"""

import getpass
import hashlib
import os

SALT_BYTES = 16


def main() -> None:
    password = getpass.getpass("New Advanced Settings password: ")
    if not password:
        raise SystemExit("Aborted: empty password.")
    if password != getpass.getpass("Confirm: "):
        raise SystemExit("Aborted: the two entries did not match.")

    salt = os.urandom(SALT_BYTES)
    digest = hashlib.sha256(salt + password.encode("utf-8")).hexdigest()

    print("\nPaste these into PasswordPopup in src/gui/main_window.py:\n")
    print(f"    SALT = bytes.fromhex('{salt.hex()}')")
    print(f"    HASH = '{digest}'")


if __name__ == "__main__":
    main()
