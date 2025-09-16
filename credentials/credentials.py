# credentials.py
# Показва всички Windows Generic Credentials с TargetName започващ с "AUTOMATION/ "
# и принтира TargetName, UserName и (по избор) паролата.

from typing import List, Tuple
import sys

try:
    import win32cred  # pywin32
except ImportError as e:
    raise SystemExit("Липсва pywin32. Инсталирай с: pip install pywin32") from e

CRED_TYPE = win32cred.CRED_TYPE_GENERIC
PREFIX = "AUTOMATION/"

def _decode_password(blob) -> str:
    """
    Windows CredentialBlob обикновено е bytes в UTF-16-LE.
    Декодираме с безопасен fallback.
    """
    if not blob:
        return ""
    if isinstance(blob, bytes):
        # основен вариант
        try:
            return blob.decode("utf-16-le")
        except UnicodeDecodeError:
            # понякога други кодировки / нестандартен запис
            try:
                return blob.decode("utf-8")
            except UnicodeDecodeError:
                return blob.decode(errors="ignore")
    # ако не е bytes, връщаме като текст
    return str(blob)

def list_automation_credentials() -> List[Tuple[str, str, str]]:
    """
    Връща списък от (target_name, username, password) за всички
    Generic Credentials с TargetName, започващ с PREFIX.
    """
    results: List[Tuple[str, str, str]] = []

    # CredEnumerate позволява филтър по TargetName
    # Използваме wildcard за всички под PREFIX.
    # https://learn.microsoft.com/windows/win32/api/wincred/nf-wincred-credenumeratea
    try:
        creds = win32cred.CredEnumerate(f"{PREFIX}*", 0)
    except win32cred.error as e:
        # ако няма нищо, CredEnumerate може да върне None или да вдигне грешка
        creds = None

    if not creds:
        return results

    for c in creds:
        target = c.get("TargetName", "")
        username = c.get("UserName", "")
        password = _decode_password(c.get("CredentialBlob", b""))

        # Някои реализации на pywin32 понякога не връщат CredentialBlob в Enumerate.
        # За по-сигурно, прочети пак конкретния запис с CredRead:
        if password == "" and target:
            try:
                one = win32cred.CredRead(target, CRED_TYPE, 0)
                password = _decode_password(one.get("CredentialBlob", b""))
                username = one.get("UserName", username)
            except win32cred.error:
                pass

        results.append((target, username, password))

    return results

def main():
    # CLI:
    #   python credentials.py           -> списък (без да показва пароли)
    #   python credentials.py --show    -> списък + пароли (внимавай!)
    show_passwords = len(sys.argv) > 1 and sys.argv[1].strip() in ("--show", "-s")

    rows = list_automation_credentials()
    if not rows:
        print(f"Няма записи с TargetName започващ с '{PREFIX}'.")
        return

    print(f"Намерени {len(rows)} credential(а) под '{PREFIX}':\n")
    for target, user, pwd in rows:
        if show_passwords:
            print(f"TargetName: {target}\n  UserName: {user}\n  Password: {pwd}\n")
        else:
            masked = "•" * 8 if pwd else ""
            print(f"TargetName: {target}\n  UserName: {user}\n  Password: {masked} (скрито - добави --show за показване)\n")

if __name__ == "__main__":
    main()
