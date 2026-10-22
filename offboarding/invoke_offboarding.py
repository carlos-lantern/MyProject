#!/usr/bin/env python3
"""
Lantern employee offboarding runbook - Python port of offboarding-exchange-actions.ps1.

Order matters: licensing uses dynamic groups driven by Entra attributes, so
attributes are only changed AFTER the mailbox is confirmed Shared.
    1. Disable account
    2. Exchange changes (shared, hide from GAL, forwarding, auto-reply)
    3. Verify mailbox is Shared (abort otherwise)
    4. Clear Entra attributes, manager, groups
    5. Intune device category
    6. Laptop return email (if label needed)

Requires a managed identity with Graph (User.ReadWrite.All, Group.ReadWrite.All,
Directory.ReadWrite.All, DeviceManagementManagedDevices.ReadWrite.All,
DeviceManagementServiceConfig.Read.All, RoleManagement.Read.Directory).

NOTE: Microsoft Graph has no REST surface for mailbox type conversion,
ForwardingSmtpAddress, HiddenFromAddressListsEnabled, or auto-reply configuration
(these are Exchange Online-only settings). Those specific steps shell out to
PowerShell + the ExchangeOnlineManagement module via subprocess; everything else
uses Microsoft Graph directly through requests/azure-identity.

Set FRESHSERVICE_API_KEY in the environment (or a secrets manager of your choice)
before running step 6. Use --what-if for a dry run.
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import time
from datetime import datetime, timedelta

import requests
from azure.identity import ManagedIdentityCredential

GRAPH_BASE = "https://graph.microsoft.com"
GRAPH_SCOPE = "https://graph.microsoft.com/.default"


def log(msg: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {msg}")


def run_pwsh(script: str) -> str:
    """Run a PowerShell snippet for the handful of Exchange Online-only cmdlets."""
    result = subprocess.run(
        ["pwsh", "-NoProfile", "-NonInteractive", "-Command", script],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(f"PowerShell step failed:\n{result.stderr}")
    return result.stdout.strip()


class GraphClient:
    def __init__(self, credential: ManagedIdentityCredential):
        token = credential.get_token(GRAPH_SCOPE).token
        self.session = requests.Session()
        self.session.headers.update(
            {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        )

    def get(self, path: str, **kwargs) -> dict:
        r = self.session.get(f"{GRAPH_BASE}{path}", **kwargs)
        r.raise_for_status()
        return r.json()

    def get_all(self, path: str, **kwargs) -> list:
        items = []
        url = f"{GRAPH_BASE}{path}"
        while url:
            r = self.session.get(url, **kwargs)
            r.raise_for_status()
            data = r.json()
            items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
            kwargs = {}
        return items

    def patch(self, path: str, body: dict) -> None:
        r = self.session.patch(f"{GRAPH_BASE}{path}", data=json.dumps(body))
        r.raise_for_status()

    def put(self, path: str, body: dict) -> None:
        r = self.session.put(f"{GRAPH_BASE}{path}", data=json.dumps(body))
        r.raise_for_status()

    def delete(self, path: str) -> None:
        r = self.session.delete(f"{GRAPH_BASE}{path}")
        if r.status_code not in (204, 404):
            r.raise_for_status()


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Lantern employee offboarding runbook")
    p.add_argument("--user-principal-name", required=True)
    p.add_argument("--last-day", required=True, help="YYYY-MM-DD")
    p.add_argument("--forward-to-email")
    p.add_argument("--retention-weeks", type=int, default=8)
    p.add_argument("--laptop-left-in-office", action="store_true")
    p.add_argument("--personal-email")
    p.add_argument("--device-purchase-date", help="YYYY-MM-DD")
    p.add_argument("--tenant-domain", default="lanternstudios.com")
    p.add_argument("--ticket-id", type=int)
    p.add_argument("--freshservice-domain", default="lanternstudios.freshservice.com")
    p.add_argument(
        "--what-if", action="store_true", help="Dry run: log actions without making changes"
    )
    return p.parse_args()


def main() -> None:
    args = parse_args()
    what_if = args.what_if
    upn = args.user_principal_name
    last_day = datetime.strptime(args.last_day, "%Y-%m-%d")
    device_purchase_date = (
        datetime.strptime(args.device_purchase_date, "%Y-%m-%d")
        if args.device_purchase_date
        else None
    )

    credential = ManagedIdentityCredential()
    graph = GraphClient(credential)

    run_pwsh(
        f"Connect-ExchangeOnline -ManagedIdentity -Organization {args.tenant_domain} -ShowBanner:$false"
    )

    user = graph.get(
        f"/v1.0/users/{upn}",
        params={"$select": "id,displayName,givenName,surname,accountEnabled"},
    )
    log(f"Offboarding {user['displayName']} ({upn})")

    # Guard: never offboard accounts holding directory roles
    member_of = graph.get_all(f"/v1.0/users/{user['id']}/transitiveMemberOf")
    if any(m.get("@odata.type") == "#microsoft.graph.directoryRole" for m in member_of):
        raise RuntimeError(f"{upn} holds directory roles; refusing to automate. Handle manually.")

    # 1. Disable account
    if not what_if:
        graph.patch(f"/v1.0/users/{user['id']}", {"accountEnabled": False})
    log("Account disabled" if not what_if else "[WhatIf] Would disable account")

    # 2. Exchange
    mbx_type = run_pwsh(f"(Get-Mailbox -Identity {upn}).RecipientTypeDetails")
    if mbx_type != "SharedMailbox" and not what_if:
        run_pwsh(f"Set-Mailbox -Identity {upn} -Type Shared")
        log("Conversion to shared requested")
    elif mbx_type != "SharedMailbox":
        log("[WhatIf] Would convert mailbox to Shared")

    if not what_if:
        run_pwsh(
            f"Set-Mailbox -Identity {upn} -CustomAttribute1 $null -HiddenFromAddressListsEnabled $true"
        )
        if args.forward_to_email:
            run_pwsh(
                f"Set-Mailbox -Identity {upn} -ForwardingSmtpAddress {args.forward_to_email} "
                f"-DeliverToMailboxAndForward $false"
            )
            fwd_name = graph.get(
                f"/v1.0/users/{args.forward_to_email}", params={"$select": "displayName"}
            )["displayName"]
            reply = (
                "Hi there,<br><br>"
                f"Thanks for reaching out! Just a heads-up, I'm Out of Office. But don't worry, "
                f"your message will be forwarded to {fwd_name} ({args.forward_to_email}), "
                "who will be happy to assist you.<br><br>Cheers,"
            )
            reply_escaped = reply.replace("'", "''")
            run_pwsh(
                f"Set-MailboxAutoReplyConfiguration -Identity {upn} -AutoReplyState Enabled "
                f"-InternalMessage '{reply_escaped}' -ExternalMessage '{reply_escaped}'"
            )
        log("Mailbox settings applied")
    else:
        log("[WhatIf] Would update mailbox settings (hidden, forwarding, auto-reply)")

    # 3. Verify shared before touching attributes (WhatIf skips the gate)
    if not what_if:
        ok = False
        for _ in range(30):
            if run_pwsh(f"(Get-Mailbox -Identity {upn}).RecipientTypeDetails") == "SharedMailbox":
                ok = True
                break
            time.sleep(20)
        if not ok:
            raise RuntimeError("Mailbox is not Shared after waiting. Aborting before attribute changes.")

        size_bytes = run_pwsh(
            f"(Get-MailboxStatistics -Identity {upn}).TotalItemSize.Value.ToBytes()"
        )
        size_gb = round(int(size_bytes) / 1_073_741_824, 2)
        hold = run_pwsh(f"(Get-Mailbox -Identity {upn}).LitigationHoldEnabled")
        if size_gb >= 50 or hold.lower() == "true":
            log(f"WARNING: Mailbox is {size_gb} GB / LitigationHold={hold} - a license may be required. Review manually.")
        log(f"Mailbox verified Shared ({size_gb} GB)")

    # 4. Entra attributes, manager, groups
    if not what_if:
        graph.patch(
            f"/v1.0/users/{user['id']}",
            {
                "companyName": "Lantern-disabled",
                "jobTitle": None,
                "department": None,
                "employeeType": None,
                "officeLocation": None,
                "state": None,
                "country": None,
            },
        )
        graph.delete(f"/v1.0/users/{user['id']}/manager/$ref")

        groups = [m for m in member_of if m.get("@odata.type") == "#microsoft.graph.group"]
        for g in groups:
            if "DynamicMembership" in (g.get("groupTypes") or []):
                continue  # removed automatically by the attribute change
            try:
                graph.delete(f"/v1.0/groups/{g['id']}/members/{user['id']}/$ref")
            except requests.HTTPError as exc:
                log(f"WARNING: Could not remove from {g.get('displayName')}: {exc}")
        log("Entra attributes cleared and groups removed")
    else:
        log("[WhatIf] Would clear Entra attributes, manager and groups")

    # 5. Intune device category
    three_years_ago = datetime.now().replace(year=datetime.now().year - 3)
    if not args.laptop_left_in_office:
        category = "Lantern-In Transit"
    elif device_purchase_date and device_purchase_date <= three_years_ago:
        category = "Lantern-retired"
    else:
        category = "Lantern-Inventory"

    devices = graph.get_all(
        "/v1.0/deviceManagement/managedDevices",
        params={"$filter": f"userPrincipalName eq '{upn}'"},
    )
    categories = graph.get("/v1.0/deviceManagement/deviceCategories")["value"]
    cat = next((c for c in categories if c["displayName"] == category), None)
    if not cat:
        log(f"WARNING: Device category '{category}' not found")
    for d in devices:
        if cat and not what_if:
            graph.put(
                f"/beta/deviceManagement/managedDevices/{d['id']}/deviceCategory/$ref",
                {"@odata.id": f"https://graph.microsoft.com/beta/deviceManagement/deviceCategories/{cat['id']}"},
            )
            log(f"{d['deviceName']} -> {category}")
        elif cat:
            log(f"[WhatIf] Would set {d['deviceName']} category to {category}")
    if not devices:
        log("WARNING: No Intune device found for user")

    # 6. Laptop return email
    if not args.laptop_left_in_office and args.personal_email:
        html = f"""<p>Hi,</p>
<p>HR has confirmed that your employment with Lantern ended today on {last_day:%B %-d, %Y}.<br>
In line with our asset-return policy, we need to collect the company-issued laptop and charger in your possession.</p>
<p><b>How the return works</b></p>
<ul>
<li>Reply with your preferred return address. We'll use this only as the "From" address on the prepaid UPS label.</li>
<li>You'll receive the label by email within one business day.</li>
<li>Place the laptop and charger in a sturdy box, affix the label, and drop the package at any UPS location.</li>
<li>Please ship within three business days of receiving the label.</li>
</ul>
<p>A few reminders:</p>
<ul>
<li>Be sure the laptop is fully powered off before packing.</li>
<li>Depending on the state of your laptop, it may be re-used or recycled. In either case, the drive will be wiped to destroy all the data on it.</li>
</ul>
<p>Thank you for helping us complete the hand-off quickly.</p>
<p>Best regards,</p>"""
        if not what_if:
            api_key = os.environ["FRESHSERVICE_API_KEY"]
            auth = requests.auth.HTTPBasicAuth(api_key, "X")
            # Replying on the ticket keeps the employee's answer (return address) on the same ticket
            # and avoids mailing from helpdesk@, which would open a new ticket.
            resp = requests.post(
                f"https://{args.freshservice_domain}/api/v2/tickets/{args.ticket_id}/reply",
                auth=auth,
                json={"body": html, "cc_emails": [args.personal_email]},
            )
            resp.raise_for_status()
            log("Return email sent via Freshservice ticket reply")
        else:
            log(f"[WhatIf] Would send laptop return reply via Freshservice ticket {args.ticket_id}")

    delete_on = (last_day + timedelta(weeks=args.retention_weeks)).strftime("%Y-%m-%d")
    log(f"DONE. Create reminder: delete shared mailbox {upn} on {delete_on}")

    run_pwsh("Disconnect-ExchangeOnline -Confirm:$false")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # noqa: BLE001 - top-level runbook error surface
        log(f"ERROR: {exc}")
        sys.exit(1)
