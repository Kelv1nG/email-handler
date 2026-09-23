"""Generate the synthetic MSG fixture; requires desktop Outlook once."""

from pathlib import Path

import win32com.client


HTML = """
<html><head><meta http-equiv="Content-Type" content="text/html; charset=windows-1252"></head><body>
<table>
  <tr><th>Item</th><th>Quantity</th></tr>
  <tr><td>Caf\u00e9</td><td>2</td></tr>
</table>
<table>
  <tr><th>Account</th><th>Amount</th></tr>
  <tr><td>A-100</td><td>125.50</td></tr>
</table>
</body></html>
"""


def main() -> None:
    destination = Path(__file__).with_name("table_message.msg").resolve()
    if destination.exists():
        raise FileExistsError(f"Refusing to overwrite fixture: {destination}")
    outlook = win32com.client.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.Subject = "Synthetic table fixture"
    message.HTMLBody = HTML
    message.SaveAs(str(destination), 9)
    print(destination)


if __name__ == "__main__":
    main()
