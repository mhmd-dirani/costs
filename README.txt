Workshop Payments Index Website (local)

What this is
- A simple single-page website (no server needed) that:
  - Opens your Excel file (client-side)
  - Lets you choose a workshop (each sheet)
  - Shows the payments table + computed total
  - Displays and edits the payment date so you can track when each entry was paid
  - Lets you add or delete rows
  - Lets you download an updated Excel file (all workshops) or CSV (current workshop)

How to use
1) Download the folder and unzip it.
2) Open index.html (double click).
3) Click 'Excel file' and choose your .xlsx.
4) Select a workshop from the dropdown.
5) Add payments using the form. (Date is optional, but matches your Excel column if you fill it in.)
6) Click 'Download updated Excel' to save your updated file.

Notes
- This does NOT edit the original Excel file directly (browsers don't allow that).
  You download a new updated Excel file.
- Your data stays on your device (no upload).

Optional (recommended)
- If your browser blocks some local features, you can run a tiny local server:
  - Windows: open CMD inside the folder then:  python -m http.server 8000
  - Then open: http://localhost:8000
