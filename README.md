# Canoe Doc Receipt Dashboard

Live dashboard tracking Q4 2025 / Q1 2026 document receipts across deals and legal entities.

## Files

| File | Purpose |
|------|---------|
| `pcap-receipt-dashboard.html` | Generated dashboard (GitHub Pages) |
| `gen_pcap_dashboard.js` | Node.js generator — builds the HTML from data |
| `parse_pcap.js` | Parses raw sqlcmd output into `pcap_data.json` |
| `pcap_data.json` | Cleaned deal/LE/doc receipt data (2,602 rows) |

## How to regenerate the dashboard

1. Export fresh data from DB → `pcap_tsv2.txt`
2. Run parser: `node parse_pcap.js`
3. Place contact files locally (see generator for expected paths)
4. Run generator: `node gen_pcap_dashboard.js`
5. Copy and push: `cp pcap-receipt-dashboard.html jira-dashboard/ && cd jira-dashboard && git add . && git commit -m "Refresh data" && git push`

## Live URL
https://abgeorge-gcmlp.github.io/jira-dashboard/pcap-receipt-dashboard.html
