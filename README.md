# Python CI/CD Runtime Assessment & Excel Reporting Pipeline

An automated end-to-end CI/CD monitoring and reporting system built with **Python**, **GitHub Actions**, and **openpyxl**.

## How It Works
1. **Automated Unit Testing (`pytest`):** Validates health-check logic and Excel spreadsheet generation on every push.
2. **Live Runtime Assessment:** Pings the deployed production web application (`ROBO-AI-AGENT`) to verify uptime (`HTTP 200`) and measure response latency.
3. **Excel Report Generation (`Assessment_Results.xlsx`):** Appends color-coded assessment rows (`PASSED` / `FAILED`) with UTC timestamps and auto-commits the updated workbook to the repository.
4. **Automated Email Dispatch:** Delivers the `.xlsx` report directly to Microsoft Outlook for immediate viewing in Excel Online.
