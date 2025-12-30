## Architecture Overview

- VirtualBox hosts a Windows domain environment
- Domain Controller provides AD, DNS, and authentication
- Endpoints simulate user workstations
- PowerShell script runs from a management host
- SMTP used for alert delivery

Flow:
CSV Inventory → Test-Connection → Status Evaluation → Email Alert
