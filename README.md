# LumberTools

PowerShell utilities for managed Windows endpoints, by Lumberstack.

## Available Tools

| Tool | Description |
|------|-------------|
| **Merge Documents** | Merge PDFs and Word documents (.doc/.docx) into a single PDF |

## Installation

1. Sync this repo to `C:\LumberTools\` on the target machine
2. Run Setup.ps1 to create Start Menu shortcuts:
   ```
   powershell.exe -ExecutionPolicy Bypass -File "C:\LumberTools\Setup.ps1"
   ```
3. Tools appear in **Start Menu > LumberTools** and can be pinned to the taskbar or desktop

## Requirements

- Windows 10/11 with PowerShell 5.1
- Microsoft Word (for tools that convert Word documents)
