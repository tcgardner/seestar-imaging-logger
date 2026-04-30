# Seestar S30 Pro Imaging Log

![CI](https://github.com/YOUR-USERNAME/seestar-s30-imaging-log/actions/workflows/ci.yml/badge.svg)
![License](https://img.shields.io/badge/license-MIT-blue.svg)

A powerful TypeScript CLI tool that processes **.fit / .fits** files from the **ZWO Seestar S30 Pro** smart telescope, intelligently groups them into **imaging sessions**, and automatically logs everything to a CSV file + your personal **Google Sheet** with nice formatting and color coding by Object Type.

## Features

- Recursive scanning of FITS files
- **Smart session grouping** (same date + object + continuous time)
- Progress bar with real-time feedback
- No-duplicate processing with persistent cache
- CSV export (create or append)
- **Google Sheets auto-export** with header formatting, frozen row, auto-resize, and **color coding** by Object Type
- Remembers last used settings
- Skips calibration folders by default

## Quick Start

```bash
npm install
npx tsx src/index.ts "/path/to/seestar/sessions" -o my-log.csv -g YOUR_SPREADSHEET_ID

 