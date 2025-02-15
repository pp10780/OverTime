# Overtime Tracker

## Description

This project is a **Python-based tool** that processes employee work logs from an Excel spreadsheet. It calculates **extra hours**, **negative hours**, and **identifies incorrect entries**, generating a new Excel report with **individual sheets for each employee**. The program also applies **conditional formatting** to highlight key data points.

## Features

- **Parses employee entry/exit times** from an Excel file.
- **Calculates extra hours** while considering lunch breaks.
- **Detects and highlights incorrect time entries**.
- **Computes negative hours** when employees leave early.
- **Generates a new Excel report** with separate sheets per employee.
- **Applies color formatting**:
  - Missing values: 🟡 Yellow
  - Overtime: 🟢 Green
  - Negative hours: 🔴 Red
  - Incorrect hours: 🔵 Blue

## Installation

### Prerequisites
Ensure you have **Python 3.x** installed along with the required libraries:

```bash
pip install pandas openpyxl
