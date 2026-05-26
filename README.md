# Localization Automation Tools

A showcase of automation tools and QA validation systems built for localization and content production workflows in the game development industry.

This repository contains demo/sanitized examples inspired by real production workflows. All company-specific names, internal data, personal information, and project-sensitive details have been removed or replaced with generic sample data.

## Technologies
- Google Apps Script (JavaScript)
- Google Sheets advanced formulas
- Regex-based validation
- HTML email automation
- Spreadsheet workflow automation
- Rich text formatting in Google Sheets

## Showcases included

### HTML Export Translation Parser

Parses an HTML table export into Google Sheets, removes HTML tags and formatting artifacts, compares source strings against a translation memory sheet, and populates existing translations into locale columns.

### Localization Consistency Validator

Checks selected locale columns against English source text and highlights potential issues with numbers, glossary terms, event names, tier names, and month translations.

### Text Change Highlighter

Compares old and new text versions and visually marks changes directly in Google Sheets: added words are highlighted in bold, removed words are marked with strikethrough.

### Invoice Reminder Email Automation

Collects active contributor emails, calculates the invoice period, retrieves working day data, and sends an HTML reminder email.

### Monthly Schedule Email Automation

Builds a formatted monthly schedule email from a Google Sheets table while preserving table styles, colors, alignments, and merged cells.

### Contributor Task Block Manager

Adds a new contributor task block into a structured spreadsheet, preserves merged cells, copies existing formulas, and generates rate lookup formulas based on contributor rates stored in another sheet.

### Calendar Schedule Builder Formulas

Complex Google Sheets formulas used to generate calendar-style schedule tables that are later processed by the monthly schedule email automation script.

## Purpose

The goal of this repository is to demonstrate practical scripting, spreadsheet automation, QA validation logic, and workflow optimization for localization and content production pipelines.