# Shipping Label Automation Tool

A lightweight Python app to automate shipping label creation from Excel-based packing lists. Built for a small business, this tool parses data and generates labels using pre-defined Excel templates.

## Why I Built This

To save time and eliminate manual errors during retail fulfillment. Labels were originally typed manually — this tool cut the process from 2+ hours to a few clicks.

## What It Does

- Reads `.xlsx` packing lists
- Parses shipment info and carton data
- Fills out Excel templates (Template 1, Template 2, etc.)
- Outputs ready-to-print `.xlsx` label sheets

## Tech Stack

- Python + Tkinter (GUI)
- OpenPyXL for Excel automation
- Regex for flexible field parsing

## Try It Out

Make sure the `templates/` folder is placed in the same directory as main.py and then run `main.py` with Python 3. Packing list and template examples included.