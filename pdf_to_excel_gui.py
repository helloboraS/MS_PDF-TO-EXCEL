
import streamlit as st
import pandas as pd
import pdfplumber
import tempfile
import os
import collections
import re

def group_words_by_line(words, y_tolerance=3):
    lines = collections.defaultdict(list)
    for word in words:
        y_center = (word["top"] + word["bottom"]) / 2
        matched = False
        for key in lines:
            if abs(key - y_center) <= y_tolerance:
                lines[key].append(word)
                matched = True
                break
        if not matched:
            lines[y_center].append(word)
    return lines

def clean_code(text):
    return re.sub(r"[^A-Za-z0-9]", "", str(text)).upper()

def extract_export_code_from_lines(lines):
    export_map = {}
    current_item = None
    for idx, line in enumerate(lines):
        if re.search(r"\bWES[-\w]+\b", line):
            current_item = re.search(r"\bWES[-\w]+\b", line).group()
        if current_item:
            for j in range(idx, min(idx + 5, len(lines))):
                line_lower = lines[j].lower()
                match = re.search(r"(export code|hs code)[:：]?\s*([\d\.]+)", line_lower)
                if match:
                    export_map[current_item] = match.group(2)
                    break
            current_item = None
    return export_map
