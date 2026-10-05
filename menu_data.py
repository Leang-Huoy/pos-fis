# -*- coding: utf-8 -*-
"""
menu_data.py
ម៉ូឌុលគ្រប់គ្រងបញ្ជីមុខម្ហូប និងកាលវិភាគអាហារូបត្ថម្ភតាមសាលារៀន (School Feeding Menus & Nutritional Scheduling)
ស្របតាមស្តង់ដារប្រព័ន្ធព័ត៌មានគ្រប់គ្រងកម្មវិធីផ្តល់អាហារតាមសាលារៀន MoEYS SFIS (https://sfis.moeys.gov.kh)
និងរចនាសម្ព័ន្ធឯកសារជាក់ស្តែង «បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm» (សន្លឹក ចំនួនសរុប ខាងកើត)
"""

import io
import re
import csv
import textwrap
import sqlite3
from datetime import datetime, date
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
try:
    import docx
except ImportError:
    docx = None
try:
    import pdfplumber
except ImportError:
    pdfplumber = None
from catalog_data import SUPPLIER_PRODUCT_CATALOG

# Khmer Day of Week mapping
KHMER_DAYS_OF_WEEK = ["ចន្ទ", "អង្គារ", "ពុធ", "ព្រហស្បតិ៍", "សុក្រ", "សៅរ៍", "អាទិត្យ"]
KHMER_WEEKDAY_MAP = {0: "ចន្ទ", 1: "អង្គារ", 2: "ពុធ", 3: "ព្រហស្បតិ៍", 4: "សុក្រ", 5: "សៅរ៍", 6: "អាទិត្យ"}

# ================= គំរូបញ្ជីមុខម្ហូបស្ដង់ដារ MoEYS SFIS =================
STANDARD_SFIS_TEMPLATES = {
    "cycle_1": {
        "id": "cycle_1",
        "name": "គំរូស្ដង់ដារ MoEYS 2026 (វដ្តទី១ - សប្ដាហ៍ទី១ និងទី៣)",
        "description": "គំរូផ្លូវការតាមឯកសារជាក់ស្តែង «បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026» និង MoEYS SFIS (ចន្ទ ដល់ សៅរ៍)",
        "dishes": [
            {
                "day": "ចន្ទ",
                "dish_name": "សម្លកកូរសាច់ជ្រូក",
                "meal_type": "អាហារពេលព្រឹក",
                "description": "សម្លកកូរសម្បូរដោយជាតិដែក វីតាមីន និងប្រូតេអ៊ីន ពីសាច់ជ្រូក ស្លឹកបាស ល្ពៅ ល្ហុងខ្ចី និងសណ្ដែកដី",
                "ingredients": [
                    {"item_name": "សាច់ជ្រូក៣ជាន់", "category": "ត្រី សាច់ ស៊ុត", "unit": "1គីឡូ", "gram_per_student": 22.0, "qty_per_100": 2.2, "note": "ប្រូតេអ៊ីន និងខ្លាញ់ល្អ"},
                    {"item_name": "ស្លឹកបាស", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 22.0, "qty_per_100": 2.2, "note": "វីតាមីន A, ជាតិដែក"},
                    {"item_name": "ផ្លែល្ពៅ", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 17.0, "qty_per_100": 1.7, "note": "វីតាមីន A, សរសៃអាហារ"},
                    {"item_name": "ល្ហុងខ្ចី", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 20.0, "qty_per_100": 2.0, "note": "សារធាតុរ៉ែ និងសរសៃអាហារ"},
                    {"item_name": "សណ្តែកដីលីង", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 6.0, "qty_per_100": 0.6, "note": "ប្រូតេអ៊ីនរុក្ខជាតិ"},
                    {"item_name": "អង្ករចម្រុះ", "category": "អង្ករ", "unit": "1គីឡូ", "gram_per_student": 100.0, "qty_per_100": 10.0, "note": "ស្បៀងគោលផ្តល់ថាមពល"},
                    {"item_name": "ប្រេងឆា", "category": "ប្រេងឆា", "unit": "1គីឡូ", "gram_per_student": 10.0, "qty_per_100": 1.0, "note": "ខ្លាញ់ផ្តល់ថាមពល"},
                    {"item_name": "អំបិលអ៊ីយូត", "category": "អំបិល", "unit": "1គីឡូ", "gram_per_student": 3.0, "qty_per_100": 0.3, "note": "សារធាតុអ៊ីយ៉ូតការពារជំងឺពកក"}
                ]
            },
            {
                "day": "អង្គារ",
                "dish_name": "បាយឆាសណ្តែកកួរនិងស៊ុតទា",
                "meal_type": "អាហារពេលព្រឹក",
                "description": "បាយឆាស៊ុតទាលាយសណ្ដែកគួរ ផ្តល់ប្រូតេអ៊ីនខ្ពស់ និងវីតាមីន ងាយស្រួលចម្អិន សិស្សចូលចិត្ត",
                "ingredients": [
                    {"item_name": "ស៊ុតទា", "category": "ត្រី សាច់ ស៊ុត", "unit": "1គ្រាប់", "gram_per_student": 0.35, "qty_per_100": 35.0, "note": "ប្រូតេអ៊ីន និងសារធាតុចិញ្ចឹមចាំបាច់"},
                    {"item_name": "សណ្តែកគួរ", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 64.0, "qty_per_100": 6.4, "note": "វីតាមីន C, ជាតិដែក, កាល់ស្យូម"},
                    {"item_name": "អង្ករចម្រុះ", "category": "អង្ករ", "unit": "1គីឡូ", "gram_per_student": 100.0, "qty_per_100": 10.0, "note": "ស្បៀងគោល"},
                    {"item_name": "ប្រេងឆា", "category": "ប្រេងឆា", "unit": "1គីឡូ", "gram_per_student": 12.0, "qty_per_100": 1.2, "note": "ខ្លាញ់សម្រាប់ឆា"},
                    {"item_name": "អំបិលអ៊ីយូត", "category": "អំបិល", "unit": "1គីឡូ", "gram_per_student": 3.0, "qty_per_100": 0.3, "note": "គ្រឿងផ្សំរសជាតិ"}
                ]
            },
            {
                "day": "ពុធ",
                "dish_name": "ស្ងោរស្ពៃសាច់ជ្រូក",
                "meal_type": "អាហារពេលព្រឹក",
                "description": "សម្លស្ងោរស្ពៃក្តៅៗ សម្បូរសារធាតុរ៉ែ វីតាមីន និងប្រូតេអ៊ីន ជួយដល់ការលូតលាស់រាងកាយ",
                "ingredients": [
                    {"item_name": "សាច់ជ្រូក៣ជាន់", "category": "ត្រី សាច់ ស៊ុត", "unit": "1គីឡូ", "gram_per_student": 22.0, "qty_per_100": 2.2, "note": "ប្រូតេអ៊ីន"},
                    {"item_name": "ស្ពៃក្រញាញ់", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 64.0, "qty_per_100": 6.4, "note": "វីតាមីន A, C និងជាតិសរសៃ"},
                    {"item_name": "អង្ករចម្រុះ", "category": "អង្ករ", "unit": "1គីឡូ", "gram_per_student": 100.0, "qty_per_100": 10.0, "note": "ស្បៀងគោល"},
                    {"item_name": "ប្រេងឆា", "category": "ប្រេងឆា", "unit": "1គីឡូ", "gram_per_student": 8.0, "qty_per_100": 0.8, "note": "បន្ថែមរសជាតិ"},
                    {"item_name": "អំបិលអ៊ីយូត", "category": "អំបិល", "unit": "1គីឡូ", "gram_per_student": 3.0, "qty_per_100": 0.3, "note": "គ្រឿងផ្សំ"}
                ]
            },
            {
                "day": "ព្រហស្បតិ៍",
                "dish_name": "សម្លម្ជូរគ្រឿងត្រីអណ្ដែង",
                "meal_type": "អាហារពេលព្រឹក",
                "description": "សម្លម្ជូរគ្រឿងប្រពៃណីខ្មែរ ប្រើប្រាស់ត្រីអណ្ដែង និងត្រកួនស្រស់ៗពីសហគមន៍",
                "ingredients": [
                    {"item_name": "ត្រីអណ្តែង", "category": "ត្រី សាច់ ស៊ុត", "unit": "1គីឡូ", "gram_per_student": 24.0, "qty_per_100": 2.4, "note": "ប្រូតេអ៊ីន និងអាស៊ីតខ្លាញ់អូមេហ្គា"},
                    {"item_name": "ត្រកួន", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 52.0, "qty_per_100": 5.2, "note": "ជាតិដែក និងវីតាមីន"},
                    {"item_name": "អង្ករចម្រុះ", "category": "អង្ករ", "unit": "1គីឡូ", "gram_per_student": 100.0, "qty_per_100": 10.0, "note": "ស្បៀងគោល"},
                    {"item_name": "ប្រេងឆា", "category": "ប្រេងឆា", "unit": "1គីឡូ", "gram_per_student": 8.0, "qty_per_100": 0.8, "note": "គ្រឿងឆា"},
                    {"item_name": "អំបិលអ៊ីយូត", "category": "អំបិល", "unit": "1គីឡូ", "gram_per_student": 3.0, "qty_per_100": 0.3, "note": "គ្រឿងផ្សំ"}
                ]
            },
            {
                "day": "សុក្រ",
                "dish_name": "សម្លម្ជូរយួនត្រីផ្ទក់ (ត្រឡាច និងប៉េងប៉ោះ)",
                "meal_type": "អាហារពេលព្រឹក",
                "description": "សម្លម្ជូរយួនត្រីផ្ទក់ជាមួយផ្លែត្រឡាច និងប៉េងប៉ោះ រសជាតិជូរអែម ស្រួលពិសា សម្បូរវីតាមីន C",
                "ingredients": [
                    {"item_name": "ត្រីផ្ទក់", "category": "ត្រី សាច់ ស៊ុត", "unit": "1គីឡូ", "gram_per_student": 30.0, "qty_per_100": 3.0, "note": "ប្រូតេអ៊ីនត្រីទឹកសាបស្រស់"},
                    {"item_name": "ផ្លែត្រឡាច", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 33.0, "qty_per_100": 3.3, "note": "ជាតិទឹក និងជាតិរ៉ែ"},
                    {"item_name": "ប៉េងប៉ោះ", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 31.0, "qty_per_100": 3.1, "note": "វីតាមីន C និងសារធាតុប្រឆាំងអុកស៊ីតកម្ម"},
                    {"item_name": "អង្ករចម្រុះ", "category": "អង្ករ", "unit": "1គីឡូ", "gram_per_student": 100.0, "qty_per_100": 10.0, "note": "ស្បៀងគោល"},
                    {"item_name": "ប្រេងឆា", "category": "ប្រេងឆា", "unit": "1គីឡូ", "gram_per_student": 8.0, "qty_per_100": 0.8, "note": "បន្ថែមរសជាតិ"},
                    {"item_name": "អំបិលអ៊ីយូត", "category": "អំបិល", "unit": "1គីឡូ", "gram_per_student": 3.0, "qty_per_100": 0.3, "note": "គ្រឿងផ្សំ"}
                ]
            },
            {
                "day": "សៅរ៍",
                "dish_name": "ឆាល្ពៅពងទា",
                "meal_type": "អាហារពេលព្រឹក",
                "description": "ឆាល្ពៅលឿងទន់ជាមួយពងទា រសជាតិផ្អែមធម្មជាតិ ផ្តល់ថាមពល និងវីតាមីន A ខ្ពស់ជំនួយភ្នែកកុមារ",
                "ingredients": [
                    {"item_name": "ស៊ុតទា", "category": "ត្រី សាច់ ស៊ុត", "unit": "1គ្រាប់", "gram_per_student": 0.35, "qty_per_100": 35.0, "note": "ប្រូតេអ៊ីន"},
                    {"item_name": "ផ្លែល្ពៅ", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 64.0, "qty_per_100": 6.4, "note": "វីតាមីន A ខ្ពស់ជំនួយភ្នែក"},
                    {"item_name": "អង្ករចម្រុះ", "category": "អង្ករ", "unit": "1គីឡូ", "gram_per_student": 100.0, "qty_per_100": 10.0, "note": "ស្បៀងគោល"},
                    {"item_name": "ប្រេងឆា", "category": "ប្រេងឆា", "unit": "1គីឡូ", "gram_per_student": 12.0, "qty_per_100": 1.2, "note": "ខ្លាញ់សម្រាប់ឆា"},
                    {"item_name": "អំបិលអ៊ីយូត", "category": "អំបិល", "unit": "1គីឡូ", "gram_per_student": 3.0, "qty_per_100": 0.3, "note": "គ្រឿងផ្សំ"}
                ]
            },
            {
                "day": "អាទិត្យ",
                "dish_name": "បបរគ្រឿងសាច់ជ្រូក និងសណ្ដែកបណ្ដុះ",
                "meal_type": "អាហារពេលព្រឹក",
                "description": "បបរគ្រឿងសាច់ជ្រូកក្តៅៗជាមួយសណ្ដែកបណ្ដុះ និងស្លឹកខ្ទឹម ងាយស្រួលញ៉ាំ ផ្តល់ថាមពលខ្ពស់",
                "ingredients": [
                    {"item_name": "សាច់ជ្រូក៣ជាន់", "category": "ត្រី សាច់ ស៊ុត", "unit": "1គីឡូ", "gram_per_student": 22.0, "qty_per_100": 2.2, "note": "ប្រូតេអ៊ីន"},
                    {"item_name": "សណ្តែកបណ្តុះ", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 30.0, "qty_per_100": 3.0, "note": "វីតាមីន C និងជាតិសរសៃ"},
                    {"item_name": "ស្លឹកខ្ទឹម", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 10.0, "qty_per_100": 1.0, "note": "គ្រឿងបន្ថែមរសជាតិ"},
                    {"item_name": "អង្ករចម្រុះ", "category": "អង្ករ", "unit": "1គីឡូ", "gram_per_student": 100.0, "qty_per_100": 10.0, "note": "ស្បៀងគោល"},
                    {"item_name": "ប្រេងឆា", "category": "ប្រេងឆា", "unit": "1គីឡូ", "gram_per_student": 8.0, "qty_per_100": 0.8, "note": "បន្ថែមរសជាតិ"},
                    {"item_name": "អំបិលអ៊ីយូត", "category": "អំបិល", "unit": "1គីឡូ", "gram_per_student": 3.0, "qty_per_100": 0.3, "note": "គ្រឿងផ្សំ"}
                ]
            }
        ]
    },
    "cycle_2": {
        "id": "cycle_2",
        "name": "គំរូស្ដង់ដារ MoEYS (វដ្តទី២ - សប្ដាហ៍ទី២ និងទី៤)",
        "description": "គំរូមុខម្ហូបឆ្លាស់ ផ្លាស់ប្តូររសជាតិ និងធានាអាហារូបត្ថម្ភចម្រុះតាមគោលការណ៍ WFP & MoEYS",
        "dishes": [
            {
                "day": "ចន្ទ",
                "dish_name": "សម្លកកូរត្រីអណ្តែង",
                "meal_type": "អាហារពេលព្រឹក",
                "description": "សម្លកកូរត្រីអណ្ដែងជាមួយស្លឹកងប់ ផ្លែល្ពៅ ននោងមូល និងសណ្ដែកដី",
                "ingredients": [
                    {"item_name": "ត្រីអណ្តែង", "category": "ត្រី សាច់ ស៊ុត", "unit": "1គីឡូ", "gram_per_student": 24.0, "qty_per_100": 2.4, "note": "ប្រូតេអ៊ីនត្រី"},
                    {"item_name": "ស្លឹកងប់", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 20.0, "qty_per_100": 2.0, "note": "វីតាមីន និងសារធាតុរ៉ែ"},
                    {"item_name": "ផ្លែល្ពៅ", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 17.0, "qty_per_100": 1.7, "note": "វីតាមីន A"},
                    {"item_name": "ននោងមូល", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 20.0, "qty_per_100": 2.0, "note": "ជាតិទឹក"},
                    {"item_name": "សណ្តែកដីលីង", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 6.0, "qty_per_100": 0.6, "note": "ប្រូតេអ៊ីនរុក្ខជាតិ"},
                    {"item_name": "អង្ករចម្រុះ", "category": "អង្ករ", "unit": "1គីឡូ", "gram_per_student": 100.0, "qty_per_100": 10.0, "note": "ស្បៀងគោល"},
                    {"item_name": "ប្រេងឆា", "category": "ប្រេងឆា", "unit": "1គីឡូ", "gram_per_student": 10.0, "qty_per_100": 1.0, "note": "ខ្លាញ់"},
                    {"item_name": "អំបិលអ៊ីយូត", "category": "អំបិល", "unit": "1គីឡូ", "gram_per_student": 3.0, "qty_per_100": 0.3, "note": "គ្រឿងផ្សំ"}
                ]
            },
            {
                "day": "អង្គារ",
                "dish_name": "ឆាននោងនិងស៊ុតទា",
                "meal_type": "អាហារពេលព្រឹក",
                "description": "ឆាននោងជ្រុងជាមួយស៊ុតទា ផ្អែមស្រួយ ងាយស្រួលរំលាយអាហារ",
                "ingredients": [
                    {"item_name": "ស៊ុតទា", "category": "ត្រី សាច់ ស៊ុត", "unit": "1គ្រាប់", "gram_per_student": 0.35, "qty_per_100": 35.0, "note": "ប្រូតេអ៊ីន"},
                    {"item_name": "ននោងជ្រុង", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 60.0, "qty_per_100": 6.0, "note": "ជាតិសរសៃ និងជាតិទឹក"},
                    {"item_name": "អង្ករចម្រុះ", "category": "អង្ករ", "unit": "1គីឡូ", "gram_per_student": 100.0, "qty_per_100": 10.0, "note": "ស្បៀងគោល"},
                    {"item_name": "ប្រេងឆា", "category": "ប្រេងឆា", "unit": "1គីឡូ", "gram_per_student": 12.0, "qty_per_100": 1.2, "note": "ខ្លាញ់"},
                    {"item_name": "អំបិលអ៊ីយូត", "category": "អំបិល", "unit": "1គីឡូ", "gram_per_student": 3.0, "qty_per_100": 0.3, "note": "គ្រឿងផ្សំ"}
                ]
            },
            {
                "day": "ពុធ",
                "dish_name": "ស្ងោរស្ពៃចង្កឹះត្រីប្រា",
                "meal_type": "អាហារពេលព្រឹក",
                "description": "ស្ងោរស្ពៃចង្កឹះត្រីប្រា ផ្តល់ជាតិខ្លាញ់ត្រីល្អ និងកាល់ស្យូមពីបន្លែស្ពៃ",
                "ingredients": [
                    {"item_name": "ត្រីប្រា", "category": "ត្រី សាច់ ស៊ុត", "unit": "1គីឡូ", "gram_per_student": 25.0, "qty_per_100": 2.5, "note": "ប្រូតេអ៊ីន និងខ្លាញ់ត្រី"},
                    {"item_name": "ស្ពៃចង្កឹះ", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 60.0, "qty_per_100": 6.0, "note": "វីតាមីន និងកាល់ស្យូម"},
                    {"item_name": "អង្ករចម្រុះ", "category": "អង្ករ", "unit": "1គីឡូ", "gram_per_student": 100.0, "qty_per_100": 10.0, "note": "ស្បៀងគោល"},
                    {"item_name": "ប្រេងឆា", "category": "ប្រេងឆា", "unit": "1គីឡូ", "gram_per_student": 8.0, "qty_per_100": 0.8, "note": "បន្ថែម"},
                    {"item_name": "អំបិលអ៊ីយូត", "category": "អំបិល", "unit": "1គីឡូ", "gram_per_student": 3.0, "qty_per_100": 0.3, "note": "គ្រឿងផ្សំ"}
                ]
            },
            {
                "day": "ព្រហស្បតិ៍",
                "dish_name": "ឆាបន្លែគ្រប់មុខសាច់ជ្រូក",
                "meal_type": "អាហារពេលព្រឹក",
                "description": "ឆាបន្លែគ្រប់មុខ (ផ្កាខាត់ណា ការ៉ុត ស្ពៃ) លាយសាច់ជ្រូក សម្បូរដោយវីតាមីនចម្រុះពណ៌",
                "ingredients": [
                    {"item_name": "សាច់ជ្រូក៣ជាន់", "category": "ត្រី សាច់ ស៊ុត", "unit": "1គីឡូ", "gram_per_student": 22.0, "qty_per_100": 2.2, "note": "ប្រូតេអ៊ីន"},
                    {"item_name": "ផ្កាខាត់ណា", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 30.0, "qty_per_100": 3.0, "note": "វីតាមីន C, K"},
                    {"item_name": "ការ៉ុត", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 25.0, "qty_per_100": 2.5, "note": "បេតាការ៉ូទីន វីតាមីន A"},
                    {"item_name": "ស្ពៃជើងទា", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 20.0, "qty_per_100": 2.0, "note": "ជាតិសរសៃ"},
                    {"item_name": "អង្ករចម្រុះ", "category": "អង្ករ", "unit": "1គីឡូ", "gram_per_student": 100.0, "qty_per_100": 10.0, "note": "ស្បៀងគោល"},
                    {"item_name": "ប្រេងឆា", "category": "ប្រេងឆា", "unit": "1គីឡូ", "gram_per_student": 12.0, "qty_per_100": 1.2, "note": "ខ្លាញ់សម្រាប់ឆា"},
                    {"item_name": "អំបិលអ៊ីយូត", "category": "អំបិល", "unit": "1គីឡូ", "gram_per_student": 3.0, "qty_per_100": 0.3, "note": "គ្រឿងផ្សំ"}
                ]
            },
            {
                "day": "សុក្រ",
                "dish_name": "សម្លម្ជូរគ្រឿងត្រីប្រា",
                "meal_type": "អាហារពេលព្រឹក",
                "description": "សម្លម្ជូរគ្រឿងត្រីប្រាជាមួយត្រួយល្ពៅ និងត្រកួនស្រស់ រសជាតិឈ្ងុយឆ្ងាញ់",
                "ingredients": [
                    {"item_name": "ត្រីប្រា", "category": "ត្រី សាច់ ស៊ុត", "unit": "1គីឡូ", "gram_per_student": 25.0, "qty_per_100": 2.5, "note": "ប្រូតេអ៊ីន"},
                    {"item_name": "ត្រួយល្ពៅ", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 30.0, "qty_per_100": 3.0, "note": "វីតាមីន និងសារធាតុរ៉ែ"},
                    {"item_name": "ត្រកួន", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 30.0, "qty_per_100": 3.0, "note": "ជាតិដែក"},
                    {"item_name": "អង្ករចម្រុះ", "category": "អង្ករ", "unit": "1គីឡូ", "gram_per_student": 100.0, "qty_per_100": 10.0, "note": "ស្បៀងគោល"},
                    {"item_name": "ប្រេងឆា", "category": "ប្រេងឆា", "unit": "1គីឡូ", "gram_per_student": 8.0, "qty_per_100": 0.8, "note": "គ្រឿង"},
                    {"item_name": "អំបិលអ៊ីយូត", "category": "អំបិល", "unit": "1គីឡូ", "gram_per_student": 3.0, "qty_per_100": 0.3, "note": "គ្រឿងផ្សំ"}
                ]
            },
            {
                "day": "សៅរ៍",
                "dish_name": "បាយឆាពងទាការ៉ុត",
                "meal_type": "អាហារពេលព្រឹក",
                "description": "បាយឆាស៊ុតទាលាយការ៉ុត និងសណ្តែកគួរ ពណ៌ស្រស់ស្អាត ទាក់ទាញចំណង់អាហារសិស្ស",
                "ingredients": [
                    {"item_name": "ស៊ុតទា", "category": "ត្រី សាច់ ស៊ុត", "unit": "1គ្រាប់", "gram_per_student": 0.35, "qty_per_100": 35.0, "note": "ប្រូតេអ៊ីន"},
                    {"item_name": "ការ៉ុត", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 30.0, "qty_per_100": 3.0, "note": "វីតាមីន A"},
                    {"item_name": "សណ្តែកគួរ", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 30.0, "qty_per_100": 3.0, "note": "វីតាមីន និងជាតិសរសៃ"},
                    {"item_name": "អង្ករចម្រុះ", "category": "អង្ករ", "unit": "1គីឡូ", "gram_per_student": 100.0, "qty_per_100": 10.0, "note": "ស្បៀងគោល"},
                    {"item_name": "ប្រេងឆា", "category": "ប្រេងឆា", "unit": "1គីឡូ", "gram_per_student": 12.0, "qty_per_100": 1.2, "note": "ខ្លាញ់"},
                    {"item_name": "អំបិលអ៊ីយូត", "category": "អំបិល", "unit": "1គីឡូ", "gram_per_student": 3.0, "qty_per_100": 0.3, "note": "គ្រឿងផ្សំ"}
                ]
            },
            {
                "day": "អាទិត្យ",
                "dish_name": "ឆាត្រកួនសាច់ជ្រូក",
                "meal_type": "អាហារពេលព្រឹក",
                "description": "ឆាត្រកួនស្រស់សាច់ជ្រូក រសជាតិឈ្ងុយឆ្ងាញ់ សម្បូរជាតិដែក និងវីតាមីន A",
                "ingredients": [
                    {"item_name": "សាច់ជ្រូក៣ជាន់", "category": "ត្រី សាច់ ស៊ុត", "unit": "1គីឡូ", "gram_per_student": 22.0, "qty_per_100": 2.2, "note": "ប្រូតេអ៊ីន"},
                    {"item_name": "ត្រកួន", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 60.0, "qty_per_100": 6.0, "note": "ជាតិដែក និងវីតាមីន"},
                    {"item_name": "អង្ករចម្រុះ", "category": "អង្ករ", "unit": "1គីឡូ", "gram_per_student": 100.0, "qty_per_100": 10.0, "note": "ស្បៀងគោល"},
                    {"item_name": "ប្រេងឆា", "category": "ប្រេងឆា", "unit": "1គីឡូ", "gram_per_student": 12.0, "qty_per_100": 1.2, "note": "ខ្លាញ់សម្រាប់ឆា"},
                    {"item_name": "អំបិលអ៊ីយូត", "category": "អំបិល", "unit": "1គីឡូ", "gram_per_student": 3.0, "qty_per_100": 0.3, "note": "គ្រឿងផ្សំ"}
                ]
            }
        ]
    }
}

# ================= គំរូមុខម្ហូបស្ដង់ដារ MoEYS SFIS តាមថ្ងៃ =================
SFIS_CATEGORIES = ["ត្រី/សាច់/ស៊ុត", "បន្លែ", "គ្រឿងទេស", "ប្រេងឆា", "អង្ករ", "អំបិល"]

SFIS_PRESET_DISHES = {
    "សម្លកកូរសាច់ជ្រូក": {
        "dish_name": "សម្លកកូរសាច់ជ្រូក",
        "meal_type": "អាហារពេលព្រឹក",
        "ingredients": [
            {"item_name": "សាច់ជ្រូក៣ជាន់", "category": "ត្រី/សាច់/ស៊ុត", "unit": "1គីឡូ", "gram_per_student": 22.0, "qty_per_100": 2.2},
            {"item_name": "ស្លឹកបាស", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 22.0, "qty_per_100": 2.2},
            {"item_name": "ផ្លែល្ពៅ", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 17.0, "qty_per_100": 1.7},
            {"item_name": "ល្ហុងខ្ចី", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 20.0, "qty_per_100": 2.0},
            {"item_name": "សណ្តែកដីលីង", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 6.0, "qty_per_100": 0.6},
        ]
    },
    "បាយឆាសណ្តែកកួរនិងស៊ុតទា": {
        "dish_name": "បាយឆាសណ្តែកកួរនិងស៊ុតទា",
        "meal_type": "អាហារពេលព្រឹក",
        "ingredients": [
            {"item_name": "ស៊ុតទា", "category": "ត្រី/សាច់/ស៊ុត", "unit": "1គ្រាប់", "gram_per_student": 0.35, "qty_per_100": 35.0},
            {"item_name": "សណ្តែកគួរ", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 64.0, "qty_per_100": 6.4},
        ]
    },
    "ស្ងោរស្ពៃសាច់ជ្រូក": {
        "dish_name": "ស្ងោរស្ពៃសាច់ជ្រូក",
        "meal_type": "អាហារពេលព្រឹក",
        "ingredients": [
            {"item_name": "សាច់ជ្រូក៣ជាន់", "category": "ត្រី/សាច់/ស៊ុត", "unit": "1គីឡូ", "gram_per_student": 22.0, "qty_per_100": 2.2},
            {"item_name": "ស្ពៃក្រញាញ់", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 64.0, "qty_per_100": 6.4},
        ]
    },
    "សម្លម្ជូរគ្រឿងត្រីអណ្ដែង": {
        "dish_name": "សម្លម្ជូរគ្រឿងត្រីអណ្ដែង",
        "meal_type": "អាហារពេលព្រឹក",
        "ingredients": [
            {"item_name": "ត្រីអណ្តែង", "category": "ត្រី/សាច់/ស៊ុត", "unit": "1គីឡូ", "gram_per_student": 24.0, "qty_per_100": 2.4},
            {"item_name": "ត្រកួន", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 52.0, "qty_per_100": 5.2},
        ]
    },
    "សម្លម្ជូរយួនត្រីផ្ទក់ (ត្រឡាច និងប៉េងប៉ោះ)": {
        "dish_name": "សម្លម្ជូរយួនត្រីផ្ទក់ (ត្រឡាច និងប៉េងប៉ោះ)",
        "meal_type": "អាហារពេលព្រឹក",
        "ingredients": [
            {"item_name": "ត្រីផ្ទក់", "category": "ត្រី/សាច់/ស៊ុត", "unit": "1គីឡូ", "gram_per_student": 30.0, "qty_per_100": 3.0},
            {"item_name": "ផ្លែត្រឡាច", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 33.0, "qty_per_100": 3.3},
            {"item_name": "ប៉េងប៉ោះ", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 31.0, "qty_per_100": 3.1},
        ]
    },
    "ឆាល្ពៅពងទា": {
        "dish_name": "ឆាល្ពៅពងទា",
        "meal_type": "អាហារពេលព្រឹក",
        "ingredients": [
            {"item_name": "ស៊ុតទា", "category": "ត្រី/សាច់/ស៊ុត", "unit": "1គ្រាប់", "gram_per_student": 0.35, "qty_per_100": 35.0},
            {"item_name": "ផ្លែល្ពៅ", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 64.0, "qty_per_100": 6.4},
        ]
    },
    "បបរគ្រឿងសាច់ជ្រូក និងសណ្ដែកបណ្ដុះ": {
        "dish_name": "បបរគ្រឿងសាច់ជ្រូក និងសណ្ដែកបណ្ដុះ",
        "meal_type": "អាហារពេលព្រឹក",
        "ingredients": [
            {"item_name": "សាច់ជ្រូក៣ជាន់", "category": "ត្រី/សាច់/ស៊ុត", "unit": "1គីឡូ", "gram_per_student": 22.0, "qty_per_100": 2.2},
            {"item_name": "សណ្តែកបណ្តុះ", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 30.0, "qty_per_100": 3.0},
            {"item_name": "ស្លឹកខ្ទឹម", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 10.0, "qty_per_100": 1.0},
        ]
    },
    "សម្លកកូរត្រីអណ្តែង": {
        "dish_name": "សម្លកកូរត្រីអណ្តែង",
        "meal_type": "អាហារពេលព្រឹក",
        "ingredients": [
            {"item_name": "ត្រីអណ្តែង", "category": "ត្រី/សាច់/ស៊ុត", "unit": "1គីឡូ", "gram_per_student": 24.0, "qty_per_100": 2.4},
            {"item_name": "ស្លឹកងប់", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 20.0, "qty_per_100": 2.0},
            {"item_name": "ផ្លែល្ពៅ", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 17.0, "qty_per_100": 1.7},
            {"item_name": "ននោងមូល", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 20.0, "qty_per_100": 2.0},
            {"item_name": "សណ្តែកដីលីង", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 6.0, "qty_per_100": 0.6},
        ]
    },
    "ឆាននោងនិងស៊ុតទា": {
        "dish_name": "ឆាននោងនិងស៊ុតទា",
        "meal_type": "អាហារពេលព្រឹក",
        "ingredients": [
            {"item_name": "ស៊ុតទា", "category": "ត្រី/សាច់/ស៊ុត", "unit": "1គ្រាប់", "gram_per_student": 0.35, "qty_per_100": 35.0},
            {"item_name": "ននោងជ្រុង", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 60.0, "qty_per_100": 6.0},
        ]
    },
    "ស្ងោរស្ពៃចង្កឹះត្រីប្រា": {
        "dish_name": "ស្ងោរស្ពៃចង្កឹះត្រីប្រា",
        "meal_type": "អាហារពេលព្រឹក",
        "ingredients": [
            {"item_name": "ត្រីប្រា", "category": "ត្រី/សាច់/ស៊ុត", "unit": "1គីឡូ", "gram_per_student": 25.0, "qty_per_100": 2.5},
            {"item_name": "ស្ពៃចង្កឹះ", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 60.0, "qty_per_100": 6.0},
        ]
    },
    "ឆាបន្លែគ្រប់មុខសាច់ជ្រូក": {
        "dish_name": "ឆាបន្លែគ្រប់មុខសាច់ជ្រូក",
        "meal_type": "អាហារពេលព្រឹក",
        "ingredients": [
            {"item_name": "សាច់ជ្រូក៣ជាន់", "category": "ត្រី/សាច់/ស៊ុត", "unit": "1គីឡូ", "gram_per_student": 22.0, "qty_per_100": 2.2},
            {"item_name": "ផ្កាខាត់ណា", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 30.0, "qty_per_100": 3.0},
            {"item_name": "ការ៉ុត", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 25.0, "qty_per_100": 2.5},
            {"item_name": "ស្ពៃជើងទា", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 20.0, "qty_per_100": 2.0},
        ]
    },
    "សម្លម្ជូរគ្រឿងត្រីប្រា": {
        "dish_name": "សម្លម្ជូរគ្រឿងត្រីប្រា",
        "meal_type": "អាហារពេលព្រឹក",
        "ingredients": [
            {"item_name": "ត្រីប្រា", "category": "ត្រី/សាច់/ស៊ុត", "unit": "1គីឡូ", "gram_per_student": 25.0, "qty_per_100": 2.5},
            {"item_name": "ត្រួយល្ពៅ", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 30.0, "qty_per_100": 3.0},
            {"item_name": "ត្រកួន", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 30.0, "qty_per_100": 3.0},
        ]
    },
    "បាយឆាពងទាការ៉ុត": {
        "dish_name": "បាយឆាពងទាការ៉ុត",
        "meal_type": "អាហារពេលព្រឹក",
        "ingredients": [
            {"item_name": "ស៊ុតទា", "category": "ត្រី/សាច់/ស៊ុត", "unit": "1គ្រាប់", "gram_per_student": 0.35, "qty_per_100": 35.0},
            {"item_name": "ការ៉ុត", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 30.0, "qty_per_100": 3.0},
            {"item_name": "សណ្តែកគួរ", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 30.0, "qty_per_100": 3.0},
        ]
    },
    "ឆាត្រកួនសាច់ជ្រូក": {
        "dish_name": "ឆាត្រកួនសាច់ជ្រូក",
        "meal_type": "អាហារពេលព្រឹក",
        "ingredients": [
            {"item_name": "សាច់ជ្រូក៣ជាន់", "category": "ត្រី/សាច់/ស៊ុត", "unit": "1គីឡូ", "gram_per_student": 22.0, "qty_per_100": 2.2},
            {"item_name": "ត្រកួន", "category": "បន្លែ", "unit": "1គីឡូ", "gram_per_student": 60.0, "qty_per_100": 6.0},
        ]
    }
}

DEFAULT_DAY_PRESET_DISHES = {
    "ចន្ទ": "សម្លកកូរសាច់ជ្រូក",
    "អង្គារ": "បាយឆាសណ្តែកកួរនិងស៊ុតទា",
    "ពុធ": "ស្ងោរស្ពៃសាច់ជ្រូក",
    "ព្រហស្បតិ៍": "សម្លម្ជូរគ្រឿងត្រីអណ្ដែង",
    "សុក្រ": "សម្លម្ជូរយួនត្រីផ្ទក់ (ត្រឡាច និងប៉េងប៉ោះ)",
    "សៅរ៍": "ឆាល្ពៅពងទា",
    "អាទិត្យ": "បបរគ្រឿងសាច់ជ្រូក និងសណ្ដែកបណ្ដុះ"
}

DEFAULT_DAY_PRESETS = {
    d: SFIS_PRESET_DISHES.get(dish, {"dish_name": dish, "meal_type": "អាហារពេលព្រឹក", "ingredients": []})
    for d, dish in DEFAULT_DAY_PRESET_DISHES.items()
}

KHMER_NUM_MAP = str.maketrans("0123456789", "០១២៣៤៥៦៧៨៩")
def to_khmer_digits(val):
    """បំប្លែងលេខអារ៉ាប់ទៅជាលេខខ្មែរ"""
    return str(val).translate(KHMER_NUM_MAP)

KHMER_MONTHS = [
    "មករា", "កុម្ភៈ", "មីនា", "មេសា", "ឧសភា", "មិថុនា",
    "កក្កដា", "សីហា", "កញ្ញា", "តុលា", "វិច្ឆិកា", "ធ្នូ"
]
KHMER_MONTH_TO_NUM = {m: i + 1 for i, m in enumerate(KHMER_MONTHS)}

def get_month_weekday_dates(year, month):
    """
    គណនាកាលបរិច្ឆេទទាំងអស់ក្នុងខែ តាមថ្ងៃនៃសប្ដាហ៍នីមួយៗ (ចន្ទ ដល់ អាទិត្យ)
    ដូចទម្រង់ជាក់ស្តែងក្នុង MoEYS SFIS
    """
    import calendar
    res = {d: [] for d in KHMER_DAYS_OF_WEEK}
    _, num_days = calendar.monthrange(year, month)
    m_name = KHMER_MONTHS[month - 1] if 1 <= month <= 12 else str(month)

    for d_num in range(1, num_days + 1):
        dt = date(year, month, d_num)
        w_name = KHMER_WEEKDAY_MAP.get(dt.weekday())
        if w_name in res:
            kh_lbl = f"{to_khmer_digits(d_num)} {m_name} {to_khmer_digits(year)}"
            res[w_name].append({
                "date": dt,
                "date_str": dt.strftime("%Y-%m-%d"),
                "day_num": d_num,
                "label": kh_lbl,
                "label_kh": kh_lbl
            })
    return res

def get_or_create_school_voucher(conn, school_name, date_str):
    """ទាញយក ឬបង្កើតលេខសក្ខីប័ត្រស្វ័យប្រវត្ត ចាប់ផ្ដើមពី 001 តាមសាលានីមួយៗ"""
    if not school_name:
        return "001"
    c = conn.cursor()
    d_str = str(date_str).strip()[:10]
    row = c.execute(
        "SELECT voucher_no FROM daily_records WHERE school_name=? AND date=? AND voucher_no IS NOT NULL AND TRIM(voucher_no) != '' LIMIT 1",
        (school_name, d_str)
    ).fetchone()
    if row and row[0]:
        return str(row[0]).strip()

    rows = c.execute(
        "SELECT DISTINCT voucher_no FROM daily_records WHERE school_name=? AND voucher_no IS NOT NULL AND TRIM(voucher_no) != ''",
        (school_name,)
    ).fetchall()
    nums = []
    for r in rows:
        val = str(r[0]).strip()
        matches = re.findall(r'\d+', val)
        if matches:
            nums.append(int(matches[-1]))
    next_val = max(nums) + 1 if nums else 1
    return f"{next_val:03d}"

def determine_phase_for_date(conn, school_name, date_val):
    """កំណត់វគ្គ (វគ្គ១ ឬ វគ្គ២) ស្វ័យប្រវត្តតាមកាលបរិច្ឆេទ"""
    c = conn.cursor()
    d_str = str(date_val)[:10]
    try:
        row = c.execute("""
            SELECT phase1_start, phase1_end, phase2_start, phase2_end
            FROM suppliers
            WHERE school_name=? OR school_name LIKE ?
            LIMIT 1
        """, (school_name, f"%{school_name}%")).fetchone()
        if not row:
            row = c.execute("SELECT phase1_start, phase1_end, phase2_start, phase2_end FROM suppliers LIMIT 1").fetchone()
        if row:
            p1_s, p1_e, p2_s, p2_e = row
            if p1_s and p1_e and str(p1_s) <= d_str <= str(p1_e):
                return "វគ្គ១"
            if p2_s and p2_e and str(p2_s) <= d_str <= str(p2_e):
                return "វគ្គ២"
    except Exception:
        pass
    try:
        day_num = int(d_str.split("-")[2])
        return "វគ្គ១" if day_num <= 15 else "វគ្គ២"
    except Exception:
        return "វគ្គ១"

def save_school_daily_requirements(conn, school_name, act_comm, act_dist, act_prov, days_menu_data, staple_data):
    """
    រក្សាទុកមុខម្ហូបប្រចាំថ្ងៃ (school_menus & menu_ingredients)
    និងបង្កើត/បញ្ចូលទិន្នន័យទៅក្នុងតារាងតម្រូវការស្បៀងប្រចាំថ្ងៃ (daily_records)
    ដោយស្វ័យប្រវត្តិសម្រាប់សាលានេះ តាមកាលបរិច្ឆេទដែលបានជ្រើសរើសក្នុងខែ
    ព្រមទាំងស្បៀងគោលទុកបានយូរនៅថ្ងៃដើមខែ។
    """
    if not school_name or school_name in ["-- ជ្រើសរើសសាលារៀន --", "គ្មានសាលា"]:
        return {"success": False, "message": "សូមជ្រើសរើសសាលារៀនជាមុនសិន!"}

    c = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    total_inserted_daily = 0
    total_dates_recorded = set()
    total_month_cost = 0.0

    # Ensure columns exist in daily_records
    c.execute("PRAGMA table_info(daily_records)")
    dr_cols = [col[1] for col in c.fetchall()]
    if "category" not in dr_cols:
        try:
            c.execute("ALTER TABLE daily_records ADD COLUMN category TEXT DEFAULT ''")
        except Exception:
            pass
    if "menu_name" not in dr_cols:
        try:
            c.execute("ALTER TABLE daily_records ADD COLUMN menu_name TEXT DEFAULT ''")
        except Exception:
            pass

    # ១. រក្សាទុកមុខម្ហូបប្រចាំសប្ដាហ៍ (៧ ថ្ងៃ) ចូលក្នុង school_menus និង menu_ingredients
    for d_name, d_cfg in days_menu_data.items():
        dish_name = d_cfg.get("dish_name", "").strip()
        if not dish_name:
            continue
        meal_type = d_cfg.get("meal_type", "អាហារពេលព្រឹក")
        target_st = int(d_cfg.get("target_students", 100))
        ingredients = d_cfg.get("ingredients", [])

        # Delete existing menu definition for this day
        old_ids = [r[0] for r in c.execute("SELECT id FROM school_menus WHERE school_name=? AND day_of_week=?", (school_name, d_name)).fetchall()]
        if old_ids:
            c.executemany("DELETE FROM menu_ingredients WHERE menu_id=?", [(oid,) for oid in old_ids])
            c.execute("DELETE FROM school_menus WHERE school_name=? AND day_of_week=?", (school_name, d_name))

        c.execute("""
            INSERT INTO school_menus (
                school_name, commune, district, province,
                menu_name, day_of_week, meal_type, target_students,
                cycle_week, notes, is_active, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'កាលវិភាគ SFIS', 'បង្កើតតាមទម្រង់ SFIS', 1, ?, ?)
        """, (school_name, act_comm, act_dist, act_prov, dish_name, d_name, meal_type, target_st, now_str, now_str))
        new_m_id = c.lastrowid

        # Insert menu ingredients
        for ing in ingredients:
            u_p = float(ing.get("unit_price") or 0)
            if u_p <= 0:
                u_p = get_active_item_price(conn, ing["item_name"], school_name, act_comm)
            g_st = float(ing.get("gram_per_student") or 0)
            unit_n = ing.get("unit", "1គីឡូ")
            
            if "គ្រាប់" in unit_n:
                d_qty = round(target_st * g_st, 1)
            else:
                d_qty = round((target_st * g_st) / 1000.0, 2)
                if d_qty <= 0:
                    d_qty = float(ing.get("total_qty") or 1.0)
            t_c = round(d_qty * u_p, 2)

            c.execute("""
                INSERT INTO menu_ingredients (
                    menu_id, item_name, category, unit,
                    gram_per_student, total_qty, unit_price, total_cost, note
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (new_m_id, ing["item_name"], ing.get("category", "បន្លែ"), unit_n, g_st, d_qty, u_p, t_c, ing.get("note", "")))

        # ២. បញ្ចូលតម្រូវការស្បៀងប្រចាំថ្ងៃ (daily_records) តាមកាលបរិច្ឆេទដែលបានធីក (Checked Dates)
        active_dates = d_cfg.get("active_dates", d_cfg.get("checked_dates", []))
        for dt_info in active_dates:
            if not dt_info.get("enabled", True):
                continue
            date_str = str(dt_info.get("date", dt_info.get("date_str", ""))).strip()
            if not date_str:
                continue
            cur_students = int(dt_info.get("students", target_st))
            v_no = get_or_create_school_voucher(conn, school_name, date_str)
            phase_val = determine_phase_for_date(conn, school_name, date_str)
            total_dates_recorded.add(date_str)

            for ing in ingredients:
                itm_name = ing["item_name"].strip()
                cat_name = ing.get("category", "បន្លែ")
                unit_n = ing.get("unit", "1គីឡូ")
                g_st = float(ing.get("gram_per_student") or 0)
                
                # If staple food is managed as monthly bulk, exclude daily duplicate
                staple_is_active = staple_data.get("include", staple_data.get("enabled", False))
                if staple_is_active and cat_name in ["អង្ករ", "ប្រេងឆា", "អំបិល"]:
                    continue

                if "គ្រាប់" in unit_n:
                    rec_qty = round(cur_students * g_st, 1)
                else:
                    rec_qty = round((cur_students * g_st) / 1000.0, 2)
                    if rec_qty <= 0:
                        rec_qty = float(ing.get("qty_per_day", ing.get("total_qty", 1.0)))

                u_p = float(ing.get("unit_price") or 0)
                if u_p <= 0:
                    u_p = get_active_item_price(conn, itm_name, school_name, act_comm)
                tot_p = round(rec_qty * u_p, 2)
                total_month_cost += tot_p

                # Upsert into daily_records
                ex_row = c.execute("""
                    SELECT id FROM daily_records 
                    WHERE school_name=? AND date=? AND item_name=?
                """, (school_name, date_str, itm_name)).fetchone()

                if ex_row:
                    c.execute("""
                        UPDATE daily_records 
                        SET quantity=?, unit_price=?, total_price=?, phase=?, voucher_no=?, consumption_date=?, category=?, menu_name=?
                        WHERE id=?
                    """, (rec_qty, u_p, tot_p, phase_val, v_no, date_str, cat_name, dish_name, ex_row[0]))
                else:
                    c.execute("""
                        INSERT INTO daily_records (
                            date, school_name, item_name, phase, quantity, unit_price, total_price, voucher_no, consumption_date, category, menu_name
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (date_str, school_name, itm_name, phase_val, rec_qty, u_p, tot_p, v_no, date_str, cat_name, dish_name))
                total_inserted_daily += 1

    # ៣. បញ្ចូលស្បៀងគោលទុកបានយូរ នៅថ្ងៃដើមខែ (Non-perishable Staples at Beginning of Month)
    if staple_data.get("include", staple_data.get("enabled", True)):
        staple_date = str(staple_data.get("delivery_date", "")).strip()
        if staple_date:
            staple_vno = get_or_create_school_voucher(conn, school_name, staple_date)
            staple_phase = determine_phase_for_date(conn, school_name, staple_date)
            total_dates_recorded.add(staple_date)

            for st_itm in staple_data.get("items", []):
                s_name = st_itm["item_name"].strip()
                s_cat = st_itm.get("category", "ស្បៀងគោល")
                s_qty = float(st_itm.get("quantity", st_itm.get("qty", 0.0)))
                s_price = float(st_itm.get("unit_price", 0.0))
                if s_price <= 0:
                    s_price = get_active_item_price(conn, s_name, school_name, act_comm)
                s_tot = round(s_qty * s_price, 2)
                total_month_cost += s_tot

                if s_qty > 0:
                    ex_st = c.execute("""
                        SELECT id FROM daily_records 
                        WHERE school_name=? AND date=? AND item_name=?
                    """, (school_name, staple_date, s_name)).fetchone()

                    if ex_st:
                        c.execute("""
                            UPDATE daily_records 
                            SET quantity=?, unit_price=?, total_price=?, phase=?, voucher_no=?, consumption_date=?, category=?, menu_name='ស្បៀងគោលប្រចាំខែ'
                            WHERE id=?
                        """, (s_qty, s_price, s_tot, staple_phase, staple_vno, staple_date, s_cat, ex_st[0]))
                    else:
                        c.execute("""
                            INSERT INTO daily_records (
                                date, school_name, item_name, phase, quantity, unit_price, total_price, voucher_no, consumption_date, category, menu_name
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'ស្បៀងគោលប្រចាំខែ')
                        """, (staple_date, school_name, s_name, staple_phase, s_qty, s_price, s_tot, staple_vno, staple_date, s_cat))
                    total_inserted_daily += 1

    conn.commit()
    return {
        "success": True,
        "total_records": total_inserted_daily,
        "total_dates": len(total_dates_recorded),
        "total_cost": total_month_cost,
        "school_name": school_name
    }


def save_school_staples_only(conn, school_name, delivery_date, staple_items):
    """
    កត់ត្រា និងរក្សាទុកស្បៀងទុកបានយូរ (អង្ករ, ប្រេងឆា, អំបិល, ទឹកត្រី) ចូលក្នុង daily_records
    """
    if not school_name or school_name in ["-- ជ្រើសរើសសាលារៀន --", "គ្មានសាលា"]:
        return {"success": False, "message": "សូមជ្រើសរើសសាលារៀនជាមុនសិន!"}
    c = conn.cursor()
    d_str = str(delivery_date).strip()[:10]
    v_no = get_or_create_school_voucher(conn, school_name, d_str)
    phase = determine_phase_for_date(conn, school_name, d_str)

    c.execute("PRAGMA table_info(daily_records)")
    dr_cols = [col[1] for col in c.fetchall()]
    if "category" not in dr_cols:
        try:
            c.execute("ALTER TABLE daily_records ADD COLUMN category TEXT DEFAULT ''")
        except Exception:
            pass
    if "menu_name" not in dr_cols:
        try:
            c.execute("ALTER TABLE daily_records ADD COLUMN menu_name TEXT DEFAULT ''")
        except Exception:
            pass

    inserted_count = 0
    total_cost = 0.0

    for it in staple_items:
        s_name = str(it.get("item_name", "")).strip()
        if not s_name:
            continue
        s_cat = str(it.get("category", "")).strip() or auto_classify_category(s_name)
        s_qty = float(it.get("quantity") or 0.0)
        s_price = float(it.get("unit_price") or 0.0)
        s_tot = round(s_qty * s_price, 2)
        total_cost += s_tot

        ex_st = c.execute(
            "SELECT id FROM daily_records WHERE school_name=? AND date=? AND item_name=?",
            (school_name, d_str, s_name)
        ).fetchone()

        if ex_st:
            c.execute("""
                UPDATE daily_records 
                SET quantity=?, unit_price=?, total_price=?, phase=?, voucher_no=?, consumption_date=?, category=?, menu_name='ស្បៀងទុកបានយូរ'
                WHERE id=?
            """, (s_qty, s_price, s_tot, phase, v_no, d_str, s_cat, ex_st[0]))
        else:
            c.execute("""
                INSERT INTO daily_records (
                    date, school_name, item_name, phase, quantity, unit_price, total_price, voucher_no, consumption_date, category, menu_name
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'ស្បៀងទុកបានយូរ')
            """, (d_str, school_name, s_name, phase, s_qty, s_price, s_tot, v_no, d_str, s_cat))
        inserted_count += 1

    conn.commit()
    return {
        "success": True,
        "school_name": school_name,
        "delivery_date": d_str,
        "inserted_count": inserted_count,
        "total_cost": total_cost,
        "voucher_no": v_no,
        "phase": phase
    }


def get_school_daily_records(conn, school_name, year, month):
    """ទាញយកទិន្នន័យតម្រូវការស្បៀងប្រចាំថ្ងៃរបស់សាលាក្នុងខែដែលបានជ្រើសរើស"""
    if not school_name or school_name in ["-- ជ្រើសរើសសាលារៀន --", "គ្មានសាលា"]:
        return []
    c = conn.cursor()
    m_str = f"{year:04d}-{month:02d}%"
    rows = c.execute("""
        SELECT date, school_name, item_name, phase, quantity, unit_price, total_price, voucher_no, consumption_date, category, menu_name
        FROM daily_records
        WHERE school_name=? AND date LIKE ?
        ORDER BY date ASC, id ASC
    """, (school_name, m_str)).fetchall()
    
    result = []
    for r in rows:
        d_val = r[0]
        try:
            dt_obj = datetime.strptime(d_val, "%Y-%m-%d")
            w_name = KHMER_WEEKDAY_MAP.get(dt_obj.weekday(), "")
        except Exception:
            w_name = ""
        result.append({
            "date": r[0],
            "day_name": w_name,
            "school_name": r[1],
            "item_name": r[2],
            "phase": r[3] or "វគ្គ១",
            "quantity": float(r[4] or 0),
            "unit_price": float(r[5] or 0),
            "total_price": float(r[6] or 0),
            "voucher_no": r[7] or "",
            "consumption_date": r[8] or r[0],
            "category": r[9] or "ផ្សេងៗ",
            "menu_name": r[10] or ""
        })
    return result

def generate_daily_requirements_excel(conn, school_name, year, month, records):
    """បង្កើតឯកសារ Excel តារាងតម្រូវការស្បៀងប្រចាំថ្ងៃតាមសាលារៀន"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "តម្រូវការស្បៀងប្រចាំថ្ងៃ"

    font_title = Font(name="Khmer OS Muol Light", size=13, bold=True, color="002060")
    font_sub = Font(name="Khmer OS Muol Light", size=10, bold=True, color="000000")
    font_header = Font(name="Khmer OS Siemreap", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="Khmer OS Siemreap", size=9)
    font_bold = Font(name="Khmer OS Siemreap", size=9, bold=True)
    
    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_total = PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='B0C4DE'),
        right=Side(style='thin', color='B0C4DE'),
        top=Side(style='thin', color='B0C4DE'),
        bottom=Side(style='thin', color='B0C4DE')
    )

    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')

    # Header
    ws.merge_cells("A1:I1")
    ws["A1"] = "ព្រះរាជាណាចក្រកម្ពុជា"
    ws["A1"].font = font_sub
    ws["A1"].alignment = align_center

    ws.merge_cells("A2:I2")
    ws["A2"] = "ជាតិ សាសនា ព្រះមហាក្សត្រ"
    ws["A2"].font = font_sub
    ws["A2"].alignment = align_center

    m_name = KHMER_MONTHS[month - 1] if 1 <= month <= 12 else str(month)
    ws.merge_cells("A4:I4")
    ws["A4"] = f"តារាងតម្រូវការស្បៀងប្រចាំថ្ងៃ - សាលាបឋមសិក្សា៖ {school_name}"
    ws["A4"].font = font_title
    ws["A4"].alignment = align_center

    ws.merge_cells("A5:I5")
    ws["A5"] = f"សម្រាប់ខែ៖ {m_name} ឆ្នាំ {year}"
    ws["A5"].font = font_bold
    ws["A5"].alignment = align_center

    headers = ["ល.រ", "កាលបរិច្ឆេទ", "ថ្ងៃនៃសប្ដាហ៍", "មុខទំនិញ/ស្បៀង", "ប្រភេទ", "បរិមាណ", "តម្លៃរាយ (៛)", "សរុបទឹកប្រាក់ (៛)", "លេខសក្ខីប័ត្រ"]
    for col_idx, h_text in enumerate(headers, 1):
        cell = ws.cell(7, col_idx, h_text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border

    ws.row_dimensions[7].height = 26

    curr_row = 8
    tot_cost = 0.0
    for idx, r in enumerate(records, 1):
        ws.cell(curr_row, 1, idx).alignment = align_center
        ws.cell(curr_row, 2, r["date"]).alignment = align_center
        ws.cell(curr_row, 3, r["day_name"]).alignment = align_center
        ws.cell(curr_row, 4, r["item_name"]).alignment = align_left
        ws.cell(curr_row, 5, r.get("category", "")).alignment = align_center
        ws.cell(curr_row, 6, r["quantity"]).alignment = align_right
        ws.cell(curr_row, 7, r["unit_price"]).alignment = align_right
        ws.cell(curr_row, 8, r["total_price"]).alignment = align_right
        ws.cell(curr_row, 9, r.get("voucher_no", "")).alignment = align_center

        for c_idx in range(1, 10):
            ws.cell(curr_row, c_idx).font = font_data
            ws.cell(curr_row, c_idx).border = thin_border
            ws.cell(curr_row, 7).number_format = "#,##0"
            ws.cell(curr_row, 8).number_format = "#,##0"
            ws.cell(curr_row, 6).number_format = "#,##0.00"

        tot_cost += r["total_price"]
        curr_row += 1

    # Total row
    ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=7)
    tot_cell = ws.cell(curr_row, 1, "សរុបថវិកាស្បៀងប្រចាំខែទាំងអស់ ៖")
    tot_cell.font = font_bold
    tot_cell.alignment = align_right
    tot_cell.fill = fill_total

    val_cell = ws.cell(curr_row, 8, tot_cost)
    val_cell.font = Font(name="Khmer OS Siemreap", size=10, bold=True, color="B91C1C")
    val_cell.alignment = align_right
    val_cell.fill = fill_total
    val_cell.number_format = "#,##0"

    ws.cell(curr_row, 9, "").fill = fill_total
    for c_idx in range(1, 10):
        ws.cell(curr_row, c_idx).border = thin_border

    col_widths = {1: 8, 2: 14, 3: 14, 4: 26, 5: 16, 6: 14, 7: 16, 8: 18, 9: 14}
    for c_idx, w in col_widths.items():
        col_letter = openpyxl.utils.get_column_letter(c_idx)
        ws.column_dimensions[col_letter].width = w

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ================= មុខងារវិភាគ និងស្រង់ទិន្នន័យពីឯកសារស្បៀង (PDF, Excel, Word, Image) =================
def auto_classify_category(item_name):
    """កំណត់ប្រភេទមុខទំនិញស្វ័យប្រវត្តិ (ត្រី/សាច់/ស៊ុត, បន្លែ, គ្រឿងទេស, អង្ករ, ប្រេងឆា, អំបិល)"""
    n = str(item_name).strip()
    if any(k in n for k in ["សាច់", "ត្រី", "ស៊ុត", "មាន់", "ទា", "ប្រហុក", "ងៀត", "ក្រៀម"]):
        return "ត្រី/សាច់/ស៊ុត"
    elif any(k in n for k in ["ស្ពៃ", "ត្រកួន", "ល្ពៅ", "ការ៉ុត", "ត្រឡាច", "ប៉េងប៉ោះ", "សណ្តែក", "ស្លឹក", "ននោង", "ខាត់ណា", "ឆៃថាវ", "ត្រប់", "ផ្សិត", "គល់ស្លឹកគ្រៃ", "រំដេង", "ម្ទេស", "ខ្ញី"]):
        return "បន្លែ"
    elif "អង្ករ" in n:
        return "អង្ករ"
    elif "ប្រេង" in n:
        return "ប្រេងឆា"
    elif "អំបិល" in n:
        return "អំបិល"
    elif any(k in n for k in ["ទឹកត្រី", "ទឹកស៊ីអ៊ីវ", "ស្ករ", "ប៊ីចេង", "ខ្ទឹម", "ម្សៅស៊ុប", "ម្ទេស"]):
        return "គ្រឿងទេស"
    return "បន្លែ"


def clean_school_name(raw):
    s = re.sub(r'\.(?:xlsx|xls|docx|doc|pdf|csv|png|jpg|jpeg|webp)$', '', str(raw), flags=re.I)
    s = s.replace('_', ' ').replace('-', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def extract_school_from_texts(cell_texts, filename="", known_schools=None):
    """សម្គាល់ឈ្មោះសាលារៀនដោយស្វ័យប្រវត្តិតាមឈ្មោះសាលាផ្លូវការ ឬតាមពាក្យគន្លឹះក្នុងឯកសារ"""
    all_str = "\n".join(str(c) for c in cell_texts if c) + "\n" + str(filename or "")
    clean_all = clean_school_name(all_str)
    compact_all = re.sub(r'[\s_\-]+', '', all_str)

    if known_schools:
        # 1. Exact or cleaned string match
        for s in sorted(known_schools, key=len, reverse=True):
            if s and (s in all_str or s in clean_all):
                return s, "ស្គាល់តាមឈ្មោះសាលាដែលមានក្នុងប្រព័ន្ធ"
        # 2. Match ignoring all spaces and punctuation
        for s in sorted(known_schools, key=len, reverse=True):
            s_compact = re.sub(r'[\s_\-]+', '', s)
            if s_compact and s_compact in compact_all:
                return s, "ស្គាល់តាមឈ្មោះសាលាដែលមានក្នុងប្រព័ន្ធ"
        # 3. Match core name (e.g. 'វត្តបូព៌' inside 'សាលាបឋមសិក្សា វត្តបូព៌')
        for s in sorted(known_schools, key=len, reverse=True):
            core = re.sub(r'^(?:សាលាបឋមសិក្សា|សាលា|បឋមសិក្សា)\s*', '', s).strip()
            if core and len(core) >= 3 and (core in all_str or core in clean_all):
                return s, "ស្គាល់តាមឈ្មោះសាលាដែលមានក្នុងប្រព័ន្ធ"

    # Fallback to regex search
    m = re.search(r'(?:សាលាបឋមសិក្សា|បឋមសិក្សា|សាលា)\s*[:៖\-_\s]?\s*([^\n\r,\|()]+?)(?:\s+(?:សម្រាប់|ខែ|ឆ្នាំ|ឃុំ|ស្រុក|ខេត្ត|\.)|[\n\r,\|()_\.]|$)', all_str)
    if m:
        raw_name = m.group(1).strip()
        c_name = clean_school_name(raw_name)
        if c_name and len(c_name) >= 2:
            full_name = f"សាលាបឋមសិក្សា {c_name}" if not c_name.startswith("សាលា") else c_name
            if known_schools:
                for s in known_schools:
                    if c_name in s or s in full_name:
                        return s, "ស្គាល់តាមឈ្មោះសាលាដែលមានក្នុងប្រព័ន្ធ"
            return full_name, "ស្គាល់តាមពាក្យគន្លឹះក្នុងឯកសារ"

    return "", "មិនបានសម្គាល់"


def extract_month_year_from_texts(cell_texts, filename=""):
    """សម្គាល់ខែ និងឆ្នាំ ពីឯកសារ"""
    all_str = "\n".join(str(c) for c in cell_texts if c) + "\n" + str(filename or "")
    found_m = 11
    found_y = 2025
    for m_idx, m_name in enumerate(KHMER_MONTHS, 1):
        if m_name in all_str:
            found_m = m_idx
            break
    y_m = re.search(r'(?:ឆ្នាំ|year)?\s*(202[4-9]|203[0-5])', all_str)
    if y_m:
        found_y = int(y_m.group(1))
    else:
        y_kh = re.search(r'(?:ឆ្នាំ\s*)?([២][០][២-៣][០-៩])', all_str)
        if y_kh:
            kh_y_str = y_kh.group(1)
            ar_y = kh_y_str.translate(str.maketrans("០១២៣៤៥៦៧៨៩", "0123456789"))
            found_y = int(ar_y)
    return found_m, found_y


def _find_table_header(raw_rows):
    """ស្វែងរកជួរឈរ header នៃតារាងស្បៀង"""
    header_idx = -1
    col_map = {}
    for idx, r in enumerate(raw_rows[:25]):
        r_str = [str(c).strip().lower() for c in r if c is not None]
        has_item = any(any(k in s for k in ["មុខទំនិញ", "ឈ្មោះទំនិញ", "ទំនិញ", "ស្បៀង", "គ្រឿងផ្សំ", "item"]) for s in r_str)
        has_qty = any(any(k in s for k in ["បរិមាណ", "ចំនួន", "qty", "quantity"]) for s in r_str)
        if has_item or (has_qty and len(r_str) >= 3):
            header_idx = idx
            for col_i, c in enumerate(r):
                if c is None:
                    continue
                s = str(c).strip().lower()
                if any(k in s for k in ["កាលបរិច្ឆេទ", "កាលបរិច្ឆេត", "កាលបរិច្ឆែទ", "ថ្ងៃទី", "date"]):
                    col_map["date"] = col_i
                elif any(k in s for k in ["មុខទំនិញ", "ឈ្មោះទំនិញ", "ទំនិញ", "ស្បៀង", "គ្រឿងផ្សំ", "item"]):
                    col_map["item_name"] = col_i
                elif any(k in s for k in ["ប្រភេទ", "category"]):
                    col_map["category"] = col_i
                elif any(k in s for k in ["ឯកតា", "unit"]):
                    col_map["unit"] = col_i
                elif any(k in s for k in ["បរិមាណ", "ចំនួន", "ទម្ងន់", "qty", "quantity"]):
                    if "date" not in col_map or col_map.get("quantity") is None:
                        col_map["quantity"] = col_i
                elif any(k in s for k in ["តម្លៃរាយ", "តម្លៃឯកតា", "តម្លៃ", "price"]):
                    col_map["unit_price"] = col_i
                elif any(k in s for k in ["សរុប", "ទឹកប្រាក់", "total", "amount"]):
                    col_map["total_price"] = col_i
                elif any(k in s for k in ["សក្ខីប័ត្រ", "ប័ណ្ណ", "voucher"]):
                    col_map["voucher_no"] = col_i
                elif any(k in s for k in ["មុខម្ហូប", "ម្ហូប", "menu"]):
                    col_map["menu_name"] = col_i
            break
    return header_idx, col_map


def _extract_rows_with_map(data_rows, col_map, def_year=2025, def_month=11):
    """បម្លែងជួរទិន្នន័យនៃតារាងជា List of Dicts"""
    items = []
    item_col = col_map.get("item_name")
    if item_col is None:
        return items
    for r in data_rows:
        if len(r) <= item_col:
            continue
        val = r[item_col]
        if not val or not str(val).strip():
            continue
        item_str = str(val).strip()
        if any(sk in item_str.lower() for sk in ["សរុប", "total", "ល.រ", "មុខទំនិញ"]):
            continue
        d_val = ""
        if "date" in col_map and len(r) > col_map["date"] and r[col_map["date"]]:
            raw_d = str(r[col_map["date"]]).strip()
            m_ymd = re.search(r'\d{4}-\d{2}-\d{2}', raw_d)
            if m_ymd:
                d_val = m_ymd.group(0)
            elif re.search(r'^\d{1,2}$', raw_d):
                d_val = f"{def_year:04d}-{def_month:02d}-{int(raw_d):02d}"
        if not d_val:
            d_val = f"{def_year:04d}-{def_month:02d}-01"
        qty = 0.0
        if "quantity" in col_map and len(r) > col_map["quantity"] and r[col_map["quantity"]]:
            try:
                q_clean = re.sub(r'[^\d.]', '', str(r[col_map["quantity"]]))
                qty = float(q_clean) if q_clean else 0.0
            except Exception:
                qty = 0.0
        price = 0.0
        if "unit_price" in col_map and len(r) > col_map["unit_price"] and r[col_map["unit_price"]]:
            try:
                p_clean = re.sub(r'[^\d.]', '', str(r[col_map["unit_price"]]))
                price = float(p_clean) if p_clean else 0.0
            except Exception:
                price = 0.0
        tot = 0.0
        if "total_price" in col_map and len(r) > col_map["total_price"] and r[col_map["total_price"]]:
            try:
                t_clean = re.sub(r'[^\d.]', '', str(r[col_map["total_price"]]))
                tot = float(t_clean) if t_clean else 0.0
            except Exception:
                tot = 0.0
        if tot == 0 and qty > 0 and price > 0:
            tot = round(qty * price, 2)
        elif price == 0 and qty > 0 and tot > 0:
            price = round(tot / qty, 2)
        cat = ""
        if "category" in col_map and len(r) > col_map["category"] and r[col_map["category"]]:
            cat = str(r[col_map["category"]]).strip()
        if not cat:
            cat = auto_classify_category(item_str)
        unit_str = "1គីឡូ"
        if "unit" in col_map and len(r) > col_map["unit"] and r[col_map["unit"]]:
            unit_str = str(r[col_map["unit"]]).strip()
        elif "ស៊ុត" in item_str or "ពង" in item_str:
            unit_str = "1គ្រាប់"
        elif "ប្រេង" in item_str or "ទឹកត្រី" in item_str or "ទឹកស៊ីអ៊ីវ" in item_str:
            unit_str = "លីត្រ"
        v_no = ""
        if "voucher_no" in col_map and len(r) > col_map["voucher_no"] and r[col_map["voucher_no"]]:
            v_no = str(r[col_map["voucher_no"]]).strip()
        m_name = ""
        if "menu_name" in col_map and len(r) > col_map["menu_name"] and r[col_map["menu_name"]]:
            m_name = str(r[col_map["menu_name"]]).strip()
        items.append({
            "date": d_val,
            "item_name": item_str,
            "category": cat,
            "unit": unit_str,
            "quantity": qty,
            "unit_price": price,
            "total_price": tot,
            "voucher_no": v_no,
            "menu_name": m_name
        })
    return items


def parse_uploaded_food_file(uploaded_file, known_schools=None, default_school=""):
    """
    វិភាគឯកសារស្បៀងដែលបានបញ្ចូល (PDF, Excel, Word, CSV, Image)
    ស្គាល់ឈ្មោះសាលា ខែ ឆ្នាំ និងទាញយកបញ្ជីមុខទំនិញស្បៀងប្រចាំថ្ងៃ
    """
    fname = getattr(uploaded_file, "name", "file.xlsx")
    ext = fname.split(".")[-1].lower() if "." in fname else ""
    file_bytes = uploaded_file.read() if hasattr(uploaded_file, "read") else uploaded_file

    cell_texts = []
    parsed_rows = []

    if ext in ["xlsx", "xls"]:
        try:
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
            sheet = wb.active
            raw_rows = []
            for r in sheet.iter_rows(values_only=True):
                if any(r):
                    raw_rows.append(list(r))
                    for c in r:
                        if c is not None:
                            cell_texts.append(str(c))
            header_idx, col_map = _find_table_header(raw_rows)
            if header_idx >= 0:
                parsed_rows = _extract_rows_with_map(raw_rows[header_idx + 1:], col_map)
        except Exception as e:
            cell_texts.append(str(e))

    elif ext == "csv":
        try:
            content_str = file_bytes.decode('utf-8', errors='ignore')
            reader = csv.reader(io.StringIO(content_str))
            raw_rows = [r for r in reader if any(r)]
            for r in raw_rows:
                for c in r:
                    if c:
                        cell_texts.append(str(c))
            header_idx, col_map = _find_table_header(raw_rows)
            if header_idx >= 0:
                parsed_rows = _extract_rows_with_map(raw_rows[header_idx + 1:], col_map)
        except Exception as e:
            cell_texts.append(str(e))

    elif ext in ["docx", "doc"]:
        if docx:
            try:
                doc = docx.Document(io.BytesIO(file_bytes))
                for p in doc.paragraphs:
                    if p.text:
                        cell_texts.append(p.text)
                for t in doc.tables:
                    raw_rows = []
                    for row in t.rows:
                        cells = [c.text.strip() for c in row.cells]
                        raw_rows.append(cells)
                        for c in cells:
                            if c:
                                cell_texts.append(c)
                    header_idx, col_map = _find_table_header(raw_rows)
                    if header_idx >= 0:
                        parsed_rows.extend(_extract_rows_with_map(raw_rows[header_idx + 1:], col_map))
            except Exception as e:
                cell_texts.append(str(e))

    elif ext == "pdf":
        if pdfplumber:
            try:
                with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                    for page in pdf.pages:
                        txt = page.extract_text()
                        if txt:
                            cell_texts.append(txt)
                        tables = page.extract_tables()
                        if tables:
                            for t in tables:
                                raw_rows = [[str(c).strip() if c is not None else "" for c in r] for r in t if any(r)]
                                header_idx, col_map = _find_table_header(raw_rows)
                                if header_idx >= 0:
                                    parsed_rows.extend(_extract_rows_with_map(raw_rows[header_idx + 1:], col_map))
            except Exception as e:
                cell_texts.append(str(e))

    elif ext in ["png", "jpg", "jpeg", "webp"]:
        cell_texts.append(fname)

    detected_school, school_conf = extract_school_from_texts(cell_texts, fname, known_schools)
    if not detected_school and default_school:
        detected_school = default_school
        school_conf = "ប្រើប្រាស់សាលាដែលកំពុងជ្រើសរើស"

    detected_month, detected_year = extract_month_year_from_texts(cell_texts, fname)

    return {
        "success": True,
        "filename": fname,
        "file_type": ext,
        "file_bytes": file_bytes,
        "detected_school": detected_school,
        "school_confidence": school_conf,
        "detected_month": detected_month,
        "detected_year": detected_year,
        "items": parsed_rows,
        "raw_text_snippet": " | ".join(cell_texts[:15]) if cell_texts else ""
    }


def init_menu_db(conn):
    """បង្កើតតារាងសម្រាប់រក្សាទុកបញ្ជីមុខម្ហូប និងគ្រឿងផ្សំ"""
    c = conn.cursor()
    c.execute("""
    CREATE TABLE IF NOT EXISTS school_menus (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        school_name TEXT NOT NULL,
        commune TEXT DEFAULT '',
        district TEXT DEFAULT '',
        province TEXT DEFAULT '',
        menu_name TEXT NOT NULL,
        day_of_week TEXT NOT NULL,
        meal_type TEXT DEFAULT 'អាហារពេលព្រឹក',
        target_students INTEGER DEFAULT 100,
        cycle_week TEXT DEFAULT 'រៀងរាល់សប្ដាហ៍',
        notes TEXT DEFAULT '',
        is_active INTEGER DEFAULT 1,
        created_at TEXT,
        updated_at TEXT
    )""")

    c.execute("""
    CREATE TABLE IF NOT EXISTS menu_ingredients (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        menu_id INTEGER NOT NULL,
        item_name TEXT NOT NULL,
        category TEXT DEFAULT '',
        unit TEXT DEFAULT '1គីឡូ',
        gram_per_student REAL DEFAULT 0,
        total_qty REAL DEFAULT 0,
        unit_price REAL DEFAULT 0,
        total_cost REAL DEFAULT 0,
        note TEXT DEFAULT '',
        FOREIGN KEY(menu_id) REFERENCES school_menus(id) ON DELETE CASCADE
    )""")

    c.execute("CREATE INDEX IF NOT EXISTS idx_school_menus_lookup ON school_menus(school_name, day_of_week)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_menu_ingredients_menu_id ON menu_ingredients(menu_id)")

    # Ensure daily_records has category and menu_name columns
    c.execute("PRAGMA table_info(daily_records)")
    dr_cols = [col[1] for col in c.fetchall()]
    if "category" not in dr_cols:
        try:
            c.execute("ALTER TABLE daily_records ADD COLUMN category TEXT DEFAULT ''")
        except Exception:
            pass
    if "menu_name" not in dr_cols:
        try:
            c.execute("ALTER TABLE daily_records ADD COLUMN menu_name TEXT DEFAULT ''")
        except Exception:
            pass

    conn.commit()


def get_active_item_price(conn, item_name, school_name="", commune=""):
    """ទាញយកតម្លៃទំនិញបច្ចុប្បន្ន អាទិភាពតាមសាលា -> តាមឃុំ -> តម្លៃគោល Benchmark"""
    c = conn.cursor()
    # ១. តាមសាលា
    if school_name:
        row = c.execute("""
            SELECT COALESCE(price_avg, price_phase1, 0)
            FROM products
            WHERE item_name=? AND school_name=? AND price_level='school'
            LIMIT 1
        """, (item_name, school_name)).fetchone()
        if row and row[0] and float(row[0]) > 0:
            return float(row[0])

    # ២. តាមឃុំ
    if commune:
        row = c.execute("""
            SELECT COALESCE(price_avg, price_phase1, 0)
            FROM products
            WHERE item_name=? AND commune=?
            LIMIT 1
        """, (item_name, commune)).fetchone()
        if row and row[0] and float(row[0]) > 0:
            return float(row[0])

    # ៣. តាម Benchmark
    row = c.execute("""
        SELECT COALESCE(base_price, 0)
        FROM benchmark_prices
        WHERE item_name=?
        LIMIT 1
    """, (item_name,)).fetchone()
    if row and row[0] and float(row[0]) > 0:
        return float(row[0])

    # ៤. តាមតម្លៃមធ្យមទូទៅក្នុង Products
    row = c.execute("""
        SELECT AVG(COALESCE(price_avg, price_phase1, 0))
        FROM products
        WHERE item_name=? AND COALESCE(price_avg, price_phase1, 0) > 0
    """, (item_name,)).fetchone()
    if row and row[0] and float(row[0]) > 0:
        return float(row[0])

    return 0.0


def apply_template_to_school(conn, school_name, template_id="cycle_1", student_count=100, overwrite=True):
    """
    អនុវត្តគំរូស្ដង់ដារ MoEYS SFIS ជូនសាលារៀនជាក់លាក់មួយ
    គណនាបរិមាណគ្រឿងផ្សំតាមចំនួនសិស្ស និងចាប់យកតម្លៃបច្ចុប្បន្នដោយស្វ័យប្រវត្តិ
    """
    if not school_name or school_name in ["-- ជ្រើសរើសសាលារៀន --", "គ្មានសាលា"]:
        return 0

    if template_id not in STANDARD_SFIS_TEMPLATES:
        template_id = "cycle_1"

    template = STANDARD_SFIS_TEMPLATES[template_id]
    c = conn.cursor()

    # ទាញយកទីតាំងសាលា
    loc_row = c.execute("""
        SELECT province, district, commune
        FROM schools
        WHERE name=?
        LIMIT 1
    """, (school_name,)).fetchone()
    prov = loc_row[0] if loc_row else ""
    dist = loc_row[1] if loc_row else ""
    comm = loc_row[2] if loc_row else ""

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if overwrite:
        # លុបមុខម្ហូបចាស់របស់សាលានេះ
        old_ids = [r[0] for r in c.execute("SELECT id FROM school_menus WHERE school_name=?", (school_name,)).fetchall()]
        if old_ids:
            c.executemany("DELETE FROM menu_ingredients WHERE menu_id=?", [(oid,) for oid in old_ids])
            c.execute("DELETE FROM school_menus WHERE school_name=?", (school_name,))

    created_count = 0
    for dish in template["dishes"]:
        # បង្កើតមុខម្ហូប
        c.execute("""
            INSERT INTO school_menus (
                school_name, commune, district, province,
                menu_name, day_of_week, meal_type, target_students,
                cycle_week, notes, is_active, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?)
        """, (
            school_name, comm, dist, prov,
            dish["dish_name"], dish["day"], dish["meal_type"], student_count,
            template["name"], dish.get("description", ""), now_str, now_str
        ))
        menu_id = c.lastrowid

        # បញ្ចូលគ្រឿងផ្សំ
        for ing in dish["ingredients"]:
            g_student = ing.get("gram_per_student", 0.0)
            unit_name = ing.get("unit", "1គីឡូ")
            
            # គណនាបរិមាណសរុបតាមចំនួនសិស្ស
            if "គ្រាប់" in unit_name:
                calc_qty = round(student_count * g_student, 1)
            else:
                calc_qty = round((student_count * g_student) / 1000.0, 2)
                if calc_qty <= 0:
                    calc_qty = round(ing.get("qty_per_100", 1.0) * (student_count / 100.0), 2)

            u_price = get_active_item_price(conn, ing["item_name"], school_name, comm)
            tot_cost = round(calc_qty * u_price, 2)

            c.execute("""
                INSERT INTO menu_ingredients (
                    menu_id, item_name, category, unit,
                    gram_per_student, total_qty, unit_price, total_cost, note
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                menu_id, ing["item_name"], ing.get("category", ""),
                unit_name, g_student, calc_qty, u_price, tot_cost, ing.get("note", "")
            ))

        created_count += 1

    conn.commit()
    return created_count


def import_menus_from_2026_workbook(conn, workbook_path="បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm"):
    """
    ស្រង់ទិន្នន័យមុខម្ហូប និងតម្រូវការស្បៀងផ្ទាល់ពីសន្លឹក «ចំនួនសរុប ខាងកើត» នៃឯកសារ Excel ២០២៦
    """
    import os
    if not os.path.exists(workbook_path):
        return {"success": False, "message": f"រកមិនឃើញឯកសារ {workbook_path}"}

    try:
        wb = openpyxl.load_workbook(workbook_path, data_only=True)
    except Exception as e:
        return {"success": False, "message": f"កំហុសក្នុងការអានឯកសារ: {str(e)}"}

    if "ចំនួនសរុប ខាងកើត" not in wb.sheetnames:
        return {"success": False, "message": "មិនមានសន្លឹក 'ចំនួនសរុប ខាងកើត' ក្នុងឯកសារឡើយ"}

    ws = wb["ចំនួនសរុប ខាងកើត"]
    school_cols = {}
    for c in range(4, ws.max_column + 1, 3):
        s_val = ws.cell(3, c).value
        if s_val and str(s_val).strip() and "សរុប" not in str(s_val):
            sch_clean = str(s_val).strip().replace("រមៀត", "រមៀត").replace("ដង្កោរ", "ដង្កោ")
            school_cols[sch_clean] = c

    c = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    imported_summary = {}
    current_day = ""
    current_dish = ""
    rows_data = []

    for r in range(4, ws.max_row + 1):
        day_val = ws.cell(r, 1).value
        dish_val = ws.cell(r, 2).value
        item_val = ws.cell(r, 3).value

        if day_val and str(day_val).strip():
            current_day = str(day_val).strip()
        if dish_val and str(dish_val).strip():
            current_dish = str(dish_val).strip()

        if item_val and str(item_val).strip():
            item_name = str(item_val).strip()
            item_name = item_name.replace("សាច់ជ្រូក ៣ជាន់", "សាច់ជ្រូក៣ជាន់")
            item_name = item_name.replace("ត្រីអណ្ដែង", "ត្រីអណ្តែង")
            
            per_school_qty = {}
            for s_name, col_idx in school_cols.items():
                qty = ws.cell(r, col_idx).value
                days_eat = ws.cell(r, col_idx + 1).value
                month_tot = ws.cell(r, col_idx + 2).value
                try:
                    qty_f = float(qty or 0)
                except Exception:
                    qty_f = 0.0
                try:
                    days_f = float(days_eat or 4)
                except Exception:
                    days_f = 4.0
                per_school_qty[s_name] = {
                    "daily_qty": qty_f,
                    "days": days_f,
                    "monthly_qty": float(month_tot or (qty_f * days_f))
                }

            rows_data.append({
                "day": current_day,
                "dish": current_dish,
                "item": item_name,
                "school_qtys": per_school_qty
            })

    # Group by school and day
    for s_name in school_cols.keys():
        loc_row = c.execute("SELECT province, district, commune FROM schools WHERE name LIKE ? LIMIT 1", (f"%{s_name}%",)).fetchone()
        prov = loc_row[0] if loc_row else ""
        dist = loc_row[1] if loc_row else ""
        comm = loc_row[2] if loc_row else ""

        old_ids = [r[0] for r in c.execute("SELECT id FROM school_menus WHERE school_name LIKE ?", (f"%{s_name}%",)).fetchall()]
        if old_ids:
            c.executemany("DELETE FROM menu_ingredients WHERE menu_id=?", [(oid,) for oid in old_ids])
            c.execute("DELETE FROM school_menus WHERE school_name LIKE ?", (f"%{s_name}%",))

        days_grouped = {}
        for row in rows_data:
            d = row["day"]
            if d not in days_grouped:
                days_grouped[d] = {
                    "dish": row["dish"],
                    "items": []
                }
            q_info = row["school_qtys"].get(s_name, {"daily_qty": 1.0, "days": 4})
            days_grouped[d]["items"].append({
                "item": row["item"],
                "qty": q_info["daily_qty"],
                "days": q_info["days"]
            })

        for d_name, d_info in days_grouped.items():
            c.execute("""
                INSERT INTO school_menus (
                    school_name, commune, district, province,
                    menu_name, day_of_week, meal_type, target_students,
                    cycle_week, notes, is_active, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, 'អាហារពេលព្រឹក', 100, 'ឯកសារ Excel 2026', 'ស្រង់ចេញពីសន្លឹក ចំនួនសរុប ខាងកើត', 1, ?, ?)
            """, (s_name, comm, dist, prov, d_info["dish"], d_name, now_str, now_str))
            m_id = c.lastrowid

            for itm in d_info["items"]:
                cat = "ត្រី សាច់ ស៊ុត" if any(k in itm["item"] for k in ["សាច់", "ត្រី", "ស៊ុត", "ពង"]) else "បន្លែ"
                unit = "1គ្រាប់" if "ស៊ុត" in itm["item"] or "ពង" in itm["item"] else "1គីឡូ"
                u_price = get_active_item_price(conn, itm["item"], s_name, comm)
                tot_cost = round(itm["qty"] * u_price, 2)

                c.execute("""
                    INSERT INTO menu_ingredients (
                        menu_id, item_name, category, unit,
                        gram_per_student, total_qty, unit_price, total_cost, note
                    ) VALUES (?, ?, ?, ?, 0, ?, ?, ?, '')
                """, (m_id, itm["item"], cat, unit, itm["qty"], u_price, tot_cost))

        imported_summary[s_name] = len(days_grouped)

    conn.commit()
    return {
        "success": True,
        "message": f"បាននាំចូលបញ្ជីមុខម្ហូបជោគជ័យសម្រាប់ {len(imported_summary)} សាលារៀន",
        "schools": imported_summary
    }


def get_school_menu_overview(conn, school_name):
    """ទាញយកមុខម្ហូបទាំងអស់របស់សាលារៀបតាមថ្ងៃនៃសប្ដាហ៍"""
    if not school_name or school_name in ["-- ជ្រើសរើសសាលារៀន --", "គ្មានសាលា"]:
        return []

    c = conn.cursor()
    day_order = {"ចន្ទ": 1, "អង្គារ": 2, "ពុធ": 3, "ព្រហស្បតិ៍": 4, "សុក្រ": 5, "សៅរ៍": 6, "អាទិត្យ": 7}
    
    rows = c.execute("""
        SELECT id, menu_name, day_of_week, meal_type, target_students, cycle_week, notes
        FROM school_menus
        WHERE school_name=?
    """, (school_name,)).fetchall()

    if not rows:
        rows = c.execute("""
            SELECT id, menu_name, day_of_week, meal_type, target_students, cycle_week, notes
            FROM school_menus
            WHERE school_name LIKE ?
        """, (f"%{school_name}%",)).fetchall()

    result = []
    for r in rows:
        m_id = r[0]
        ing_rows = c.execute("""
            SELECT id, item_name, category, unit, gram_per_student, total_qty, unit_price, total_cost, note
            FROM menu_ingredients
            WHERE menu_id=?
            ORDER BY id ASC
        """, (m_id,)).fetchall()

        ingredients = []
        tot_day_cost = 0.0
        for ing in ing_rows:
            tot_day_cost += (ing[7] or 0.0)
            ingredients.append({
                "id": ing[0],
                "item_name": ing[1],
                "category": ing[2],
                "unit": ing[3],
                "gram_per_student": ing[4],
                "total_qty": ing[5],
                "unit_price": ing[6],
                "total_cost": ing[7],
                "note": ing[8]
            })

        result.append({
            "id": r[0],
            "menu_name": r[1],
            "day_of_week": r[2],
            "meal_type": r[3],
            "target_students": r[4],
            "cycle_week": r[5],
            "notes": r[6],
            "total_day_cost": tot_day_cost,
            "ingredients": ingredients
        })

    result.sort(key=lambda x: day_order.get(x["day_of_week"], 99))
    return result


def get_menu_by_day(conn, school_name, day_of_week):
    """ស្វែងរកមុខម្ហូបតាមថ្ងៃសម្រាប់សាលាជាក់លាក់"""
    if not school_name or school_name in ["គ្មានសាលា", "-- ជ្រើសរើសសាលារៀន --"] or not day_of_week:
        return None

    c = conn.cursor()
    row = c.execute("""
        SELECT id, menu_name, day_of_week, meal_type, target_students
        FROM school_menus
        WHERE school_name=? AND day_of_week=?
        LIMIT 1
    """, (school_name, day_of_week)).fetchone()

    if not row:
        row = c.execute("""
            SELECT id, menu_name, day_of_week, meal_type, target_students
            FROM school_menus
            WHERE school_name LIKE ? AND day_of_week=?
            LIMIT 1
        """, (f"%{school_name}%", day_of_week)).fetchone()

    if not row:
        return None

    m_id = row[0]
    ing_rows = c.execute("""
        SELECT item_name, category, unit, total_qty, unit_price, total_cost
        FROM menu_ingredients
        WHERE menu_id=?
    """, (m_id,)).fetchall()

    return {
        "id": row[0],
        "menu_name": row[1],
        "day_of_week": row[2],
        "meal_type": row[3],
        "target_students": row[4],
        "ingredients": [
            {
                "item_name": ir[0],
                "category": ir[1],
                "unit": ir[2],
                "quantity": ir[3],
                "unit_price": ir[4],
                "total_price": ir[5]
            }
            for ir in ing_rows
        ]
    }


def calculate_school_monthly_matrix(conn, school_name, days_per_month=4):
    """
    គណនាតារាងតម្រូវការស្បៀងប្រចាំខែតាមគំរូសន្លឹក «ចំនួនសរុប ខាងកើត»
    ជួរឈរ៖ ល.រ, ថ្ងៃ, មុខម្ហូប, មុខទំនិញ, ប្រភេទ, បរិមាណប្រចាំថ្ងៃ, ចំនួនថ្ងៃហូប, សរុប១ខែ, តម្លៃរាយ, សរុបទឹកប្រាក់
    """
    if not school_name or school_name in ["គ្មានសាលា", "-- ជ្រើសរើសសាលារៀន --"]:
        return {
            "rows": [],
            "grand_total_cost": 0.0,
            "total_items": 0
        }
    menus = get_school_menu_overview(conn, school_name)
    matrix_rows = []
    
    row_no = 1
    grand_total_cost = 0.0

    for m in menus:
        day_str = m["day_of_week"]
        dish_str = m["menu_name"]
        
        for idx, ing in enumerate(m["ingredients"]):
            d_qty = float(ing["total_qty"] or 0)
            u_prc = float(ing["unit_price"] or 0)
            month_qty = round(d_qty * days_per_month, 2)
            month_cost = round(month_qty * u_prc, 2)
            grand_total_cost += month_cost

            matrix_rows.append({
                "ល.រ": row_no,
                "ថ្ងៃ": day_str if idx == 0 else "",
                "មុខម្ហូប": dish_str if idx == 0 else "",
                "មុខទំនិញ": ing["item_name"],
                "ប្រភេទ": ing["category"],
                "ឯកតា": ing["unit"],
                "បរិមាណប្រចាំថ្ងៃ": d_qty,
                "ចំនួនថ្ងៃហូប": days_per_month,
                "សរុប១ខែ": month_qty,
                "តម្លៃរាយ (៛)": u_prc,
                "សរុបទឹកប្រាក់ (៛)": month_cost
            })
            row_no += 1

    return {
        "rows": matrix_rows,
        "grand_total_cost": grand_total_cost,
        "total_items": len(matrix_rows)
    }


def generate_official_menu_excel(conn, school_name, month_name="មីនា", year_num=2026, days_per_month=4):
    """
    បង្កើតឯកសារ Excel ផ្លូវការស្របតាមសន្លឹក «ចំនួនសរុប ខាងកើត» និងទម្រង់ស្ដង់ដារ MoEYS SFIS
    """
    if not school_name or school_name in ["-- ជ្រើសរើសសាលារៀន --", "គ្មានសាលា"]:
        wb = openpyxl.Workbook()
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    c = conn.cursor()
    loc_row = c.execute("""
        SELECT province, district, commune, village
        FROM schools WHERE name=? OR name LIKE ? LIMIT 1
    """, (school_name, f"%{school_name}%")).fetchone()
    
    prov = loc_row[0] if loc_row else "សៀមរាប"
    dist = loc_row[1] if loc_row else "ស្រីស្នំ"
    comm = loc_row[2] if loc_row else "ស្លែងស្ពាន"

    matrix_res = calculate_school_monthly_matrix(conn, school_name, days_per_month)
    rows = matrix_res["rows"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "តារាងមុខម្ហូប និងតម្រូវការ"

    font_title = Font(name="Khmer OS Muol Light", size=14, bold=True, color="002060")
    font_sub = Font(name="Khmer OS Muol Light", size=11, bold=True, color="000000")
    font_header = Font(name="Khmer OS Siemreap", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="Khmer OS Siemreap", size=10)
    font_bold = Font(name="Khmer OS Siemreap", size=10, bold=True)
    
    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_day_header = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    fill_total = PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='B0C4DE'),
        right=Side(style='thin', color='B0C4DE'),
        top=Side(style='thin', color='B0C4DE'),
        bottom=Side(style='thin', color='B0C4DE')
    )

    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')

    # Header section
    ws.merge_cells("A1:K1")
    ws["A1"] = "ព្រះរាជាណាចក្រកម្ពុជា"
    ws["A1"].font = font_sub
    ws["A1"].alignment = align_center

    ws.merge_cells("A2:K2")
    ws["A2"] = "ជាតិ សាសនា ព្រះមហាក្សត្រ"
    ws["A2"].font = font_sub
    ws["A2"].alignment = align_center

    ws.merge_cells("A4:K4")
    ws["A4"] = "តារាងមុខម្ហូបប្រចាំសប្ដាហ៍ និង តារាងតម្រូវការបន្លែ ត្រី សាច់ ស៊ុត"
    ws["A4"].font = font_title
    ws["A4"].alignment = align_center

    ws["A5"] = f"សាលាបឋមសិក្សា៖ {school_name}"
    ws["A5"].font = font_bold
    ws["E5"] = f"ឃុំ/សង្កាត់៖ {comm}"
    ws["E5"].font = font_bold
    ws["H5"] = f"ក្រុង/ស្រុក៖ {dist}"
    ws["H5"].font = font_bold
    ws["K5"] = f"ខេត្ត៖ {prov}"
    ws["K5"].font = font_bold

    ws["A6"] = f"សម្រាប់ខែ៖ {month_name} ឆ្នាំ {year_num}"
    ws["A6"].font = font_bold

    headers = ["ល.រ", "ថ្ងៃនៃសប្ដាហ៍", "មុខម្ហូប", "មុខទំនិញ/គ្រឿងផ្សំ", "ប្រភេទ", "ឯកតា", "បរិមាណ/ថ្ងៃ", "ចំនួនថ្ងៃហូប", "សរុប ១ខែ", "តម្លៃរាយ (៛)", "សរុបទឹកប្រាក់ (៛)"]
    for col_idx, h_text in enumerate(headers, 1):
        cell = ws.cell(8, col_idx, h_text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border

    ws.row_dimensions[8].height = 28

    curr_row = 9
    for r in rows:
        ws.cell(curr_row, 1, r["ល.រ"]).alignment = align_center
        ws.cell(curr_row, 2, r["ថ្ងៃ"]).alignment = align_center
        ws.cell(curr_row, 3, r["មុខម្ហូប"]).alignment = align_left
        ws.cell(curr_row, 4, r["មុខទំនិញ"]).alignment = align_left
        ws.cell(curr_row, 5, r["ប្រភេទ"]).alignment = align_center
        ws.cell(curr_row, 6, r["ឯកតា"]).alignment = align_center
        ws.cell(curr_row, 7, r["បរិមាណប្រចាំថ្ងៃ"]).alignment = align_right
        ws.cell(curr_row, 8, r["ចំនួនថ្ងៃហូប"]).alignment = align_center
        ws.cell(curr_row, 9, r["សរុប១ខែ"]).alignment = align_right
        ws.cell(curr_row, 10, r["តម្លៃរាយ (៛)"]).alignment = align_right
        ws.cell(curr_row, 11, r["សរុបទឹកប្រាក់ (៛)"]).alignment = align_right

        ws.cell(curr_row, 10).number_format = '#,##0'
        ws.cell(curr_row, 11).number_format = '#,##0'

        for c_idx in range(1, 12):
            c_cell = ws.cell(curr_row, c_idx)
            c_cell.font = font_data
            c_cell.border = thin_border
            if r["ថ្ងៃ"]:
                c_cell.fill = fill_day_header

        ws.row_dimensions[curr_row].height = 22
        curr_row += 1

    # Total Row
    ws.merge_cells(f"A{curr_row}:I{curr_row}")
    tot_cell = ws.cell(curr_row, 1, "សរុបថវិកាស្បៀងប្រចាំខែទាំងអស់ ៖")
    tot_cell.font = font_bold
    tot_cell.alignment = align_right

    ws.cell(curr_row, 10, "")
    sum_cell = ws.cell(curr_row, 11, f"=SUM(K9:K{curr_row-1})")
    sum_cell.font = font_bold
    sum_cell.alignment = align_right
    sum_cell.number_format = '#,##0 "៛"'

    for c_idx in range(1, 12):
        cell = ws.cell(curr_row, c_idx)
        cell.fill = fill_total
        cell.border = thin_border

    ws.row_dimensions[curr_row].height = 26

    # Signatures
    sig_row = curr_row + 3
    ws.merge_cells(f"A{sig_row}:D{sig_row}")
    ws[f"A{sig_row}"] = "បានឃើញ និងឯកភាព\nប្រធាន គមស (នាយកសាលា)"
    ws[f"A{sig_row}"].font = font_bold
    ws[f"A{sig_row}"].alignment = align_center

    ws.merge_cells(f"H{sig_row}:K{sig_row}")
    ws[f"H{sig_row}"] = "អ្នករៀបចំតារាងមុខម្ហូប\n(ចុងភៅ / នាយឃ្លាំង)"
    ws[f"H{sig_row}"].font = font_bold
    ws[f"H{sig_row}"].alignment = align_center

    col_widths = {
        'A': 6, 'B': 14, 'C': 26, 'D': 24, 'E': 14,
        'F': 10, 'G': 14, 'H': 14, 'I': 14, 'J': 16, 'K': 20
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ================= HTML RENDERING HELPER =================
def render_clean_html(html_str):
    """
    បង្ហាញ HTML ដោយសុវត្ថិភាព ធានាថាមិនមាន leading 4+ spaces ដែលបណ្តាលឱ្យ CommonMark បង្ហាញជាផ្ទាំងកូដ Error
    """
    dedented = textwrap.dedent(html_str).strip()
    cleaned_lines = [re.sub(r'^[ ]{4,}', '  ', line) for line in dedented.split('\n')]
    st.markdown('\n'.join(cleaned_lines), unsafe_allow_html=True)


# ================= STREAMLIT UI RENDER FUNCTION =================
def render_school_menu_section(conn, cursor, user_prov, user_dist, user_comm, user_school, is_admin,
                                get_scoped_district_choices, get_scoped_commune_choices, get_scoped_schools, get_school_location_info):
    """
    ផ្ទាំងគ្រប់គ្រងបញ្ជីមុខម្ហូប និងកាលវិភាគអាហារូបត្ថម្ភតាមសាលានីមួយៗ
    ស្របតាមស្តង់ដារ MoEYS SFIS (https://sfis.moeys.gov.kh)
    """
    init_menu_db(conn)

    col_t1, col_t2 = st.columns([5, 1])
    with col_t1:
        st.title("🍲 គំរូ និងបញ្ជីមុខម្ហូបតាមសាលា (MoEYS SFIS)")
        st.caption("ប្រព័ន្ធព័ត៌មានគ្រប់គ្រងកម្មវិធីផ្តល់អាហារតាមសាលារៀន MoEYS SFIS (https://sfis.moeys.gov.kh) | តារាងមុខម្ហូបប្រចាំសប្ដាហ៍ និងតម្រូវការស្បៀងតាមសាលារៀន")
    with col_t2:
        st.write("")
        if st.button("🔄 Refresh", key="btn_ref_menu_page", use_container_width=True, help="Refresh ទំព័រមុខម្ហូប"):
            st.rerun()

    render_clean_html("""
    <div style="background: linear-gradient(135deg, #1e3a8a 0%, #0369a1 100%); color: white; padding: 14px 20px; border-radius: 12px; margin-bottom: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.08);">
        <div style="font-size: 1.05rem; font-weight: bold; margin-bottom: 4px;">
            🇰🇭 កម្មវិធីផ្តល់អាហារតាមសាលារៀនដោយប្រើប្រាស់កសិផលក្នុងសហគមន៍ (HGSF) - ក្រសួងអប់រំ យុវជន និងកីឡា
        </div>
        <div style="font-size: 0.9rem; opacity: 0.92;">
            យោងតាមគោលការណ៍ណែនាំស្តីពីអាហារូបត្ថម្ភរបស់ប្រព័ន្ធ <b><a href="https://sfis.moeys.gov.kh/users/sign_in" target="_blank" style="color: #fef08a; text-decoration: underline;">MoEYS SFIS</a></b> និងកម្មវិធីស្បៀងអាហារពិភពលោក (WFP)៖ មុខម្ហូបប្រចាំសប្ដាហ៍ត្រូវបានរៀបចំជាវដ្តវិលជុំ (ចន្ទ ដល់ សៅរ៍) ដោយប្រើប្រាស់បន្លែ ត្រី សាច់ ស៊ុត ស្រស់ៗពីកសិករក្នុងសហគមន៍ ធានាបាននូវតុល្យភាពសារធាតុចិញ្ចឹមគ្រប់គ្រាន់សម្រាប់កុមារ។
        </div>
    </div>
    """)

    # ១. ជួរជ្រើសរើសទីតាំងតៗគ្នា (Cascading: Province -> District -> Commune -> School)
    col_p, col_d, col_c, col_s = st.columns([1, 1, 1, 1.2])

    with col_p:
        if not is_admin and user_prov:
            sel_prov = st.selectbox("ខេត្ត", [user_prov], key="menu_prov")
        else:
            prov_list = [r[0] for r in cursor.execute("SELECT DISTINCT province FROM locations WHERE province IS NOT NULL AND province != '' ORDER BY province").fetchall()]
            sel_prov = st.selectbox("ខេត្ត", ["-- ទាំងអស់ --"] + prov_list, key="menu_prov")

    with col_d:
        p_for_d = sel_prov if sel_prov != "-- ទាំងអស់ --" else (user_prov if not is_admin else None)
        dist_list = get_scoped_district_choices(p_for_d, prefix_all=(is_admin or not user_dist))
        if not is_admin and user_dist:
            sel_dist = st.selectbox("ក្រុង/ស្រុក", [user_dist], key="menu_dist")
        else:
            if "menu_dist" in st.session_state and st.session_state["menu_dist"] not in dist_list:
                st.session_state["menu_dist"] = dist_list[0] if dist_list else "-- ទាំងអស់ --"
            sel_dist = st.selectbox("ក្រុង/ស្រុក", dist_list, key="menu_dist")

    with col_c:
        d_for_c = sel_dist if sel_dist != "-- ទាំងអស់ --" else (user_dist if not is_admin else None)
        comm_list = get_scoped_commune_choices(p_for_d, d_for_c, prefix_all=(is_admin or not user_comm))
        if not is_admin and user_comm:
            sel_comm = st.selectbox("ឃុំ/សង្កាត់", [user_comm], key="menu_comm")
        else:
            if "menu_comm" in st.session_state and st.session_state["menu_comm"] not in comm_list:
                st.session_state["menu_comm"] = comm_list[0] if comm_list else "-- ទាំងអស់ --"
            sel_comm = st.selectbox("ឃុំ/សង្កាត់", comm_list, key="menu_comm")

    with col_s:
        c_for_s = sel_comm if sel_comm != "-- ទាំងអស់ --" else (user_comm if not is_admin else None)
        school_options = get_scoped_schools(p_for_d, d_for_c, c_for_s, prefix_all=False)
        clean_schools = [s for s in school_options if s not in ["-- ទាំងអស់ --", "គ្មានសាលា", "-- ជ្រើសរើសសាលារៀន --"]]
        school_select_list = ["-- ជ្រើសរើសសាលារៀន --"] + clean_schools if clean_schools else ["-- ជ្រើសរើសសាលារៀន --"]
        
        def_idx = 0
        if user_school and user_school in school_select_list:
            def_idx = school_select_list.index(user_school)

        if "menu_school" in st.session_state and st.session_state["menu_school"] not in school_select_list:
            st.session_state["menu_school"] = school_select_list[def_idx]

        sel_school = st.selectbox("សាលារៀន", school_select_list, key="menu_school")

    is_school_selected = bool(sel_school and sel_school not in ["-- ជ្រើសរើសសាលារៀន --", "គ្មានសាលា"])

    if is_school_selected:
        prov_db, dist_db, comm_db, vill_db = get_school_location_info(sel_school)
        act_prov = prov_db or (sel_prov if sel_prov != "-- ទាំងអស់ --" else "")
        act_dist = dist_db or (sel_dist if sel_dist != "-- ទាំងអស់ --" else "")
        act_comm = comm_db or (sel_comm if sel_comm != "-- ទាំងអស់ --" else "")
        school_menus = get_school_menu_overview(conn, sel_school)
        dish_count = len(school_menus)
        tot_week_cost = sum(m["total_day_cost"] for m in school_menus)
        tot_month_est = tot_week_cost * 4.0
        target_st = school_menus[0]["target_students"] if school_menus else 100
        display_school_title = sel_school
        badge_html = '<span style="background: #dcfce7; color: #166534; padding: 6px 12px; border-radius: 20px; font-weight: bold; font-size: 0.85rem;">✅ បានកំណត់មុខម្ហូបរួចរាល់</span>' if dish_count >= 7 else '<span style="background: #fef3c7; color: #92400e; padding: 6px 12px; border-radius: 20px; font-weight: bold; font-size: 0.85rem;">⚠️ មិនទាន់គ្រប់ ៧ ថ្ងៃ</span>' if dish_count > 0 else '<span style="background: #fee2e2; color: #991b1b; padding: 6px 12px; border-radius: 20px; font-weight: bold; font-size: 0.85rem;">❌ ពុំទាន់មានមុខម្ហូប</span>'
    else:
        act_prov = sel_prov if sel_prov != "-- ទាំងអស់ --" else ""
        act_dist = sel_dist if sel_dist != "-- ទាំងអស់ --" else ""
        act_comm = sel_comm if sel_comm != "-- ទាំងអស់ --" else ""
        school_menus = []
        dish_count = 0
        tot_week_cost = 0.0
        tot_month_est = 0.0
        target_st = 100
        display_school_title = "គំរូទទេ (សូមជ្រើសរើស ខេត្ត ស្រុក ឃុំ សាលារៀន ខាងលើ)"
        badge_html = '<span style="background: #f1f5f9; color: #475569; padding: 6px 12px; border-radius: 20px; font-weight: bold; font-size: 0.85rem;">💡 គំរូតារាងទទេ</span>'

    # បង្ហាញកាតសង្ខេបព័ត៌មានសាលា
    render_clean_html(f"""
    <div style="background: white; border: 1px solid #e2e8f0; border-radius: 12px; padding: 16px 20px; margin-bottom: 20px; box-shadow: 0 2px 6px rgba(0,0,0,0.03);">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
            <div>
                <span style="font-size: 1.25rem; font-weight: bold; color: #1e293b;">🏫 សាលាបឋមសិក្សា៖ <span style="color: #0284c7;">{display_school_title}</span></span>
                <span style="font-size: 0.9rem; color: #64748b; margin-left: 12px;">📍 ឃុំ៖ <b>{act_comm or 'មិនទាន់បញ្ជាក់'}</b> | ស្រុក៖ <b>{act_dist or 'មិនទាន់បញ្ជាក់'}</b> | ខេត្ត៖ <b>{act_prov or 'មិនទាន់បញ្ជាក់'}</b></span>
            </div>
            <div>
                {badge_html}
            </div>
        </div>
    </div>
    """)

    # KPI Metrics
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    with kpi1:
        st.metric("🍲 មុខម្ហូបក្នុងសប្ដាហ៍", f"{dish_count} មុខ", f"{dish_count}/៧ ថ្ងៃ" if (is_school_selected and dish_count < 7) else "ពេញលេញ" if is_school_selected else "គំរូទទេ")
    with kpi2:
        st.metric("👥 សិស្សទទួលទានគោលដៅ", f"{target_st:,} នាក់", "គណនាស្វ័យប្រវត្ត")
    with kpi3:
        st.metric("💰 ថវិកាស្បៀង ១សប្ដាហ៍", f"{tot_week_cost:,.0f} ៛", f"{(tot_week_cost/target_st if target_st else 0):,.0f} ៛/សិស្ស/សប្ដាហ៍" if is_school_selected else "គំរូទទេ")
    with kpi4:
        st.metric("📅 ថវិកាស្បៀង ១ខែ (៤សប្ដាហ៍)", f"{tot_month_est:,.0f} ៛", "ប៉ាន់ស្មានតាមមុខម្ហូប" if is_school_selected else "គំរូទទេ")

    # ៧ ផ្ទាំងបញ្ជា (Tabs)
    tab_builder, tab_daily_rec, tab_cards, tab_seed, tab_manual, tab_matrix, tab_export = st.tabs([
        "📝 បង្កើតមុខម្ហូបតាមថ្ងៃ MoEYS SFIS",
        "📋 តារាងតម្រូវការស្បៀងប្រចាំថ្ងៃ (Daily Records)",
        "📅 កាលវិភាគមុខម្ហូបប្រចាំសប្ដាហ៍ (Weekly Schedule)",
        "⚡ អនុវត្តគំរូស្ដង់ដារ MoEYS SFIS (One-Click Seed)",
        "➕ បញ្ចូល / កែសម្រួលមុខម្ហូប (Add & Edit Dishes)",
        "📊 តារាងតម្រូវការស្បៀងប្រចាំខែ (Monthly Matrix)",
        "🖨️ បោះពុម្ព & ទាញយកឯកសារផ្លូវការ (Excel / Print)",
    ])

    # ================= TAB 1: បង្កើតមុខម្ហូបតាមថ្ងៃ MoEYS SFIS =================
    with tab_builder:
        builder_title = f" (សាលា៖ {sel_school})" if is_school_selected else " (គំរូទទេ)"
        st.subheader(f"📝 បង្កើតមុខម្ហូបតាមថ្ងៃ និងគណនាតម្រូវការស្បៀង{builder_title}")
        st.caption("រៀបចំមុខម្ហូបប្រចាំថ្ងៃនៃសប្ដាហ៍ (៧ ថ្ងៃ៖ ចន្ទ-អាទិត្យ) ជ្រើសរើសកាលបរិច្ឆេទក្នុងខែ កែសម្រួលមុខទំនិញ បរិមាណ នាំចូលឯកសារស្បៀង និងកត់ត្រាស្បៀងទុកបានយូរ ស្របតាមប្រព័ន្ធ MoEYS SFIS")

        # ---------------- 📤 នាំចូលឯកសារស្បៀង (Upload File) ----------------
        with st.expander("📤 នាំចូលឯកសារស្បៀងពីខាងក្រៅ (Upload File: PDF, Excel, Word, CSV, រូបភាព)", expanded=False):
            st.markdown("<div style='font-size: 0.92rem; font-weight: bold; color: #1e3a8a; margin-bottom: 6px;'>ប្រព័ន្ធសម្គាល់ឈ្មោះសាលារៀន ស្រង់កាលបរិច្ឆេទ មុខទំនិញ បរិមាណ និងតម្លៃដោយស្វ័យប្រវត្តិ រួចបញ្ចូលទៅក្នុងបញ្ជីតម្រូវការស្បៀងប្រចាំថ្ងៃ</div>", unsafe_allow_html=True)

            all_known_schools = sorted(list(set(
                [r[0] for r in conn.cursor().execute("SELECT DISTINCT school_name FROM suppliers WHERE school_name IS NOT NULL AND school_name != ''").fetchall()] +
                [r[0] for r in conn.cursor().execute("SELECT DISTINCT school_name FROM school_menus WHERE school_name IS NOT NULL AND school_name != ''").fetchall()] +
                [r[0] for r in conn.cursor().execute("SELECT DISTINCT school_name FROM daily_records WHERE school_name IS NOT NULL AND school_name != ''").fetchall()] +
                ([sel_school] if is_school_selected else [])
            )))

            uploaded_food_file = st.file_uploader(
                "📂 ជ្រើសរើស ឬទម្លាក់ឯកសារនៅទីនេះ (គាំទ្រ PDF, Excel .xlsx/.xls, Word .docx, CSV, រូបភាព .png/.jpg/.jpeg)",
                type=["pdf", "xlsx", "xls", "docx", "doc", "png", "jpg", "jpeg", "csv"],
                key="uploader_food_file_input"
            )

            if uploaded_food_file is not None:
                parsed_data = parse_uploaded_food_file(
                    uploaded_food_file,
                    known_schools=all_known_schools,
                    default_school=sel_school if is_school_selected else ""
                )

                det_s = parsed_data["detected_school"] or (sel_school if is_school_selected else (all_known_schools[0] if all_known_schools else ""))
                det_m = parsed_data["detected_month"]
                det_y = parsed_data["detected_year"]

                render_clean_html(f"""
                <div style="background: #f0fdf4; border: 1.5px solid #86efac; border-radius: 10px; padding: 14px 18px; margin-top: 12px; margin-bottom: 15px;">
                    <div style="display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px;">
                        <div>
                            <span style="font-size: 1.15rem; font-weight: bold; color: #166534;">🏫 សាលារៀនដែលបានសម្គាល់៖ <span style="color: #0369a1;">{det_s}</span></span>
                            <span style="background: #dcfce7; color: #15803d; padding: 3px 10px; border-radius: 12px; font-size: 0.82rem; font-weight: bold; margin-left: 10px;">✅ {parsed_data['school_confidence']}</span>
                        </div>
                        <div style="font-size: 0.88rem; color: #475569;">
                            ឯកសារ៖ <b>{parsed_data['filename']}</b> (ប្រភេទ <b>.{parsed_data['file_type'].upper()}</b>)
                        </div>
                    </div>
                </div>
                """)

                col_cf1, col_cf2, col_cf3 = st.columns([1.8, 1, 1])
                with col_cf1:
                    sch_opts = [det_s] + [s for s in all_known_schools if s != det_s]
                    confirmed_school = st.selectbox("🏫 បញ្ជាក់ឈ្មោះសាលាដែលត្រូវទទួលស្បៀង", sch_opts if sch_opts else ["-- ជ្រើសរើសសាលារៀន --"], index=0, key="sel_confirmed_import_school")
                with col_cf2:
                    cf_month_idx = det_m - 1 if 1 <= det_m <= 12 else 10
                    confirmed_month_name = st.selectbox("📅 ខែ", KHMER_MONTHS, index=cf_month_idx, key="sel_confirmed_import_month")
                    confirmed_m_num = KHMER_MONTH_TO_NUM.get(confirmed_month_name, det_m)
                with col_cf3:
                    confirmed_year = st.number_input("ឆ្នាំ", min_value=2024, max_value=2035, value=det_y, step=1, key="sel_confirmed_import_year")

                if parsed_data["file_type"] in ["png", "jpg", "jpeg", "webp"]:
                    st.image(parsed_data["file_bytes"], caption=f"🖼️ រូបភាពឯកសារ៖ {parsed_data['filename']}", use_container_width=True)
                    if not parsed_data["items"]:
                        st.info("💡 រូបភាពត្រូវបានផ្ទុកជោគជ័យ! លោកអ្នកអាចចុចប៊ូតុងខាងក្រោមដើម្បីបង្កើតតារាងស្បៀងស្ដង់ដារ MoEYS SFIS សម្រាប់សាលានេះ រួចកែសម្រួលតាមរូបភាពបានភ្លាមៗ។")
                        if st.button("⚡ បង្កើតតារាងស្បៀងស្ដង់ដារ MoEYS SFIS តាមរូបភាពនេះ", key="btn_seed_from_image", type="primary"):
                            sample_items = []
                            for d_k in ["ចន្ទ", "អង្គារ", "ពុធ", "ព្រហស្បតិ៍", "សុក្រ", "សៅរ៍", "អាទិត្យ"]:
                                p_info = DEFAULT_DAY_PRESETS.get(d_k, {})
                                for ig in p_info.get("ingredients", []):
                                    sample_items.append({
                                        "date": f"{confirmed_year:04d}-{confirmed_m_num:02d}-05",
                                        "item_name": ig["item_name"],
                                        "category": ig.get("category", "បន្លែ"),
                                        "unit": ig.get("unit", "1គីឡូ"),
                                        "quantity": round(float(ig.get("gram_per_student", 20.0)) * 100 / 1000.0, 2) or 2.0,
                                        "unit_price": get_active_item_price(conn, ig["item_name"], confirmed_school),
                                        "total_price": 0.0,
                                        "voucher_no": "001",
                                        "menu_name": p_info.get("dish_name", "")
                                    })
                            parsed_data["items"] = sample_items
                            st.session_state["cached_import_items"] = sample_items
                            st.rerun()

                import_items = st.session_state.get("cached_import_items", parsed_data["items"])
                if import_items:
                    tot_up_cnt = len(import_items)
                    tot_up_dates = len(set(it["date"] for it in import_items))
                    tot_up_cost = sum(it.get("total_price", 0.0) or (it.get("quantity", 0.0) * it.get("unit_price", 0.0)) for it in import_items)

                    um1, um2, um3 = st.columns(3)
                    with um1:
                        st.metric("📦 ចំនួនមុខទំនិញដែលបានស្រង់", f"{tot_up_cnt} ជួរ")
                    with um2:
                        st.metric("📅 ចំនួនកាលបរិច្ឆេទ", f"{tot_up_dates} ថ្ងៃ")
                    with um3:
                        st.metric("💰 ថវិកាសរុបប៉ាន់ស្មាន", f"{tot_up_cost:,.0f} ៛")

                    st.markdown("<div style='font-size: 0.9rem; font-weight: bold; color: #0f172a; margin-top: 10px; margin-bottom: 6px;'>📋 ទិន្នន័យស្បៀងដែលបានស្រង់ (អាចកែប្រែបានក្នុងតារាងផ្ទាល់)៖</div>", unsafe_allow_html=True)
                    df_editor_data = pd.DataFrame([
                        {
                            "កាលបរិច្ឆេទ": it["date"],
                            "មុខទំនិញ/ស្បៀង": it["item_name"],
                            "ប្រភេទ": it.get("category") or auto_classify_category(it["item_name"]),
                            "ឯកតា": it.get("unit", "1គីឡូ"),
                            "បរិមាណ": float(it.get("quantity") or 0.0),
                            "តម្លៃរាយ (៛)": float(it.get("unit_price") or get_active_item_price(conn, it["item_name"], confirmed_school)),
                            "សរុប (៛)": float(it.get("total_price") or (float(it.get("quantity") or 0.0) * float(it.get("unit_price") or 0.0))),
                            "មុខម្ហូប": it.get("menu_name", ""),
                            "លេខសក្ខីប័ត្រ": it.get("voucher_no", "")
                        }
                        for it in import_items
                    ])
                    edited_df = st.data_editor(
                        df_editor_data,
                        use_container_width=True,
                        num_rows="dynamic",
                        key="editor_upload_food_table"
                    )

                    st.markdown("---")
                    col_imp_exec, col_imp_sp = st.columns([2.5, 2])
                    with col_imp_exec:
                        if st.button(
                            f"📥 បញ្ចូលទៅក្នុងបញ្ជីតម្រូវការស្បៀងប្រចាំថ្ងៃ (សាលា៖ {confirmed_school})",
                            type="primary",
                            use_container_width=True,
                            key="btn_exec_import_to_daily_records"
                        ):
                            if not confirmed_school or confirmed_school in ["-- ជ្រើសរើសសាលារៀន --", "គ្មានសាលា"]:
                                st.warning("⚠️ សូមបញ្ជាក់ឈ្មោះសាលារៀនជាមុនសិន!")
                            else:
                                c_imp = conn.cursor()
                                inserted_count = 0
                                total_imported_cost = 0.0

                                for _, row_val in edited_df.iterrows():
                                    r_date = str(row_val["កាលបរិច្ឆេទ"]).strip()[:10]
                                    r_item = str(row_val["មុខទំនិញ/ស្បៀង"]).strip()
                                    if not r_item or not r_date:
                                        continue
                                    r_cat = str(row_val.get("ប្រភេទ", "")).strip() or auto_classify_category(r_item)
                                    r_qty = float(row_val.get("បរិមាណ") or 0.0)
                                    r_price = float(row_val.get("តម្លៃរាយ (៛)") or 0.0)
                                    r_tot = float(row_val.get("សរុប (៛)") or (r_qty * r_price))
                                    r_menu = str(row_val.get("មុខម្ហូប", "")).strip()
                                    r_vno = str(row_val.get("លេខសក្ខីប័ត្រ", "")).strip() or get_or_create_school_voucher(conn, confirmed_school, r_date)
                                    r_phase = determine_phase_for_date(conn, confirmed_school, r_date)

                                    total_imported_cost += r_tot

                                    ex_rec = c_imp.execute("""
                                        SELECT id FROM daily_records
                                        WHERE school_name=? AND date=? AND item_name=?
                                    """, (confirmed_school, r_date, r_item)).fetchone()

                                    if ex_rec:
                                        c_imp.execute("""
                                            UPDATE daily_records
                                            SET quantity=?, unit_price=?, total_price=?, phase=?, voucher_no=?, consumption_date=?, category=?, menu_name=?
                                            WHERE id=?
                                        """, (r_qty, r_price, r_tot, r_phase, r_vno, r_date, r_cat, r_menu, ex_rec[0]))
                                    else:
                                        c_imp.execute("""
                                            INSERT INTO daily_records (
                                                date, school_name, item_name, phase, quantity, unit_price, total_price, voucher_no, consumption_date, category, menu_name
                                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                                        """, (r_date, confirmed_school, r_item, r_phase, r_qty, r_price, r_tot, r_vno, r_date, r_cat, r_menu))
                                    inserted_count += 1

                                conn.commit()
                                st.success(f"""
                                🎉 **នាំចូលទិន្នន័យស្បៀងជោគជ័យ!**
                                - 🏫 សាលាបឋមសិក្សា៖ **{confirmed_school}**
                                - 📋 ចំនួនទិន្នន័យស្បៀងបានបញ្ចូល៖ **{inserted_count} ជួរ** ចូលក្នុងតារាងតម្រូវការស្បៀងប្រចាំថ្ងៃ (Daily Records)
                                - 💰 ថវិកាសរុប៖ **{total_imported_cost:,.0f} ៛**
                                """)
                                st.balloons()
                                st.rerun()
                else:
                    st.warning("⚠️ មិនទាន់មានទិន្នន័យស្បៀងដែលបានស្រង់នៅឡើយទេ។ សូមពិនិត្យមើលឯកសារដែលបានបញ្ចូល ឬជ្រើសរើសឯកសារផ្សេង។")

        # ---------------- កាលបរិច្ឆេទ និងព័ត៌មានទូទៅនៃខែ ----------------
        col_b_m, col_b_y, col_b_st, col_b_act = st.columns([1.2, 1, 1.2, 1.6])
        with col_b_m:
            sel_b_month_name = st.selectbox("📅 ជ្រើសរើសខែ", KHMER_MONTHS, index=10, key="bldr_month_sel")
            b_month_num = KHMER_MONTH_TO_NUM.get(sel_b_month_name, 11)
        with col_b_y:
            b_year_num = st.number_input("ឆ្នាំ", min_value=2024, max_value=2035, value=2025, step=1, key="bldr_year_num")
        with col_b_st:
            b_target_st = st.number_input("👥 ចំនួនសិស្ស (នាក់)", min_value=1, max_value=5000, value=target_st or 100, step=10, key="bldr_target_st")
        with col_b_act:
            st.write("")
            if st.button("⚡ ផ្ទុកគំរូស្ដង់ដារ MoEYS SFIS (៧ ថ្ងៃ)", use_container_width=True, key="btn_bldr_seed_defaults"):
                if not is_school_selected:
                    st.warning("⚠️ សូមជ្រើសរើសសាលារៀនជាមុនសិនដើម្បីផ្ទុកគំរូ!")
                else:
                    for d_k, d_v in DEFAULT_DAY_PRESETS.items():
                        st.session_state[f"bldr_dish_name_{sel_school}_{d_k}"] = d_v["dish_name"]
                        st.session_state[f"bldr_ings_{sel_school}_{d_k}"] = [
                            {
                                "category": ig.get("category", "បន្លែ"),
                                "item_name": ig["item_name"],
                                "unit": ig.get("unit", "1គីឡូ"),
                                "gram_per_student": float(ig.get("gram_per_student", 20.0)),
                                "qty_per_100": float(ig.get("qty_per_100", 2.0)),
                                "unit_price": get_active_item_price(conn, ig["item_name"], sel_school, act_comm)
                            }
                            for ig in d_v["ingredients"]
                        ]
                    st.success("✅ បានផ្ទុកគំរូស្ដង់ដារ MoEYS SFIS ទាំង ៧ ថ្ងៃជោគជ័យ!")
                    st.rerun()

        # ប្រសិនបើមិនទាន់ជ្រើសរើសសាលា បង្ហាញត្រឹមគំរូតារាងទទេ
        if not is_school_selected:
            st.info("💡 **គំរូតារាងទទេ៖** សូមជ្រើសរើស **ខេត្ត ស្រុក ឃុំ និង សាលារៀន** ខាងលើ ដើម្បីបង្កើត និងកែសម្រួលមុខម្ហូបប្រចាំថ្ងៃនៃសប្ដាហ៍។")
            empty_bldr_df = pd.DataFrame(columns=[
                "កាលបរិច្ឆេទ", "ថ្ងៃនៃសប្ដាហ៍", "មុខម្ហូប", "ប្រភេទស្បៀង", "មុខទំនិញ", "ឯកតា", "បរិមាណ", "តម្លៃរាយ (៛)", "សរុប (៛)"
            ])
            st.dataframe(empty_bldr_df, use_container_width=True, hide_index=True)
            builder_days_payload = {}
        else:
            month_weekday_dates = get_month_weekday_dates(b_year_num, b_month_num)
            builder_days_payload = {}
            catalog_item_names = [p["name"] for p in SUPPLIER_PRODUCT_CATALOG]

            for d_name in KHMER_DAYS_OF_WEEK:
                def_preset = DEFAULT_DAY_PRESETS.get(d_name, {"dish_name": "សម្លកកូរសាច់ជ្រូក", "ingredients": []})
                existing_d_menu = next((m for m in school_menus if m["day_of_week"] == d_name), None)

                day_ing_key = f"bldr_ings_{sel_school}_{d_name}"
                day_dish_key = f"bldr_dish_name_{sel_school}_{d_name}"
                day_custom_dates_key = f"custom_dates_{sel_school}_{d_name}_{b_year_num}_{b_month_num}"

                if day_custom_dates_key not in st.session_state:
                    st.session_state[day_custom_dates_key] = []

                if day_ing_key not in st.session_state:
                    if existing_d_menu and existing_d_menu.get("ingredients"):
                        st.session_state[day_ing_key] = [
                            {
                                "category": ig.get("category") or auto_classify_category(ig.get("item_name", "")),
                                "item_name": ig.get("item_name"),
                                "unit": ig.get("unit") or "1គីឡូ",
                                "gram_per_student": float(ig.get("gram_per_student") or 20.0),
                                "qty_per_100": float(ig.get("total_qty") or 2.0) / max(1, (b_target_st / 100.0)),
                                "unit_price": float(ig.get("unit_price") or 0.0)
                            }
                            for ig in existing_d_menu["ingredients"]
                        ]
                    else:
                        st.session_state[day_ing_key] = [
                            {
                                "category": ig.get("category", "បន្លែ"),
                                "item_name": ig["item_name"],
                                "unit": ig.get("unit", "1គីឡូ"),
                                "gram_per_student": float(ig.get("gram_per_student", 20.0)),
                                "qty_per_100": float(ig.get("qty_per_100", 2.0)),
                                "unit_price": get_active_item_price(conn, ig["item_name"], sel_school, act_comm)
                            }
                            for ig in def_preset.get("ingredients", [])
                        ]

                if day_dish_key not in st.session_state:
                    st.session_state[day_dish_key] = existing_d_menu["menu_name"] if existing_d_menu else def_preset["dish_name"]

                render_clean_html(f"""
                <div style="background-color: #cbe3e7; border: 1.5px solid #8cb9c5; border-radius: 8px; padding: 10px 16px; margin-top: 14px; margin-bottom: 6px;">
                    <div style="font-weight: bold; font-size: 1.1rem; color: #0f172a;">
                        📅 ថ្ងៃ{d_name}
                    </div>
                </div>
                """)

                col_d_left, col_d_mid, col_d_right = st.columns([1.3, 1.4, 2.5])

                with col_d_left:
                    use_preset = st.checkbox("ជ្រើសមុខម្ហូបមានស្រាប់", value=True, key=f"chk_preset_{d_name}")
                    if use_preset:
                        preset_keys = list(SFIS_PRESET_DISHES.keys())
                        cur_val = st.session_state[day_dish_key]
                        p_idx = preset_keys.index(cur_val) if cur_val in preset_keys else 0
                        chosen_p = st.selectbox(
                            "មុខម្ហូប MoEYS SFIS",
                            preset_keys,
                            index=p_idx,
                            key=f"sel_preset_{d_name}",
                            label_visibility="collapsed"
                        )
                        if chosen_p != st.session_state[day_dish_key]:
                            st.session_state[day_dish_key] = chosen_p
                            st.session_state[day_ing_key] = [
                                {
                                    "category": ig.get("category", "បន្លែ"),
                                    "item_name": ig["item_name"],
                                    "unit": ig.get("unit", "1គីឡូ"),
                                    "gram_per_student": float(ig.get("gram_per_student", 20.0)),
                                    "qty_per_100": float(ig.get("qty_per_100", 2.0)),
                                    "unit_price": get_active_item_price(conn, ig["item_name"], sel_school, act_comm)
                                }
                                for ig in SFIS_PRESET_DISHES[chosen_p]["ingredients"]
                            ]
                            st.rerun()

                    cur_dish_name = st.text_input(
                        "ឈ្មោះមុខម្ហូប",
                        value=st.session_state[day_dish_key],
                        key=f"inp_dish_{d_name}",
                        label_visibility="collapsed"
                    )
                    st.session_state[day_dish_key] = cur_dish_name

                with col_d_mid:
                    st.markdown("<div style='font-size: 0.85rem; font-weight: bold; color: #1e3a8a; margin-bottom: 4px;'>កាលបរិច្ឆេទ និងចំនួនសិស្ស៖</div>", unsafe_allow_html=True)
                    dates_for_day = month_weekday_dates.get(d_name, []) + st.session_state[day_custom_dates_key]
                    day_active_dates = []

                    for dt in dates_for_day:
                        c_dt1, c_dt2 = st.columns([2.3, 1.2])
                        with c_dt1:
                            chk_dt = st.checkbox(dt["label"], value=True, key=f"dt_chk_{d_name}_{dt['date_str']}")
                        with c_dt2:
                            st_cnt = st.number_input(
                                "",
                                min_value=1,
                                max_value=5000,
                                value=b_target_st,
                                step=1,
                                key=f"st_cnt_{d_name}_{dt['date_str']}",
                                label_visibility="collapsed"
                            )
                        if chk_dt:
                            day_active_dates.append({
                                "date": dt["date_str"],
                                "label": dt["label"],
                                "students": st_cnt
                            })

                    with st.popover(f"➕ បន្ថែមកាលបរិច្ឆេទថ្ងៃ{d_name}"):
                        st.caption(f"ជ្រើសរើសកាលបរិច្ឆេទជាក់លាក់បន្ថែមសម្រាប់ថ្ងៃ{d_name}")
                        new_custom_d = st.date_input("កាលបរិច្ឆេទ", value=date(b_year_num, b_month_num, 1), key=f"inp_dt_pick_{d_name}")
                        if st.button(f"បញ្ចូលកាលបរិច្ឆេទនេះចូលថ្ងៃ{d_name}", key=f"btn_confirm_add_dt_{d_name}"):
                            c_d_str = new_custom_d.strftime("%Y-%m-%d")
                            c_lbl = f"{to_khmer_digits(new_custom_d.day)} {KHMER_MONTHS[new_custom_d.month - 1]} {to_khmer_digits(new_custom_d.year)}"
                            if not any(x["date_str"] == c_d_str for x in st.session_state[day_custom_dates_key]):
                                st.session_state[day_custom_dates_key].append({
                                    "date_str": c_d_str,
                                    "label": c_lbl,
                                    "day_num": new_custom_d.day
                                })
                                st.rerun()

                with col_d_right:
                    st.markdown("<div style='font-size: 0.85rem; font-weight: bold; color: #1e3a8a; margin-bottom: 4px;'>គ្រឿងផ្សំ និងប្រភេទស្បៀង (បញ្ចូល/កែប្រែបាន)៖</div>", unsafe_allow_html=True)
                    cur_ings = st.session_state[day_ing_key]
                    processed_ings = []
                    del_idx = None

                    for idx, ig in enumerate(cur_ings):
                        c_c1, c_c2, c_c3, c_c4 = st.columns([1.2, 1.8, 1.0, 0.4])
                        with c_c1:
                            cat_list = SFIS_CATEGORIES
                            c_idx = cat_list.index(ig["category"]) if ig["category"] in cat_list else 0
                            sel_cat = st.selectbox(
                                "",
                                cat_list,
                                index=c_idx,
                                key=f"cat_{d_name}_{idx}",
                                label_visibility="collapsed"
                            )
                            ig["category"] = sel_cat

                        with c_c2:
                            cur_i_name = ig["item_name"]
                            item_opts = [cur_i_name] + [n for n in catalog_item_names if n != cur_i_name]
                            sel_item = st.selectbox(
                                "",
                                item_opts,
                                index=0,
                                key=f"item_{d_name}_{idx}",
                                label_visibility="collapsed"
                            )
                            ig["item_name"] = sel_item

                        with c_c3:
                            u_match = next((p["unit"] for p in SUPPLIER_PRODUCT_CATALOG if p["name"] == sel_item), ig.get("unit", "1គីឡូ"))
                            g_std = ig.get("gram_per_student", 20.0)
                            if "គ្រាប់" in u_match:
                                calc_qty = round(b_target_st * g_std, 1)
                            else:
                                calc_qty = round((b_target_st * g_std) / 1000.0, 2)
                                if calc_qty <= 0:
                                    calc_qty = round(ig.get("qty_per_100", 1.0) * (b_target_st / 100.0), 2)
                            
                            qty_val = st.number_input(
                                "",
                                min_value=0.01,
                                value=max(0.01, float(ig.get("qty_per_day", calc_qty))),
                                step=0.5,
                                key=f"qty_{d_name}_{idx}",
                                label_visibility="collapsed"
                            )
                            ig["qty_per_day"] = qty_val

                        with c_c4:
                            if st.button("🗑️", key=f"btn_del_ig_{d_name}_{idx}", help=f"លុបមុខទំនិញ «{cur_i_name}»"):
                                del_idx = idx

                        u_price = get_active_item_price(conn, sel_item, sel_school, act_comm)
                        processed_ings.append({
                            "category": sel_cat,
                            "item_name": sel_item,
                            "unit": u_match,
                            "gram_per_student": g_std,
                            "qty_per_day": qty_val,
                            "unit_price": u_price,
                            "total_cost": qty_val * u_price
                        })

                    if del_idx is not None:
                        st.session_state[day_ing_key].pop(del_idx)
                        st.rerun()

                    c_act_l, c_act_r = st.columns([1, 1.4])
                    with c_act_l:
                        with st.popover(f"➕ ជ្រើសមុខទំនិញថ្មី"):
                            pop_cat = st.selectbox("ប្រភេទស្បៀង", SFIS_CATEGORIES, key=f"pop_cat_{d_name}")
                            pop_item = st.selectbox("ឈ្មោះទំនិញ (៦១ មុខ)", catalog_item_names, key=f"pop_item_{d_name}")
                            pop_unit = next((p["unit"] for p in SUPPLIER_PRODUCT_CATALOG if p["name"] == pop_item), "1គីឡូ")
                            pop_qty = st.number_input("បរិមាណ", min_value=0.1, value=2.0 if "គីឡូ" in pop_unit else 35.0, key=f"pop_qty_{d_name}")
                            if st.button(f"បញ្ចូលមុខទំនិញនេះ", key=f"btn_pop_add_{d_name}"):
                                st.session_state[day_ing_key].append({
                                    "category": pop_cat,
                                    "item_name": pop_item,
                                    "unit": pop_unit,
                                    "gram_per_student": (pop_qty * 1000.0 / b_target_st) if "គីឡូ" in pop_unit else (pop_qty / b_target_st),
                                    "qty_per_100": (pop_qty / (b_target_st / 100.0)),
                                    "qty_per_day": pop_qty,
                                    "unit_price": get_active_item_price(conn, pop_item, sel_school, act_comm)
                                })
                                st.rerun()

                    with c_act_r:
                        if st.button(f"បន្ថែមមុខម្ហូបសម្រាប់ថ្ងៃ{d_name} ＋", key=f"btn_add_item_fast_{d_name}", use_container_width=True):
                            st.session_state[day_ing_key].append({
                                "category": "បន្លែ",
                                "item_name": "ស្ពៃក្រញាញ់",
                                "unit": "1គីឡូ",
                                "gram_per_student": 20.0,
                                "qty_per_100": 2.0,
                                "qty_per_day": 2.0,
                                "unit_price": get_active_item_price(conn, "ស្ពៃក្រញាញ់", sel_school, act_comm)
                            })
                            st.rerun()

                builder_days_payload[d_name] = {
                    "day_name": d_name,
                    "dish_name": cur_dish_name,
                    "meal_type": "អាហារពេលព្រឹក",
                    "target_students": b_target_st,
                    "ingredients": processed_ings,
                    "active_dates": day_active_dates
                }

        # ================= ស្បៀងទុកបានយូរ (ផ្គត់ផ្គង់នៅថ្ងៃដើមខែ) =================
        st.markdown("---")
        render_clean_html("""
        <div style="background: linear-gradient(135deg, #f0fdf4 0%, #e0f2fe 100%); border: 1.5px solid #7dd3fc; border-radius: 12px; padding: 18px 22px; margin-top: 15px; margin-bottom: 20px;">
            <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 8px;">
                <span style="font-size: 1.5rem;">🌾</span>
                <span style="font-size: 1.15rem; font-weight: bold; color: #0369a1;">ស្បៀងទុកបានយូរ (ផ្គត់ផ្គង់នៅថ្ងៃដើមខែ)</span>
            </div>
            <div style="font-size: 0.88rem; color: #334155;">
                យោងតាមគោលការណ៍ណែនាំ <b>MoEYS SFIS</b> និងកម្មវិធីស្បៀងអាហារពិភពលោក (WFP)៖ មុខទំនិញស្បៀងគោលដែលទុកបានយូរ (<b>អង្ករ, ប្រេងឆា, អំបិលអ៊ីយ៉ូត, ទឹកត្រី/ទឹកស៊ីអ៊ីវ</b>) ត្រូវបានផ្គត់ផ្គង់ និងបញ្ជាទិញជាដុំ <b>នៅថ្ងៃដើមខែ</b> សម្រាប់ទុកប្រើប្រាស់ពេញមួយខែ។ នៅពេលចុចរក្សាទុក ប្រព័ន្ធនឹងបង្កើតកំណត់ត្រាតម្រូវការស្បៀងប្រចាំថ្ងៃ (Daily Records) នៅថ្ងៃដើមខែនេះដោយស្វ័យប្រវត្តិ។
            </div>
        </div>
        """)

        col_st1, col_st2, col_st3 = st.columns([1.2, 1.2, 2.5])
        with col_st1:
            include_staples = st.checkbox("✅ កត់ត្រាស្បៀងទុកបានយូរ", value=True, key="chk_include_staples")
        with col_st2:
            def_start_date = f"{b_year_num:04d}-{b_month_num:02d}-01"
            staple_deliv_date = st.text_input("📅 កាលបរិច្ឆេទចែកផ្ដល់ (ដើមខែ)", value=def_start_date, key="inp_staple_deliv_date")
        with col_st3:
            staple_feeding_days = st.number_input("ចំនួនថ្ងៃហូបអាហារក្នុងខែ (សម្រាប់គណនាស្បៀងទុកបានយូរ)", min_value=1, max_value=31, value=24, step=1, key="inp_staple_feeding_days")

        rice_kg = round((b_target_st * 100.0 * staple_feeding_days) / 1000.0, 1)
        oil_lit = round((b_target_st * 10.0 * staple_feeding_days) / 1000.0, 1)
        salt_kg = round((b_target_st * 3.0 * staple_feeding_days) / 1000.0, 1)
        sauce_lit = round((b_target_st * 10.0 * staple_feeding_days) / 1000.0, 1)

        p_rice = get_active_item_price(conn, "អង្ករ", sel_school if is_school_selected else "", act_comm) or 3000.0
        p_oil = get_active_item_price(conn, "ប្រេងឆា", sel_school if is_school_selected else "", act_comm) or 8500.0
        p_salt = get_active_item_price(conn, "អំបិល", sel_school if is_school_selected else "", act_comm) or 1500.0
        p_sauce = get_active_item_price(conn, "ទឹកត្រី", sel_school if is_school_selected else "", act_comm) or 3500.0

        st.markdown("<div style='font-size: 0.9rem; font-weight: bold; color: #0f172a; margin-bottom: 8px;'>បញ្ជីបរិមាណ និងតម្លៃស្បៀងទុកបានយូរប្រចាំខែ៖</div>", unsafe_allow_html=True)
        col_sp1, col_sp2, col_sp3, col_sp4 = st.columns(4)
        with col_sp1:
            fin_rice_kg = st.number_input("🌾 អង្ករ (គីឡូក្រាម)", min_value=0.0, value=float(rice_kg), step=5.0, key="inp_rice_kg")
            fin_rice_pr = st.number_input("តម្លៃអង្ករ (៛/គ.ក)", min_value=0.0, value=float(p_rice), step=100.0, key="inp_rice_pr")
        with col_sp2:
            fin_oil_lit = st.number_input("🍳 ប្រេងឆា (លីត្រ)", min_value=0.0, value=float(oil_lit), step=1.0, key="inp_oil_lit")
            fin_oil_pr = st.number_input("តម្លៃប្រេងឆា (៛/លីត្រ)", min_value=0.0, value=float(p_oil), step=100.0, key="inp_oil_pr")
        with col_sp3:
            fin_salt_kg = st.number_input("🧂 អំបិលអ៊ីយ៉ូត (គីឡូក្រាម)", min_value=0.0, value=float(salt_kg), step=0.5, key="inp_salt_kg")
            fin_salt_pr = st.number_input("តម្លៃអំបិល (៛/គ.ក)", min_value=0.0, value=float(p_salt), step=100.0, key="inp_salt_pr")
        with col_sp4:
            fin_sauce_lit = st.number_input("🫙 ទឹកត្រី/ទឹកស៊ីអ៊ីវ (លីត្រ)", min_value=0.0, value=float(sauce_lit), step=1.0, key="inp_sauce_lit")
            fin_sauce_pr = st.number_input("តម្លៃទឹកត្រី (៛/លីត្រ)", min_value=0.0, value=float(p_sauce), step=100.0, key="inp_sauce_pr")

        staples_payload = {
            "include": include_staples,
            "delivery_date": staple_deliv_date,
            "items": [
                {"item_name": "អង្ករ", "category": "អង្ករ", "unit": "1គីឡូ", "quantity": fin_rice_kg, "unit_price": fin_rice_pr},
                {"item_name": "ប្រេងឆា", "category": "ប្រេងឆា", "unit": "លីត្រ", "quantity": fin_oil_lit, "unit_price": fin_oil_pr},
                {"item_name": "អំបិល", "category": "អំបិល", "unit": "1គីឡូ", "quantity": fin_salt_kg, "unit_price": fin_salt_pr},
                {"item_name": "ទឹកត្រី", "category": "គ្រឿងទេស", "unit": "លីត្រ", "quantity": fin_sauce_lit, "unit_price": fin_sauce_pr},
            ]
        }

        tot_staple_cost = sum(it["quantity"] * it["unit_price"] for it in staples_payload["items"]) if include_staples else 0
        render_clean_html(f"""
        <div style="background: #f1f5f9; padding: 8px 16px; border-radius: 6px; font-weight: bold; color: #1e3a8a; text-align: right; margin-top: 8px; margin-bottom: 20px;">
            សរុបថវិកាស្បៀងទុកបានយូរ (១ខែ) ៖ <span style="color: #0369a1; font-size: 1.05rem;">{tot_staple_cost:,.0f} ៛</span>
        </div>
        """)

        # ប៊ូតុងរក្សាទុកស្បៀងទុកបានយូរ និងប៊ូតុងរក្សាទុកកាលវិភាគមុខម្ហូប
        col_btn_st, col_btn_wk = st.columns([1.5, 2.2])
        with col_btn_st:
            if st.button("💾 រក្សាទុកស្បៀងទុកបានយូរ", type="primary", use_container_width=True, key="btn_save_staples_only"):
                if not is_school_selected:
                    st.warning("⚠️ សូមជ្រើសរើសសាលារៀនជាមុនសិនដើម្បីរក្សាទុកស្បៀងទុកបានយូរ!")
                else:
                    st_res = save_school_staples_only(
                        conn=conn,
                        school_name=sel_school,
                        delivery_date=staple_deliv_date,
                        staple_items=staples_payload["items"]
                    )
                    if st_res["success"]:
                        st.success(f"""
                        🎉 **រក្សាទុកស្បៀងទុកបានយូរជោគជ័យ!**
                        - 🏫 សាលាបឋមសិក្សា៖ **{sel_school}**
                        - 📅 កាលបរិច្ឆេទចែកផ្ដល់៖ **{st_res['delivery_date']}**
                        - 📋 ចំនួនមុខទំនិញ៖ **{st_res['inserted_count']} មុខ** (អង្ករ, ប្រេងឆា, អំបិល, ទឹកត្រី)
                        - 💰 ថវិកាសរុប៖ **{st_res['total_cost']:,.0f} ៛**
                        - 📑 លេខសក្ខីប័ត្រ៖ **{st_res['voucher_no']}** | វគ្គ៖ **{st_res['phase']}**
                        - ទិន្នន័យត្រូវបានបញ្ចូលទៅក្នុង **តារាងតម្រូវការស្បៀងប្រចាំថ្ងៃ (Daily Records)** រួចរាល់!
                        """)
                        st.balloons()
                        st.rerun()

        with col_btn_wk:
            if is_school_selected and builder_days_payload:
                if st.button(
                    f"💾 រក្សាទុកកាលវិភាគមុខម្ហូប និងបញ្ចូលទៅក្នុង Daily Records (សាលា៖ {sel_school})",
                    use_container_width=True,
                    key="btn_save_builder_and_daily_records"
                ):
                    save_res = save_school_daily_requirements(
                        conn=conn,
                        school_name=sel_school,
                        commune=act_comm,
                        district=act_dist,
                        province=act_prov,
                        days_menu_data=builder_days_payload,
                        staple_data=staples_payload
                    )
                    if save_res["success"]:
                        st.success(f"""
                        🎉 **រក្សាទុក និងបញ្ចូលទិន្នន័យស្បៀងជោគជ័យ!**
                        - 🏫 សាលាបឋមសិក្សា៖ **{sel_school}**
                        - 📅 ចំនួនថ្ងៃផ្គត់ផ្គង់ស្បៀង៖ **{save_res['total_dates']} ថ្ងៃ**
                        - 📋 ចំនួនទិន្នន័យស្បៀងបានបញ្ចូល៖ **{save_res['total_records']} ជួរ** ចូលក្នុងតារាងតម្រូវការស្បៀងប្រចាំថ្ងៃ (Daily Records)
                        - 💰 ថវិកាស្បៀងសរុបប្រចាំខែ៖ **{save_res['total_cost']:,.0f} ៛**
                        """)
                        st.balloons()
                        st.rerun()

    # ================= TAB 2: តារាងតម្រូវការស្បៀងប្រចាំថ្ងៃ =================
    with tab_daily_rec:
        dr_title = f" (សាលា៖ {sel_school})" if is_school_selected else " (គំរូទទេ)"
        st.subheader(f"📋 តារាងតម្រូវការស្បៀងប្រចាំថ្ងៃ{dr_title}")
        st.caption("ទិន្នន័យតម្រូវការស្បៀងដែលបានបញ្ចូល និងគណនាដោយស្វ័យប្រវត្តិតាមថ្ងៃនីមួយៗ ស្របតាមប្រព័ន្ធ MoEYS SFIS")

        col_dr1, col_dr2, col_dr3 = st.columns([1.2, 1, 2])
        with col_dr1:
            sel_dr_month = st.selectbox("📅 ជ្រើសរើសខែ", KHMER_MONTHS, index=10, key="sel_dr_month_view")
            dr_m_num = KHMER_MONTH_TO_NUM.get(sel_dr_month, 11)
        with col_dr2:
            sel_dr_year = st.number_input("ឆ្នាំ", min_value=2024, max_value=2035, value=2025, step=1, key="sel_dr_year_view")
        with col_dr3:
            st.write("")
            records = get_school_daily_records(conn, sel_school, sel_dr_year, dr_m_num) if is_school_selected else []
            if records:
                excel_daily_bytes = generate_daily_requirements_excel(conn, sel_school, sel_dr_year, dr_m_num, records)
                st.download_button(
                    label=f"📥 ទាញយកតារាងតម្រូវការស្បៀងជា Excel (.xlsx)",
                    data=excel_daily_bytes,
                    file_name=f"តម្រូវការស្បៀងប្រចាំថ្ងៃ_{sel_school}_{sel_dr_month}_{sel_dr_year}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                    key="btn_dl_daily_excel"
                )

        if not is_school_selected:
            st.info("💡 **គំរូតារាងទទេ៖** សូមជ្រើសរើស **ខេត្ត ស្រុក ឃុំ និង សាលារៀន** ខាងលើ ដើម្បីពិនិត្យមើលតារាងតម្រូវការស្បៀងប្រចាំថ្ងៃ។")
            empty_dr_df = pd.DataFrame(columns=[
                "កាលបរិច្ឆេទ", "ថ្ងៃនៃសប្ដាហ៍", "មុខម្ហូប", "ប្រភេទ", "មុខទំនិញ/ស្បៀង", "បរិមាណ", "តម្លៃរាយ (៛)", "សរុបទឹកប្រាក់ (៛)", "វគ្គ", "លេខសក្ខីប័ត្រ"
            ])
            st.dataframe(empty_dr_df, use_container_width=True, hide_index=True)
        elif not records:
            st.info(f"💡 **គំរូតារាងទទេ៖** ពុំទាន់មានទិន្នន័យតម្រូវការស្បៀងប្រចាំថ្ងៃសម្រាប់សាលា «**{sel_school}**» ក្នុងខែ {sel_dr_month} ឆ្នាំ {sel_dr_year} នៅឡើយទេ។")
            empty_dr_df = pd.DataFrame(columns=[
                "កាលបរិច្ឆេទ", "ថ្ងៃនៃសប្ដាហ៍", "មុខម្ហូប", "ប្រភេទ", "មុខទំនិញ/ស្បៀង", "បរិមាណ", "តម្លៃរាយ (៛)", "សរុបទឹកប្រាក់ (៛)", "វគ្គ", "លេខសក្ខីប័ត្រ"
            ])
            st.dataframe(empty_dr_df, use_container_width=True, hide_index=True)
            st.markdown("""
            លោកអ្នកអាចបញ្ចូលទិន្នន័យបានតាមពីរវិធី៖
            1. ចុចផ្ទាំង **📝 បង្កើតមុខម្ហូបតាមថ្ងៃ MoEYS SFIS** ដើម្បីបង្កើតមុខម្ហូប និងកាលបរិច្ឆេទក្នុងខែ
            2. ប្រើប្រាស់មុខងារ **📤 នាំចូលឯកសារស្បៀង (Upload File)** ក្នុងផ្ទាំងបង្កើតមុខម្ហូប ដើម្បីបញ្ចូលឯកសារ Excel, PDF, Word ឬ រូបភាព
            """)
        else:
            tot_days_cnt = len(set(r["date"] for r in records))
            tot_items_cnt = len(records)
            tot_weight_kg = sum(r["quantity"] for r in records if "គីឡូ" in r.get("item_name", "") or r.get("category") in ["បន្លែ", "ត្រី/សាច់/ស៊ុត", "អង្ករ"])
            tot_req_cost = sum(r["total_price"] for r in records)

            m1, m2, m3, m4 = st.columns(4)
            with m1:
                st.metric("📅 ចំនួនថ្ងៃផ្គត់ផ្គង់", f"{tot_days_cnt} ថ្ងៃ")
            with m2:
                st.metric("📦 ចំនួនមុខទំនិញសរុប", f"{tot_items_cnt} ជួរ")
            with m3:
                st.metric("⚖️ ទម្ងន់ស្បៀង (ប្រហាក់ប្រហែល)", f"{tot_weight_kg:,.1f} គ.ក")
            with m4:
                st.metric("💰 ថវិកាស្បៀងសរុប", f"{tot_req_cost:,.0f} ៛")

            col_f1, col_f2 = st.columns([1.5, 1.5])
            with col_f1:
                avail_dates = ["-- ទាំងអស់ --"] + sorted(list(set(r["date"] for r in records)))
                filter_date = st.selectbox("🔍 ចម្រោះតាមកាលបរិច្ឆេទ", avail_dates, key="flt_dr_date")
            with col_f2:
                avail_cats = ["-- ទាំងអស់ --"] + sorted(list(set(r["category"] for r in records if r["category"])))
                filter_cat = st.selectbox("🔍 ចម្រោះតាមប្រភេទស្បៀង", avail_cats, key="flt_dr_cat")

            filtered_records = records
            if filter_date != "-- ទាំងអស់ --":
                filtered_records = [r for r in filtered_records if r["date"] == filter_date]
            if filter_cat != "-- ទាំងអស់ --":
                filtered_records = [r for r in filtered_records if r["category"] == filter_cat]

            df_display = pd.DataFrame([
                {
                    "កាលបរិច្ឆេទ": r["date"],
                    "ថ្ងៃនៃសប្ដាហ៍": r["day_name"],
                    "មុខម្ហូប": r["menu_name"],
                    "ប្រភេទ": r["category"],
                    "មុខទំនិញ/ស្បៀង": r["item_name"],
                    "បរិមាណ": r["quantity"],
                    "តម្លៃរាយ (៛)": f"{r['unit_price']:,.0f}",
                    "សរុបទឹកប្រាក់ (៛)": f"{r['total_price']:,.0f}",
                    "វគ្គ": r["phase"],
                    "លេខសក្ខីប័ត្រ": r["voucher_no"]
                }
                for r in filtered_records
            ])
            st.dataframe(df_display, use_container_width=True, hide_index=True)

            with st.expander("⚙️ ជម្រើសលុបទិន្នន័យខែនេះ ដើម្បីបង្កើតឡើងវិញ"):
                st.warning("⚠️ ប្រសិនបើលោកអ្នកចង់លុបទិន្នន័យស្បៀងប្រចាំថ្ងៃនៃខែនេះទាំងអស់ដើម្បីបញ្ចូលថ្មី សូមចុចប៊ូតុងខាងក្រោម៖")
                if st.button(f"🗑️ លុបទិន្នន័យស្បៀងខែ {sel_dr_month} ឆ្នាំ {sel_dr_year} របស់សាលានេះ", type="secondary", key="btn_clear_dr_month"):
                    c_del = conn.cursor()
                    m_pat = f"{sel_dr_year:04d}-{dr_m_num:02d}%"
                    c_del.execute("DELETE FROM daily_records WHERE school_name=? AND date LIKE ?", (sel_school, m_pat))
                    conn.commit()
                    st.success(f"✅ បានលុបទិន្នន័យខែ {sel_dr_month} ឆ្នាំ {sel_dr_year} រួចរាល់!")
                    st.rerun()

    # ================= TAB 3: កាលវិភាគមុខម្ហូបប្រចាំសប្ដាហ៍ =================
    with tab_cards:
        if not is_school_selected or not school_menus:
            if not is_school_selected:
                st.info("💡 **គំរូទទេ៖** សូមជ្រើសរើស **ខេត្ត ស្រុក ឃុំ និង សាលារៀន** ខាងលើ ដើម្បីមើលកាលវិភាគមុខម្ហូបប្រចាំសប្ដាហ៍។")
            else:
                st.info(f"💡 **គំរូទទេ៖** សាលាបឋមសិក្សា «**{sel_school}**» មិនទាន់មានទិន្នន័យបញ្ជីមុខម្ហូបនៅឡើយទេ!")
                if st.button("🚀 អនុវត្តគំរូស្ដង់ដារ MoEYS 2026 (៧ ថ្ងៃ) ជូនសាលានេះភ្លាមៗ", key="btn_quick_seed_tab1", type="primary"):
                    apply_template_to_school(conn, sel_school, template_id="cycle_1", student_count=100)
                    st.success(f"✅ បានកំណត់គំរូស្ដង់ដារ MoEYS ជូនសាលា {sel_school} ជោគជ័យ!")
                    st.rerun()

            day_colors = {
                "ចន្ទ": ("#eff6ff", "#1d4ed8", "🟦"),
                "អង្គារ": ("#fdf2f8", "#be185d", "🟪"),
                "ពុធ": ("#f0fdf4", "#15803d", "🟩"),
                "ព្រហស្បតិ៍": ("#fffbeb", "#b45309", "🟧"),
                "សុក្រ": ("#f0f9ff", "#0369a1", "🩵"),
                "សៅរ៍": ("#faf5ff", "#7e22ce", "🟣"),
                "អាទិត្យ": ("#fff1f2", "#be123c", "🔴"),
            }

            for row_idx, days_chunk in enumerate([["ចន្ទ", "អង្គារ", "ពុធ", "ព្រហស្បតិ៍"], ["សុក្រ", "សៅរ៍", "អាទិត្យ"]]):
                cols = st.columns(len(days_chunk))
                for col_idx, d_name in enumerate(days_chunk):
                    with cols[col_idx]:
                        render_clean_html(f"""
                        <div style="background-color: #f8fafc; border: 1.5px dashed #cbd5e1; border-radius: 12px; padding: 20px; text-align: center; min-height: 200px; display: flex; flex-direction: column; justify-content: center; align-items: center; margin-bottom: 12px;">
                            <div style="font-size: 1.8rem; margin-bottom: 6px;">🍽️</div>
                            <div style="font-weight: bold; color: #64748b; margin-bottom: 4px;">ថ្ងៃ{d_name}</div>
                            <div style="font-size: 0.85rem; color: #94a3b8;">មិនទាន់មានមុខម្ហូប (គំរូទទេ)</div>
                        </div>
                        """)
        else:
            st.subheader(f"📅 កាលវិភាគមុខម្ហូបប្រចាំសប្ដាហ៍ (សាលាបឋមសិក្សា៖ {sel_school})")
            
            day_colors = {
                "ចន្ទ": ("#eff6ff", "#1d4ed8", "🟦"),
                "អង្គារ": ("#fdf2f8", "#be185d", "🟪"),
                "ពុធ": ("#f0fdf4", "#15803d", "🟩"),
                "ព្រហស្បតិ៍": ("#fffbeb", "#b45309", "🟧"),
                "សុក្រ": ("#f0f9ff", "#0369a1", "🩵"),
                "សៅរ៍": ("#faf5ff", "#7e22ce", "🟣"),
                "អាទិត្យ": ("#fff1f2", "#be123c", "🔴"),
            }

            menus_by_day = {m["day_of_week"]: m for m in school_menus}
            
            for row_idx, days_chunk in enumerate([["ចន្ទ", "អង្គារ", "ពុធ", "ព្រហស្បតិ៍"], ["សុក្រ", "សៅរ៍", "អាទិត្យ"]]):
                cols = st.columns(len(days_chunk))
                for col_idx, d_name in enumerate(days_chunk):
                    with cols[col_idx]:
                        bg_c, text_c, icon = day_colors.get(d_name, ("#f8fafc", "#334155", "⚪"))
                        m_obj = menus_by_day.get(d_name)
                        if m_obj:
                            pills_html = "".join([
                                f'<div style="display: flex; justify-content: space-between; font-size: 0.82rem; margin-bottom: 4px; background: white; padding: 4px 8px; border-radius: 6px; border: 1px solid #e2e8f0;">'
                                f'<span>{"🥩" if ("សាច់" in ing["category"] or "ត្រី" in ing["category"] or "ស៊ុត" in ing["category"]) else "🥬" if "បន្លែ" in ing["category"] else "🌾" if "អង្ករ" in ing["category"] else "🍳" if "ប្រេង" in ing["category"] else "🧂"} <b>{ing["item_name"]}</b></span>'
                                f'<span><b>{ing["total_qty"]}</b> {ing["unit"]} <span style="color: #64748b; font-size: 0.75rem;">({ing["total_cost"]:,.0f}៛)</span></span>'
                                f'</div>'
                                for ing in m_obj['ingredients']
                            ])

                            card_html = f"""
                            <div style="background-color: {bg_c}; border: 1.5px solid {text_c}40; border-radius: 12px; padding: 14px; margin-bottom: 15px; min-height: 280px; box-shadow: 0 2px 5px rgba(0,0,0,0.04);">
                                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                                    <span style="font-weight: bold; font-size: 1rem; color: {text_c};">{icon} ថ្ងៃ{d_name}</span>
                                    <span style="background: white; border: 1px solid {text_c}30; color: {text_c}; font-size: 0.78rem; padding: 2px 8px; border-radius: 12px; font-weight: bold;">{m_obj["meal_type"]}</span>
                                </div>
                                <div style="font-size: 1.15rem; font-weight: bold; color: #0f172a; margin-bottom: 6px;">{m_obj["menu_name"]}</div>
                                <div style="font-size: 0.82rem; color: #475569; margin-bottom: 10px;">👥 សិស្ស៖ <b>{m_obj["target_students"]} នាក់</b> | វដ្ត៖ {m_obj["cycle_week"]}</div>
                                <div style="border-top: 1px dashed {text_c}40; padding-top: 8px; margin-bottom: 8px;">
                                    <div style="font-size: 0.8rem; font-weight: bold; color: #334155; margin-bottom: 6px;">🥗 គ្រឿងផ្សំ ({len(m_obj["ingredients"])} មុខ)៖</div>
                                    {pills_html}
                                </div>
                                <div style="border-top: 1.5px solid {text_c}60; padding-top: 8px; margin-top: 10px; display: flex; justify-content: space-between; align-items: center;">
                                    <span style="font-size: 0.85rem; font-weight: bold; color: #334155;">💰 សរុបប្រចាំថ្ងៃ៖</span>
                                    <span style="font-size: 1rem; font-weight: bold; color: {text_c};">{m_obj["total_day_cost"]:,.0f} ៛</span>
                                </div>
                            </div>
                            """
                            render_clean_html(card_html)
                        else:
                            empty_html = f"""
                            <div style="background-color: #f8fafc; border: 1.5px dashed #cbd5e1; border-radius: 12px; padding: 20px; text-align: center; min-height: 280px; display: flex; flex-direction: column; justify-content: center; align-items: center;">
                                <div style="font-size: 2rem; margin-bottom: 8px;">🍽️</div>
                                <div style="font-weight: bold; color: #64748b; margin-bottom: 4px;">ថ្ងៃ{d_name}</div>
                                <div style="font-size: 0.85rem; color: #94a3b8;">មិនទាន់មានមុខម្ហូប</div>
                            </div>
                            """
                            render_clean_html(empty_html)

    # ================= TAB 4: អនុវត្តគំរូស្ដង់ដារ MoEYS SFIS =================
    with tab_seed:
        st.subheader("⚡ អនុវត្តគំរូស្ដង់ដារ MoEYS SFIS ជូនសាលារៀន")
        st.write("ជ្រើសរើសគំរូស្ដង់ដារដែលបានសិក្សាស្រាវជ្រាវពីប្រព័ន្ធព័ត៌មាន **MoEYS SFIS** និងឯកសារ **បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026** ដើម្បីបញ្ចូលជូនសាលានេះដោយស្វ័យប្រវត្តិ។")

        col_tpl1, col_tpl2 = st.columns([1.5, 1])
        with col_tpl1:
            tpl_choices = {
                "cycle_1": "គំរូស្ដង់ដារ MoEYS 2026 (វដ្តទី១ - សប្ដាហ៍ទី១ និងទី៣)",
                "cycle_2": "គំរូស្ដង់ដារ MoEYS (វដ្តទី២ - សប្ដាហ៍ទី២ និងទី៤ - មុខម្ហូបឆ្លាស់)",
            }
            chosen_tpl_id = st.radio("ជ្រើសរើសគំរូវដ្តមុខម្ហូប៖", list(tpl_choices.keys()), format_func=lambda x: tpl_choices[x], horizontal=True)

        with col_tpl2:
            seed_students = st.number_input("ចំនួនសិស្សទទួលទានអាហារ (នាក់)", min_value=10, max_value=2000, value=target_st or 100, step=10, key="inp_seed_students")

        tpl_data = STANDARD_SFIS_TEMPLATES[chosen_tpl_id]
        st.info(f"📋 **{tpl_data['name']}** ៖ {tpl_data['description']}")

        preview_rows = []
        for dish in tpl_data["dishes"]:
            ing_names = []
            for ig in dish["ingredients"]:
                if "គ្រាប់" in ig["unit"]:
                    q = round(seed_students * ig["gram_per_student"], 1)
                else:
                    q = round((seed_students * ig["gram_per_student"]) / 1000.0, 2)
                    if q <= 0:
                        q = round(ig.get("qty_per_100", 1.0) * (seed_students / 100.0), 2)
                ing_names.append(f"{ig['item_name']} ({q} {ig['unit']})")

            preview_rows.append({
                "ថ្ងៃ": dish["day"],
                "ឈ្មោះមុខម្ហូប": dish["dish_name"],
                "ប្រភេទអាហារ": dish["meal_type"],
                "គ្រឿងផ្សំគណនាតាមចំនួនសិស្ស": " + ".join(ing_names),
                "ចំនួនមុខ": len(dish["ingredients"])
            })

        st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)

        col_btn_apply, col_btn_sp = st.columns([2, 3])
        with col_btn_apply:
            apply_label = f"🚀 អនុវត្តគំរូនេះជូនសាលា «{sel_school}» ({seed_students} នាក់)" if is_school_selected else "🚀 អនុវត្តគំរូនេះជូនសាលារៀន"
            if st.button(apply_label, type="primary", use_container_width=True, key="btn_apply_tpl_exec"):
                if not is_school_selected:
                    st.warning("⚠️ សូមជ្រើសរើសសាលារៀនជាមុនសិនដើម្បីអនុវត្តគំរូស្ដង់ដារ!")
                else:
                    cnt = apply_template_to_school(conn, sel_school, template_id=chosen_tpl_id, student_count=seed_students, overwrite=True)
                    st.success(f"✅ បានកំណត់ និងគណនាបរិមាណស្បៀងទាំង {cnt} ថ្ងៃ ជូនសាលា {sel_school} ជោគជ័យ!")
                    st.rerun()

    # ================= TAB 5: បញ្ចូល / កែសម្រួលមុខម្ហូបដោយដៃ =================
    with tab_manual:
        manual_title = f" (សាលា៖ {sel_school})" if is_school_selected else " (គំរូទទេ)"
        st.subheader(f"➕ បញ្ចូល ឬកែសម្រួលមុខម្ហូប{manual_title}")
        
        if not is_school_selected:
            st.info("💡 **គំរូទទេ៖** សូមជ្រើសរើស **ខេត្ត ស្រុក ឃុំ និង សាលារៀន** ខាងលើ ដើម្បីបញ្ចូល ឬកែសម្រួលមុខម្ហូប។")
            df_empty_manual = pd.DataFrame(columns=["មុខទំនិញ", "ប្រភេទ", "ឯកតា", "បរិមាណសរុប", "តម្លៃរាយ (៛)", "សរុប (៛)", "ចំណាំ"])
            st.dataframe(df_empty_manual, use_container_width=True, hide_index=True)
        else:
            col_m_day, col_m_name, col_m_type = st.columns([1, 1.5, 1])
            with col_m_day:
                edit_day = st.selectbox("ជ្រើសរើសថ្ងៃនៃសប្ដាហ៍", KHMER_DAYS_OF_WEEK, key="inp_edit_menu_day")

            existing_day_menu = next((m for m in school_menus if m["day_of_week"] == edit_day), None)
            
            with col_m_name:
                def_m_name = existing_day_menu["menu_name"] if existing_day_menu else "សម្លកកូរសាច់ជ្រូក"
                popular_dishes = [
                    "សម្លកកូរសាច់ជ្រូក", "បាយឆាសណ្តែកកួរនិងស៊ុតទា", "ស្ងោរស្ពៃសាច់ជ្រូក",
                    "សម្លម្ជូរគ្រឿងត្រីអណ្ដែង", "សម្លម្ជូរយួនត្រីផ្ទក់ (ត្រឡាច និងប៉េងប៉ោះ)", "ឆាល្ពៅពងទា",
                    "សម្លកកូរត្រីអណ្តែង", "ឆាននោងនិងស៊ុតទា", "ស្ងោរស្ពៃចង្កឹះត្រីប្រា",
                    "ឆាបន្លែគ្រប់មុខសាច់ជ្រូក", "សម្លម្ជូរគ្រឿងត្រីប្រា", "បាយឆាពងទាការ៉ុត",
                    "➕ វាយឈ្មោះមុខម្ហូបថ្មី..."
                ]
                sel_dish_choice = st.selectbox("ជ្រើសរើសឈ្មោះមុខម្ហូប (MoEYS SFIS)", popular_dishes, index=(popular_dishes.index(def_m_name) if def_m_name in popular_dishes else len(popular_dishes)-1), key="sel_preset_dish")
                if sel_dish_choice == "➕ វាយឈ្មោះមុខម្ហូបថ្មី...":
                    dish_name_input = st.text_input("វាយឈ្មោះមុខម្ហូបថ្មី", value=def_m_name if def_m_name not in popular_dishes else "", key="inp_custom_dish_name").strip()
                else:
                    dish_name_input = sel_dish_choice

            with col_m_type:
                meal_type_input = st.selectbox("ប្រភេទពេលអាហារ", ["អាហារពេលព្រឹក", "អាហារថ្ងៃត្រង់"], key="inp_meal_type")

            col_st_cnt, col_note = st.columns([1, 2])
            with col_st_cnt:
                target_st_input = st.number_input("ចំនួនសិស្សទទួលទាន (នាក់)", min_value=1, value=existing_day_menu["target_students"] if existing_day_menu else (target_st or 100), key="inp_dish_target_st")
            with col_note:
                note_input = st.text_input("ចំណាំ / ការណែនាំចម្អិន", value=existing_day_menu["notes"] if existing_day_menu else "", key="inp_dish_notes")

            st.markdown("#### 🥗 បញ្ជីគ្រឿងផ្សំក្នុងមុខម្ហូបនេះ")
            
            state_key = f"draft_ing_{sel_school}_{edit_day}"
            if state_key not in st.session_state or st.session_state.get(f"last_day_viewed_{sel_school}") != edit_day:
                if existing_day_menu:
                    st.session_state[state_key] = [
                        {
                            "item_name": ing["item_name"],
                            "category": ing["category"],
                            "unit": ing["unit"],
                            "total_qty": float(ing["total_qty"] or 0),
                            "unit_price": float(ing["unit_price"] or 0),
                            "total_cost": float(ing["total_cost"] or 0),
                            "note": ing["note"] or ""
                        }
                        for ing in existing_day_menu["ingredients"]
                    ]
                else:
                    st.session_state[state_key] = []
                st.session_state[f"last_day_viewed_{sel_school}"] = edit_day

            with st.expander("➕ បន្ថែមគ្រឿងផ្សំចូលក្នុងមុខម្ហូបនេះ", expanded=True):
                col_ing1, col_ing2, col_ing3, col_ing4, col_ing5 = st.columns([2, 1, 1, 1.2, 1])
                with col_ing1:
                    cat_item_names = [p["name"] for p in SUPPLIER_PRODUCT_CATALOG]
                    ing_item_sel = st.selectbox("មុខទំនិញ (៦១ មុខស្ដង់ដារ)", cat_item_names + ["➕ វាយឈ្មោះទំនិញថ្មី..."], key="inp_add_ing_name")
                    if ing_item_sel == "➕ វាយឈ្មោះទំនិញថ្មី...":
                        final_ing_name = st.text_input("ឈ្មោះទំនិញថ្មី", key="inp_custom_ing_name").strip()
                    else:
                        final_ing_name = ing_item_sel

                with col_ing2:
                    cat_match = next((p["category"] for p in SUPPLIER_PRODUCT_CATALOG if p["name"] == final_ing_name), "បន្លែ")
                    ing_cat = st.selectbox("ប្រភេទ", ["បន្លែ", "ត្រី សាច់ ស៊ុត", "អង្ករ", "ប្រេងឆា", "អំបិល"], index=(["បន្លែ", "ត្រី សាច់ ស៊ុត", "អង្ករ", "ប្រេងឆា", "អំបិល"].index(cat_match) if cat_match in ["បន្លែ", "ត្រី សាច់ ស៊ុត", "អង្ករ", "ប្រេងឆា", "អំបិល"] else 0), key="inp_add_ing_cat")

                with col_ing3:
                    unit_match = next((p["unit"] for p in SUPPLIER_PRODUCT_CATALOG if p["name"] == final_ing_name), "1គីឡូ")
                    ing_unit = st.selectbox("ឯកតា", ["1គីឡូ", "1គ្រាប់", "លីត្រ", "កញ្ចប់"], index=(0 if "គីឡូ" in unit_match else 1 if "គ្រាប់" in unit_match else 0), key="inp_add_ing_unit")

                with col_ing4:
                    ing_qty = st.number_input("បរិមាណសរុប (គ.ក/គ្រាប់)", min_value=0.1, value=2.0 if "គីឡូ" in ing_unit else 35.0, step=0.5, key="inp_add_ing_qty")

                with col_ing5:
                    def_price = get_active_item_price(conn, final_ing_name, sel_school, act_comm)
                    ing_price = st.number_input("តម្លៃរាយ (៛)", min_value=0, value=int(round(def_price)), step=100, key="inp_add_ing_price")

                if st.button("➕ បញ្ចូលគ្រឿងផ្សំនេះ", key="btn_add_ing_draft"):
                    if final_ing_name:
                        st.session_state[state_key].append({
                            "item_name": final_ing_name,
                            "category": ing_cat,
                            "unit": ing_unit,
                            "total_qty": ing_qty,
                            "unit_price": ing_price,
                            "total_cost": round(ing_qty * ing_price, 2),
                            "note": ""
                        })
                        st.rerun()

            current_draft = st.session_state.get(state_key, [])
            if current_draft:
                df_draft = pd.DataFrame(current_draft)
                st.write(f"📋 បញ្ជីគ្រឿងផ្សំបច្ចុប្បន្នសម្រាប់ថ្ងៃ{edit_day} (សរុប {len(current_draft)} មុខ)៖")
                sub_cost = sum(x["total_cost"] for x in current_draft)
                st.caption(f"💰 ទឹកប្រាក់សរុបប្រចាំថ្ងៃ៖ **{sub_cost:,.0f} ៛**")
                
                edited_df = st.data_editor(
                    df_draft,
                    use_container_width=True,
                    num_rows="dynamic",
                    key=f"editor_draft_ing_{sel_school}_{edit_day}"
                )
            else:
                st.warning(f"⚠️ មុខម្ហូបថ្ងៃ{edit_day} មិនទាន់មានគ្រឿងផ្សំនៅឡើយទេ។ សូមបន្ថែមគ្រឿងផ្សំខាងលើ!")

            col_save_m, col_del_m, col_sp_m = st.columns([1.5, 1.2, 3])
            with col_save_m:
                if st.button(f"💾 រក្សាទុកមុខម្ហូបថ្ងៃ{edit_day}", type="primary", use_container_width=True, key="btn_save_dish_manual"):
                    if not dish_name_input:
                        st.error("❌ សូមវាយឈ្មោះមុខម្ហូប!")
                    else:
                        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        c = conn.cursor()
                        
                        old_ids = [r[0] for r in c.execute("SELECT id FROM school_menus WHERE school_name=? AND day_of_week=?", (sel_school, edit_day)).fetchall()]
                        if old_ids:
                            c.executemany("DELETE FROM menu_ingredients WHERE menu_id=?", [(oid,) for oid in old_ids])
                            c.execute("DELETE FROM school_menus WHERE school_name=? AND day_of_week=?", (sel_school, edit_day))

                        c.execute("""
                            INSERT INTO school_menus (
                                school_name, commune, district, province,
                                menu_name, day_of_week, meal_type, target_students,
                                cycle_week, notes, is_active, created_at, updated_at
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'មុខម្ហូបផ្ទាល់ខ្លួន', ?, 1, ?, ?)
                        """, (sel_school, act_comm, act_dist, act_prov, dish_name_input, edit_day, meal_type_input, target_st_input, note_input, now_str, now_str))
                        new_m_id = c.lastrowid

                        items_to_save = edited_df.to_dict('records') if (edited_df is not None and not edited_df.empty) else current_draft
                        for r_ing in items_to_save:
                            u_p = float(r_ing.get("unit_price") or 0)
                            q_v = float(r_ing.get("total_qty") or 0)
                            t_c = round(q_v * u_p, 2)
                            c.execute("""
                                INSERT INTO menu_ingredients (
                                    menu_id, item_name, category, unit,
                                    gram_per_student, total_qty, unit_price, total_cost, note
                                ) VALUES (?, ?, ?, ?, 0, ?, ?, ?, ?)
                            """, (new_m_id, r_ing["item_name"], r_ing.get("category", ""), r_ing.get("unit", "1គីឡូ"), q_v, u_p, t_c, r_ing.get("note", "")))

                        conn.commit()
                        st.session_state[state_key] = items_to_save
                        st.success(f"✅ បានរក្សាទុកមុខម្ហូប «{dish_name_input}» សម្រាប់ថ្ងៃ{edit_day} ជោគជ័យ!")
                        st.rerun()

            with col_del_m:
                if existing_day_menu:
                    if st.button(f"🗑️ លុបមុខម្ហូបថ្ងៃ{edit_day}", key="btn_del_day_menu", use_container_width=True):
                        c = conn.cursor()
                        c.execute("DELETE FROM menu_ingredients WHERE menu_id=?", (existing_day_menu["id"],))
                        c.execute("DELETE FROM school_menus WHERE id=?", (existing_day_menu["id"],))
                        conn.commit()
                        st.success(f"🗑️ បានលុបមុខម្ហូបថ្ងៃ{edit_day} រួចរាល់!")
                        st.rerun()

    # ================= TAB 6: តារាងតម្រូវការស្បៀងប្រចាំខែ =================
    with tab_matrix:
        matrix_title = f" (សាលា៖ {sel_school})" if is_school_selected else " (គំរូទទេ)"
        st.subheader(f"📊 តារាងតម្រូវការបន្លែ ត្រី សាច់ ស៊ុត ប្រចាំខែ{matrix_title}")
        st.caption("គំរូទម្រង់ស្ដង់ដារដូចក្នុងឯកសារ «បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm» (សន្លឹក ចំនួនសរុប ខាងកើត)")

        col_days_opt, col_stat_opt = st.columns([1, 2])
        with col_days_opt:
            days_per_month_inp = st.number_input("ចំនួនថ្ងៃហូបក្នុង១ខែ (ថ្ងៃ)", min_value=1, max_value=10, value=4, step=1, key="inp_matrix_days")

        matrix_data = calculate_school_monthly_matrix(conn, sel_school, days_per_month=days_per_month_inp) if is_school_selected else {"rows": [], "grand_total_cost": 0.0, "total_items": 0}
        rows_m = matrix_data["rows"]

        if not is_school_selected or not rows_m:
            if not is_school_selected:
                st.info("💡 **គំរូតារាងទទេ៖** សូមជ្រើសរើស **ខេត្ត ស្រុក ឃុំ និង សាលារៀន** ខាងលើ ដើម្បីមើលតារាងតម្រូវការស្បៀងប្រចាំខែ។")
            else:
                st.info(f"💡 **គំរូតារាងទទេ៖** សាលា «**{sel_school}**» មិនទាន់មានមុខម្ហូបសម្រាប់គណនាតារាងតម្រូវការស្បៀងនៅឡើយទេ។")
            empty_m_df = pd.DataFrame(columns=[
                "ល.រ", "ថ្ងៃ", "មុខម្ហូប", "មុខទំនិញ", "ប្រភេទ", "ឯកតា", "បរិមាណប្រចាំថ្ងៃ", "ចំនួនថ្ងៃហូប", "សរុប១ខែ", "តម្លៃរាយ (៛)", "សរុបទឹកប្រាក់ (៛)"
            ])
            st.dataframe(empty_m_df, use_container_width=True, hide_index=True)
        else:
            df_m = pd.DataFrame(rows_m)
            
            sm_c1, sm_c2, sm_c3 = st.columns(3)
            with sm_c1:
                st.metric("📦 មុខទំនិញសរុប", f"{len(df_m)} ជួរទំនិញ")
            with sm_c2:
                tot_kg = sum(r["សរុប១ខែ"] for r in rows_m if "គីឡូ" in r["ឯកតា"])
                tot_eggs = sum(r["សរុប១ខែ"] for r in rows_m if "គ្រាប់" in r["ឯកតា"])
                st.metric("⚖️ ទម្ងន់ស្បៀងសរុប/ខែ", f"{tot_kg:,.1f} គ.ក", f"+ ស៊ុតទា {tot_eggs:,.0f} គ្រាប់")
            with sm_c3:
                st.metric("💰 ថវិកាសរុបប្រចាំខែ", f"{matrix_data['grand_total_cost']:,.0f} ៛")

            st.dataframe(
                df_m,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "ល.រ": st.column_config.NumberColumn("ល.រ", width="small"),
                    "បរិមាណប្រចាំថ្ងៃ": st.column_config.NumberColumn("បរិមាណ/ថ្ងៃ", format="%.2f"),
                    "សរុប១ខែ": st.column_config.NumberColumn("សរុប១ខែ", format="%.2f"),
                    "តម្លៃរាយ (៛)": st.column_config.NumberColumn("តម្លៃរាយ (៛)", format="%,.0f ៛"),
                    "សរុបទឹកប្រាក់ (៛)": st.column_config.NumberColumn("សរុបទឹកប្រាក់ (៛)", format="%,.0f ៛"),
                }
            )

            st.markdown("#### 🥗 ការបែងចែកថវិកាតាមក្រុមចំណីអាហារ")
            cat_groups = {}
            for r in rows_m:
                cat = r["ប្រភេទ"] or "ផ្សេងៗ"
                cat_groups[cat] = cat_groups.get(cat, 0.0) + r["សរុបទឹកប្រាក់ (៛)"]

            if cat_groups:
                num_cols = min(len(cat_groups), 5)
                cat_cols = st.columns(num_cols)
                for i, (cat_name, cat_amt) in enumerate(cat_groups.items()):
                    with cat_cols[i % num_cols]:
                        pct = (cat_amt / matrix_data['grand_total_cost'] * 100) if matrix_data['grand_total_cost'] else 0
                        st.metric(f"🏷️ {cat_name}", f"{cat_amt:,.0f} ៛", f"{pct:.1f}% នៃថវិកា")

    # ================= TAB 7: បោះពុម្ព & ទាញយកឯកសារផ្លូវការ =================
    with tab_export:
        exp_title = f" (សាលា៖ {sel_school})" if is_school_selected else " (គំរូទទេ)"
        st.subheader(f"🖨️ បោះពុម្ព និងទាញយកតារាងមុខម្ហូបផ្លូវការ{exp_title}")
        
        col_ex1, col_ex2, col_ex3 = st.columns([1, 1, 1.5])
        with col_ex1:
            exp_month = st.selectbox("ខែ", ["មករា", "កុម្ភៈ", "មីនា", "មេសា", "ឧសភា", "មិថុនា", "កក្កដា", "សីហា", "កញ្ញា", "តុលា", "វិច្ឆិកា", "ធ្នូ"], index=2, key="sel_exp_month")
        with col_ex2:
            exp_year = st.number_input("ឆ្នាំ", min_value=2024, max_value=2035, value=2026, key="inp_exp_year")

        with col_ex3:
            st.write("")
            st.write("")
            if is_school_selected:
                excel_bytes = generate_official_menu_excel(conn, sel_school, month_name=exp_month, year_num=exp_year, days_per_month=4)
                st.download_button(
                    label=f"📥 ទាញយកជា Excel ផ្លូវការ (.xlsx)",
                    data=excel_bytes,
                    file_name=f"តារាងមុខម្ហូប_សាលា_{sel_school}_{exp_month}_{exp_year}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                    key="btn_dl_menu_excel"
                )
            else:
                st.button("📥 ទាញយកជា Excel ផ្លូវការ (.xlsx)", disabled=True, use_container_width=True, key="btn_dl_menu_excel_dis")

        # Print Preview
        st.markdown("---")
        st.markdown("#### 📄 ទិដ្ឋភាពបោះពុម្ព (Print Preview)")
        
        if not is_school_selected:
            st.info("💡 **គំរូតារាងទទេ៖** សូមជ្រើសរើស **ខេត្ត ស្រុក ឃុំ និង សាលារៀន** ខាងលើ ដើម្បីបោះពុម្ព និងទាញយកឯកសារផ្លូវការ។ ខាងក្រោមជាទម្រង់គំរូតារាងទទេ៖")
            matrix_res_p = {"rows": [], "grand_total_cost": 0.0}
            p_rows = []
        else:
            matrix_res_p = calculate_school_monthly_matrix(conn, sel_school, days_per_month=4)
            p_rows = matrix_res_p["rows"]
        
        rows_html = ""
        if not p_rows:
            rows_html = """
            <tr>
                <td colspan="11" style="text-align: center; padding: 18px; color: #64748b; font-style: italic;">
                    (គំរូតារាងទទេ - ពុំទាន់មានទិន្នន័យមុខម្ហូប)
                </td>
            </tr>
            """
        else:
            for pr in p_rows:
                day_td = f'<td style="text-align: center; font-weight: bold; background: #e0f2fe;">{pr["ថ្ងៃ"]}</td>' if pr["ថ្ងៃ"] else '<td></td>'
                dish_td = f'<td style="font-weight: bold; background: #f8fafc;">{pr["មុខម្ហូប"]}</td>' if pr["មុខម្ហូប"] else '<td></td>'
                rows_html += f"""
                <tr>
                    <td style="text-align: center;">{pr['ល.រ']}</td>
                    {day_td}
                    {dish_td}
                    <td>{pr['មុខទំនិញ']}</td>
                    <td style="text-align: center;">{pr['ប្រភេទ']}</td>
                    <td style="text-align: center;">{pr['ឯកតា']}</td>
                    <td style="text-align: right;">{pr['បរិមាណប្រចាំថ្ងៃ']:.2f}</td>
                    <td style="text-align: center;">{pr['ចំនួនថ្ងៃហូប']}</td>
                    <td style="text-align: right; font-weight: bold;">{pr['សរុប១ខែ']:.2f}</td>
                    <td style="text-align: right;">{pr['តម្លៃរាយ (៛)']:,.0f}</td>
                    <td style="text-align: right; font-weight: bold; color: #0369a1;">{pr['សរុបទឹកប្រាក់ (៛)']:,.0f}</td>
                </tr>
                """

        preview_school = sel_school if is_school_selected else "............................................"
        preview_comm = act_comm if act_comm else "...................."
        preview_dist = act_dist if act_dist else "...................."
        preview_prov = act_prov if act_prov else "...................."

        preview_html = f"""
        <div style="background: white; border: 1px solid #cbd5e1; border-radius: 8px; padding: 24px; font-family: 'Khmer OS Siemreap', 'Siemreap', sans-serif; box-shadow: 0 4px 15px rgba(0,0,0,0.06);">
            <div style="text-align: center; margin-bottom: 20px;">
                <div style="font-family: 'Khmer OS Muol Light', sans-serif; font-size: 1.15rem; font-weight: bold; color: #1e3a8a;">ព្រះរាជាណាចក្រកម្ពុជា</div>
                <div style="font-family: 'Khmer OS Muol Light', sans-serif; font-size: 1rem; font-weight: bold; color: #1e3a8a;">ជាតិ សាសនា ព្រះមហាក្សត្រ</div>
                <div style="font-size: 1.25rem; font-weight: bold; color: #0f172a; margin-top: 15px;">តារាងមុខម្ហូបប្រចាំសប្ដាហ៍ និង តារាងតម្រូវការបន្លែ ត្រី សាច់ ស៊ុត</div>
                <div style="font-size: 0.95rem; color: #475569; margin-top: 5px;">
                    សាលាបឋមសិក្សា៖ <b>{preview_school}</b> | ឃុំ៖ <b>{preview_comm}</b> | ស្រុក៖ <b>{preview_dist}</b> | ខេត្ត៖ <b>{preview_prov}</b>
                </div>
                <div style="font-size: 0.9rem; color: #0284c7; font-weight: bold; margin-top: 4px;">សម្រាប់ខែ {exp_month} ឆ្នាំ {exp_year}</div>
            </div>

            <table style="width: 100%; border-collapse: collapse; font-size: 0.88rem;">
                <thead>
                    <tr style="background-color: #1e40af; color: white;">
                        <th style="border: 1px solid #94a3b8; padding: 8px;">ល.រ</th>
                        <th style="border: 1px solid #94a3b8; padding: 8px;">ថ្ងៃ</th>
                        <th style="border: 1px solid #94a3b8; padding: 8px;">មុខម្ហូប</th>
                        <th style="border: 1px solid #94a3b8; padding: 8px;">មុខទំនិញ/គ្រឿងផ្សំ</th>
                        <th style="border: 1px solid #94a3b8; padding: 8px;">ប្រភេទ</th>
                        <th style="border: 1px solid #94a3b8; padding: 8px;">ឯកតា</th>
                        <th style="border: 1px solid #94a3b8; padding: 8px;">បរិមាណ/ថ្ងៃ</th>
                        <th style="border: 1px solid #94a3b8; padding: 8px;">ថ្ងៃហូប</th>
                        <th style="border: 1px solid #94a3b8; padding: 8px;">សរុប១ខែ</th>
                        <th style="border: 1px solid #94a3b8; padding: 8px;">តម្លៃរាយ (៛)</th>
                        <th style="border: 1px solid #94a3b8; padding: 8px;">សរុបទឹកប្រាក់ (៛)</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_html}
                    <tr style="background-color: #fee2e2; font-weight: bold;">
                        <td colspan="10" style="border: 1px solid #94a3b8; padding: 10px; text-align: right;">សរុបថវិកាស្បៀងប្រចាំខែទាំងអស់ ៖</td>
                        <td style="border: 1px solid #94a3b8; padding: 10px; text-align: right; color: #b91c1c; font-size: 1rem;">{matrix_res_p['grand_total_cost']:,.0f} ៛</td>
                    </tr>
                </tbody>
            </table>

            <div style="display: flex; justify-content: space-between; margin-top: 40px; padding: 0 40px;">
                <div style="text-align: center;">
                    <div style="font-weight: bold;">បានឃើញ និងឯកភាព</div>
                    <div style="font-size: 0.85rem; color: #475569;">ប្រធាន គមស (នាយកសាលា)</div>
                    <div style="margin-top: 60px; border-bottom: 1px dotted #94a3b8; width: 160px; margin-left: auto; margin-right: auto;"></div>
                </div>
                <div style="text-align: center;">
                    <div style="font-weight: bold;">អ្នករៀបចំតារាងមុខម្ហូប</div>
                    <div style="font-size: 0.85rem; color: #475569;">(ចុងភៅ / នាយឃ្លាំង)</div>
                    <div style="margin-top: 60px; border-bottom: 1px dotted #94a3b8; width: 160px; margin-left: auto; margin-right: auto;"></div>
                </div>
            </div>
        </div>
        """
        render_clean_html(preview_html)
