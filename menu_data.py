# -*- coding: utf-8 -*-
"""
menu_data.py
ម៉ូឌុលគ្រប់គ្រងបញ្ជីមុខម្ហូប និងកាលវិភាគអាហារូបត្ថម្ភតាមសាលារៀន (School Feeding Menus & Nutritional Scheduling)
ស្របតាមស្តង់ដារប្រព័ន្ធព័ត៌មានគ្រប់គ្រងកម្មវិធីផ្តល់អាហារតាមសាលារៀន MoEYS SFIS (https://sfis.moeys.gov.kh)
និងរចនាសម្ព័ន្ធឯកសារជាក់ស្តែង «បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm» (សន្លឹក ចំនួនសរុប ខាងកើត)
"""

import io
import re
import sqlite3
from datetime import datetime, date
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from catalog_data import SUPPLIER_PRODUCT_CATALOG

# Khmer Day of Week mapping
KHMER_DAYS_OF_WEEK = ["ចន្ទ", "អង្គារ", "ពុធ", "ព្រហស្បតិ៍", "សុក្រ", "សៅរ៍"]
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
            }
        ]
    }
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

    st.markdown("""
    <div style="background: linear-gradient(135deg, #1e3a8a 0%, #0369a1 100%); color: white; padding: 14px 20px; border-radius: 12px; margin-bottom: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.08);">
        <div style="font-size: 1.05rem; font-weight: bold; margin-bottom: 4px;">
            🇰🇭 កម្មវិធីផ្តល់អាហារតាមសាលារៀនដោយប្រើប្រាស់កសិផលក្នុងសហគមន៍ (HGSF) - ក្រសួងអប់រំ យុវជន និងកីឡា
        </div>
        <div style="font-size: 0.9rem; opacity: 0.92;">
            យោងតាមគោលការណ៍ណែនាំស្តីពីអាហារូបត្ថម្ភរបស់ប្រព័ន្ធ <b><a href="https://sfis.moeys.gov.kh/users/sign_in" target="_blank" style="color: #fef08a; text-decoration: underline;">MoEYS SFIS</a></b> និងកម្មវិធីស្បៀងអាហារពិភពលោក (WFP)៖ មុខម្ហូបប្រចាំសប្ដាហ៍ត្រូវបានរៀបចំជាវដ្តវិលជុំ (ចន្ទ ដល់ សៅរ៍) ដោយប្រើប្រាស់បន្លែ ត្រី សាច់ ស៊ុត ស្រស់ៗពីកសិករក្នុងសហគមន៍ ធានាបាននូវតុល្យភាពសារធាតុចិញ្ចឹមគ្រប់គ្រាន់សម្រាប់កុមារ។
        </div>
    </div>
    """, unsafe_allow_html=True)

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
            sel_dist = st.selectbox("ក្រុង/ស្រុក", dist_list, key="menu_dist")

    with col_c:
        d_for_c = sel_dist if sel_dist != "-- ទាំងអស់ --" else (user_dist if not is_admin else None)
        comm_list = get_scoped_commune_choices(p_for_d, d_for_c, prefix_all=(is_admin or not user_comm))
        if not is_admin and user_comm:
            sel_comm = st.selectbox("ឃុំ/សង្កាត់", [user_comm], key="menu_comm")
        else:
            sel_comm = st.selectbox("ឃុំ/សង្កាត់", comm_list, key="menu_comm")

    with col_s:
        c_for_s = sel_comm if sel_comm != "-- ទាំងអស់ --" else (user_comm if not is_admin else None)
        school_options = get_scoped_schools(p_for_d, d_for_c, c_for_s, prefix_all=False)
        if not school_options:
            school_options = ["គ្មានសាលា"]
        sel_school = st.selectbox("សាលារៀន", school_options, key="menu_school")

    if sel_school == "គ្មានសាលា" or not sel_school:
        st.warning("⚠️ សូមជ្រើសរើសសាលារៀនដើម្បីមើល ឬបញ្ចូលបញ្ជីមុខម្ហូប!")
        return

    # ទាញយកទីតាំងជាក់ស្តែងរបស់សាលា
    prov_db, dist_db, comm_db, vill_db = get_school_location_info(sel_school)
    act_prov = prov_db or (sel_prov if sel_prov != "-- ទាំងអស់ --" else "")
    act_dist = dist_db or (sel_dist if sel_dist != "-- ទាំងអស់ --" else "")
    act_comm = comm_db or (sel_comm if sel_comm != "-- ទាំងអស់ --" else "")

    # ទាញយកទិន្នន័យមុខម្ហូបបច្ចុប្បន្នរបស់សាលា
    school_menus = get_school_menu_overview(conn, sel_school)
    dish_count = len(school_menus)
    tot_week_cost = sum(m["total_day_cost"] for m in school_menus)
    tot_month_est = tot_week_cost * 4.0
    target_st = school_menus[0]["target_students"] if school_menus else 100

    # បង្ហាញកាតសង្ខេបព័ត៌មានសាលា
    st.markdown(f"""
    <div style="background: white; border: 1px solid #e2e8f0; border-radius: 12px; padding: 16px 20px; margin-bottom: 20px; box-shadow: 0 2px 6px rgba(0,0,0,0.03);">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
            <div>
                <span style="font-size: 1.25rem; font-weight: bold; color: #1e293b;">🏫 សាលាបឋមសិក្សា៖ <span style="color: #0284c7;">{sel_school}</span></span>
                <span style="font-size: 0.9rem; color: #64748b; margin-left: 12px;">📍 ឃុំ៖ <b>{act_comm or 'មិនទាន់បញ្ជាក់'}</b> | ស្រុក៖ <b>{act_dist or 'មិនទាន់បញ្ជាក់'}</b> | ខេត្ត៖ <b>{act_prov or 'មិនទាន់បញ្ជាក់'}</b></span>
            </div>
            <div>
                {'<span style="background: #dcfce7; color: #166534; padding: 6px 12px; border-radius: 20px; font-weight: bold; font-size: 0.85rem;">✅ បានកំណត់មុខម្ហូបរួចរាល់</span>' if dish_count >= 6 else '<span style="background: #fef3c7; color: #92400e; padding: 6px 12px; border-radius: 20px; font-weight: bold; font-size: 0.85rem;">⚠️ មិនទាន់គ្រប់ ៦ ថ្ងៃ</span>' if dish_count > 0 else '<span style="background: #fee2e2; color: #991b1b; padding: 6px 12px; border-radius: 20px; font-weight: bold; font-size: 0.85rem;">❌ ពុំទាន់មានមុខម្ហូប</span>'}
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # KPI Metrics
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    with kpi1:
        st.metric("🍲 មុខម្ហូបក្នុងសប្ដាហ៍", f"{dish_count} មុខ", f"{dish_count}/៦ ថ្ងៃ" if dish_count < 6 else "ពេញលេញ")
    with kpi2:
        st.metric("👥 សិស្សទទួលទានគោលដៅ", f"{target_st:,} នាក់", "គណនាស្វ័យប្រវត្ត")
    with kpi3:
        st.metric("💰 ថវិកាស្បៀង ១សប្ដាហ៍", f"{tot_week_cost:,.0f} ៛", f"{(tot_week_cost/target_st if target_st else 0):,.0f} ៛/សិស្ស/សប្ដាហ៍")
    with kpi4:
        st.metric("📅 ថវិកាស្បៀង ១ខែ (៤សប្ដាហ៍)", f"{tot_month_est:,.0f} ៛", "ប៉ាន់ស្មានតាមមុខម្ហូប")

    # ៥ ផ្ទាំងបញ្ជា (Tabs)
    tab_cards, tab_seed, tab_manual, tab_matrix, tab_export = st.tabs([
        "📅 កាលវិភាគមុខម្ហូបប្រចាំសប្ដាហ៍ (Weekly Schedule)",
        "⚡ អនុវត្តគំរូស្ដង់ដារ MoEYS SFIS (One-Click Seed)",
        "➕ បញ្ចូល / កែសម្រួលមុខម្ហូប (Add & Edit Dishes)",
        "📊 តារាងតម្រូវការស្បៀងប្រចាំខែ (Monthly Matrix)",
        "🖨️ បោះពុម្ព & ទាញយកឯកសារផ្លូវការ (Excel / Print)",
    ])

    # ================= TAB 1: កាលវិភាគមុខម្ហូបប្រចាំសប្ដាហ៍ =================
    with tab_cards:
        if not school_menus:
            st.info(f"💡 សាលាបឋមសិក្សា «**{sel_school}**» មិនទាន់មានទិន្នន័យបញ្ជីមុខម្ហូបនៅឡើយទេ!")
            st.markdown("""
            លោកអ្នកអាចជ្រើសរើសជម្រើសមួយក្នុងចំណោមពីរខាងក្រោម៖
            1. ចុចផ្ទាំង **⚡ អនុវត្តគំរូស្ដង់ដារ MoEYS SFIS** ដើម្បីបញ្ចូលមុខម្ហូបផ្លូវការទាំង ៦ ថ្ងៃភ្លាមៗក្នុង ១ ឃ្លីក
            2. ចុចផ្ទាំង **➕ បញ្ចូល / កែសម្រួលមុខម្ហូប** ដើម្បីបញ្ចូលមុខម្ហូបដោយដៃផ្ទាល់តាមការចង់បាន
            """)
            if st.button("🚀 អនុវត្តគំរូស្ដង់ដារ MoEYS 2026 (វដ្តទី១) ជូនសាលានេះភ្លាមៗ", key="btn_quick_seed_tab1", type="primary"):
                apply_template_to_school(conn, sel_school, template_id="cycle_1", student_count=100)
                st.success(f"✅ បានកំណត់គំរូស្ដង់ដារ MoEYS ជូនសាលា {sel_school} ជោគជ័យ!")
                st.rerun()
        else:
            st.subheader(f"📅 កាលវិភាគមុខម្ហូបប្រចាំសប្ដាហ៍ (សាលាបឋមសិក្សា៖ {sel_school})")
            
            # បង្ហាញជា Grid កាត ៦ ថ្ងៃនៃសប្ដាហ៍
            day_colors = {
                "ចន្ទ": ("#eff6ff", "#1d4ed8", "🟦"),
                "អង្គារ": ("#fdf2f8", "#be185d", "🟪"),
                "ពុធ": ("#f0fdf4", "#15803d", "🟩"),
                "ព្រហស្បតិ៍": ("#fffbeb", "#b45309", "🟧"),
                "សុក្រ": ("#f0f9ff", "#0369a1", "🩵"),
                "សៅរ៍": ("#faf5ff", "#7e22ce", "🟣"),
            }

            # 2 Rows of 3 columns
            menus_by_day = {m["day_of_week"]: m for m in school_menus}
            
            for row_idx, days_chunk in enumerate([["ចន្ទ", "អង្គារ", "ពុធ"], ["ព្រហស្បតិ៍", "សុក្រ", "សៅរ៍"]]):
                cols = st.columns(3)
                for col_idx, d_name in enumerate(days_chunk):
                    with cols[col_idx]:
                        bg_c, text_c, icon = day_colors.get(d_name, ("#f8fafc", "#334155", "⚪"))
                        m_obj = menus_by_day.get(d_name)
                        if m_obj:
                            st.markdown(f"""
                            <div style="background-color: {bg_c}; border: 1.5px solid {text_c}40; border-radius: 12px; padding: 14px; margin-bottom: 15px; min-height: 280px; box-shadow: 0 2px 5px rgba(0,0,0,0.04);">
                                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                                    <span style="font-weight: bold; font-size: 1rem; color: {text_c};">{icon} ថ្ងៃ{d_name}</span>
                                    <span style="background: white; border: 1px solid {text_c}30; color: {text_c}; font-size: 0.78rem; padding: 2px 8px; border-radius: 12px; font-weight: bold;">{m_obj['meal_type']}</span>
                                </div>
                                <div style="font-size: 1.15rem; font-weight: bold; color: #0f172a; margin-bottom: 6px;">{m_obj['menu_name']}</div>
                                <div style="font-size: 0.82rem; color: #475569; margin-bottom: 10px;">👥 សិស្ស៖ <b>{m_obj['target_students']} នាក់</b> | វដ្ត៖ {m_obj['cycle_week']}</div>
                                <div style="border-top: 1px dashed {text_c}40; padding-top: 8px; margin-bottom: 8px;">
                                    <div style="font-size: 0.8rem; font-weight: bold; color: #334155; margin-bottom: 4px;">🥗 គ្រឿងផ្សំ ({len(m_obj['ingredients'])} មុខ)៖</div>
                            """, unsafe_allow_html=True)
                            
                            # Ingredients Pills
                            for ing in m_obj['ingredients']:
                                cat_badge = "🥩" if ing['category'] == "ត្រី សាច់ ស៊ុត" else "🥬" if ing['category'] == "បន្លែ" else "🌾" if ing['category'] == "អង្ករ" else "🍳" if ing['category'] == "ប្រេងឆា" else "🧂"
                                st.markdown(f"""
                                    <div style="display: flex; justify-content: space-between; font-size: 0.82rem; margin-bottom: 3px; background: white; padding: 3px 8px; border-radius: 6px; border: 1px solid #e2e8f0;">
                                        <span>{cat_badge} <b>{ing['item_name']}</b></span>
                                        <span><b>{ing['total_qty']}</b> {ing['unit']} <span style="color: #64748b; font-size: 0.75rem;">({ing['total_cost']:,.0f}៛)</span></span>
                                    </div>
                                """, unsafe_allow_html=True)

                            st.markdown(f"""
                                </div>
                                <div style="border-top: 1.5px solid {text_c}60; padding-top: 6px; margin-top: 10px; display: flex; justify-content: space-between; align-items: center;">
                                    <span style="font-size: 0.85rem; font-weight: bold; color: #334155;">💰 សរុបប្រចាំថ្ងៃ៖</span>
                                    <span style="font-size: 1rem; font-weight: bold; color: {text_c};">{m_obj['total_day_cost']:,.0f} ៛</span>
                                </div>
                            </div>
                            """, unsafe_allow_html=True)
                        else:
                            st.markdown(f"""
                            <div style="background-color: #f8fafc; border: 1.5px dashed #cbd5e1; border-radius: 12px; padding: 20px; text-align: center; min-height: 280px; display: flex; flex-direction: column; justify-content: center; align-items: center;">
                                <div style="font-size: 2rem; margin-bottom: 8px;">🍽️</div>
                                <div style="font-weight: bold; color: #64748b; margin-bottom: 4px;">ថ្ងៃ{d_name}</div>
                                <div style="font-size: 0.85rem; color: #94a3b8;">មិនទាន់មានមុខម្ហូប</div>
                            </div>
                            """, unsafe_allow_html=True)

    # ================= TAB 2: អនុវត្តគំរូស្ដង់ដារ MoEYS SFIS =================
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

        # Preview table of dishes in this template
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
            if st.button(f"🚀 អនុវត្តគំរូនេះជូនសាលា «{sel_school}» ({seed_students} នាក់)", type="primary", use_container_width=True, key="btn_apply_tpl_exec"):
                cnt = apply_template_to_school(conn, sel_school, template_id=chosen_tpl_id, student_count=seed_students, overwrite=True)
                st.success(f"✅ បានកំណត់ និងគណនាបរិមាណស្បៀងទាំង {cnt} ថ្ងៃ ជូនសាលា {sel_school} ជោគជ័យ!")
                st.rerun()

    # ================= TAB 3: បញ្ចូល / កែសម្រួលមុខម្ហូបដោយដៃ =================
    with tab_manual:
        st.subheader(f"➕ បញ្ចូល ឬកែសម្រួលមុខម្ហូប (សាលា៖ {sel_school})")
        
        col_m_day, col_m_name, col_m_type = st.columns([1, 1.5, 1])
        with col_m_day:
            edit_day = st.selectbox("ជ្រើសរើសថ្ងៃនៃសប្ដាហ៍", KHMER_DAYS_OF_WEEK, key="inp_edit_menu_day")

        # ពិនិត្យមើលថាតើថ្ងៃនេះមានមុខម្ហូបស្រាប់ឬនៅ
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
        
        # Session state for draft ingredients
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

        # Form to add an ingredient
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

        # Display current draft ingredients table
        current_draft = st.session_state.get(state_key, [])
        if current_draft:
            df_draft = pd.DataFrame(current_draft)
            st.write(f"📋 បញ្ជីគ្រឿងផ្សំបច្ចុប្បន្នសម្រាប់ថ្ងៃ{edit_day} (សរុប {len(current_draft)} មុខ)៖")
            
            # Show summary
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
                    # Update or insert school_menus
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    c = conn.cursor()
                    
                    # Delete existing menu for this day
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

                    # Save ingredients
                    for r_ing in current_draft:
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

    # ================= TAB 4: តារាងតម្រូវការស្បៀងប្រចាំខែ =================
    with tab_matrix:
        st.subheader(f"📊 តារាងតម្រូវការបន្លែ ត្រី សាច់ ស៊ុត ប្រចាំខែ (សាលា៖ {sel_school})")
        st.caption("គំរូទម្រង់ស្ដង់ដារដូចក្នុងឯកសារ «បញ្ជីមុខម្ហូបដែលត្រូវបញ្ជាទិញ_2026.xlsm» (សន្លឹក ចំនួនសរុប ខាងកើត)")

        col_days_opt, col_stat_opt = st.columns([1, 2])
        with col_days_opt:
            days_per_month_inp = st.number_input("ចំនួនថ្ងៃហូបក្នុង១ខែ (ថ្ងៃ)", min_value=1, max_value=10, value=4, step=1, key="inp_matrix_days")

        matrix_data = calculate_school_monthly_matrix(conn, sel_school, days_per_month=days_per_month_inp)
        rows_m = matrix_data["rows"]

        if not rows_m:
            st.warning("⚠️ សាលានេះមិនទាន់មានមុខម្ហូបសម្រាប់គណនាតារាងតម្រូវការស្បៀងនៅឡើយទេ!")
        else:
            df_m = pd.DataFrame(rows_m)
            
            # Summary Metrics
            sm_c1, sm_c2, sm_c3 = st.columns(3)
            with sm_c1:
                st.metric("📦 មុខទំនិញសរុប", f"{len(df_m)} ជួរទំនិញ")
            with sm_c2:
                tot_kg = sum(r["សរុប១ខែ"] for r in rows_m if "គីឡូ" in r["ឯកតា"])
                tot_eggs = sum(r["សរុប១ខែ"] for r in rows_m if "គ្រាប់" in r["ឯកតា"])
                st.metric("⚖️ ទម្ងន់ស្បៀងសរុប/ខែ", f"{tot_kg:,.1f} គ.ក", f"+ ស៊ុតទា {tot_eggs:,.0f} គ្រាប់")
            with sm_c3:
                st.metric("💰 ថវិកាសរុបប្រចាំខែ", f"{matrix_data['grand_total_cost']:,.0f} ៛")

            # Data Table
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

            # Category Breakdown
            st.markdown("#### 🥗 ការបែងចែកថវិកាតាមក្រុមចំណីអាហារ")
            cat_groups = {}
            for r in rows_m:
                cat = r["ប្រភេទ"] or "ផ្សេងៗ"
                cat_groups[cat] = cat_groups.get(cat, 0.0) + r["សរុបទឹកប្រាក់ (៛)"]

            cat_cols = st.columns(len(cat_groups))
            for i, (cat_name, cat_amt) in enumerate(cat_groups.items()):
                with cat_cols[i]:
                    pct = (cat_amt / matrix_data['grand_total_cost'] * 100) if matrix_data['grand_total_cost'] else 0
                    st.metric(f"🏷️ {cat_name}", f"{cat_amt:,.0f} ៛", f"{pct:.1f}% នៃថវិកា")

    # ================= TAB 5: បោះពុម្ព & ទាញយកឯកសារផ្លូវការ =================
    with tab_export:
        st.subheader(f"🖨️ បោះពុម្ព និងទាញយកតារាងមុខម្ហូបផ្លូវការ (សាលា៖ {sel_school})")
        
        col_ex1, col_ex2, col_ex3 = st.columns([1, 1, 1.5])
        with col_ex1:
            exp_month = st.selectbox("ខែ", ["មករា", "កុម្ភៈ", "មីនា", "មេសា", "ឧសភា", "មិថុនា", "កក្កដា", "សីហា", "កញ្ញា", "តុលា", "វិច្ឆិកា", "ធ្នូ"], index=2, key="sel_exp_month")
        with col_ex2:
            exp_year = st.number_input("ឆ្នាំ", min_value=2024, max_value=2035, value=2026, key="inp_exp_year")

        with col_ex3:
            st.write("")
            st.write("")
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

        # Print Preview
        st.markdown("---")
        st.markdown("#### 📄 ទិដ្ឋភាពបោះពុម្ព (Print Preview)")
        
        matrix_res_p = calculate_school_monthly_matrix(conn, sel_school, days_per_month=4)
        p_rows = matrix_res_p["rows"]
        
        rows_html = ""
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

        preview_html = f"""
        <div style="background: white; border: 1px solid #cbd5e1; border-radius: 8px; padding: 30px; font-family: 'Khmer OS Siemreap', 'Siemreap', sans-serif; box-shadow: 0 4px 15px rgba(0,0,0,0.06);">
            <div style="text-align: center; margin-bottom: 20px;">
                <div style="font-family: 'Khmer OS Muol Light', sans-serif; font-size: 1.15rem; font-weight: bold; color: #1e3a8a;">ព្រះរាជាណាចក្រកម្ពុជា</div>
                <div style="font-family: 'Khmer OS Muol Light', sans-serif; font-size: 1rem; font-weight: bold; color: #1e3a8a;">ជាតិ សាសនា ព្រះមហាក្សត្រ</div>
                <div style="font-size: 1.25rem; font-weight: bold; color: #0f172a; margin-top: 15px;">តារាងមុខម្ហូបប្រចាំសប្ដាហ៍ និង តារាងតម្រូវការបន្លែ ត្រី សាច់ ស៊ុត</div>
                <div style="font-size: 0.95rem; color: #475569; margin-top: 5px;">
                    សាលាបឋមសិក្សា៖ <b>{sel_school}</b> | ឃុំ៖ <b>{act_comm}</b> | ស្រុក៖ <b>{act_dist}</b> | ខេត្ត៖ <b>{act_prov}</b>
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
        st.markdown(preview_html, unsafe_allow_html=True)
