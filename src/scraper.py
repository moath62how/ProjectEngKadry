"""
Web scraper module for fetching engineer syndicate data from data.eea.org.eg
"""

import requests
from bs4 import BeautifulSoup
import re


URL = "https://data.eea.org.eg/lastpaid.aspx"


def get_engineer_syndicate(national_id: str) -> str:
    """
    Fetches the engineer sub-syndicate (النقابة الفرعية) using the Egyptian National ID.
    
    :param national_id: 14-digit Egyptian national number
    :return: Sub-syndicate name (string)
    :raises ValueError: if input validation fails
    :raises Exception: if request fails or data not found
    """

    # ------------------ Validation ------------------
    if not isinstance(national_id, str):
        raise ValueError("National number must be a string.")

    if not re.fullmatch(r"\d{14}", national_id):
        raise ValueError("National number must be exactly 14 digits.")

    # ------------------ Session ------------------
    session = requests.Session()
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Content-Type": "application/x-www-form-urlencoded"
    }

    # ------------------ Step 1: GET ------------------
    r = session.get(URL, headers=headers, timeout=15)
    r.raise_for_status()

    soup = BeautifulSoup(r.text, "html.parser")

    def get_val(name):
        tag = soup.find("input", {"name": name})
        return tag["value"] if tag else ""

    # ------------------ Step 2: POST ------------------
    payload = {
        "__EVENTTARGET": "",
        "__EVENTARGUMENT": "",
        "__LASTFOCUS": "",
        "__VIEWSTATE": get_val("__VIEWSTATE"),
        "__VIEWSTATEGENERATOR": get_val("__VIEWSTATEGENERATOR"),
        "__EVENTVALIDATION": get_val("__EVENTVALIDATION"),
        "txtdat": get_val("txtdat"),
        "TextBox1": "",
        "TextBox2": "",
        "OldRefID": "",
        "TextBox3": "",
        "NationalNumber": national_id,
        "btnSearch": "بحث"
    }

    res = session.post(URL, data=payload, headers=headers, timeout=15)
    res.raise_for_status()

    soup = BeautifulSoup(res.text, "html.parser")

    # ------------------ Step 3: Extract Data ------------------
    synd = soup.find("input", {"id": "txtSynd"})
    if not synd or not synd.get("value"):
        raise Exception("No data found for this national number.")

    return synd["value"].strip()


def get_engineer_syndicate_safe(national_id: str) -> dict:
    """
    Safe wrapper around get_engineer_syndicate that returns a dict with status.
    
    :param national_id: 14-digit Egyptian national number
    :return: Dictionary with 'success', 'data', and optionally 'error' keys
    """
    try:
        syndicate = get_engineer_syndicate(national_id)
        return {
            "success": True,
            "national_id": national_id,
            "syndicate": syndicate
        }
    except ValueError as e:
        return {
            "success": False,
            "national_id": national_id,
            "error": f"Validation Error: {str(e)}"
        }
    except requests.exceptions.RequestException as e:
        return {
            "success": False,
            "national_id": national_id,
            "error": f"Network Error: {str(e)}"
        }
    except Exception as e:
        return {
            "success": False,
            "national_id": national_id,
            "error": str(e)
        }