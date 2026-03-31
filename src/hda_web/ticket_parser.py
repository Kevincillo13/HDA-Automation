import json
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

def _get_attribute_or_empty(driver: WebDriver, by: By, value: str, attribute: str = 'value') -> str:
    """Safely gets an attribute from an element, returning 'Empty' if it fails."""
    try:
        element = driver.find_element(by, value)
        attr_value = element.get_attribute(attribute)
        return attr_value.strip() if attr_value else "Empty"
    except NoSuchElementException:
        return "Empty"
    except Exception:
        return "Empty"

def _get_json_field_or_empty(driver: WebDriver, by: By, value: str, key: str) -> str:
    """Safely gets and parses a JSON field from an element's value attribute."""
    try:
        json_string = _get_attribute_or_empty(driver, by, value)
        if not json_string or json_string == "Empty":
            return "Empty"
        data = json.loads(json_string)
        if isinstance(data, list) and data:
            return data[0].get(key, "Empty")
        elif isinstance(data, dict):
            return data.get(key, "Empty")
        return "Empty"
    except Exception:
        return "Empty"

def extract_ticket_data(driver: WebDriver) -> dict:
    """
    Extracts all necessary data from the ticket page using Selenium.

    Args:
        driver: The Selenium WebDriver instance.

    Returns:
        A dictionary containing the extracted data.
    """
    # Wait for a reliable element to be present to ensure the page is loaded
    wait = WebDriverWait(driver, 30)
    subject_selector = (By.CSS_SELECTOR, 'input[componentid$="_C9C"]')
    wait.until(ec.presence_of_element_located(subject_selector))

    data = {
        "Company": _get_json_field_or_empty(driver, By.CSS_SELECTOR, 'input[name="C6C"]', 'value'),
        "Amount": _get_attribute_or_empty(driver, By.CSS_SELECTOR, 'input[componentid$="_PAYRQ05"]'),
        "Invoice Date": _get_attribute_or_empty(driver, By.CSS_SELECTOR, 'input[componentid$="_PAYRQ02"]'),
        "Invoice Number": _get_attribute_or_empty(driver, By.CSS_SELECTOR, 'input[componentid$="_PAYRQ03"]'),
        "Vendor #": _get_attribute_or_empty(driver, By.CSS_SELECTOR, 'input[componentid$="_PAYRQ01"]'),
        "Currency": _get_attribute_or_empty(driver, By.CSS_SELECTOR, 'input[componentid$="_PAYRQ06"]'),
        "POR #": _get_attribute_or_empty(driver, By.CSS_SELECTOR, 'input[componentid$="_PAYRQ04"]'),
        "Payable to": _get_attribute_or_empty(driver, By.CSS_SELECTOR, 'input[componentid$="_PAYRQ07"]'),
        "Address": _get_attribute_or_empty(driver, By.CSS_SELECTOR, 'input[componentid$="_PAYRQ19"]'),
        "City/State": _get_attribute_or_empty(driver, By.CSS_SELECTOR, 'input[componentid$="_PAYRQ08"]'),
        "Vendor Contact": _get_attribute_or_empty(driver, By.CSS_SELECTOR, 'input[componentid$="_PAYRQ09"]'),
        "Payment method": _get_json_field_or_empty(driver, By.CSS_SELECTOR, 'input[name="PAYRQ10"]', 'value'),
        "Subject": _get_attribute_or_empty(driver, By.CSS_SELECTOR, 'input[componentid$="_C9C"]'),
        "Cost/Profit center": _get_attribute_or_empty(driver, By.CSS_SELECTOR, 'input[componentid$="_PRCOST01"]'),
        "GL Account": _get_attribute_or_empty(driver, By.CSS_SELECTOR, 'input[componentid$="_PRGLACC01"]'),
        "WBS Element": _get_attribute_or_empty(driver, By.CSS_SELECTOR, 'input[componentid$="_PRWBS01"]'),
        "Distribution AMT": _get_attribute_or_empty(driver, By.CSS_SELECTOR, 'input[componentid$="_PRAMT01"]'),
        "Brand code": _get_attribute_or_empty(driver, By.CSS_SELECTOR, 'input[componentid$="_PRBRAND01"]'),
    }

    # Placeholders for data not directly available via simple selectors
    data["Id"] = "Empty"
    data["Created"] = "Empty"

    return data
