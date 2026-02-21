import json
import time
from typing import Optional, Dict, Any
import os
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.common.exceptions import (
    TimeoutException,
    ElementClickInterceptedException,
    StaleElementReferenceException,
)

URL = "https://portal.sw.nat.gov.tw/APGQ/LoginFree?request_locale=zh_TW&breadCrumbs=JTdCJTIyYnJlYWRDcnVtYnMlMjIlM0ElNUIlN0IlMjJuYW1lJTIyJTNBJTIyJUU1JTg1JThEJUU4JUFEJTg5JUU2JTlGJUE1JUU4JUE5JUEyJUU2JTlDJThEJUU1JThCJTk5JTIyJTJDJTIydXJsJTIyJTNBJTIyJTIyJTdEJTJDJTdCJTIybmFtZSUyMiUzQSUyMiVFNSU4NSVCNiVFNCVCQiU5NiVFNyU5QiVCOCVFOSU5NyU5QyVFNiU5RiVBNSVFOCVBOSVBMiUyMiUyQyUyMnVybCUyMiUzQSUyMmNoYW5nZU1lbnVVcmwyKCclRTUlODUlQjYlRTQlQkIlOTYlRTclOUIlQjglRTklOTclOUMlRTYlOUYlQTUlRTglQTklQTInJTJDJ0FQR1FfNicpJTIyJTdEJTJDJTdCJTIybmFtZSUyMiUzQSUyMihHQzMzNSklRTglQjMlQkMlRTglQjIlQjclRTklODAlQjIlRTUlOEYlQTMlRTYlOTYlQjAlRTUlQjAlOEYlRTUlOUUlOEIlRTYlQjElQkQlRTYlQTklOUYlRTglQkIlOEElRTklODAlODAlRTklODIlODQlRTglQjIlQTglRTclODklQTklRTclQTglODUlRTclOTQlQjMlRTglQUIlOEIlRTYlQTElODglRTQlQkIlQjYlRTklODAlQjIlRTUlQkElQTYlRTYlOUYlQTUlRTglQTklQTIlMjIlMkMlMjJ1cmwlMjIlM0ElMjJvcGVuTWVudSgnJTJGQVBHUSUyRkdDMzM1JyklMjIlN0QlMkMlN0IlN0QlMkMlN0IlN0QlNUQlMkMlMjJwYXRoVXJsJTIyJTNBJTIyJTIzTUVOVV9BUEdRJTJDJTIzTUVOVV9BUEdRXzYlMkMlMkZBUEdRJTJGR0MzMzUlMjIlN0Q="
QUERY_ENDPOINT_SUFFIX = "/APGQ/GC335!query"

def get_chromedriver_path() -> str:
        if getattr(sys, "frozen", False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))

        driver_path = os.path.join(base_dir, "chromedriver.exe")
        if not os.path.exists(driver_path):
            raise FileNotFoundError(
                f"chromedriver.exe not found at:\n{driver_path}\n"
                "Put chromedriver.exe next to the app."
            )
        return driver_path

class GC335Client:
    """
    Keep ONE Chrome session open and query many plates.
    Robust against BlockUI overlay click intercept + slow responses.
    """

    def __init__(self, headless: bool = False, timeout: int = 40):
        self.timeout = timeout

        options = webdriver.ChromeOptions()
        options.set_capability("goog:loggingPrefs", {"performance": "ALL"})
        if headless:
            options.add_argument("--headless=new")

        self.driver = webdriver.Chrome(
            service=Service(get_chromedriver_path()),
            options=options,
        )
        self.wait = WebDriverWait(self.driver, timeout)

        # Enable CDP network once
        self.driver.execute_cdp_cmd("Network.enable", {})

        # Open page once
        self.driver.get(URL)
        self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

        # Select 新車車牌 once
        radio = self.wait.until(EC.presence_of_element_located((By.ID, "inlineRadio2")))
        self.driver.execute_script("arguments[0].click();", radio)

        # Cache input element
        self.plate_input = self._get_plate_input()

        # Clear old logs so we don't match previous requests
        _ = self.driver.get_log("performance")

    def close(self):
        try:
            self.driver.quit()
        except Exception:
            pass

    # -----------------------
    # Helpers (overlay + click)
    # -----------------------
    def _wait_blockui_gone(self, timeout: Optional[int] = None):
        """
        Wait until BlockUI overlay is not present OR not displayed.
        """
        t = timeout or self.timeout
        wait = WebDriverWait(self.driver, t)

        def overlay_gone(_):
            overlays = self.driver.find_elements(By.CSS_SELECTOR, "div.blockUI.blockOverlay, div.blockOverlay")
            if not overlays:
                return True
            return all(not o.is_displayed() for o in overlays)

        wait.until(overlay_gone)

    def _safe_click(self, locator, timeout: Optional[int] = None):
        """
        Wait overlay gone -> wait clickable -> click.
        If intercepted, wait overlay again then JS click.
        """
        t = timeout or self.timeout
        self._wait_blockui_gone(timeout=t)

        el = WebDriverWait(self.driver, t).until(EC.element_to_be_clickable(locator))
        try:
            el.click()
            return
        except ElementClickInterceptedException:
            self._wait_blockui_gone(timeout=t)
            el = WebDriverWait(self.driver, t).until(EC.element_to_be_clickable(locator))
            self.driver.execute_script("arguments[0].click();", el)

    def _get_plate_input(self):
        return self.wait.until(
            EC.visibility_of_element_located((By.NAME, "gc335FormBean.newLicenseNo"))
        )

    def _clear_performance_logs(self):
        try:
            _ = self.driver.get_log("performance")
        except Exception:
            pass

    # -----------------------
    # Main query methods
    # -----------------------
    def query_plate_raw(self, plate: str) -> Dict[str, Any]:
        """
        Returns the full JSON response from GC335!query (adds ok=True/False).
        """
        plate = str(plate).strip()
        if not plate:
            return {"ok": False, "error": "empty plate"}

        # Clear performance logs right before we click, so we only read fresh events
        self._clear_performance_logs()

        # Fill input (re-find if stale)
        try:
            self.plate_input.clear()
            self.plate_input.send_keys(plate)
        except StaleElementReferenceException:
            self.plate_input = self._get_plate_input()
            self.plate_input.clear()
            self.plate_input.send_keys(plate)

        # Click 查詢 safely (fixes ElementClickInterceptedException)
        search_btn = (By.ID, "searchButton")
        # fallback locator if ID ever changes
        search_btn_fallback = (By.XPATH, "//button[contains(.,'查詢')] | //input[@value='查詢']")

        try:
            self._safe_click(search_btn, timeout=30)
        except Exception:
            self._safe_click(search_btn_fallback, timeout=30)

        # After clicking, wait until overlay finishes again (query in progress)
        try:
            self._wait_blockui_gone(timeout=30)
        except Exception:
            # not fatal; continue
            pass

        # Find matching requestId by waiting for responseReceived (more reliable than requestWillBeSent only)
        request_id = None
        deadline = time.time() + 30

        while time.time() < deadline and request_id is None:
            logs = []
            try:
                logs = self.driver.get_log("performance")
            except Exception:
                pass

            for entry in logs:
                try:
                    msg = json.loads(entry["message"])["message"]
                except Exception:
                    continue

                method = msg.get("method")
                params = msg.get("params", {})

                # Prefer responseReceived: means response exists
                if method == "Network.responseReceived":
                    resp = params.get("response", {})
                    url = resp.get("url", "")
                    if url.endswith(QUERY_ENDPOINT_SUFFIX):
                        request_id = params.get("requestId")
                        break

                # Fallback: requestWillBeSent
                if method == "Network.requestWillBeSent":
                    req = params.get("request", {})
                    url = req.get("url", "")
                    if url.endswith(QUERY_ENDPOINT_SUFFIX):
                        request_id = params.get("requestId")
                        # don't break immediately; responseReceived is better, but keep it as fallback
            if request_id is None:
                time.sleep(0.1)

        if not request_id:
            return {"ok": False, "error": "could not find GC335!query requestId", "plate": plate}

        # Fetch response body (retry a bit; sometimes body isn't ready immediately)
        body = None
        for _ in range(200):  # up to ~20s
            try:
                resp = self.driver.execute_cdp_cmd("Network.getResponseBody", {"requestId": request_id})
                body = resp.get("body")
                if body:
                    break
            except Exception:
                pass
            time.sleep(0.1)

        if not body:
            return {"ok": False, "error": "no response body", "plate": plate, "requestId": request_id}

        try:
            data = json.loads(body)
            data["ok"] = True
            return data
        except Exception as e:
            return {"ok": False, "error": f"json decode failed: {e}", "plate": plate}

    def query_plate_first_row(self, plate: str) -> Dict[str, Any]:
        """
        Returns gridModel[0] (your usable row dict) or an ok=False dict.
        """
        data = self.query_plate_raw(plate)
        if not data.get("ok"):
            return data

        gm = data.get("gridModel")
        if isinstance(gm, list) and gm:
            return gm[0]

        return {"ok": False, "error": "gridModel empty", "plate": plate}
