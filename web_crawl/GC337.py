import json
import time
from typing import Optional, Dict, Any
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    StaleElementReferenceException,
    TimeoutException,
)
from webdriver_manager.chrome import ChromeDriverManager


URL = "http://portal.sw.nat.gov.tw/APGQ/LoginFree?request_locale=zh_TW&breadCrumbs=JTdCJTIyYnJlYWRDcnVtYnMlMjIlM0ElNUIlN0IlMjJuYW1lJTIyJTNBJTIyJUU1JTg1JThEJUU4JUFEJTg5JUU2JTlGJUE1JUU4JUE5JUEyJUU2JTlDJThEJUU1JThCJTk5JTIyJTJDJTIydXJsJTIyJTNBJTIyJTIyJTdEJTJDJTdCJTIybmFtZSUyMiUzQSUyMiVFNSU4NSVCNiVFNCVCQiU5NiVFNyU5QiVCOCVFOSU5NyU5QyVFNiU5RiVBNSVFOCVBOSVBMiUyMiUyQyUyMnVybCUyMiUzQSUyMmNoYW5nZU1lbnVVcmwyKCclRTUlODUlQjYlRTQlQkIlOTYlRTclOUIlQjglRTklOTclOUMlRTYlOUYlQTUlRTglQTklQTInJTJDJ0FQR1FfNicpJTIyJTdEJTJDJTdCJTIybmFtZSUyMiUzQSUyMihHQzMzNyklRTklODAlODAlRTklODIlODQlRTklODAlQjIlRTUlOEYlQTMlRTklOUIlQkIlRTUlOEIlOTUlRTglQkIlOEElRTglQkMlOUIlRTglQjIlQTglRTclODklQTklRTclQTglODUlRTclOTQlQjMlRTglQUIlOEIlRTYlQTElODglRTQlQkIlQjYlRTklODAlQjIlRTUlQkElQTYlRTYlOUYlQTUlRTglQTklQTIlMjIlMkMlMjJ1cmwlMjIlM0ElMjJvcGVuTWVudSgnJTJGQVBHUSUyRkdDMzM3JyklMjIlN0QlMkMlN0IlN0QlMkMlN0IlN0QlNUQlMkMlMjJwYXRoVXJsJTIyJTNBJTIyJTIzTUVOVV9BUEdRJTJDJTIzTUVOVV9BUEdRXzYlMkMlMkZBUEdRJTJGR0MzMzclMjIlN0Q="

QUERY_ENDPOINT_SUFFIX = "/APGQ/GC337!query"


class GC337Client:
    def __init__(self, headless: bool = False, timeout: int = 40):
        options = webdriver.ChromeOptions()
        options.set_capability("goog:loggingPrefs", {"performance": "ALL"})
        if headless:
            options.add_argument("--headless=new")

        self.driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=options,
        )
        self.wait = WebDriverWait(self.driver, timeout)

        self.driver.execute_cdp_cmd("Network.enable", {})
        self.driver.get(URL)
        self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        _ = self.driver.get_log("performance")

    def close(self):
        try:
            self.driver.quit()
        except Exception:
            pass

    # -------------------------
    # NEW: overlay-safe helpers
    # -------------------------
    def _wait_overlay_gone(self, timeout: int = 20):
        """
        Wait until blockUI overlay is invisible (prevents click intercepted).
        """
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.invisibility_of_element_located(
                    (By.CSS_SELECTOR, "div.blockUI.blockOverlay, div.blockUI.blockMsg")
                )
            )
        except TimeoutException:
            # Overlay might be opacity:0 but still present; we'll still retry clicks.
            pass

    def _js_click(self, element):
        self.driver.execute_script("arguments[0].click();", element)

    def _safe_click(self, element, retries: int = 5):
        """
        Try normal click; if intercepted, wait overlay and retry; fall back to JS click.
        """
        last_err = None
        for _ in range(retries):
            try:
                self._wait_overlay_gone()
                self.wait.until(EC.element_to_be_clickable(element))
                element.click()
                return
            except ElementClickInterceptedException as e:
                last_err = e
                self._wait_overlay_gone()
                time.sleep(0.2)
            except StaleElementReferenceException as e:
                last_err = e
                time.sleep(0.2)
            except Exception as e:
                last_err = e
                time.sleep(0.2)

        # final fallback
        try:
            self._wait_overlay_gone()
            self._js_click(element)
            return
        except Exception:
            raise last_err if last_err else RuntimeError("click failed for unknown reason")

    # -------------------------
    # Page element finders
    # -------------------------
    def _find_plate_input(self):
        candidates = [
            (By.ID, "licenseNo"),
            (By.NAME, "gc337FormBean.licenseNo"),
            (By.CSS_SELECTOR, "input#licenseNo"),
            (By.CSS_SELECTOR, "input[name='gc337FormBean.licenseNo']"),
            (By.XPATH, "//input[@id='licenseNo' or @name='gc337FormBean.licenseNo']"),
        ]
        last_err = None
        for by, sel in candidates:
            try:
                self._wait_overlay_gone()
                return self.wait.until(EC.element_to_be_clickable((by, sel)))
            except Exception as e:
                last_err = e
        raise RuntimeError(f"plate input not found/clickable: {last_err}")

    def _find_query_button(self):
        return self.wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(.,'查詢')] | //input[@value='查詢']")
            )
        )

    # -------------------------
    # Network capture
    # -------------------------
    def _find_request_id(self, deadline_seconds: int = 20) -> Optional[str]:
        deadline = time.time() + deadline_seconds
        while time.time() < deadline:
            for entry in self.driver.get_log("performance"):
                msg = json.loads(entry["message"]).get("message", {})
                if msg.get("method") != "Network.requestWillBeSent":
                    continue
                params = msg.get("params", {})
                req = params.get("request", {})
                if req.get("url", "").endswith(QUERY_ENDPOINT_SUFFIX):
                    return params.get("requestId")
            time.sleep(0.1)
        return None

    def _get_response_body(self, request_id: str) -> Optional[str]:
        for _ in range(120):
            try:
                resp = self.driver.execute_cdp_cmd(
                    "Network.getResponseBody", {"requestId": request_id}
                )
                body = resp.get("body")
                if body:
                    return body
            except Exception:
                pass
            time.sleep(0.1)
        return None

    # -------------------------
    # Public query APIs
    # -------------------------
    def query_plate_raw(self, plate: str) -> Dict[str, Any]:
        plate = str(plate).strip().upper()
        if not plate:
            return {"ok": False, "error": "empty plate"}

        # clear perf log buffer so we only capture this request
        _ = self.driver.get_log("performance")

        self._wait_overlay_gone()

        plate_input = self._find_plate_input()
        try:
            plate_input.click()
        except ElementClickInterceptedException:
            self._wait_overlay_gone()
            self._js_click(plate_input)

        # clear + type
        plate_input.send_keys(Keys.CONTROL, "a")
        plate_input.send_keys(Keys.BACKSPACE)
        plate_input.send_keys(plate)

        # click query
        self._wait_overlay_gone()
        btn = self._find_query_button()
        try:
            btn.click()
        except ElementClickInterceptedException:
            self._wait_overlay_gone()
            self._js_click(btn)

        # sometimes overlay briefly appears; wait it out
        time.sleep(0.1)
        self._wait_overlay_gone(timeout=20)

        request_id = self._find_request_id()
        if not request_id:
            return {"ok": False, "error": "request not found", "plate": plate}

        body = self._get_response_body(request_id)
        if not body:
            return {"ok": False, "error": "empty response", "plate": plate}

        try:
            data = json.loads(body)
            data["ok"] = True
            return data
        except Exception as e:
            return {"ok": False, "error": f"json decode failed: {e}", "plate": plate}

    # pick latest crtDate
    def query_plate_first_row(self, plate: str) -> Dict[str, Any]:
        data = self.query_plate_raw(plate)
        if not data.get("ok"):
            return data

        gm = data.get("gridModel")
        if not isinstance(gm, list) or not gm:
            return {"ok": False, "error": "gridModel empty", "plate": plate}

        def parse_crtdate(row: Dict[str, Any]) -> datetime:
            return datetime.strptime(row["crtDate"], "%Y/%m/%d")

        try:
            latest_row = max(gm, key=parse_crtdate)
            latest_row["ok"] = True
            return latest_row
        except Exception as e:
            return {
                "ok": False,
                "error": f"failed to select latest crtDate: {e}",
                "plate": plate,
                "gridModel": gm,
            }
