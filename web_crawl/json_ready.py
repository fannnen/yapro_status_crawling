import json
import time
from pathlib import Path
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

from webdriver_manager.chrome import ChromeDriverManager


URL = "https://portal.sw.nat.gov.tw/APGQ/LoginFree?request_locale=zh_TW&breadCrumbs=JTdCJTIyYnJlYWRDcnVtYnMlMjIlM0ElNUIlN0IlMjJuYW1lJTIyJTNBJTIyJUU1JTg1JThEJUU4JUFEJTg5JUU2JTlGJUE1JUU4JUE5JUEyJUU2JTlDJThEJUU1JThCJTk5JTIyJTJDJTIydXJsJTIyJTNBJTIyJTIyJTdEJTJDJTdCJTIybmFtZSUyMiUzQSUyMiVFNSU4NSVCNiVFNCVCQiU5NiVFNyU5QiVCOCVFOSU5NyU5QyVFNiU5RiVBNSVFOCVBOSVBMiUyMiUyQyUyMnVybCUyMiUzQSUyMmNoYW5nZU1lbnVVcmwyKCclRTUlODUlQjYlRTQlQkIlOTYlRTclOUIlQjglRTklOTclOUMlRTYlOUYlQTUlRTglQTklQTInJTJDJ0FQR1FfNicpJTIyJTdEJTJDJTdCJTIybmFtZSUyMiUzQSUyMihHQzMzNSklRTglQjMlQkMlRTglQjIlQjclRTklODAlQjIlRTUlOEYlQTMlRTYlOTYlQjAlRTUlQjAlOEYlRTUlOUUlOEIlRTYlQjElQkQlRTYlQTklOUYlRTglQkIlOEElRTklODAlODAlRTklODIlODQlRTglQjIlQTglRTclODklQTklRTclQTglODUlRTclOTQlQjMlRTglQUIlOEIlRTYlQTElODglRTQlQkIlQjYlRTklODAlQjIlRTUlQkElQTYlRTYlOUYlQTUlRTglQTklQTIlMjIlMkMlMjJ1cmwlMjIlM0ElMjJvcGVuTWVudSgnJTJGQVBHUSUyRkdDMzM1JyklMjIlN0QlMkMlN0IlN0QlMkMlN0IlN0QlNUQlMkMlMjJwYXRoVXJsJTIyJTNBJTIyJTIzTUVOVV9BUEdRJTJDJTIzTUVOVV9BUEdRXzYlMkMlMkZBUEdRJTJGR0MzMzUlMjIlN0Q="

QUERY_ENDPOINT_SUFFIX = "/APGQ/GC335!query"


def _script_dir() -> Path:
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path.cwd()


def _safe(s: str) -> str:
    return "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in s)


def search_and_capture_json(plate: str) -> dict:
    # Enable performance logs so we can read Network events
    options = webdriver.ChromeOptions()
    options.set_capability("goog:loggingPrefs", {"performance": "ALL"})
    # options.add_argument("--headless=new")  # optional

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options,
    )
    wait = WebDriverWait(driver, 40)

    try:
        # Enable CDP Network so we can fetch response bodies by requestId
        driver.execute_cdp_cmd("Network.enable", {})

        driver.get(URL)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

        # Select 新車車牌
        radio = wait.until(EC.presence_of_element_located((By.ID, "inlineRadio2")))
        driver.execute_script("arguments[0].click();", radio)

        # Enter plate
        plate_input = wait.until(
            EC.visibility_of_element_located((By.NAME, "gc335FormBean.newLicenseNo"))
        )
        plate_input.clear()
        plate_input.send_keys(plate)

        # Click 查詢
        wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(.,'查詢')] | //input[@value='查詢']")
            )
        ).click()

        # Now capture the XHR response for GC335!query
        request_id = None
        post_data = None

        deadline = time.time() + 25  # seconds to find the request + response
        while time.time() < deadline and request_id is None:
            logs = driver.get_log("performance")
            for entry in logs:
                msg = json.loads(entry["message"])["message"]
                method = msg.get("method")
                params = msg.get("params", {})

                if method == "Network.requestWillBeSent":
                    req = params.get("request", {})
                    url = req.get("url", "")
                    if url.endswith(QUERY_ENDPOINT_SUFFIX):
                        request_id = params.get("requestId")
                        post_data = req.get("postData")
                        break

            if request_id is None:
                time.sleep(0.2)

        if request_id is None:
            return {
                "plate": plate,
                "ok": False,
                "error": "Could not find Network request for GC335!query in performance logs",
            }

        # Wait a bit for response body to be ready, then fetch it
        body = None
        for _ in range(60):
            try:
                resp = driver.execute_cdp_cmd("Network.getResponseBody", {"requestId": request_id})
                body = resp.get("body")
                if body:
                    break
            except Exception:
                pass
            time.sleep(0.2)

        if not body:
            return {
                "plate": plate,
                "ok": False,
                "error": "Found requestId but could not fetch response body (Network.getResponseBody failed)",
                "requestId": request_id,
                "postData": post_data,
            }

        # Parse JSON (server says text/json)
        data = json.loads(body)
        return data

    except TimeoutException:
        return {"plate": plate, "ok": False, "error": "Timed out during page interaction"}
    finally:
        driver.quit()


def search_and_get(license_plate: str) -> dict:
    result = search_and_capture_json(license_plate)
    json_data = result["gridModel"][0]
    return json_data
