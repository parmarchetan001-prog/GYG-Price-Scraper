import asyncio
from datetime import datetime, timedelta
import os
import re
import pandas as pd
from playwright.async_api import async_playwright


# ============================================================
# CONFIGURATION
# ============================================================

INPUT_EXCEL = "WB_competitor_input - Copy.xlsx"
OUTPUT_EXCEL = "WB_competitor_input_Prices.xlsx"

START_DATE_STR = "2026-09-08"
END_DATE_STR = "2026-11-30"

MIN_MATCH_SCORE = 0.70

HEADLESS = False

PAGE_TIMEOUT = 60000

DEBUG_FOLDER = "debug"


# ============================================================
# DATE GENERATION
# ============================================================

def generate_date_labels(start_str, end_str):

    labels = []

    start_dt = datetime.strptime(start_str, "%Y-%m-%d")
    end_dt = datetime.strptime(end_str, "%Y-%m-%d")

    current_dt = start_dt

    while current_dt <= end_dt:

        day = current_dt.strftime("%d").lstrip("0")

        labels.append({
            "iso": current_dt.strftime("%Y-%m-%d"),
            "gyg_label": current_dt.strftime(
                f"%A, {day} %B %Y"
            )
        })

        current_dt += timedelta(days=1)

    return labels


# ============================================================
# TEXT HELPERS
# ============================================================

def normalize_text(text):

    if not text:
        return ""

    text = str(text).lower()

    replacements = {
        "&": " and ",
        "+": " ",
        "-": " ",
        "/": " ",
        "|": " ",
        "–": " ",
        "—": " ",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    text = re.sub(r"\s+", " ", text)

    return text.strip()


def clean_tokens(text):

    text = normalize_text(text)

    words = re.findall(
        r"[a-z0-9]+",
        text
    )

    ignore = {
        "with",
        "the",
        "from",
        "and",
        "for",
        "tour",
        "ticket",
        "tickets",
        "of",
        "to",
        "in",
        "by",
        "per",
        "person",
        "experience",
        "activity",
        "option",
        "entry",
        "london",
    }

    return {
        x
        for x in words
        if x not in ignore
        and len(x) > 1
    }


def calculate_token_match_score(
    target,
    card_text
):

    target_tokens = clean_tokens(target)
    card_tokens = clean_tokens(card_text)

    if not target_tokens:
        return 0.0

    matched = (
        target_tokens.intersection(card_tokens)
    )

    return len(matched) / len(target_tokens)


def calculate_exact_phrase_score(
    target,
    card_text
):

    target = normalize_text(target)
    card_text = normalize_text(card_text)

    if not target:
        return 0.0

    if target in card_text:
        return 1.0

    return 0.0


def calculate_combined_score(
    target,
    card_text
):

    token_score = calculate_token_match_score(
        target,
        card_text
    )

    phrase_score = calculate_exact_phrase_score(
        target,
        card_text
    )

    if phrase_score >= 0.9:
        return 1.0

    return (
        token_score * 0.75
        +
        phrase_score * 0.25
    )


# ============================================================
# PRICE HELPERS
# ============================================================

PRICE_REGEX = re.compile(
    r"(?:£|₹|€|$|USD|GBP|INR|EUR)\s*"
    r"\d[\d,]*(?:\.\d{1,2})?",
    re.IGNORECASE
)


def parse_price_value(text):

    if not text:
        return None

    match = PRICE_REGEX.search(
        str(text)
    )

    if not match:
        return None

    raw = match.group(0)

    number = re.sub(
        r"[^0-9.]",
        "",
        raw
    )

    try:
        return float(number)
    except Exception:
        return None


def get_price_currency(text):

    if not text:
        return None

    text = str(text)

    if "£" in text or re.search(
        r"\bGBP\b",
        text,
        re.IGNORECASE
    ):
        return "GBP"

    if "₹" in text or re.search(
        r"\bINR\b",
        text,
        re.IGNORECASE
    ):
        return "INR"

    if "€" in text or re.search(
        r"\bEUR\b",
        text,
        re.IGNORECASE
    ):
        return "EUR"

    if "$" in text or re.search(
        r"\bUSD\b",
        text,
        re.IGNORECASE
    ):
        return "USD"

    return None


def format_price(
    value,
    currency=None
):

    if value is None:
        return None

    if currency == "INR":
        return f"₹{value:,.0f}"

    if currency == "EUR":
        return f"€{value:,.2f}"

    if currency == "USD":
        return f"${value:,.2f}"

    return f"£{value:,.2f}"


# ============================================================
# STRIKETHROUGH
# ============================================================

async def is_strikethrough(element):

    try:

        return bool(
            await element.evaluate(
                """
                el => {

                    let node = el;

                    while (node) {

                        const style =
                            window.getComputedStyle(node);

                        if (
                            style.textDecoration &&
                            style.textDecoration
                                .toLowerCase()
                                .includes("line-through")
                        ) {
                            return true;
                        }

                        node = node.parentElement;
                    }

                    return false;
                }
                """
            )
        )

    except Exception:

        return False


# ============================================================
# FIND OPTION CARDS
# ============================================================

async def find_option_cards(page):

    selectors = [

        "[data-test-id='activity-option-card']",

        "[data-test-id*='activity-option']",

        "[data-test-id*='option-card']",

        "[class*='activity-option-card']",

        "[class*='activity-option']",

        ".activity-options__item",

        "[class*='option-card']",

        "[class*='option-item']",

        "[class*='option__container']",

        "[class*='activity-option-container']",
    ]

    cards = []

    seen = set()

    for selector in selectors:

        try:

            locator = page.locator(
                selector
            )

            count = await locator.count()

            for i in range(count):

                card = locator.nth(i)

                try:

                    if not await card.is_visible():
                        continue

                    text = await card.inner_text(
                        timeout=2000
                    )

                except Exception:
                    continue

                text = " ".join(
                    text.split()
                )

                if not text:
                    continue

                fingerprint = (
                    text[:1200],
                    len(text)
                )

                if fingerprint in seen:
                    continue

                seen.add(fingerprint)

                cards.append(card)

        except Exception:
            continue

    return cards


# ============================================================
# FIND BEST OPTION
# ============================================================

async def find_best_matching_option(
    page,
    target_product
):

    cards = await find_option_cards(
        page
    )

    if not cards:

        return None, 0.0, None

    print(
        f"      [OPTION SCAN] "
        f"Found {len(cards)} visible option cards."
    )

    best_card = None
    best_score = 0.0
    best_text = None

    for card in cards:

        try:

            text = await card.inner_text()

        except Exception:
            continue

        if not text:
            continue

        score = calculate_combined_score(
            target_product,
            text
        )

        normalized_target = normalize_text(
            target_product
        )

        normalized_card = normalize_text(
            text
        )

        if normalized_target in normalized_card:

            score = 1.0

        if score > best_score:

            best_score = score
            best_card = card
            best_text = text

    return (
        best_card,
        best_score,
        best_text
    )


# ============================================================
# FORCE GBP CURRENCY
# ============================================================

async def set_gbp_currency(page):

    print(
        "      [CURRENCY] Checking currency..."
    )

    try:

        currency_input = page.locator(
            "#footer-currency-selector"
        )

        if await currency_input.count() == 0:

            print(
                "      [CURRENCY] "
                "Currency selector not found."
            )

            return False

        if not await currency_input.is_visible(
            timeout=1500
        ):

            try:

                await currency_input.scroll_into_view_if_needed()

            except Exception:
                pass

        current_value = ""

        try:

            current_value = await currency_input.input_value()
        except Exception:
            pass

        print(
            f"      [CURRENCY] Current: "
            f"{current_value}"
        )

        if (
            "pound" in current_value.lower()
            or "£" in current_value
        ):

            print(
                "      [CURRENCY] "
                "Already GBP."
            )

            return True

        await currency_input.click(
            force=True
        )

        await page.wait_for_timeout(
            700
        )

        # ----------------------------------------------------
        # Look for British Pound option
        # ----------------------------------------------------

        option_selectors = [

            "text=British Pound",

            "text=British pound",

            "text=GBP",

            "text=£",
        ]

        clicked = False

        for selector in option_selectors:

            try:

                options = page.locator(
                    selector
                )

                count = await options.count()

                for i in range(count):

                    option = options.nth(i)

                    try:

                        if not await option.is_visible(
                            timeout=300
                        ):
                            continue

                    except Exception:
                        continue

                    try:

                        option_text = (
                            await option.inner_text()
                        ).strip()

                    except Exception:

                        option_text = ""

                    if (
                        "pound" not in
                        option_text.lower()
                        and "gbp" not in
                        option_text.lower()
                        and "£" not in option_text
                    ):
                        continue

                    try:

                        await option.click(
                            force=True
                        )

                        clicked = True

                        break

                    except Exception:
                        continue

                if clicked:
                    break

            except Exception:
                continue

        if not clicked:

            print(
                "      [CURRENCY] "
                "British Pound option not found."
            )

            try:

                await page.keyboard.press(
                    "Escape"
                )

            except Exception:
                pass

            return False

        print(
            "      [CURRENCY] "
            "GBP selected."
        )

        # Currency change can reload/update the page
        await page.wait_for_timeout(
            2500
        )

        return True

    except Exception as e:

        print(
            f"      [CURRENCY] "
            f"Error: {e}"
        )

        return False


# ============================================================
# OPEN DATE PICKER
# ============================================================

async def open_date_picker(page):

    selectors = [

        "button:has-text('Select date')",

        "[aria-label*='Select date']",

        "[data-test-id*='date']",

        ".input-label-wrapper",

        "[class*='input-label']",
    ]

    for selector in selectors:

        try:

            locator = page.locator(
                selector
            ).first

            if await locator.is_visible(
                timeout=1200
            ):

                await locator.click(
                    force=True
                )

                await page.wait_for_timeout(
                    700
                )

                return True

        except Exception:
            continue

    return False


# ============================================================
# SELECT DATE
# ============================================================

async def select_date(
    page,
    date_label
):

    selectors = [

        f"[aria-label='{date_label}']",

        f"button[aria-label='{date_label}']",

        f".c-datepicker-day__container[aria-label='{date_label}']",

        f"[data-date='{date_label}']",
    ]

    for month_attempt in range(14):

        for selector in selectors:

            try:

                element = page.locator(
                    selector
                ).first

                if await element.is_visible(
                    timeout=500
                ):

                    await element.click(
                        force=True
                    )

                    await page.wait_for_timeout(
                        800
                    )

                    return True

            except Exception:
                continue

        next_selectors = [

            "button[aria-label='Next month']",

            ".c-datepicker__nav-button--next",

            "[data-test-id*='arrow-right']",

            "button:has-text('Next')",
        ]

        moved = False

        for selector in next_selectors:

            try:

                button = page.locator(
                    selector
                ).first

                if await button.is_visible(
                    timeout=500
                ):

                    await button.click(
                        force=True
                    )

                    await page.wait_for_timeout(
                        700
                    )

                    moved = True

                    break

            except Exception:
                continue

        if not moved:
            break

    return False


# ============================================================
# CHECK AVAILABILITY
# ============================================================

async def click_check_availability(page):

    selectors = [

        "button:has-text('Check availability')",

        "[data-test-id*='check-availability']",
    ]

    for selector in selectors:

        try:

            button = page.locator(
                selector
            ).first

            if await button.is_visible(
                timeout=1500
            ):

                await button.click(
                    force=True
                )

                try:

                    await page.wait_for_load_state(
                        "networkidle",
                        timeout=8000
                    )

                except Exception:
                    pass

                await page.wait_for_timeout(
                    2500
                )

                return True

        except Exception:
            continue

    return False


# ============================================================
# ACCEPT COOKIES
# ============================================================

async def accept_cookies(page):

    selectors = [

        "#onetrust-accept-btn-handler",

        "button:has-text('Accept all')",

        "button:has-text('Accept')",
    ]

    for selector in selectors:

        try:

            button = page.locator(
                selector
            ).first

            if await button.is_visible(
                timeout=1000
            ):

                await button.click(
                    force=True
                )

                await page.wait_for_timeout(
                    500
                )

                return True

        except Exception:
            continue

    return False


# ============================================================
# SCROLL / LAZY LOAD
# ============================================================

async def trigger_lazy_render(page):

    try:

        await page.evaluate(
            """
            async () => {

                window.scrollTo(
                    0,
                    document.body.scrollHeight
                );

                await new Promise(
                    r => setTimeout(r, 800)
                );

                window.scrollTo(
                    0,
                    document.body.scrollHeight / 2
                );

                await new Promise(
                    r => setTimeout(r, 400)
                );

                window.scrollTo(
                    0,
                    0
                );

                await new Promise(
                    r => setTimeout(r, 500)
                );
            }
            """
        )

    except Exception:
        pass


# ============================================================
# GET AVAILABLE TIMES
# ============================================================

async def get_available_times(
    card,
    page
):

    print(
        "      [TIME SLOTS] "
        "Reading time slots directly..."
    )

    selectors = [

        "span[data-v-ebc78423]",

        "[data-test-id*='time']",

        "[data-testid*='time']",

        "[class*='time-slot']",

        "[class*='timeslot']",

        "button[class*='time']",
    ]

    time_pattern = re.compile(
        r"\b\d{1,2}:\d{2}(?:\s*[AP]M)?\b",
        re.IGNORECASE
    )

    found = []
    seen = set()

    # --------------------------------------------------------
    # FIRST: MATCHED CARD ONLY
    # --------------------------------------------------------

    for selector in selectors:

        try:

            elements = card.locator(
                selector
            )

            count = await elements.count()

            for i in range(count):

                element = elements.nth(i)

                try:

                    if not await element.is_visible(
                        timeout=300
                    ):
                        continue

                    text = (
                        await element.inner_text()
                        or ""
                    ).strip()

                except Exception:
                    continue

                match = time_pattern.search(
                    text
                )

                if not match:
                    continue

                time_value = (
                    match.group(0).strip()
                )

                key = normalize_text(
                    time_value
                )

                if key in seen:
                    continue

                seen.add(key)

                found.append(
                    time_value
                )

        except Exception:
            continue

    # --------------------------------------------------------
    # PAGE FALLBACK
    # --------------------------------------------------------

    if not found:

        try:

            elements = page.locator(
                "span[data-v-ebc78423]"
            )

            count = await elements.count()

            print(
                f"      [TIME DOM] "
                f"Found {count} time elements."
            )

            for i in range(count):

                element = elements.nth(i)

                try:

                    if not await element.is_visible(
                        timeout=300
                    ):
                        continue

                    text = (
                        await element.inner_text()
                        or ""
                    ).strip()

                except Exception:
                    continue

                match = time_pattern.search(
                    text
                )

                if not match:
                    continue

                time_value = (
                    match.group(0).strip()
                )

                key = normalize_text(
                    time_value
                )

                if key in seen:
                    continue

                seen.add(key)

                found.append(
                    time_value
                )

        except Exception:
            pass

    # --------------------------------------------------------
    # SORT
    # --------------------------------------------------------

    def sort_key(value):

        value = value.upper().strip()

        ampm = None

        if "AM" in value:
            ampm = "AM"

        elif "PM" in value:
            ampm = "PM"

        clean = (
            value
            .replace("AM", "")
            .replace("PM", "")
            .strip()
        )

        try:

            hour, minute = map(
                int,
                clean.split(":")
            )

            if ampm == "AM" and hour == 12:
                hour = 0

            if ampm == "PM" and hour != 12:
                hour += 12

            return hour, minute

        except Exception:

            return 99, 99

    found.sort(
        key=sort_key
    )

    print(
        f"      [TIME SLOTS] "
        f"Found {len(found)} slots:"
    )

    for value in found:

        print(
            f"          {value}"
        )

    return found


# ============================================================
# CLICK SPECIFIC TIME
# ============================================================

async def click_specific_time(
    page,
    time_string
):

    print(
        f"      [CLICK] "
        f"Looking for {time_string}"
    )

    selectors = [

        "span[data-v-ebc78423]",

        "[data-test-id*='time']",

        "[data-testid*='time']",

        "[class*='time-slot']",

        "[class*='timeslot']",

        "button[class*='time']",
    ]

    target = None

    wanted = normalize_text(
        time_string
    )

    for selector in selectors:

        try:

            elements = page.locator(
                selector
            )

            count = await elements.count()

            for i in range(count):

                element = elements.nth(i)

                try:

                    if not await element.is_visible(
                        timeout=300
                    ):
                        continue

                    text = (
                        await element.inner_text()
                        or ""
                    ).strip()

                except Exception:
                    continue

                match = re.search(
                    r"\b\d{1,2}:\d{2}(?:\s*[AP]M)?\b",
                    text,
                    re.IGNORECASE
                )

                if not match:
                    continue

                actual = match.group(0).strip()

                if normalize_text(actual) != wanted:
                    continue

                target = element

                break

            if target:
                break

        except Exception:
            continue

    if not target:

        print(
            f"      [CLICK] "
            f"Could not find {time_string}"
        )

        return False

    try:

        await target.scroll_into_view_if_needed()

    except Exception:
        pass

    try:

        await target.click(
            force=True,
            timeout=5000
        )

        print(
            f"      [CLICK] "
            f"{time_string} selected."
        )

        return True

    except Exception as e:

        print(
            f"      [CLICK] "
            f"Normal click failed: {e}"
        )

        try:

            await target.evaluate(
                "el => el.click()"
            )

            print(
                f"      [CLICK] "
                f"{time_string} selected via JS."
            )

            return True

        except Exception:

            return False


# ============================================================
# GET PRICE FROM MATCHED OPTION ONLY
# ============================================================

async def get_selected_time_price(
    page,
    wanted_time,
    target_product
):

    print(
        f"      [PRICE] "
        f"Finding price for {wanted_time}..."
    )

    # Wait for Vue to update the selected time
    await page.wait_for_timeout(
        1500
    )

    # ========================================================
    # VERY IMPORTANT
    #
    # Re-find the matched option AFTER the time click.
    #
    # We do NOT use:
    #
    #     page.locator("span.text-atom--title-2")
    #
    # because that finds prices belonging to other options,
    # including the £4.95 digital-guide price.
    # ========================================================

    matched_card = None

    for attempt in range(5):

        try:

            (
                matched_card,
                score,
                card_text
            ) = await find_best_matching_option(
                page,
                target_product
            )

            if (
                matched_card
                and score >= MIN_MATCH_SCORE
            ):

                print(
                    f"      [PRICE] "
                    f"Matched option after time click. "
                    f"Score={score:.2f}"
                )

                break

        except Exception:
            pass

        await page.wait_for_timeout(
            700
        )

    if not matched_card:

        print(
            "      [PRICE] "
            "Could not re-find matched option."
        )

        return None

    # ========================================================
    # EXACT GYG PRICE CONTAINER
    #
    # HTML observed:
    #
    # <span class="activity-option-price">
    #     <span class="text-atom text-atom--title-2 ...">
    #         <span>₹12,387</span>
    #     </span>
    # </span>
    #
    # This is the correct price.
    # ========================================================

    price_selectors = [

        ".activity-option-price",

        "span.activity-option-price",

        ".activity-option-price-wrapper__price-wrapper "
        ".activity-option-price",

        "ins .activity-option-price",
    ]

    for selector in price_selectors:

        try:

            elements = matched_card.locator(
                selector
            )

            count = await elements.count()

            if count == 0:
                continue

            for i in range(count):

                element = elements.nth(i)

                try:

                    if not await element.is_visible(
                        timeout=500
                    ):
                        continue

                    text = (
                        await element.inner_text()
                        or ""
                    ).strip()

                    text = " ".join(
                        text.split()
                    )

                except Exception:
                    continue

                if not text:
                    continue

                # ------------------------------------------------
                # Ignore original/struck prices
                # ------------------------------------------------

                if await is_strikethrough(
                    element
                ):
                    continue

                value = parse_price_value(
                    text
                )

                if value is None:
                    continue

                currency = get_price_currency(
                    text
                )

                price = format_price(
                    value,
                    currency
                )

                print(
                    f"      [PRICE FOUND] "
                    f"{wanted_time}: {price}"
                )

                return price

        except Exception:
            continue

    # ========================================================
    # DIRECT TITLE-2 SEARCH INSIDE MATCHED CARD ONLY
    # ========================================================

    try:

        elements = matched_card.locator(
            "span.text-atom--title-2"
        )

        count = await elements.count()

        print(
            f"      [PRICE] "
            f"Matched card has "
            f"{count} title-2 elements."
        )

        for i in range(count):

            element = elements.nth(i)

            try:

                if not await element.is_visible(
                    timeout=500
                ):
                    continue

                text = (
                    await element.inner_text()
                    or ""
                ).strip()

            except Exception:
                continue

            value = parse_price_value(
                text
            )

            if value is None:
                continue

            if await is_strikethrough(
                element
            ):
                continue

            currency = get_price_currency(
                text
            )

            price = format_price(
                value,
                currency
            )

            print(
                f"      [PRICE FOUND] "
                f"{wanted_time}: {price}"
            )

            return price

    except Exception:
        pass

    # ========================================================
    # JAVASCRIPT DIRECT SEARCH
    # INSIDE MATCHED OPTION ONLY
    # ========================================================

    try:

        result = await matched_card.evaluate(
            """
            card => {

                const selectors = [
                    ".activity-option-price",
                    "span.text-atom--title-2"
                ];

                const regex =
                    /(?:£|₹|€|\\$|GBP|INR|EUR|USD)\\s*\\d[\\d,]*(?:\\.\\d{1,2})?/i;

                for (const selector of selectors) {

                    const elements =
                        card.querySelectorAll(selector);

                    for (const element of elements) {

                        const style =
                            window.getComputedStyle(element);

                        if (
                            style.display === "none" ||
                            style.visibility === "hidden"
                        ) {
                            continue;
                        }

                        let text =
                            element.innerText ||
                            element.textContent ||
                            "";

                        text =
                            text
                            .replace(/\\s+/g, " ")
                            .trim();

                        if (!text) {
                            continue;
                        }

                        const match =
                            text.match(regex);

                        if (!match) {
                            continue;
                        }

                        /*
                         * Do not return struck-through price.
                         */

                        let node = element;
                        let struck = false;

                        while (node) {

                            const nodeStyle =
                                window.getComputedStyle(node);

                            if (
                                nodeStyle.textDecoration &&
                                nodeStyle.textDecoration
                                    .toLowerCase()
                                    .includes("line-through")
                            ) {
                                struck = true;
                                break;
                            }

                            node =
                                node.parentElement;
                        }

                        if (struck) {
                            continue;
                        }

                        return match[0];
                    }
                }

                return null;
            }
            """
        )

        if result:

            value = parse_price_value(
                result
            )

            if value is not None:

                currency = get_price_currency(
                    result
                )

                price = format_price(
                    value,
                    currency
                )

                print(
                    f"      [PRICE FOUND] "
                    f"{wanted_time}: {price}"
                )

                return price

    except Exception as e:

        print(
            f"      [PRICE] "
            f"JavaScript search failed: {e}"
        )

    print(
        f"      [WARNING] "
        f"No price found inside matched option "
        f"for {wanted_time}"
    )

    return None


# ============================================================
# EXTRACT ALL TIME SLOT PRICES
# ============================================================

async def extract_all_timeslot_prices(
    page,
    card,
    target_product
):

    times = await get_available_times(
        card,
        page
    )

    if not times:

        price = await get_selected_time_price(
            page,
            "",
            target_product
        )

        if price:

            return (
                f"Standard Rate: {price}"
            )

        return None

    results = []

    seen = set()

    # ========================================================
    # EVERY TIME SLOT
    # ========================================================

    for time_string in times:

        print(
            "\n"
            f"      -> Processing {time_string}"
        )

        # ----------------------------------------------------
        # Click
        # ----------------------------------------------------

        clicked = await click_specific_time(
            page,
            time_string
        )

        if not clicked:

            print(
                f"      [WARNING] "
                f"Could not select {time_string}"
            )

            continue

        # ----------------------------------------------------
        # Price
        # ----------------------------------------------------

        price = await get_selected_time_price(
            page,
            time_string,
            target_product
        )

        if price:

            result = (
                f"{time_string}: {price}"
            )

            if result not in seen:

                seen.add(result)

                results.append(
                    result
                )

                print(
                    f"      [RESULT] {result}"
                )

        else:

            print(
                f"      [WARNING] "
                f"No price found for "
                f"{time_string}"
            )

            await save_price_debug_html(
                page,
                time_string
            )

    # ========================================================
    # RETURN
    # ========================================================

    if results:

        return " | ".join(
            results
        )

    return None


# ============================================================
# SHOW ALL OPTIONS
# ============================================================

async def click_show_all_options(page):

    selectors = [

        "button:has-text('Show all options')",

        "span:has-text('Show all options')",

        "text=Show all options",
    ]

    clicked = False

    for selector in selectors:

        try:

            elements = page.locator(
                selector
            )

            count = await elements.count()

            for i in range(count):

                element = elements.nth(i)

                try:

                    if not await element.is_visible(
                        timeout=400
                    ):
                        continue

                except Exception:
                    continue

                target = element

                try:

                    tag = await element.evaluate(
                        "el => el.tagName.toLowerCase()"
                    )

                    if tag == "span":

                        parent = element.locator(
                            "xpath=ancestor::button[1]"
                        )

                        if await parent.count():

                            target = parent.first

                except Exception:
                    pass

                try:

                    await target.click(
                        force=True
                    )

                    clicked = True

                    await page.wait_for_timeout(
                        1000
                    )

                except Exception:
                    continue

        except Exception:
            continue

    return clicked


# ============================================================
# EXPAND OPTION SECTIONS
# ============================================================

async def expand_all_option_sections(page):

    await click_show_all_options(
        page
    )

    selectors = [

        "button:has-text('Show more')",

        "button:has-text('Show details')",

        "button:has-text('See more')",

        "button:has-text('More options')",

        "[class*='show-more']",
    ]

    for selector in selectors:

        try:

            elements = page.locator(
                selector
            )

            count = await elements.count()

            for i in range(count):

                try:

                    element = elements.nth(i)

                    if not await element.is_visible():
                        continue

                    if not await element.is_enabled():
                        continue

                    await element.click(
                        force=True
                    )

                    await page.wait_for_timeout(
                        300
                    )

                except Exception:
                    continue

        except Exception:
            continue


# ============================================================
# SOLD OUT
# ============================================================

def is_sold_out(text):

    if not text:
        return False

    text = normalize_text(
        text
    )

    phrases = [
        "sold out",
        "unavailable",
        "fully booked",
        "no options available",
        "not available",
    ]

    return any(
        x in text
        for x in phrases
    )


# ============================================================
# DEBUG HTML
# ============================================================

async def save_price_debug_html(
    page,
    time_string
):

    try:

        os.makedirs(
            DEBUG_FOLDER,
            exist_ok=True
        )

        safe = re.sub(
            r"[^A-Za-z0-9_-]",
            "_",
            time_string
        )

        path = os.path.join(
            DEBUG_FOLDER,
            f"after_time_{safe}.html"
        )

        html = await page.content()

        with open(
            path,
            "w",
            encoding="utf-8"
        ) as f:

            f.write(html)

        print(
            f"      [DEBUG] "
            f"Saved: {path}"
        )

    except Exception:
        pass


# ============================================================
# SCRAPE ONE DATE
# ============================================================

async def scrape_one_date(
    page,
    url,
    date_info,
    target_product
):

    print(
        f"    -> Date: "
        f"{date_info['gyg_label']} | "
        f"Target: {target_product}"
    )

    try:

        # ====================================================
        # OPEN PAGE
        # ====================================================

        await page.goto(
            url,
            wait_until="domcontentloaded",
            timeout=PAGE_TIMEOUT
        )

        await page.wait_for_timeout(
            3000
        )

        await accept_cookies(
            page
        )

        # ====================================================
        # FORCE GBP
        # ====================================================

        await set_gbp_currency(
            page
        )

        # Currency selection may update the DOM
        await page.wait_for_timeout(
            1500
        )

        # ====================================================
        # DATE PICKER
        # ====================================================

        if not await open_date_picker(
            page
        ):

            print(
                "      [WARNING] "
                "Date picker not found."
            )

            return "Date Picker Not Found"

        # ====================================================
        # SELECT DATE
        # ====================================================

        if not await select_date(
            page,
            date_info["gyg_label"]
        ):

            print(
                "      [WARNING] "
                "Date unavailable."
            )

            return "Sold Out / Date Unavailable"

        # ====================================================
        # CHECK AVAILABILITY
        # ====================================================

        await click_check_availability(
            page
        )

        await page.wait_for_timeout(
            2500
        )

        # ====================================================
        # RENDER OPTIONS
        # ====================================================

        best_card = None
        best_score = 0.0
        best_text = None

        for attempt in range(5):

            print(
                f"      [RENDER ATTEMPT "
                f"{attempt + 1}/5]"
            )

            await trigger_lazy_render(
                page
            )

            await expand_all_option_sections(
                page
            )

            await page.wait_for_timeout(
                1000
            )

            (
                best_card,
                best_score,
                best_text
            ) = await find_best_matching_option(
                page,
                target_product
            )

            if (
                best_card
                and best_score >= MIN_MATCH_SCORE
            ):

                break

            await page.wait_for_timeout(
                1200
            )

        # ====================================================
        # MATCH CHECK
        # ====================================================

        if not best_card:

            print(
                "      [WARNING] "
                "No option cards found."
            )

            return "Option Text Not Found"

        print(
            f"      [BEST MATCH] "
            f"Score = {best_score:.2f}"
        )

        if best_score < MIN_MATCH_SCORE:

            return "Option Text Not Found"

        if is_sold_out(
            best_text
        ):

            return "Option Sold Out"

        # ====================================================
        # EXTRACT TIMES + PRICES
        # ====================================================

        result = await extract_all_timeslot_prices(
            page,
            best_card,
            target_product
        )

        if result:

            print(
                f"      [FINAL PRICE] "
                f"{result}"
            )

            return result

        return "Price Not Found"

    except Exception as e:

        print(
            f"      [EXCEPTION] "
            f"{type(e).__name__}: {e}"
        )

        try:

            os.makedirs(
                DEBUG_FOLDER,
                exist_ok=True
            )

            path = os.path.join(
                DEBUG_FOLDER,
                "exception_page.html"
            )

            html = await page.content()

            with open(
                path,
                "w",
                encoding="utf-8"
            ) as f:

                f.write(html)

        except Exception:
            pass

        return "Extraction Error"


# ============================================================
# CREATE PAGE
# ============================================================

async def create_page(context):

    page = await context.new_page()

    page.set_default_timeout(
        12000
    )

    return page

# ============================================================
# MAIN
# ============================================================

async def main():

    # ========================================================
    # INPUT CHECK
    # ========================================================

    if not os.path.exists(INPUT_EXCEL):
        print(f"ERROR: {INPUT_EXCEL} not found.")
        return

    # ========================================================
    # READ EXCEL
    # ========================================================

    df = pd.read_excel(INPUT_EXCEL)

    required = [
        "Competitor",
        "Product",
        "Target_Product",
        "URL"
    ]

    missing = [
        x for x in required
        if x not in df.columns
    ]

    if missing:
        print("ERROR: Missing columns:")
        print(missing)
        return

    # ========================================================
    # DATES
    # ========================================================

    dates = generate_date_labels(
        START_DATE_STR,
        END_DATE_STR
    )

    # --------------------------------------------------------
    # OUTPUT FORMAT
    # One row per Product + Date
    # --------------------------------------------------------

    output_rows = []

    for _, row in df.iterrows():
        for date_info in dates:
            output_rows.append({
                "Competitor": row["Competitor"],
                "Product": row["Product"],
                "Target_Product": row["Target_Product"],
                "URL": row["URL"],
                "Date": date_info["iso"],
                "Price": None
            })

    results_df = pd.DataFrame(output_rows)

    os.makedirs(
        DEBUG_FOLDER,
        exist_ok=True
    )

    # Save the blank structure immediately so the output file exists.
    results_df.to_excel(
        OUTPUT_EXCEL,
        index=False
    )

    # ========================================================
    # PLAYWRIGHT
    # ========================================================

    async with async_playwright() as p:

        browser = await p.chromium.launch(
            headless=HEADLESS
        )

        context = await browser.new_context(
            viewport={
                "width": 1366,
                "height": 950
            },
            locale="en-GB",
            timezone_id="Europe/London",
            user_agent=(
                "Mozilla/5.0 "
                "(Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 "
                "(KHTML, like Gecko) "
                "Chrome/131.0.0.0 "
                "Safari/537.36"
            )
        )

        page = await create_page(context)

        # ====================================================
        # PRODUCTS
        # ====================================================

        for product_index, row in df.iterrows():

            target_url = str(row["URL"]).strip()
            target_product = str(row["Target_Product"]).strip()

            print("\n" + "=" * 90)
            print(
                f"[{product_index + 1}/{len(df)}] "
                f"{target_product}"
            )
            print("=" * 90)

            # ------------------------------------------------
            # URL MISSING
            # ------------------------------------------------

            if not target_url or target_url.lower() == "nan":

                for date_info in dates:
                    mask = (
                        (results_df["Product"] == row["Product"]) &
                        (results_df["Target_Product"] == row["Target_Product"]) &
                        (results_df["URL"] == row["URL"]) &
                        (results_df["Date"] == date_info["iso"])
                    )
                    results_df.loc[mask, "Price"] = "URL Missing"

                results_df.to_excel(
                    OUTPUT_EXCEL,
                    index=False
                )
                continue

            # ------------------------------------------------
            # DATES
            # ------------------------------------------------

            for date_info in dates:

                print(
                    "\n"
                    f"      DATE ROW: {date_info['iso']}"
                )

                # ---------------------------------------------
                # PAGE RECOVERY
                # ---------------------------------------------

                try:
                    if page.is_closed():
                        page = await create_page(context)
                except Exception:
                    page = await create_page(context)

                # ---------------------------------------------
                # SCRAPE
                # ---------------------------------------------

                result = await scrape_one_date(
                    page,
                    target_url,
                    date_info,
                    target_product
                )

                # ---------------------------------------------
                # UPDATE CORRECT OUTPUT ROW
                # ---------------------------------------------

                mask = (
                    (results_df["Product"] == row["Product"]) &
                    (results_df["Target_Product"] == row["Target_Product"]) &
                    (results_df["URL"] == row["URL"]) &
                    (results_df["Date"] == date_info["iso"])
                )

                results_df.loc[mask, "Price"] = result

                # ---------------------------------------------
                # SAVE AFTER EVERY DATE
                # ---------------------------------------------

                try:
                    results_df.to_excel(
                        OUTPUT_EXCEL,
                        index=False
                    )

                    print(
                        f"      [SAVED] {OUTPUT_EXCEL}"
                    )

                except Exception as e:
                    print(
                        f"      [SAVE WARNING] {e}"
                    )

        # ====================================================
        # CLOSE
        # ====================================================

        try:
            await context.close()
        except Exception:
            pass

        try:
            await browser.close()
        except Exception:
            pass

    # ========================================================
    # FINAL SAVE
    # ========================================================

    results_df.to_excel(
        OUTPUT_EXCEL,
        index=False
    )

    print("\n" + "=" * 90)
    print("EXECUTION FINISHED")
    print(f"Output: {OUTPUT_EXCEL}")
    print("=" * 90)


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":
    asyncio.run(main())
