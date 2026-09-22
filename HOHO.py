import asyncio
from datetime import datetime, timedelta
import os
import re
import pandas as pd
from playwright.async_api import async_playwright
from zoneinfo import ZoneInfo

# ============================================================
# CONFIGURATION
# ============================================================

INPUT_EXCEL = "HOHO_competitor_input.xlsx"
OUTPUT_EXCEL = "HOHO_competitor_Prices.xlsx"

LONDON_TZ = ZoneInfo("Europe/London")

START_DATE_STR = os.getenv(
    "HOHO_START_DATE",
    datetime.now(LONDON_TZ).strftime("%Y-%m-%d")
)

END_DATE_STR = os.getenv(
    "HOHO_END_DATE",
    "2026-10-31"
)

# Minimum score required to accept a product option
MIN_MATCH_SCORE = 0.70

# ============================================================


def generate_date_labels(start_str, end_str):
    """
    Generate dates in the exact format normally used by GYG.
    """

    labels = []

    start_dt = datetime.strptime(
        start_str,
        "%Y-%m-%d"
    )

    end_dt = datetime.strptime(
        end_str,
        "%Y-%m-%d"
    )

    current_dt = start_dt

    while current_dt <= end_dt:

        day_str = current_dt.strftime(
            "%d"
        ).lstrip("0")

        readable_label = current_dt.strftime(
            f"%A, {day_str} %B %Y"
        )

        labels.append({
            "iso": current_dt.strftime("%Y-%m-%d"),
            "gyg_label": readable_label
        })

        current_dt += timedelta(days=1)

    return labels


# ============================================================
# TEXT NORMALISATION
# ============================================================

def normalize_text(text):

    if not text:
        return ""

    text = text.lower()

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

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


def clean_tokens(text):

    text = normalize_text(text)

    words = re.findall(
        r"[a-z0-9]+",
        text
    )

    ignore_set = {
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
        word
        for word in words
        if word not in ignore_set
        and len(word) > 1
    }


def calculate_token_match_score(
    target_str,
    card_str
):

    target_tokens = clean_tokens(
        target_str
    )

    card_tokens = clean_tokens(
        card_str
    )

    if not target_tokens:
        return 0.0

    matched_tokens = (
        target_tokens.intersection(
            card_tokens
        )
    )

    return (
        len(matched_tokens)
        / len(target_tokens)
    )


def calculate_exact_phrase_score(
    target_str,
    card_str
):

    target = normalize_text(
        target_str
    )

    card = normalize_text(
        card_str
    )

    if not target:
        return 0.0

    if target in card:
        return 1.0

    target_words = target.split()

    if len(target_words) >= 2:

        compact_target = " ".join(
            target_words
        )

        if compact_target in card:
            return 0.9

    return 0.0


def calculate_combined_score(
    target_str,
    card_str
):

    token_score = (
        calculate_token_match_score(
            target_str,
            card_str
        )
    )

    phrase_score = (
        calculate_exact_phrase_score(
            target_str,
            card_str
        )
    )

    if phrase_score >= 0.9:
        return 1.0

    return (
        token_score * 0.75
        +
        phrase_score * 0.25
    )


# ============================================================
# PRICE EXTRACTION
# ============================================================

PRICE_REGEX = re.compile(
    r"(?:£|GBP\s*)\s*\d[\d,]*(?:\.\d{1,2})?",
    re.IGNORECASE
)


def parse_price_value(price_text):

    if not price_text:
        return None

    match = PRICE_REGEX.search(
        price_text
    )

    if not match:
        return None

    value = re.sub(
        r"[^0-9.]",
        "",
        match.group(0)
    )

    try:
        return float(value)

    except Exception:
        return None


def format_price(value):

    if value is None:
        return None

    return f"£{value:,.2f}"


def extract_price_candidates_from_text(
    text
):

    if not text:
        return []

    results = []

    for match in PRICE_REGEX.finditer(
        text
    ):

        raw = match.group(0)

        value = parse_price_value(
            raw
        )

        if value is not None:

            results.append({
                "raw": raw.strip(),
                "value": value,
                "position": match.start()
            })

    return results


# ============================================================
# PRICE FROM MATCHED CARD
# ============================================================

async def extract_best_price_from_card(
    card
):

    price_selectors = [

        "[data-test-id='activity-option-price']",

        "[data-test-id*='price']",

        "[class*='price-actual']",

        "[class*='price-amount']",

        "[class*='activity-option-price']",

        "[class*='option-price']",

        "[class*='price-wrapper']",

        "[class*='price']",

        "strong",

        "span",

        "div",
    ]

    candidates = []

    seen_texts = set()

    for selector in price_selectors:

        try:

            elements = card.locator(
                selector
            )

            count = await elements.count()

            for i in range(count):

                element = elements.nth(i)

                try:

                    if not await element.is_visible():
                        continue

                    text = await element.inner_text(
                        timeout=1000
                    )

                except Exception:
                    continue

                if not text:
                    continue

                text = " ".join(
                    text.split()
                )

                if text in seen_texts:
                    continue

                seen_texts.add(
                    text
                )

                prices = (
                    extract_price_candidates_from_text(
                        text
                    )
                )

                for price in prices:

                    if price["value"] <= 0:
                        continue

                    # ----------------------------------------
                    # Ignore crossed-out prices
                    # ----------------------------------------

                    try:

                        decoration = (
                            await element.evaluate(
                                """
                                el => {
                                    let node = el;

                                    while (node) {
                                        const style =
                                            window.getComputedStyle(node);

                                        if (
                                            style.textDecoration
                                            .includes('line-through')
                                        ) {
                                            return 'strikethrough';
                                        }

                                        node =
                                            node.parentElement;
                                    }

                                    return '';
                                }
                                """
                            )
                        )

                        if (
                            decoration
                            == "strikethrough"
                        ):
                            continue

                    except Exception:
                        pass

                    candidates.append({

                        "value":
                            price["value"],

                        "raw":
                            price["raw"],

                        "selector":
                            selector,

                        "text":
                            text,
                    })

        except Exception:
            continue

    # --------------------------------------------------------
    # Remove duplicates
    # --------------------------------------------------------

    unique = {}

    for candidate in candidates:

        key = (
            candidate["value"],
            candidate["raw"]
        )

        if key not in unique:
            unique[key] = candidate

    candidates = list(
        unique.values()
    )

    if candidates:

        # IMPORTANT:
        # Do not use the LAST price found.
        candidates.sort(
            key=lambda x: x["value"]
        )

        return format_price(
            candidates[0]["value"]
        )

    # --------------------------------------------------------
    # Final card-text fallback
    # --------------------------------------------------------

    try:

        card_text = await card.inner_text()

        prices = (
            extract_price_candidates_from_text(
                card_text
            )
        )

        if prices:

            prices.sort(
                key=lambda x: x["value"]
            )

            return format_price(
                prices[0]["value"]
            )

    except Exception:
        pass

    return None


# ============================================================
# FIND OPTION CARDS
# ============================================================

async def find_option_cards(
    page
):

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
                        timeout=1500
                    )

                except Exception:
                    continue

                if not text:
                    continue

                fingerprint = (
                    text[:800],
                    len(text)
                )

                if fingerprint in seen:
                    continue

                seen.add(
                    fingerprint
                )

                cards.append(
                    card
                )

        except Exception:
            continue

    return cards


# ============================================================
# NEW: CLICK "SHOW ALL OPTIONS"
# ============================================================

async def click_show_all_options(
    page
):
    """
    GYG sometimes displays only a limited number of
    option cards initially.

    The screenshot shows:

        <button>
            ...
            <span>Show all options</span>
        </button>

    Therefore we specifically find the text and click
    the parent button.

    We repeat the process because some pages can have
    more than one expandable option section.
    """

    clicked_any = False

    # --------------------------------------------------------
    # Exact text locators
    # --------------------------------------------------------

    selectors = [

        # Span shown in user's screenshot
        "span:has-text('Show all options')",

        # Button itself
        "button:has-text('Show all options')",

        # Text locator fallback
        "text=Show all options",

        # Case-insensitive-ish CSS fallback
        "[class*='show-all']",
    ]

    for selector in selectors:

        try:

            locator = page.locator(
                selector
            )

            count = await locator.count()

            if count == 0:
                continue

            for i in range(count):

                element = locator.nth(i)

                try:

                    if not await element.is_visible(
                        timeout=500
                    ):
                        continue

                except Exception:
                    continue

                # ------------------------------------------------
                # If this is the SPAN, find the parent button.
                # This matches the DOM shown in your screenshot.
                # ------------------------------------------------

                try:

                    tag_name = await element.evaluate(
                        "el => el.tagName.toLowerCase()"
                    )

                except Exception:
                    tag_name = ""

                target = element

                if tag_name == "span":

                    try:

                        parent_button = (
                            element.locator(
                                "xpath=ancestor::button[1]"
                            )
                        )

                        if await parent_button.count() > 0:

                            target = (
                                parent_button.first
                            )

                    except Exception:
                        pass

                # ------------------------------------------------
                # Make sure button is visible/enabled
                # ------------------------------------------------

                try:

                    if not await target.is_visible():
                        continue

                    if not await target.is_enabled():
                        continue

                except Exception:
                    continue

                # ------------------------------------------------
                # Get text BEFORE click
                # ------------------------------------------------

                try:

                    before_text = (
                        await target.inner_text()
                    )

                except Exception:
                    before_text = ""

                print(
                    f"      [SHOW ALL OPTIONS] "
                    f"Clicking: {before_text.strip()}"
                )

                # ------------------------------------------------
                # Click
                # ------------------------------------------------

                try:

                    await target.scroll_into_view_if_needed()

                    await target.click(
                        force=True
                    )

                    clicked_any = True

                    # Give GYG time to render cards.
                    await page.wait_for_timeout(
                        1000
                    )

                except Exception as e:

                    print(
                        f"      [WARNING] "
                        f"Could not click Show all options: {e}"
                    )

        except Exception:
            continue

    return clicked_any


# ============================================================
# EXPAND OTHER OPTION SECTIONS
# ============================================================

async def expand_all_option_sections(
    page
):

    # --------------------------------------------------------
    # FIRST: Show ALL options
    # --------------------------------------------------------

    await click_show_all_options(
        page
    )

    # --------------------------------------------------------
    # Other expandable controls
    # --------------------------------------------------------

    selectors = [

        "button[class*='accordion']",

        "[class*='show-more']",

        "button:has-text('Show more')",

        "button:has-text('Show details')",

        "button:has-text('See more')",

        "button:has-text('More options')",

        "[class*='option-heading']",
    ]

    for selector in selectors:

        try:

            elements = page.locator(
                selector
            )

            count = await elements.count()

            for i in range(count):

                try:

                    button = elements.nth(i)

                    if (
                        await button.is_visible()
                        and
                        await button.is_enabled()
                    ):

                        # Avoid clicking Show All again here.
                        try:

                            button_text = (
                                await button.inner_text()
                            )

                            if (
                                "show all options"
                                in normalize_text(
                                    button_text
                                )
                            ):
                                continue

                        except Exception:
                            pass

                        await button.click(
                            force=True
                        )

                        await page.wait_for_timeout(
                            200
                        )

                except Exception:
                    continue

        except Exception:
            continue


# ============================================================
# SOLD OUT CHECK
# ============================================================

def is_sold_out(
    text
):

    if not text:
        return False

    text = normalize_text(
        text
    )

    phrases = [

        "sold out",

        "unavailable",

        "no options available",

        "not available",

        "fully booked",
    ]

    return any(
        phrase in text
        for phrase in phrases
    )


# ============================================================
# FIND BEST MATCH
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

    for index, card in enumerate(cards):

        try:

            card_text = await card.inner_text()

        except Exception:
            continue

        if not card_text:
            continue

        score = calculate_combined_score(
            target_product,
            card_text
        )

        normalized_target = normalize_text(
            target_product
        )

        normalized_card = normalize_text(
            card_text
        )

        if normalized_target in normalized_card:

            score = max(
                score,
                1.0
            )

        if score > best_score:

            best_score = score

            best_card = card

            best_text = card_text

    return (
        best_card,
        best_score,
        best_text
    )


# ============================================================
# OPEN DATE PICKER
# ============================================================

async def open_date_picker(
    page
):

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
                timeout=1000
            ):

                await locator.click(
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
# SELECT DATE
# ============================================================

async def select_date(
    page,
    date_label
):

    date_selectors = [

        (
            ".c-datepicker-day__container"
            f"[aria-label='{date_label}']"
        ),

        f"[aria-label='{date_label}']",

        f"button[aria-label='{date_label}']",

        f"[data-date='{date_label}']",
    ]

    max_months_to_scroll = 14

    for month_attempt in range(
        max_months_to_scroll
    ):

        # ----------------------------------------------------
        # Try exact date
        # ----------------------------------------------------

        for selector in date_selectors:

            try:

                date_element = (
                    page.locator(
                        selector
                    ).first
                )

                if await date_element.is_visible(
                    timeout=500
                ):

                    await date_element.click(
                        force=True
                    )

                    await page.wait_for_timeout(
                        600
                    )

                    return True

            except Exception:
                continue

        # ----------------------------------------------------
        # Next month
        # ----------------------------------------------------

        next_selectors = [

            "button[aria-label='Next month']",

            ".c-datepicker__nav-button--next",

            "[data-test-id*='arrow-right']",

            "button:has-text('Next')",
        ]

        moved = False

        for selector in next_selectors:

            try:

                next_button = (
                    page.locator(
                        selector
                    ).first
                )

                if await next_button.is_visible(
                    timeout=500
                ):

                    await next_button.click(
                        force=True
                    )

                    await page.wait_for_timeout(
                        500
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

async def click_check_availability(
    page
):

    selectors = [

        "button:has-text('Check availability')",

        "[data-test-id*='check-availability']",
    ]

    for selector in selectors:

        try:

            button = (
                page.locator(
                    selector
                ).first
            )

            if await button.is_visible(
                timeout=1500
            ):

                await button.click(
                    force=True
                )

                try:

                    await page.wait_for_load_state(
                        "networkidle",
                        timeout=5000
                    )

                except Exception:
                    pass

                await page.wait_for_timeout(
                    1500
                )

                return True

        except Exception:
            continue

    return False


# ============================================================
# MAIN SCRAPER FOR ONE DATE
# ============================================================

async def scrape_gyg_target_option(
    page,
    target_url,
    date_info,
    target_product
):

    date_label = date_info[
        "gyg_label"
    ]

    print(
        f"    -> Date: {date_label} | "
        f"Target: {target_product}"
    )

    try:

        # ----------------------------------------------------
        # LOAD PAGE
        # ----------------------------------------------------

        await page.goto(
            target_url,
            wait_until="domcontentloaded",
            timeout=40000
        )

        await page.wait_for_timeout(
            1800
        )

        # ----------------------------------------------------
        # COOKIE
        # ----------------------------------------------------

        cookie_selectors = [

            "#onetrust-accept-btn-handler",

            "button:has-text('Accept')",

            "button:has-text('Accept all')",
        ]

        for selector in cookie_selectors:

            try:

                button = (
                    page.locator(
                        selector
                    ).first
                )

                if await button.is_visible(
                    timeout=1000
                ):

                    await button.click(
                        force=True
                    )

                    await page.wait_for_timeout(
                        300
                    )

                    break

            except Exception:
                continue

        # ----------------------------------------------------
        # DATE PICKER
        # ----------------------------------------------------

        opened = await open_date_picker(
            page
        )

        if not opened:

            print(
                "      [WARNING] "
                "Date picker not found."
            )

            return "Date Picker Not Found"

        # ----------------------------------------------------
        # SELECT DATE
        # ----------------------------------------------------

        selected = await select_date(
            page,
            date_label
        )

        if not selected:

            print(
                "      [WARNING] "
                "Date unavailable."
            )

            return "Sold Out / Date Unavailable"

        # ----------------------------------------------------
        # CHECK AVAILABILITY
        # ----------------------------------------------------

        await click_check_availability(
            page
        )

        await page.wait_for_timeout(
            1800
        )

        # ----------------------------------------------------
        # RENDERING / OPTION EXPANSION
        # ----------------------------------------------------

        for attempt in range(4):

            print(
                f"      [RENDER ATTEMPT "
                f"{attempt + 1}/4]"
            )

            # -----------------------------------------------
            # VERY IMPORTANT:
            # Click Show all options BEFORE scanning cards.
            # -----------------------------------------------

            await expand_all_option_sections(
                page
            )

            # Extra wait after Show all options.
            await page.wait_for_timeout(
                1000
            )

            # -----------------------------------------------
            # Scan cards AFTER expansion.
            # -----------------------------------------------

            (
                best_card,
                best_score,
                best_card_text
            ) = await find_best_matching_option(
                page,
                target_product
            )

            if best_card:

                print(
                    f"      [BEST MATCH] "
                    f"Score = {best_score:.2f}"
                )

                # -------------------------------------------
                # MATCH THRESHOLD
                # -------------------------------------------

                if (
                    best_score
                    >= MIN_MATCH_SCORE
                ):

                    if is_sold_out(
                        best_card_text
                    ):

                        return "Option Sold Out"

                    # ---------------------------------------
                    # PRICE
                    # ---------------------------------------

                    price = (
                        await
                        extract_best_price_from_card(
                            best_card
                        )
                    )

                    if price:

                        print(
                            f"      [PRICE FOUND] "
                            f"{price}"
                        )

                        return price

                    print(
                        "      [WARNING] "
                        "Matched option found but "
                        "no reliable price found."
                    )

                else:

                    print(
                        f"      [WARNING] "
                        f"Match score {best_score:.2f} "
                        f"is below required "
                        f"{MIN_MATCH_SCORE:.2f}"
                    )

            else:

                print(
                    "      [WARNING] "
                    "No option cards found."
                )

            # ------------------------------------------------
            # RETRY
            # ------------------------------------------------

            if attempt < 3:

                await page.wait_for_timeout(
                    1500
                )

        # ----------------------------------------------------
        # NOT FOUND
        # ----------------------------------------------------

        print(
            f"      [NOTICE] "
            f"Target option '{target_product}' "
            f"not found in active listings."
        )

        return "Option Text Not Found"

    except Exception as e:

        print(
            f"      [EXCEPTION] "
            f"{type(e).__name__}: {e}"
        )

        return "Extraction Error"


# ============================================================
# MAIN
# ============================================================

async def main():

    # --------------------------------------------------------
    # CHECK INPUT
    # --------------------------------------------------------

    if not os.path.exists(
        INPUT_EXCEL
    ):

        print(
            f"ERROR: Could not locate "
            f"'{INPUT_EXCEL}'"
        )

        return

    # --------------------------------------------------------
    # READ EXCEL
    # --------------------------------------------------------

    input_df = pd.read_excel(
        INPUT_EXCEL
    )

    required_cols = [

        "Competitor",

        "Product",

        "Target_Product",

        "URL",
    ]

    missing_cols = [

        col
        for col in required_cols
        if col not in input_df.columns
    ]

    if missing_cols:

        print(
            "ERROR: Excel is missing columns:"
        )

        print(
            missing_cols
        )

        return

    # --------------------------------------------------------
    # DATES
    # --------------------------------------------------------

    target_dates_pool = (
        generate_date_labels(
            START_DATE_STR,
            END_DATE_STR
        )
    )

    print()
    print(
        "======================================================"
    )

    print(
        "STARTING GYG PRICE SCRAPER"
    )

    print(
        f"Input: {INPUT_EXCEL}"
    )

    print(
        f"Output: {OUTPUT_EXCEL}"
    )

    print(
        f"Dates: "
        f"{START_DATE_STR} -> "
        f"{END_DATE_STR}"
    )

    print(
        f"Products: {len(input_df)}"
    )

    print(
        "======================================================"
    )

    # --------------------------------------------------------
    # PLAYWRIGHT
    # --------------------------------------------------------

    async with async_playwright() as p:

        browser = await p.chromium.launch(
    headless=True
)

        context = await browser.new_context(

            locale="en-GB",

            viewport={
                "width": 1366,
                "height": 768
            },

            user_agent=(
                "Mozilla/5.0 "
                "(Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 "
                "(KHTML, like Gecko) "
                "Chrome/124.0.0.0 "
                "Safari/537.36"
            ),

            timezone_id="Europe/London"
        )

        scraped_data_list = []

        fetch_date_today = (
            datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )
        )

        # ----------------------------------------------------
        # PRODUCT LOOP
        # ----------------------------------------------------

        for idx, row in input_df.iterrows():

            competitor = str(
                row["Competitor"]
            ).strip()

            product_group = str(
                row["Product"]
            ).strip()

            target_product = str(
                row["Target_Product"]
            ).strip()

            url = str(
                row["URL"]
            ).strip()

            print()
            print(
                "--------------------------------------------------"
            )

            print(
                f"ROW {idx + 1}/"
                f"{len(input_df)}"
            )

            print(
                f"Competitor: {competitor}"
            )

            print(
                f"Target: {target_product}"
            )

            # ------------------------------------------------
            # CLEAN URL
            # ------------------------------------------------

            clean_url = url.split(
                "?"
            )[0]

            target_url = (
                f"{clean_url}?currency=GBP"
            )

            page = await context.new_page()

            # ------------------------------------------------
            # DATE LOOP
            # ------------------------------------------------

            for date_info in (
                target_dates_pool
            ):

                price_result = (
                    await scrape_gyg_target_option(
                        page,
                        target_url,
                        date_info,
                        target_product
                    )
                )

                scraped_data_list.append({

                    "Fetch Date":
                        fetch_date_today,

                    "Competitor":
                        competitor,

                    "Product Group":
                        product_group,

                    "Target Product Match":
                        target_product,

                    "Travel Date":
                        date_info[
                            "gyg_label"
                        ],

                    "ISO Date":
                        date_info[
                            "iso"
                        ],

                    "Price Extraction":
                        price_result,

                    "URL":
                        url
                })

                await asyncio.sleep(
                    0.5
                )

            await page.close()

            await asyncio.sleep(
                0.8
            )

        # ----------------------------------------------------
        # CLOSE BROWSER
        # ----------------------------------------------------

        await browser.close()

        # ----------------------------------------------------
        # SAVE EXCEL
        # ----------------------------------------------------

        df_out = pd.DataFrame(
            scraped_data_list
        )

        df_out.to_excel(
            OUTPUT_EXCEL,
            index=False
        )

        print()
        print(
            "======================================================"
        )

        print(
            "[SUCCESS] Pricing data saved to:"
        )

        print(
            OUTPUT_EXCEL
        )

        print(
            "======================================================"
        )


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    asyncio.run(
        main()
    )
