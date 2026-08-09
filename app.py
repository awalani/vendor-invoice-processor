import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import difflib

from PIL import Image, ImageOps, ImageSequence
import pytesseract
import pymupdf


# ============================================================
# PAGE
# ============================================================

st.set_page_config(
    page_title="Vendor Invoice Processor",
    page_icon="🧾",
    layout="wide",
)


# ============================================================
# RULES
# ============================================================

PACK_TYPES = {
    "Cigarettes (10)": 10,
    "Dip / Smokeless (5)": 5,
    "Each (1)": 1,
}


# ============================================================
# HELPERS
# ============================================================

def text_value(value):

    if value is None:
        return ""

    if (
        isinstance(value, float)
        and np.isnan(value)
    ):
        return ""

    return str(value).strip()


def to_number(value):

    if value is None:
        return None

    if isinstance(
        value,
        (
            int,
            float,
            np.integer,
            np.floating,
        ),
    ):

        if pd.isna(value):
            return None

        return float(value)

    cleaned = (
        str(value)
        .replace("$", "")
        .replace(",", "")
        .strip()
    )

    if (
        not cleaned
        or cleaned.lower() == "nan"
    ):
        return None

    try:
        return float(cleaned)

    except ValueError:
        return None


def margin_percent(cost, retail):

    cost = to_number(cost)
    retail = to_number(retail)

    if (
        cost is None
        or retail is None
        or retail <= 0
    ):
        return np.nan

    return round(
        (
            (retail - cost)
            / retail
        )
        * 100,
        1,
    )


def normalize_name(value):

    value = (
        text_value(value)
        .upper()
    )

    value = re.sub(
        r"[^A-Z0-9 ]+",
        " ",
        value,
    )

    return re.sub(
        r"\s+",
        " ",
        value,
    ).strip()


# ============================================================
# CLOVER
# ============================================================

def detect_clover_header_index(
    clover_bytes,
):

    preview = pd.read_excel(
        io.BytesIO(
            clover_bytes
        ),
        sheet_name="Items",
        header=None,
        nrows=15,
        dtype=object,
    )

    for index, row in (
        preview.iterrows()
    ):

        values = {
            text_value(value)
            for value
            in row.tolist()
        }

        if (
            "Clover ID"
            in values
            and "Name"
            in values
            and "Price"
            in values
            and "Cost"
            in values
        ):

            return int(index)

    raise ValueError(
        "Could not locate the Clover Items header row."
    )


def load_clover(
    clover_bytes,
):

    header_index = (
        detect_clover_header_index(
            clover_bytes
        )
    )

    df = pd.read_excel(
        io.BytesIO(
            clover_bytes
        ),
        sheet_name="Items",
        header=header_index,
        dtype=object,
    )

    df.columns = [
        text_value(column)
        for column
        in df.columns
    ]

    required = [
        "Clover ID",
        "Name",
        "Price",
        "Cost",
    ]

    missing = [
        column
        for column
        in required
        if column
        not in df.columns
    ]

    if missing:

        raise ValueError(
            "Clover Items sheet is missing: "
            + ", ".join(
                missing
            )
        )

    df = (
        df[
            df["Name"].notna()
        ]
        .copy()
        .reset_index(drop=True)
    )

    df["_Name Clean"] = (
        df["Name"]
        .apply(
            normalize_name
        )
    )

    return df


# ============================================================
# RECEIPT IMAGE / PDF
# ============================================================

def prepare_image(image):

    image = (
        ImageOps.exif_transpose(
            image
        )
    )

    image = (
        image
        .convert("L")
    )

    image = (
        ImageOps.autocontrast(
            image
        )
    )

    width, height = (
        image.size
    )

    # Keep OCR readable without making
    # very long scans unnecessarily huge.

    target_width = 1800

    if width != target_width:

        scale = (
            target_width
            / width
        )

        new_height = int(
            height * scale
        )

        image = image.resize(
            (
                target_width,
                new_height,
            ),
            Image.Resampling.LANCZOS,
        )

    return image


def receipt_to_images(
    uploaded_file,
):

    file_bytes = (
        uploaded_file
        .getvalue()
    )

    name = (
        uploaded_file
        .name
        .lower()
    )

    # ---------------- PDF ----------------

    if name.endswith(".pdf"):

        document = pymupdf.open(
            stream=file_bytes,
            filetype="pdf",
        )

        images = []

        for page in document:

            pix = page.get_pixmap(
                matrix=pymupdf.Matrix(
                    1.5,
                    1.5,
                ),
                alpha=False,
            )

            image = Image.frombytes(
                "RGB",
                (
                    pix.width,
                    pix.height,
                ),
                pix.samples,
            )

            images.append(
                image
            )

        document.close()

        return images

    # ---------------- IMAGE / TIFF ----------------

    image = Image.open(
        io.BytesIO(
            file_bytes
        )
    )

    images = []

    for frame in ImageSequence.Iterator(
        image
    ):

        images.append(
            frame.copy()
        )

    return images


# ============================================================
# OCR
# ============================================================

def read_receipt(
    uploaded_file,
):

    images = (
        receipt_to_images(
            uploaded_file
        )
    )

    page_texts = []

    for (
        page_number,
        image,
    ) in enumerate(
        images,
        start=1,
    ):

        prepared = (
            prepare_image(
                image
            )
        )

        text = (
            pytesseract
            .image_to_string(
                prepared,
                lang="eng",
                config=(
                    "--oem 1 "
                    "--psm 6"
                ),
            )
        )

        page_texts.append(
            (
                f"--- PAGE "
                f"{page_number} ---\n"
                f"{text}"
            )
        )

    return "\n".join(
        page_texts
    )


# ============================================================
# COSTCO PARSER
# ============================================================

def parse_costco_receipt(
    ocr_text,
):

    rows = []

    for raw_line in (
        ocr_text.splitlines()
    ):

        line = re.sub(
            r"\s+",
            " ",
            raw_line,
        ).strip()

        if not line:
            continue

        # Costco item number:
        # usually long numeric identifier.

        item_match = re.search(
            r"(?<!\d)"
            r"(\d{8,12})"
            r"(?!\d)",
            line,
        )

        # Find dollar-style amounts.
        amount_matches = list(
            re.finditer(
                r"(?<!\d)"
                r"(\d{1,5}[.,]\d{2})"
                r"(?!\d)",
                line,
            )
        )

        if (
            item_match is None
            or not amount_matches
        ):
            continue

        # Last dollar amount on the line
        # is treated as the item amount.

        amount_match = (
            amount_matches[-1]
        )

        amount_text = (
            amount_match
            .group(1)
            .replace(",", ".")
        )

        try:

            amount = float(
                amount_text
            )

        except ValueError:

            continue

        item_number = (
            item_match
            .group(1)
        )

        description = (
            line[
                item_match.end()
                :
                amount_match.start()
            ]
            .strip(
                " -:$"
            )
        )

        if not description:

            description = (
                "UNKNOWN"
            )

        upper_description = (
            description.upper()
        )

        # Protect against totals accidentally
        # being treated as merchandise.

        blocked_words = [
            "TOTAL",
            "SUBTOTAL",
            "TAX",
            "CHANGE",
            "TENDER",
        ]

        if any(
            word
            in upper_description
            for word
            in blocked_words
        ):

            continue

        rows.append(
            {
                "Item Number":
                    item_number,

                "Description":
                    description,

                "Vendor Package Cost":
                    round(
                        amount,
                        2,
                    ),
            }
        )

    if not rows:

        return pd.DataFrame(
            columns=[
                "Item Number",
                "Description",
                "Qty",
                "Vendor Package Cost",
                "Pack Type",
            ]
        )

    df = pd.DataFrame(
        rows
    )

    # Costco often prints identical cartons
    # as repeated receipt lines.
    # Group identical lines into quantity.

    grouped = (
        df.groupby(
            [
                "Item Number",
                "Description",
                "Vendor Package Cost",
            ],
            dropna=False,
        )
        .size()
        .reset_index(
            name="Qty"
        )
    )

    # Costco V1 is cigarette-focused.
    # User can change any line to dip.

    grouped[
        "Pack Type"
    ] = "Cigarettes (10)"

    return grouped[
        [
            "Item Number",
            "Description",
            "Qty",
            "Vendor Package Cost",
            "Pack Type",
        ]
    ]


# ============================================================
# RECEIPT RECONCILIATION
# ============================================================

def calculated_receipt_total(
    df,
):

    if (
        df is None
        or df.empty
    ):
        return 0.0

    qty = pd.to_numeric(
        df["Qty"],
        errors="coerce",
    ).fillna(0)

    cost = pd.to_numeric(
        df[
            "Vendor Package Cost"
        ],
        errors="coerce",
    ).fillna(0)

    return round(
        float(
            (
                qty
                * cost
            ).sum()
        ),
        2,
    )


def calculated_item_count(
    df,
):

    if (
        df is None
        or df.empty
    ):
        return 0

    qty = pd.to_numeric(
        df["Qty"],
        errors="coerce",
    ).fillna(0)

    return int(
        qty.sum()
    )


# ============================================================
# COST CALCULATION
# ============================================================

def calculate_unit_costs(
    df,
):

    result = (
        df.copy()
    )

    result[
        "Pack Size"
    ] = (
        result[
            "Pack Type"
        ]
        .map(
            PACK_TYPES
        )
        .fillna(1)
        .astype(int)
    )

    package_cost = (
        pd.to_numeric(
            result[
                "Vendor Package Cost"
            ],
            errors="coerce",
        )
    )

    result[
        "Calculated Clover Unit Cost"
    ] = (
        package_cost
        / result[
            "Pack Size"
        ]
    ).round(3)

    return result


# ============================================================
# CLOVER MATCHING
# ============================================================

def suggested_matches(
    description,
    clover_df,
    search_text="",
    limit=10,
):

    target = normalize_name(
        search_text
        if search_text.strip()
        else description
    )

    scored = []

    for (
        index,
        row,
    ) in clover_df.iterrows():

        candidate = (
            row[
                "_Name Clean"
            ]
        )

        if not candidate:
            continue

        if (
            target
            and target
            in candidate
        ):

            score = 1.0

        else:

            score = (
                difflib
                .SequenceMatcher(
                    None,
                    target,
                    candidate,
                )
                .ratio()
            )

        scored.append(
            (
                score,
                index,
            )
        )

    scored.sort(
        reverse=True
    )

    return [
        index
        for (
            _,
            index,
        )
        in scored[:limit]
    ]


def clover_label(
    row,
):

    name = text_value(
        row.get(
            "Name"
        )
    )

    retail = to_number(
        row.get(
            "Price"
        )
    )

    cost = to_number(
        row.get(
            "Cost"
        )
    )

    clover_id = text_value(
        row.get(
            "Clover ID"
        )
    )

    retail_text = (
        f"${retail:.2f}"
        if retail
        is not None
        else "—"
    )

    cost_text = (
        f"${cost:.2f}"
        if cost
        is not None
        else "—"
    )

    return (
        f"{name}"
        f"  |  Retail {retail_text}"
        f"  |  Cost {cost_text}"
        f"  |  {clover_id}"
    )


# ============================================================
# SCREEN
# ============================================================

st.title(
    "Vendor Invoice Processor"
)

st.caption(
    "Costco V1 — OCR, reconciliation, "
    "cost conversion, Clover matching "
    "and margin review."
)


# ============================================================
# INPUTS
# ============================================================

with st.container(
    border=True
):

    st.subheader(
        "1. Upload current files"
    )

    left, right = (
        st.columns(2)
    )

    with left:

        clover_file = (
            st.file_uploader(
                "Current Clover inventory export",
                type=[
                    "xlsx"
                ],
            )
        )

    with right:

        receipt_file = (
            st.file_uploader(
                "Costco receipt / scan",
                type=[
                    "jpg",
                    "jpeg",
                    "png",
                    "pdf",
                    "tif",
                    "tiff",
                ],
            )
        )


with st.container(
    border=True
):

    st.subheader(
        "2. Receipt controls"
    )

    left, right = (
        st.columns(2)
    )

    with left:

        expected_total = (
            st.number_input(
                "Final receipt total",
                min_value=0.0,
                step=0.01,
                format="%.2f",
            )
        )

    with right:

        expected_count_text = (
            st.text_input(
                "Total items/cartons (optional)",
                placeholder=(
                    "Example: 6"
                ),
            )
        )


ready = (
    clover_file
    is not None
    and receipt_file
    is not None
    and expected_total
    > 0
)


# ============================================================
# PROCESS
# ============================================================

if st.button(
    "Read & Reconcile Receipt",
    type="primary",
    use_container_width=True,
    disabled=not ready,
):

    try:

        with st.status(
            "Reading Costco receipt...",
            expanded=True,
        ) as status:

            status.write(
                "Reading current Clover inventory..."
            )

            clover_df = (
                load_clover(
                    clover_file
                    .getvalue()
                )
            )

            status.write(
                "Reading receipt image / pages..."
            )

            raw_text = (
                read_receipt(
                    receipt_file
                )
            )

            status.write(
                "Finding Costco merchandise lines..."
            )

            lines_df = (
                parse_costco_receipt(
                    raw_text
                )
            )

            st.session_state[
                "costco_clover"
            ] = clover_df

            st.session_state[
                "costco_ocr"
            ] = raw_text

            st.session_state[
                "costco_lines"
            ] = lines_df

            st.session_state[
                "costco_total"
            ] = float(
                expected_total
            )

            st.session_state[
                "costco_count"
            ] = (
                expected_count_text
                .strip()
            )

            status.update(
                label="Receipt read",
                state="complete",
                expanded=False,
            )

    except Exception as exc:

        st.error(
            f"Receipt processing stopped: {exc}"
        )


# ============================================================
# OCR REVIEW
# ============================================================

if (
    "costco_lines"
    in st.session_state
):

    st.divider()

    st.subheader(
        "3. Review extracted lines"
    )

    st.caption(
        "Correct OCR mistakes here. "
        "Rows can be added manually if OCR missed one."
    )

    edited_df = (
        st.data_editor(
            st.session_state[
                "costco_lines"
            ],
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={

                "Item Number":
                    st.column_config.TextColumn(
                        "Costco Item #",
                        required=True,
                    ),

                "Description":
                    st.column_config.TextColumn(
                        "Description",
                        required=True,
                        width="large",
                    ),

                "Qty":
                    st.column_config.NumberColumn(
                        "Qty",
                        min_value=1,
                        step=1,
                        required=True,
                    ),

                "Vendor Package Cost":
                    st.column_config.NumberColumn(
                        "Package Cost",
                        min_value=0.0,
                        format="$%.2f",
                        required=True,
                    ),

                "Pack Type":
                    st.column_config.SelectboxColumn(
                        "Pack Type",
                        options=list(
                            PACK_TYPES.keys()
                        ),
                        required=True,
                    ),
            },
            key="receipt_editor",
        )
    )

    st.session_state[
        "costco_lines"
    ] = edited_df


    # ========================================================
    # TOTAL RECONCILIATION
    # ========================================================

    calculated_total = (
        calculated_receipt_total(
            edited_df
        )
    )

    calculated_count = (
        calculated_item_count(
            edited_df
        )
    )

    user_total = (
        st.session_state[
            "costco_total"
        ]
    )

    difference = round(
        calculated_total
        - user_total,
        2,
    )

    m1, m2, m3 = (
        st.columns(3)
    )

    m1.metric(
        "Your receipt total",
        f"${user_total:,.2f}",
    )

    m2.metric(
        "App calculated total",
        f"${calculated_total:,.2f}",
    )

    m3.metric(
        "Difference",
        f"${difference:,.2f}",
    )

    totals_match = (
        abs(difference)
        <= 0.01
    )

    if totals_match:

        st.success(
            "Receipt total matches."
        )

    else:

        st.error(
            "Receipt does NOT match. "
            "Correct the extracted lines "
            "before continuing."
        )


    # ========================================================
    # OPTIONAL ITEM COUNT
    # ========================================================

    entered_count = None

    count_text = (
        st.session_state
        .get(
            "costco_count",
            "",
        )
    )

    if count_text:

        try:

            entered_count = int(
                count_text
            )

        except ValueError:

            st.warning(
                "Optional item count must "
                "be a whole number."
            )

    if (
        entered_count
        is not None
    ):

        if (
            entered_count
            == calculated_count
        ):

            st.success(
                f"Item count also matches: "
                f"{calculated_count}."
            )

        else:

            st.warning(
                f"Item count mismatch. "
                f"You entered {entered_count}; "
                f"the app calculated "
                f"{calculated_count}."
            )


    # ========================================================
    # RAW OCR
    # ========================================================

    with st.expander(
        "Show raw OCR text"
    ):

        st.text(
            st.session_state[
                "costco_ocr"
            ]
        )


    # ========================================================
    # PRICING / MARGINS
    # ========================================================

    if totals_match:

        st.divider()

        st.subheader(
            "4. Clover cost & margin review"
        )

        cost_df = (
            calculate_unit_costs(
                edited_df
            )
        )

        clover_df = (
            st.session_state[
                "costco_clover"
            ]
        )

        for (
            row_index,
            item,
        ) in cost_df.iterrows():

            description = (
                text_value(
                    item[
                        "Description"
                    ]
                )
            )

            item_number = (
                text_value(
                    item[
                        "Item Number"
                    ]
                )
            )

            package_cost = (
                to_number(
                    item[
                        "Vendor Package Cost"
                    ]
                )
            )

            new_cost = (
                to_number(
                    item[
                        "Calculated Clover Unit Cost"
                    ]
                )
            )

            with st.container(
                border=True
            ):

                st.markdown(
                    f"### {description}"
                )

                a, b, c = (
                    st.columns(3)
                )

                a.metric(
                    "Costco Item #",
                    item_number,
                )

                b.metric(
                    "Vendor package cost",
                    f"${package_cost:,.2f}",
                )

                c.metric(
                    "Calculated Clover unit cost",
                    f"${new_cost:,.3f}",
                )


                # ============================================
                # CLOVER SEARCH
                # ============================================

                search_text = (
                    st.text_input(
                        "Search Clover",
                        placeholder=description,
                        key=(
                            f"search_"
                            f"{row_index}_"
                            f"{item_number}"
                        ),
                    )
                )

                match_indexes = (
                    suggested_matches(
                        description,
                        clover_df,
                        search_text,
                        limit=10,
                    )
                )

                option_map = {
                    "— Select Clover item —":
                        None
                }

                for match_index in (
                    match_indexes
                ):

                    clover_row = (
                        clover_df.loc[
                            match_index
                        ]
                    )

                    option_map[
                        clover_label(
                            clover_row
                        )
                    ] = match_index

                selected_label = (
                    st.selectbox(
                        "Suggested Clover matches",
                        options=list(
                            option_map.keys()
                        ),
                        key=(
                            f"select_"
                            f"{row_index}_"
                            f"{item_number}"
                        ),
                    )
                )

                selected_index = (
                    option_map[
                        selected_label
                    ]
                )

                if (
                    selected_index
                    is None
                ):

                    st.info(
                        "Select the correct "
                        "Clover item."
                    )

                    continue


                # ============================================
                # CURRENT CLOVER VALUES
                # ============================================

                clover_row = (
                    clover_df.loc[
                        selected_index
                    ]
                )

                old_cost = (
                    to_number(
                        clover_row.get(
                            "Cost"
                        )
                    )
                )

                retail = (
                    to_number(
                        clover_row.get(
                            "Price"
                        )
                    )
                )

                old_margin = (
                    margin_percent(
                        old_cost,
                        retail,
                    )
                )

                new_margin = (
                    margin_percent(
                        new_cost,
                        retail,
                    )
                )

                if (
                    not pd.isna(
                        old_margin
                    )
                    and not pd.isna(
                        new_margin
                    )
                ):

                    margin_change = round(
                        new_margin
                        - old_margin,
                        1,
                    )

                else:

                    margin_change = (
                        np.nan
                    )


                # ============================================
                # METRICS
                # ============================================

                p1, p2, p3, p4 = (
                    st.columns(4)
                )

                p1.metric(
                    "Current Clover cost",
                    (
                        f"${old_cost:,.3f}"
                        if old_cost
                        is not None
                        else "—"
                    ),
                )

                p2.metric(
                    "Current retail",
                    (
                        f"${retail:,.2f}"
                        if retail
                        is not None
                        else "—"
                    ),
                )

                p3.metric(
                    "Old margin",
                    (
                        f"{old_margin:.1f}%"
                        if not pd.isna(
                            old_margin
                        )
                        else "—"
                    ),
                )

                p4.metric(
                    "New margin",
                    (
                        f"{new_margin:.1f}%"
                        if not pd.isna(
                            new_margin
                        )
                        else "—"
                    ),
                    delta=(
                        f"{margin_change:+.1f} pts"
                        if not pd.isna(
                            margin_change
                        )
                        else None
                    ),
                )


                # ============================================
                # BUSINESS RULES
                # ============================================

                if (
                    package_cost
                    is not None
                    and package_cost
                    <= 0
                ):

                    st.error(
                        "Zero/free promotional line. "
                        "Do NOT overwrite the existing "
                        "Clover cost."
                    )

                elif (
                    old_cost
                    is None
                ):

                    st.warning(
                        "Clover has no current cost. "
                        "Proposed cost is the calculated "
                        "unit cost."
                    )

                elif (
                    new_cost
                    > old_cost
                    + 0.004
                ):

                    st.warning(
                        "COST INCREASED — proposed "
                        "Clover cost uses the new cost. "
                        "Retail remains unchanged. "
                        "Review the new margin before "
                        "deciding whether retail should "
                        "increase."
                    )

                elif (
                    new_cost
                    < old_cost
                    - 0.004
                ):

                    st.info(
                        "COST DECREASED — possible "
                        "promotion. Proposed cost uses "
                        "the lower positive invoice cost. "
                        "Retail remains unchanged and "
                        "will never be lowered "
                        "automatically."
                    )

                else:

                    st.success(
                        "Cost is effectively unchanged. "
                        "Retail remains unchanged."
                    )


        st.info(
            "Costco V1 deliberately stops here. "
            "We first prove that OCR, receipt totals, "
            "pack conversion, Clover matching and "
            "margin calculations are correct. "
            "Then we add the safe Clover output file."
        )
