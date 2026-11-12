
import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(
    page_title="Excel Reconciliation",
    layout="wide"
)

st.title("Excel Reconciliation App")
st.caption("Books vs GST Portal reconciliation")


# =====================================================
# 1. PORTAL COLUMN DEFINITIONS
# =====================================================

STARRED_FIELDS = [
    "Trade/Legal name*",
    "Invoice type*",
    "Invoice Date*",
    "Taxable Value (₹)*",
    "Integrated Tax(₹)*",
    "Central Tax(₹)*",
    "State/UT Tax(₹)*",
    "Cess(₹)*",
    "ITC Availability*",
]

OPTIONAL_PORTAL_FIELDS = [
    "Invoice Value(₹)",
    "Place of supply",
    "Supply Attract Reverse Charge",
    "GSTR-1/1A/IFF/GSTR-5 Period",
    "GSTR-1/1A/IFF/GSTR-5 Filing Date",
    "Reason",
    "Applicable % of Tax Rate",
    "Source",
    "IRN",
    "IRN Date",
]


# =====================================================
# 2. HELPER FUNCTIONS
# =====================================================

def clean_header(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_header(value):
    """Ignore spaces, punctuation, asterisks and case."""
    return re.sub(
        r"[^a-z0-9]",
        "",
        clean_header(value).lower()
    )


def find_column(columns, candidates):
    """Find an actual column using flexible header matching."""

    normalized = {
        normalize_header(col): col
        for col in columns
    }

    for candidate in candidates:
        key = normalize_header(candidate)

        if key in normalized:
            return normalized[key]

    return None


def detect_header_row(raw, file_type):
    """Find the real table header within the first 50 rows."""

    if file_type == "portal":
        expected = [
            "GSTIN of supplier",
            "Invoice number",
            "Trade/Legal name",
            "Invoice type",
            "Invoice Date",
            "Invoice Value",
            "Taxable Value",
            "Integrated Tax",
            "Central Tax",
            "State/UT Tax",
            "Cess",
            "ITC Availability",
        ]
    else:
        expected = [
            "Books GST No.",
            "Books Bill No.",
            "Books Date",
            "Party Name (Books/Portal)",
            "TV",
            "IGST",
            "CGST",
            "SGST",
            "Cess",
        ]

    expected = {
        normalize_header(x)
        for x in expected
    }

    best_row = None
    best_score = 0

    for row_index in range(min(50, len(raw))):

        values = [
            normalize_header(value)
            for value in raw.iloc[row_index]
            if pd.notna(value)
        ]

        score = sum(
            1 for value in values
            if value in expected
        )

        if score > best_score:
            best_score = score
            best_row = row_index

    if best_row is None or best_score < 2:
        raise ValueError(
            f"Could not identify the {file_type} table header. "
            "Please check the Excel file."
        )

    return best_row


def read_excel(file, file_type):
    """Read Excel and automatically detect its header row."""

    file.seek(0)

    raw = pd.read_excel(
        file,
        header=None,
        dtype=object
    )

    header_index = detect_header_row(raw, file_type)

    headers = [
        clean_header(value)
        for value in raw.iloc[header_index]
    ]

    # Remove blank and Unnamed columns
    keep = [
        i for i, header in enumerate(headers)
        if header
        and not header.lower().startswith("unnamed")
    ]

    headers = [headers[i] for i in keep]

    data = raw.iloc[header_index + 1:, keep].copy()
    data.columns = headers

    # Keep original Excel row numbers even if blank rows are removed.
    source_excel_rows = (data.index + 1).tolist()
    nonblank_mask = ~data.isna().all(axis=1)

    df = data.loc[nonblank_mask].reset_index(drop=True)
    df.attrs["header_excel_row"] = header_index + 1
    df.attrs["excel_rows"] = [
        row_number
        for row_number, keep_row in zip(source_excel_rows, nonblank_mask)
        if keep_row
    ]

    return df


def text_value(value):
    if pd.isna(value):
        return ""

    value = str(value).strip()

    if value.lower() in ["nan", "none", "nat"]:
        return ""

    return value


def gst_value(value):
    return re.sub(
        r"\s+",
        "",
        text_value(value).upper()
    )


def invoice_key(value):
    """
    Normalize invoice numbers.

    Examples:
    832 -> 832
    NA-832/25-26 -> 832
    ROYAL/25-26/0610 -> 610
    MH1170105826 -> 1170105826
    """

    value = text_value(value).upper()

    if not value:
        return ""

    value = re.sub(r"\s+", "", value)

    # Remove financial year
    value = re.sub(r"\d{2}-\d{2}", "", value)

    numbers = re.findall(r"\d+", value)

    if numbers:
        return str(int(numbers[-1]))

    return value


def amount(value):
    if pd.isna(value):
        return 0.0

    value = str(value).strip()

    if not value:
        return 0.0

    negative = (
        value.startswith("(")
        and value.endswith(")")
    )

    value = (
        value.replace(",", "")
        .replace("₹", "")
        .replace("(", "")
        .replace(")", "")
        .replace(" ", "")
    )

    try:
        result = float(value)

        return -abs(result) if negative else result

    except (ValueError, TypeError):
        return 0.0


def date_value(value):
    if pd.isna(value):
        return ""

    try:
        return pd.to_datetime(value).strftime("%Y-%m-%d")
    except Exception:
        return text_value(value)


def is_amount_field(field):
    name = normalize_header(field)

    return any(x in name for x in [
        "taxablevalue",
        "integratedtax",
        "centraltax",
        "stateuttax",
        "cess",
        "invoicevalue",
    ])


def is_date_field(field):
    return "date" in normalize_header(field)


def compare_values(book_value, portal_value, portal_field):

    if is_amount_field(portal_field):
        return amount(book_value), amount(portal_value)

    if is_date_field(portal_field):
        return date_value(book_value), date_value(portal_value)

    return (
        text_value(book_value).upper(),
        text_value(portal_value).upper()
    )


def calculate_difference(book_value, portal_value, field):

    a, b = compare_values(
        book_value,
        portal_value,
        field
    )

    if is_amount_field(field):
        return round(a - b, 2)

    return "MATCH" if a == b else "DIFFERENT"


def total_tax(row, columns):
    total = 0.0

    for column in columns:
        if column and column in row.index:
            total += amount(row[column])

    return round(total, 2)


# =====================================================
# 3. UPLOAD FILES
# =====================================================

st.header("1. Upload Excel Files")

col1, col2 = st.columns(2)

with col1:
    books_file = st.file_uploader(
        "Upload Books Excel",
        type=["xlsx"],
        key="books"
    )

with col2:
    portal_file = st.file_uploader(
        "Upload Portal Excel",
        type=["xlsx"],
        key="portal"
    )

if not books_file or not portal_file:
    st.info("Please upload both Excel files.")
    st.stop()


try:
    books_df = read_excel(books_file, "books")
    portal_df = read_excel(portal_file, "portal")

except Exception as e:
    st.error(f"Error reading Excel files: {e}")
    st.stop()


books_columns = list(books_df.columns)
portal_columns = list(portal_df.columns)

st.success("Both Excel files loaded successfully.")

col1, col2 = st.columns(2)

with col1:
    st.info(
        f"Books header detected at Excel row "
        f"{books_df.attrs['header_excel_row']}"
    )

with col2:
    st.info(
        f"Portal header detected at Excel row "
        f"{portal_df.attrs['header_excel_row']}"
    )


# =====================================================
# 4. PRIMARY MATCHING COLUMN SELECTION
# =====================================================

st.header("2. Primary Matching Columns")

books_gst_default = find_column(
    books_columns,
    ["Books GST No.", "GSTIN", "GST No"]
)

books_invoice_default = find_column(
    books_columns,
    ["Books Bill No.", "Bill No", "Invoice No"]
)

portal_gst_default = find_column(
    portal_columns,
    ["GSTIN of supplier"]
)

portal_invoice_default = find_column(
    portal_columns,
    ["Invoice number"]
)


def get_index(options, selected):
    if selected in options:
        return options.index(selected)
    return 0


c1, c2 = st.columns(2)

with c1:
    books_gst_col = st.selectbox(
        "Books GSTIN column",
        books_columns,
        index=get_index(books_columns, books_gst_default)
    )

with c2:
    portal_gst_col = st.selectbox(
        "Portal GSTIN column",
        portal_columns,
        index=get_index(portal_columns, portal_gst_default)
    )

c1, c2 = st.columns(2)

with c1:
    books_invoice_col = st.selectbox(
        "Books Invoice Number column",
        books_columns,
        index=get_index(books_columns, books_invoice_default)
    )

with c2:
    portal_invoice_col = st.selectbox(
        "Portal Invoice Number column",
        portal_columns,
        index=get_index(portal_columns, portal_invoice_default)
    )


# =====================================================
# 5. COMPARISON COLUMN MAPPING
# =====================================================

st.header("3. Comparison Column Mapping")

st.write(
    "Select which Books column should be compared "
    "against each Portal column."
)

st.caption(
    "Starred Portal fields are included automatically. "
    "Columns without a corresponding Books field can be left unselected."
)

# Suggested Books column names for each Portal field
DEFAULT_MAPPING = {
    "Trade/Legal name*": [
        "Party Name (Books/Portal)",
        "Party Name",
        "Vendor Name",
    ],
    "Invoice type*": [
        "Invoice Type",
        "Voucher Type",
        "Document Type",
    ],
    "Invoice Date*": [
        "Books Date",
        "Date",
    ],
    "Invoice Value(₹)": [
        "Invoice Value",
    ],
    "Taxable Value (₹)*": [
        "TV",
        "Taxable Value",
    ],
    "Integrated Tax(₹)*": [
        "IGST",
    ],
    "Central Tax(₹)*": [
        "CGST",
    ],
    "State/UT Tax(₹)*": [
        "SGST",
    ],
    "Cess(₹)*": [
        "Cess",
    ],
    "ITC Availability*": [
        "ITC Availability",
    ],
}

# Resolve actual Portal headers, even if punctuation or * differs
resolved_starred = {}

for expected in STARRED_FIELDS:
    actual = find_column(portal_columns, [expected])

    if actual:
        resolved_starred[expected] = actual

st.subheader("Automatically identified starred Portal fields")

if resolved_starred:
    st.write(list(resolved_starred.values()))
else:
    st.warning(
        "No starred Portal fields were automatically identified."
    )

missing_starred = [
    field for field in STARRED_FIELDS
    if field not in resolved_starred
]

if missing_starred:
    st.warning(
        "These expected Portal fields were not found: "
        + ", ".join(missing_starred)
    )

with st.expander("Add optional Portal fields"):
    optional_available = [
        actual
        for expected in OPTIONAL_PORTAL_FIELDS
        if (actual := find_column(portal_columns, [expected]))
    ]

    additional_fields = st.multiselect(
        "Select additional Portal columns",
        options=optional_available
    )

# Only use actual Portal headers
comparison_portal_fields = (
    list(resolved_starred.values())
    + additional_fields
)

mapping = {}

st.subheader("Choose Books columns")

for actual_portal_col in comparison_portal_fields:

    # Find the expected field name for suggested mapping
    expected_name = actual_portal_col

    for expected, actual in resolved_starred.items():
        if actual == actual_portal_col:
            expected_name = expected
            break

    suggestions = DEFAULT_MAPPING.get(
        expected_name,
        []
    )

    suggested = find_column(
        books_columns,
        suggestions
    )

    options = ["-- Do not compare --"] + books_columns

    default_index = (
        options.index(suggested)
        if suggested in options
        else 0
    )

    selected_books_col = st.selectbox(
        f"Books column for Portal: {actual_portal_col}",
        options,
        index=default_index,
        key=f"map_{actual_portal_col}"
    )

    if selected_books_col != "-- Do not compare --":
        mapping[actual_portal_col] = selected_books_col


st.write(
    f"Number of mapped comparison fields: {len(mapping)}"
)

tolerance = st.number_input(
    "Amount Difference Tolerance (₹)",
    min_value=0.0,
    value=0.01,
    step=0.01
)


# =====================================================
# 6. RECONCILIATION
# =====================================================

st.header("4. Reconciliation")

if st.button("Reconcile Files", type="primary"):

    results = []
    matched_portal_rows = set()

    # Convert each DataFrame to records once. This avoids the overhead
    # of repeatedly creating pandas Series with iterrows().
    books_records = books_df.to_dict("records")
    portal_records = portal_df.to_dict("records")

    books_excel_rows = books_df.attrs.get(
        "excel_rows",
        list(range(books_df.attrs["header_excel_row"] + 1,
                   books_df.attrs["header_excel_row"] + 1 + len(books_df)))
    )
    portal_excel_rows = portal_df.attrs.get(
        "excel_rows",
        list(range(portal_df.attrs["header_excel_row"] + 1,
                   portal_df.attrs["header_excel_row"] + 1 + len(portal_df)))
    )

    # Build two Portal indexes:
    # 1) GSTIN + invoice number for the preferred exact match.
    # 2) Invoice number alone for the "Needs Review" fallback.
    portal_index = {}
    portal_invoice_index = {}

    progress = st.progress(0, text="Building Portal invoice indexes...")

    for portal_i, portal_row in enumerate(portal_records):
        gst = gst_value(portal_row.get(portal_gst_col))
        invoice = invoice_key(portal_row.get(portal_invoice_col))

        if not invoice:
            continue

        portal_invoice_index.setdefault(invoice, []).append(portal_i)

        if gst:
            key = (gst, invoice)
            portal_index.setdefault(key, []).append(portal_i)

    # Tax columns for total tax calculation.
    books_tax_columns = [
        find_column(books_columns, ["IGST"]),
        find_column(books_columns, ["CGST"]),
        find_column(books_columns, ["SGST"]),
        find_column(books_columns, ["Cess"]),
    ]

    portal_tax_columns = [
        find_column(portal_columns, ["Integrated Tax(₹)*"]),
        find_column(portal_columns, ["Central Tax(₹)*"]),
        find_column(portal_columns, ["State/UT Tax(₹)*"]),
        find_column(portal_columns, ["Cess(₹)*"]),
    ]

    def record_total_tax(row, columns):
        total = 0.0
        for column in columns:
            if column:
                total += amount(row.get(column))
        return round(total, 2)

    total_books = len(books_records)
    update_every = max(1, total_books // 100)

    for books_i, books_row in enumerate(books_records):

        books_gst = gst_value(books_row.get(books_gst_col))
        books_invoice = invoice_key(books_row.get(books_invoice_col))
        key = (books_gst, books_invoice)

        # Exact GSTIN + invoice match.
        candidates = [
            i for i in portal_index.get(key, [])
            if i not in matched_portal_rows
        ] if books_gst and books_invoice else []

        portal_i = None
        portal_row = None

        if len(candidates) == 1:
            portal_i = candidates[0]
            portal_row = portal_records[portal_i]
            matched_portal_rows.add(portal_i)
            status = "Matched"

        elif len(candidates) > 1:
            status = "Ambiguous Match"

        else:
            # Fast fallback lookup by invoice number only.
            # A unique invoice with a different GSTIN is not treated
            # as an exact match; it is flagged for manual review.
            possible = [
                i for i in portal_invoice_index.get(books_invoice, [])
                if i not in matched_portal_rows
            ] if books_invoice else []

            if len(possible) == 1:
                portal_i = possible[0]
                portal_row = portal_records[portal_i]
                matched_portal_rows.add(portal_i)
                status = "Needs Review"

            elif len(possible) > 1:
                status = "Ambiguous Match"

            else:
                status = "Only in Books"

        output = {
            "Books Excel Row": books_excel_rows[books_i],
            "Portal Excel Row": (
                portal_excel_rows[portal_i]
                if portal_i is not None else ""
            ),
            "Status": status,
        }

        # Preserve all original Books columns.
        for col in books_columns:
            output[f"Books | {col}"] = books_row.get(col)

        # Preserve all original Portal columns and compare mapped fields.
        if portal_row is not None:
            for col in portal_columns:
                output[f"Portal | {col}"] = portal_row.get(col)

            differences = []

            for portal_col, books_col in mapping.items():
                books_value = books_row.get(books_col)
                portal_value = portal_row.get(portal_col)

                diff = calculate_difference(
                    books_value,
                    portal_value,
                    portal_col
                )

                output[f"Difference | {portal_col}"] = diff

                if isinstance(diff, (int, float)):
                    if abs(diff) > tolerance:
                        differences.append(portal_col)
                elif diff == "DIFFERENT":
                    differences.append(portal_col)

            books_tax = record_total_tax(books_row, books_tax_columns)
            portal_tax = record_total_tax(portal_row, portal_tax_columns)

            output["Books Total Tax"] = books_tax
            output["Portal Total Tax"] = portal_tax
            output["Total Tax Difference"] = round(
                books_tax - portal_tax, 2
            )

            if status == "Matched":
                output["Status"] = (
                    "Mismatch" if differences else "Exact Match"
                )

        else:
            for col in portal_columns:
                output[f"Portal | {col}"] = ""

            for portal_col in mapping:
                output[f"Difference | {portal_col}"] = ""

            output["Books Total Tax"] = record_total_tax(
                books_row, books_tax_columns
            )
            output["Portal Total Tax"] = 0
            output["Total Tax Difference"] = output["Books Total Tax"]

        results.append(output)

        if (books_i + 1) % update_every == 0 or books_i + 1 == total_books:
            progress.progress(
                min((books_i + 1) / max(total_books, 1), 1.0),
                text=f"Reconciling Books rows: {books_i + 1:,} / {total_books:,}"
            )

    # Add all unmatched Portal rows.
    for portal_i, portal_row in enumerate(portal_records):
        if portal_i in matched_portal_rows:
            continue

        output = {
            "Books Excel Row": "",
            "Portal Excel Row": portal_excel_rows[portal_i],
            "Status": "Only on Portal",
        }

        for col in books_columns:
            output[f"Books | {col}"] = ""

        for col in portal_columns:
            output[f"Portal | {col}"] = portal_row.get(col)

        for portal_col in mapping:
            output[f"Difference | {portal_col}"] = ""

        output["Books Total Tax"] = 0
        output["Portal Total Tax"] = record_total_tax(
            portal_row, portal_tax_columns
        )
        output["Total Tax Difference"] = -output["Portal Total Tax"]

        results.append(output)

    progress.empty()

    result_df = pd.DataFrame(results)
    st.session_state["reconciliation_results"] = result_df
    st.success("Reconciliation completed!")


# =====================================================
# 7. DISPLAY RESULTS
# =====================================================

if "reconciliation_results" in st.session_state:

    result_df = st.session_state["reconciliation_results"]

    st.header("5. Reconciliation Results")

    counts = result_df["Status"].value_counts()

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "Total Rows",
        len(result_df)
    )

    c2.metric(
        "Exact Matches",
        counts.get("Exact Match", 0)
    )

    c3.metric(
        "Mismatches",
        counts.get("Mismatch", 0)
    )

    c4.metric(
        "Only in Books",
        counts.get("Only in Books", 0)
    )

    st.write("### Status Summary")

    st.dataframe(
        counts.rename("Number of Invoices").to_frame(),
        use_container_width=True
    )

    selected_status = st.multiselect(
        "Filter by Status",
        options=list(result_df["Status"].unique()),
        default=list(result_df["Status"].unique())
    )

    displayed = result_df[
        result_df["Status"].isin(selected_status)
    ]

    st.dataframe(
        displayed,
        use_container_width=True,
        height=500
    )

    # Download Excel report
    buffer = io.BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl"
    ) as writer:

        result_df.to_excel(
            writer,
            index=False,
            sheet_name="Reconciliation"
        )

        counts.rename(
            "Number of Invoices"
        ).to_frame().to_excel(
            writer,
            sheet_name="Summary"
        )

    st.download_button(
        label="Download Reconciliation Report",
        data=buffer.getvalue(),
        file_name="Reconciliation_Report.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )
