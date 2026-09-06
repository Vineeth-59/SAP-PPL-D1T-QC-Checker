import io, re
import pandas as pd

FIELD_MAPPING = {
    "currency": "Currency",
    "country": "Country/Region Key",
    "country/region key": "Country/Region Key",
    "description": "Company Name",
    "company name": "Company Name",
    "city": "City",
    "address": "Address",
    "language key": "Language Key",
    "name 1": "Company Name",
}

def clean(v):
    if pd.isna(v): return ""
    return str(v).strip()

def norm(v):
    s = clean(v).replace("\n", " ")
    return re.sub(r"\s+", " ", s).casefold()

def find_d1t_column(field, columns):
    target = FIELD_MAPPING.get(clean(field).casefold())
    if target in columns: return target
    for c in columns:
        if norm(c) == norm(target or field): return c
    return None

def run_qc(ppl, d1t, company, tcode, stream="All"):
    p = ppl.copy()
    d = d1t.copy()
    p["T-Code"] = p["T-Code"].astype(str).str.strip()
    p = p[p["T-Code"] == str(tcode).strip()].copy()
    if stream != "All":
        p = p[p["Value Stream"].astype(str).str.strip() == str(stream).strip()].copy()

    d["Company Code"] = d["Company Code"].astype(str).str.strip()
    actual = d[d["Company Code"] == str(company).strip()].copy()
    rows = []

    for _, r in p.iterrows():
        field = clean(r["Configuration key field NAMES (identifier)"])
        expected = clean(r["New Org Value"])
        common = dict(
            Step=clean(r["Step"]), **{"T-Code":clean(r["T-Code"])},
            **{"IMG-Activity":clean(r["IMG-Activity"])},
            **{"Configuration Key Field":field},
            **{"Template Value":clean(r["Template Value"])},
            Rule=clean(r["Rule"]), **{"Expected Value":expected}
        )

        if not expected:
            common.update(**{"Actual Value":""}, Status="NOT CHECKABLE",
                          Reason="PPL has no New Org Value for this row.")
            rows.append(common); continue

        if actual.empty:
            common.update(**{"Actual Value":""}, Status="FAIL",
                          Reason="Company Code not found in D1T export.")
            rows.append(common); continue

        dcol = find_d1t_column(field, d.columns)
        if not dcol:
            common.update(**{"Actual Value":""}, Status="NOT CHECKABLE",
                          Reason=f"No D1T column mapping found for PPL field '{field}'.")
            rows.append(common); continue

        av = clean(actual.iloc[0][dcol])
        ok = norm(expected) == norm(av)
        common.update(**{"Actual Value":av},
                      Status="PASS" if ok else "FAIL",
                      Reason="Expected and actual values match." if ok else
                             f"Value mismatch. PPL expects '{expected}', D1T has '{av}'.")
        rows.append(common)

    result = pd.DataFrame(rows)
    total = int((result["Status"] != "NOT CHECKABLE").sum()) if len(result) else 0
    passed = int((result["Status"] == "PASS").sum()) if len(result) else 0
    failed = int((result["Status"] == "FAIL").sum()) if len(result) else 0
    nc = int((result["Status"] == "NOT CHECKABLE").sum()) if len(result) else 0
    return result, {"total":total,"pass":passed,"fail":failed,"not_checkable":nc,
                    "pass_rate":(passed/total*100 if total else 0)}

def make_excel(result, summary):
    out=io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        result.to_excel(w,index=False,sheet_name="QC Results")
        pd.DataFrame([summary]).to_excel(w,index=False,sheet_name="Summary")
    return out.getvalue()
