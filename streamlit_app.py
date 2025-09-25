import io, zipfile
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Enterprise Structure Generator", page_icon="📊", layout="wide")
st.title("Enterprise Structure Generator — Excel Only")

st.markdown("""
Upload up to **4 Oracle export ZIPs** (any order):
- `Manage General Ledger` (Ledgers)
- `Manage Legal Entities` (Legal Entities)
- `Assign Legal Entities` (Ledger↔LE mapping)
- `Manage Business Units` (Business Units)
""")

uploads = st.file_uploader("Drop your ZIPs here", type="zip", accept_multiple_files=True)

def read_csv_from_zip(zf, name):
    if name not in zf.namelist():
        return None
    with zf.open(name) as fh:
        return pd.read_csv(fh, dtype=str)

if not uploads:
    st.info("Upload your ZIPs to generate the Excel.")
else:
    # collectors
    ledger_names = set()            # GL_PRIMARY_LEDGER.csv :: ORA_GL_PRIMARY_LEDGER_CONFIG.Name
    legal_entity_names = set()      # XLE_ENTITY_PROFILE.csv :: Name
    ledger_to_idents = {}           # ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv :: GL_LEDGER.Name -> {LegalEntityIdentifier}
    ident_to_le_name = {}           # ORA_GL_JOURNAL_CONFIG_DETAIL.csv     :: LegalEntityIdentifier -> ObjectName
    bu_rows = []                    # FUN_BUSINESS_UNIT.csv :: Name, PrimaryLedgerName, LegalEntityName

    # scan all uploaded zips
    for up in uploads:
        try:
            z = zipfile.ZipFile(up)
        except Exception as e:
            st.error(f"Could not open `{up.name}` as a ZIP: {e}")
            continue

        # Ledgers
        df = read_csv_from_zip(z, "GL_PRIMARY_LEDGER.csv")
        if df is not None:
            col = "ORA_GL_PRIMARY_LEDGER_CONFIG.Name"
            if col in df.columns:
                ledger_names |= set(df[col].dropna().map(str).str.strip())
            else:
                st.warning(f"`GL_PRIMARY_LEDGER.csv` missing `{col}`. Found: {list(df.columns)}")

        # Legal Entities
        df = read_csv_from_zip(z, "XLE_ENTITY_PROFILE.csv")
        if df is not None:
            col = "Name"
            if col in df.columns:
                legal_entity_names |= set(df[col].dropna().map(str).str.strip())
            else:
                st.warning(f"`XLE_ENTITY_PROFILE.csv` missing `Name`. Found: {list(df.columns)}")

        # Ledger ↔ LE identifier
        df = read_csv_from_zip(z, "ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv")
        if df is not None:
            need = ["GL_LEDGER.Name", "LegalEntityIdentifier"]
            missing = [c for c in need if c not in df.columns]
            if missing:
                st.warning(f"`ORA_LEGAL_ENTITY_BAL_SEG_VAL_DEF.csv` missing {missing}. Found: {list(df.columns)}")
            else:
                for _, r in df[need].dropna().iterrows():
                    led = str(r["GL_LEDGER.Name"]).strip()
                    ident = str(r["LegalEntityIdentifier"]).strip()
                    if led and ident:
                        ledger_to_idents.setdefault(led, set()).add(ident)

        # Identifier ↔ LE name
        df = read_csv_from_zip(z, "ORA_GL_JOURNAL_CONFIG_DETAIL.csv")
        if df is not None:
            need = ["LegalEntityIdentifier", "ObjectName"]
            missing = [c for c in need if c not in df.columns]
            if missing:
                st.warning(f"`ORA_GL_JOURNAL_CONFIG_DETAIL.csv` missing {missing}. Found: {list(df.columns)}")
            else:
                for _, r in df[need].dropna().iterrows():
                    ident = str(r["LegalEntityIdentifier"]).strip()
                    obj = str(r["ObjectName"]).strip()
                    if ident:
                        ident_to_le_name[ident] = obj

        # Business Units
        df = read_csv_from_zip(z, "FUN_BUSINESS_UNIT.csv")
        if df is not None:
            need = ["Name", "PrimaryLedgerName", "LegalEntityName"]
            missing = [c for c in need if c not in df.columns]
            if missing:
                st.warning(f"`FUN_BUSINESS_UNIT.csv` missing {missing}. Found: {list(df.columns)}")
            else:
                for c in need:
                    df[c] = df[c].astype(str).map(lambda x: x.strip() if x else "")
                bu_rows += df[need].to_dict(orient="records")

    # build mappings
    ledger_to_le_names = {}
    for led, idents in ledger_to_idents.items():
        for ident in idents:
            le_name = ident_to_le_name.get(ident, "").strip()
            if le_name:
                ledger_to_le_names.setdefault(led, set()).add(le_name)

    le_to_ledgers = {}
    for led, le_set in ledger_to_le_names.items():
        for le in le_set:
            le_to_ledgers.setdefault(le, set()).add(led)

    # Build final rows: Ledger Name | Legal Entity | Business Unit
    rows = []
    seen_triples = set()
    seen_ledgers_with_bu = set()
    seen_les_with_bu = set()

    # 1) BU-driven rows with smart back-fill
    for r in bu_rows:
        bu = r["Name"]
        led = r["PrimaryLedgerName"] if r["PrimaryLedgerName"] in ledger_names else ""
        le  = r["LegalEntityName"]  if r["LegalEntityName"]  in legal_entity_names else ""

        # back-fill ledger from LE if missing and unique
        if not led and le and le in le_to_ledgers and len(le_to_ledgers[le]) == 1:
            led = next(iter(le_to_ledgers[le]))
        # back-fill LE from ledger if missing and unique
        if not le and led and led in ledger_to_le_names and len(ledger_to_le_names[led]) == 1:
            le = next(iter(ledger_to_le_names[led]))

        rows.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": bu})
        seen_triples.add((led, le, bu))
        if led: seen_ledgers_with_bu.add(led)
        if le:  seen_les_with_bu.add(le)

    # 2) Ledger–LE pairs with no BU
    seen_pairs = {(a, b) for (a, b, _) in seen_triples}
    for led, le_set in ledger_to_le_names.items():
        if not le_set:
            if led not in seen_ledgers_with_bu:
                rows.append({"Ledger Name": led, "Legal Entity": "", "Business Unit": ""})
            continue
        for le in le_set:
            if (led, le) not in seen_pairs:
                rows.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": ""})

    # 3) Orphan ledgers (in master list) with no mapping & no BU
    for led in sorted(ledger_names - set(ledger_to_le_names.keys()) - seen_ledgers_with_bu):
        rows.append({"Ledger Name": led, "Legal Entity": "", "Business Unit": ""})

    # 4) Orphan LEs (in master list) with no BU; back-fill ledger if uniquely known
    for le in sorted(legal_entity_names - seen_les_with_bu):
        if le in le_to_ledgers and len(le_to_ledgers[le]) == 1:
            led = next(iter(le_to_ledgers[le]))
        else:
            led = ""
        rows.append({"Ledger Name": led, "Legal Entity": le, "Business Unit": ""})

    df = pd.DataFrame(rows).drop_duplicates().reset_index(drop=True)

    # sort: group by Ledger (non-empty first), then LE, then BU
    df["__LedgerEmpty"] = (df["Ledger Name"] == "").astype(int)
    df = df.sort_values(["__LedgerEmpty","Ledger Name","Legal Entity","Business Unit"],
                        ascending=[True, True, True, True]).drop(columns="__LedgerEmpty").reset_index(drop=True)

    # assignment counter
    df.insert(0, "Assignment", range(1, len(df)+1))

    st.success(f"Built {len(df)} assignment rows.")
    st.dataframe(df, use_container_width=True, height=450)

    # Excel download
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Ledger_LE_BU_Assignments")

    st.download_button(
        "⬇️ Download Excel (EnterpriseStructure.xlsx)",
        data=excel_buf.getvalue(),
        file_name="EnterpriseStructure.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
# ======= DRAW.IO DIAGRAM (safe block; runs only after df exists) =======
if "df" in locals() and isinstance(df, pd.DataFrame) and not df.empty:
    import xml.etree.ElementTree as ET, zlib, base64, uuid

    def _make_drawio_xml(df: pd.DataFrame) -> str:
    import xml.etree.ElementTree as ET, uuid

    # ---------- 1) Build relationships ----------
    ledgers = [x for x in df["Ledger Name"].dropna().unique() if x]
    les     = [x for x in df["Legal Entity"].dropna().unique() if x]
    bus     = [x for x in df["Business Unit"].dropna().unique() if x]

    le_to_bus = {}
    led_to_le = {}

    for _, r in df.iterrows():
        led = str(r["Ledger Name"]).strip() if r["Ledger Name"] else ""
        le  = str(r["Legal Entity"]).strip() if r["Legal Entity"] else ""
        bu  = str(r["Business Unit"]).strip() if r["Business Unit"] else ""

        if le and bu:
            le_to_bus.setdefault(le, set()).add(bu)
        if led and le:
            led_to_le.setdefault(led, set()).add(le)

    # ---------- 2) Compute positions (center parents above children) ----------
    base_x, step, W, H = 40, 220, 180, 60
    Y_LEDGER, Y_LE, Y_BU = 40, 240, 440

    bu_x, le_x, led_x = {}, {}, {}
    next_slot = 0

    # Bottom row: BUs placed evenly
    for name in sorted(bus):
        bu_x[name] = base_x + next_slot * step
        next_slot += 1

    # Middle row: LEs centered above their BUs (avg of their children)
    for name in sorted(les):
        kids = sorted(le_to_bus.get(name, []))
        if kids:
            xs = [bu_x[k] for k in kids if k in bu_x]
            le_x[name] = int(sum(xs) / len(xs)) if xs else base_x + next_slot * step
        else:
            le_x[name] = base_x + next_slot * step
            next_slot += 1

    # Top row: Ledgers centered above their LEs (avg of their children)
    for name in sorted(ledgers):
        kids = sorted(led_to_le.get(name, []))
        if kids:
            xs = [le_x[k] for k in kids if k in le_x]
            led_x[name] = int(sum(xs) / len(xs)) if xs else base_x + next_slot * step
        else:
            led_x[name] = base_x + next_slot * step
            next_slot += 1

    # ---------- 3) Build mxfile ----------
    mxfile  = ET.Element("mxfile", attrib={"host": "app.diagrams.net"})
    diagram = ET.SubElement(mxfile, "diagram", attrib={"id": str(uuid.uuid4()), "name": "Enterprise Structure"})
    model   = ET.SubElement(diagram, "mxGraphModel")
    root    = ET.SubElement(model, "root")
    ET.SubElement(root, "mxCell", attrib={"id": "0"})
    ET.SubElement(root, "mxCell", attrib={"id": "1", "parent": "0"})

    # Styles — colors aligned to your sample (peach/orange for Ledgers, light yellow for LEs, soft green for BUs)
    S_LEDGER = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE6CC;strokeColor=#D79B00;fontSize=12;"
    S_LE     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFF2CC;strokeColor=#A68000;fontSize=12;"
    S_BU     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#D5E8D4;strokeColor=#82B366;fontSize=12;"
    # Upward orthogonal arrows, center-to-center
    S_EDGE   = "endArrow=block;rounded=1;edgeStyle=orthogonalEdgeStyle;exitX=0.5;exitY=0;entryX=0.5;entryY=1;"

    def add_vertex(label, style, x, y, w=W, h=H):
        cid = uuid.uuid4().hex[:8]
        v = ET.SubElement(root, "mxCell",
                          attrib={"id": cid, "value": label, "style": style, "vertex": "1", "parent": "1"})
        ET.SubElement(v, "mxGeometry",
                      attrib={"x": str(x), "y": str(y), "width": str(w), "height": str(h), "as": "geometry"})
        return cid

    def add_edge(src_id, tgt_id, label=""):
        eid = uuid.uuid4().hex[:8]
        e = ET.SubElement(root, "mxCell",
                          attrib={"id": eid, "value": label, "style": S_EDGE,
                                  "edge": "1", "parent": "1", "source": src_id, "target": tgt_id})
        ET.SubElement(e, "mxGeometry", attrib={"relative": "1", "as": "geometry"})

    id_map = {}

    # Place vertices (BUs bottom, LEs middle, Ledgers top)
    for name in bus:
        cid = add_vertex(name, S_BU, bu_x[name], Y_BU); id_map[("B", name)] = cid
    for name in les:
        cid = add_vertex(name, S_LE, le_x[name], Y_LE); id_map[("E", name)] = cid
    for name in ledgers:
        cid = add_vertex(name, S_LEDGER, led_x[name], Y_LEDGER); id_map[("L", name)] = cid

    # Edges flow upward: BU -> LE, LE -> Ledger
    added = set()
    for _, r in df.iterrows():
        led = str(r["Ledger Name"]).strip() if r["Ledger Name"] else ""
        le  = str(r["Legal Entity"]).strip() if r["Legal Entity"] else ""
        bu  = str(r["Business Unit"]).strip() if r["Business Unit"] else ""

        if le and bu and ("B", bu) in id_map and ("E", le) in id_map:
            k = ("B2E", bu, le)
            if k not in added:
                add_edge(id_map[("B", bu)], id_map[("E", le)]); added.add(k)

        if led and le and ("E", le) in id_map and ("L", led) in id_map:
            k = ("E2L", le, led)
            if k not in added:
                add_edge(id_map[("E", le)], id_map[("L", led)]); added.add(k)

    # ---------- 4) Legend (top-left) ----------
    def add_text(text, x, y):
        tid = uuid.uuid4().hex[:8]
        t = ET.SubElement(root, "mxCell",
                          attrib={"id": tid, "value": text,
                                  "style": "text;html=1;align=left;verticalAlign=middle;resizable=0;autosize=1;",
                                  "vertex": "1", "parent": "1"})
        ET.SubElement(t, "mxGeometry", attrib={"x": str(x), "y": str(y), "width": "80", "height": "20", "as": "geometry"})
        return tid

    # Legend frame
    LBOX_W, LBOX_H = 180, 120
    lid = add_vertex("Legend", "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFFFFF;strokeColor=#666666;fontSize=12;",
                     10, 10, LBOX_W, LBOX_H)
    # Swatches + labels
    sw_y = 40
    def legend_row(color_style, label, y):
        sw = add_vertex("", f"rounded=1;whiteSpace=wrap;html=1;fillColor={color_style};strokeColor=#666666;", 20, y, 28, 18)
        add_text(label, 56, y-1)

    legend_row("#FFE6CC", "Ledger",       sw_y)
    legend_row("#FFF2CC", "Legal Entity", sw_y+28)
    legend_row("#D5E8D4", "Business Unit",sw_y+56)

    return ET.tostring(mxfile, encoding="utf-8", method="xml").decode("utf-8")
