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
# ---- draw.io (diagrams.net) one-click diagram + .drawio download ----
import xml.etree.ElementTree as ET, zlib, base64, uuid

def make_drawio_xml(df: pd.DataFrame) -> str:
    # Collect unique nodes by layer
    ledgers = [x for x in df["Ledger Name"].dropna().unique() if x]
    les     = [x for x in df["Legal Entity"].dropna().unique() if x]
    bus     = [x for x in df["Business Unit"].dropna().unique() if x]

    # Assign positions (simple 3-row layout)
    def layer_positions(items, y, x_step=220, w=180, h=60):
        pos = {}
        for i, name in enumerate(items):
            pos[name] = {"x": 40 + i * x_step, "y": y, "w": w, "h": h}
        return pos

    pos_ledger = layer_positions(ledgers, y=40)
    pos_le     = layer_positions(les,     y=240)
    pos_bu     = layer_positions(bus,     y=440)

    # Prepare mxfile skeleton
    mxfile = ET.Element("mxfile", attrib={"host": "app.diagrams.net"})
    diagram = ET.SubElement(mxfile, "diagram", attrib={"id": str(uuid.uuid4()), "name": "Enterprise Structure"})
    model = ET.SubElement(diagram, "mxGraphModel")
    root = ET.SubElement(model, "root")
    ET.SubElement(root, "mxCell", id="0")
    ET.SubElement(root, "mxCell", id="1", parent="0")

    def add_vertex(cell_id, label, style, geom):
        c = ET.SubElement(root, "mxCell", id=cell_id, value=label, style=style, vertex="1", parent="1")
        g = ET.SubElement(c, "mxGeometry", as_="geometry")
        g.set("x", str(geom["x"])); g.set("y", str(geom["y"]))
        g.set("width", str(geom["w"])); g.set("height", str(geom["h"]))
        return c

    def add_edge(edge_id, src, tgt, label=""):
        c = ET.SubElement(root, "mxCell", id=edge_id, value=label,
                          style="endArrow=block;edgeStyle=elbowEdgeStyle;rounded=1;", edge="1", parent="1", source=src, target=tgt)
        ET.SubElement(c, "mxGeometry", relative="1", as_="geometry")

    # Node styles
    S_LEDGER = "rounded=1;whiteSpace=wrap;html=1;fillColor=#e3f2fd;strokeColor=#1565c0;fontSize=12;"
    S_LE     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#fff3e0;strokeColor=#ef6c00;fontSize=12;"
    S_BU     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#e8f5e9;strokeColor=#2e7d32;fontSize=12;"

    # Build vertex ids
    id_map = {}
    for name, geom in pos_ledger.items():
        vid = f"L::{uuid.uuid4().hex[:8]}"
        id_map(("L", name)) = vid
        add_vertex(vid, name, S_LEDGER, geom)
    for name, geom in pos_le.items():
        vid = f"E::{uuid.uuid4().hex[:8]}"
        id_map(("E", name)) = vid
        add_vertex(vid, name, S_LE, geom)
    for name, geom in pos_bu.items():
        vid = f"B::{uuid.uuid4().hex[:8]}"
        id_map(("B", name)) = vid
        add_vertex(vid, name, S_BU, geom)

    # Edges: Ledger→LE and LE→BU
    added = set()
    for _, r in df.iterrows():
        led, le, bu = r["Ledger Name"], r["Legal Entity"], r["Business Unit"]
        if led and le:
            key = ("L2E", led, le)
            if key not in added and ("L", led) in id_map and ("E", le) in id_map:
                add_edge(f"e::{uuid.uuid4().hex[:8]}", id_map[("L", led)], id_map[("E", le)], "")
                added.add(key)
        if le and bu:
            key = ("E2B", le, bu)
            if key not in added and ("E", le) in id_map and ("B", bu) in id_map:
                add_edge(f"e::{uuid.uuid4().hex[:8]}", id_map[("E", le)], id_map[("B", bu)], "")
                added.add(key)

    # Serialize
    xml = ET.tostring(mxfile, encoding="utf-8", method="xml").decode("utf-8")
    return xml

def drawio_url_from_xml(xml: str) -> str:
    # draw.io expects raw DEFLATE (no zlib header) then base64
    payload = zlib.compress(xml.encode("utf-8"), level=9)[2:-4]  # strip zlib header & checksum
    b64 = base64.b64encode(payload).decode("ascii")
    return f"https://app.diagrams.net/?title=EnterpriseStructure.drawio#R{b64}"

# Build XML + offer downloads/links
xml_str = make_drawio_xml(df)

# Downloadable .drawio file
st.download_button("⬇️ Download diagram (.drawio)", data=xml_str.encode("utf-8"),
                   file_name="EnterpriseStructure.drawio", mime="application/xml")

# One-click "Open in draw.io" link
drawio_link = drawio_url_from_xml(xml_str)
st.markdown(f"[🔗 Open in draw.io (preview)]({drawio_link})")
st.caption("This opens the diagram in diagrams.net; click **File → Save** to persist to Drive/GitHub/Desktop.")
