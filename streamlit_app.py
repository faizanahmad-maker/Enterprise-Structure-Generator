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
# ===================== DRAW.IO DIAGRAM BLOCK (paste below the Excel button) =====================
# ======= DRAW.IO DIAGRAM (bus-style org chart) =======
import xml.etree.ElementTree as ET, zlib, base64, uuid

def _make_drawio_xml_bus(df: pd.DataFrame) -> str:
    # --- constants / layout ---
    W, H = 180, 48
    X_STEP = 230
    Y_LEDGER, Y_LE, Y_BU = 80, 220, 360
    BUS_Y = (Y_LEDGER + Y_LE) / 2.0          # shared horizontal run for Ledger←LE edges
    PAD_GROUP = 60                            # extra spacing between ledgers
    RIGHT_PAD = 160                           # spacing before the "Unassigned" area

    # styles (match palette)
    S_LEDGER = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE6E6;strokeColor=#C86868;fontSize=12;"
    S_LE     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFE2C2;strokeColor=#A66000;fontSize=12;"
    S_BU     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFF1B3;strokeColor=#B38F00;fontSize=12;"

    # orthogonal edges; special bus variant for LE→Ledger
    S_EDGE_OTHER = (
        "endArrow=block;rounded=1;"
        "edgeStyle=orthogonalEdgeStyle;"
        "orthogonal=1;jettySize=auto;"
        "exitX=0.5;exitY=0;entryX=0.5;entryY=1;"
    )
    S_EDGE_LEDGER = (
        "endArrow=block;rounded=1;"
        "edgeStyle=orthogonalEdgeStyle;elbow=horizontal;"
        "orthogonal=1;jettySize=auto;"
        "exitX=0.5;exitY=0;entryX=0.5;entryY=1;"
    )

    # --- normalize input ---
    df = df[["Ledger Name", "Legal Entity", "Business Unit"]].copy()
    for c in df.columns:
        df[c] = df[c].fillna("").map(str).str.strip()

    # mappings (disambiguate LE by ledger)
    ledgers = sorted([x for x in df["Ledger Name"].unique() if x])
    # {ledger: sorted unique (ledger, le) keys}
    led_to_les = {}
    # {(ledger, le): sorted unique BUs}
    le_to_bus = {}

    for _, r in df.iterrows():
        L, E, B = r["Ledger Name"], r["Legal Entity"], r["Business Unit"]
        if L and E:
            key = (L, E)
            led_to_les.setdefault(L, set()).add(key)
        if E and B:
            # tie BU list to the (ledger, le) *if* L present so edge can be drawn; if not present,
            # we will still create unassigned nodes below
            if L:
                key = (L, E)
                le_to_bus.setdefault(key, set()).add(B)

    # unassigned LEs: rows with no Ledger but LE present
    unassigned_les = sorted(
        set(df.loc[(df["Ledger Name"] == "") & (df["Legal Entity"] != ""), "Legal Entity"].unique())
    )
    # unassigned BUs: rows with BU present but no Ledger/LE combo to attach
    # A BU is considered assigned if it appears in any row with both L and E.
    assigned_bus = set(df.loc[(df["Ledger Name"] != "") & (df["Legal Entity"] != "") & (df["Business Unit"] != ""), "Business Unit"])
    all_bus = set(df.loc[df["Business Unit"] != "", "Business Unit"])
    unassigned_bus = sorted(all_bus - assigned_bus)

    # coerce to lists (sorted within)
    led_to_les = {L: sorted(v, key=lambda k: k[1]) for L, v in led_to_les.items()}
    le_to_bus  = {k: sorted(v) for k, v in le_to_bus.items()}

    # --- compute x-positions (pack by ledger, then LEs, then BUs) ---
    # allocate horizontal slots; each BU occupies a slot; if no BUs, the LE takes one slot.
    next_x = 40
    led_x, le_x, bu_x = {}, {}, {}

    for L in ledgers:
        les = led_to_les.get(L, [])
        start_x = next_x
        # place BUs (leafs)
        for key in les:
            buses = le_to_bus.get(key, [])
            if buses:
                for b in buses:
                    bu_x[b] = next_x
                    next_x += X_STEP
            else:
                # LE with no BUs still needs space
                le_x[key] = next_x
                next_x += X_STEP
        # place LEs centered over their child BUs, or keep prior slot if no BUs
        for key in les:
            if key in le_x:
                continue
            buses = le_to_bus.get(key, [])
            if buses:
                xs = [bu_x[b] for b in buses]
                le_x[key] = int(sum(xs) / len(xs))
        # if a ledger had no LEs, give it one slot
        if not les:
            led_x[L] = next_x
            next_x += X_STEP + PAD_GROUP
        else:
            # center the ledger over all of its LEs (or their BUs)
            xs = []
            for key in les:
                # LE’s own x (maybe from “no BUs” case) or avg of children already set above
                xs.append(le_x[key])
            led_x[L] = int(sum(xs) / len(xs))
            next_x += PAD_GROUP

    # after all ledgers, add space then unassigned areas
    next_x += RIGHT_PAD
    for e in unassigned_les:
        le_x[("UNASSIGNED", e)] = next_x
        next_x += X_STEP
    next_x += PAD_GROUP
    for b in unassigned_bus:
        bu_x[b] = next_x
        next_x += X_STEP

    # --- build XML skeleton ---
    mxfile  = ET.Element("mxfile", attrib={"host": "app.diagrams.net"})
    diagram = ET.SubElement(mxfile, "diagram", attrib={"id": str(uuid.uuid4()), "name": "Enterprise Structure"})
    model   = ET.SubElement(diagram, "mxGraphModel", attrib={
        "dx": "1284", "dy": "682", "grid": "1", "gridSize": "10",
        "page": "1", "pageWidth": "1920", "pageHeight": "1080",
        "background": "#ffffff"   # clean white page
    })
    root    = ET.SubElement(model, "root")
    ET.SubElement(root, "mxCell", attrib={"id": "0"})
    ET.SubElement(root, "mxCell", attrib={"id": "1", "parent": "0"})

    def add_vertex(label, style, x, y, w=W, h=H):
        vid = uuid.uuid4().hex[:8]
        c = ET.SubElement(root, "mxCell", attrib={"id": vid, "value": label, "style": style, "vertex": "1", "parent": "1"})
        ET.SubElement(c, "mxGeometry", attrib={"x": str(int(x)), "y": str(int(y)), "width": str(w), "height": str(h), "as": "geometry"})
        return vid

    def add_edge(src, tgt, style=S_EDGE_OTHER, bus_y=None):
        eid = uuid.uuid4().hex[:8]
        c = ET.SubElement(root, "mxCell", attrib={
            "id": eid, "value": "", "style": style, "edge": "1", "parent": "1",
            "source": src, "target": tgt
        })
        g = ET.SubElement(c, "mxGeometry", attrib={"relative": "1", "as": "geometry"})
        if bus_y is not None:
            arr = ET.SubElement(g, "Array", attrib={"as": "points"})
            # one waypoint with only Y forces a shared horizontal run at that Y
            ET.SubElement(arr, "mxPoint", attrib={"y": str(int(bus_y))})

    # --- create vertices & ids ---
    id_map = {}  # ("L", L) | ("E", L, E) | ("B", B)

    for L in ledgers:
        id_map[("L", L)] = add_vertex(L, S_LEDGER, led_x[L], Y_LEDGER)

    for L in ledgers:
        for key in led_to_les.get(L, []):
            x = le_x[key]
            id_map[("E", L, key[1])] = add_vertex(key[1], S_LE, x, Y_LE)

    for b, x in bu_x.items():
        id_map[("B", b)] = add_vertex(b, S_BU, x, Y_BU)

    # unassigned LEs
    for e in unassigned_les:
        x = le_x[("UNASSIGNED", e)]
        id_map[("E", "UNASSIGNED", e)] = add_vertex(e, S_LE, x, Y_LE)

    # --- edges ---
    drawn = set()
    # BU -> LE (only when both present)
    for _, r in df.iterrows():
        L, E, B = r["Ledger Name"], r["Legal Entity"], r["Business Unit"]
        if B and E and L and (("B", B) in id_map) and (("E", L, E) in id_map):
            k = ("B2E", B, L, E)
            if k not in drawn:
                add_edge(id_map[("B", B)], id_map[("E", L, E)], style=S_EDGE_OTHER)  # standard orthogonal
                drawn.add(k)

    # LE -> Ledger (bus-style with shared BUS_Y)
    for _, r in df.iterrows():
        L, E = r["Ledger Name"], r["Legal Entity"]
        if L and E and (("E", L, E) in id_map) and (("L", L) in id_map):
            k = ("E2L", L, E)
            if k not in drawn:
                add_edge(id_map[("E", L, E)], id_map[("L", L)], style=S_EDGE_LEDGER, bus_y=BUS_Y)
                drawn.add(k)

    # --- legend (top-left) ---
    def add_legend(x=30, y=10):
        # background panel
        panel_w, panel_h = 210, 120
        p = add_vertex("", "rounded=1;fillColor=#FFFFFF;strokeColor=#CBD5E1;", x, y, panel_w, panel_h)
        # title (top aligned)
        title = ET.SubElement(root, "mxCell", attrib={
            "id": uuid.uuid4().hex[:8], "value": "Legend",
            "style": "text;verticalAlign=top;align=left;fontSize=13;fontStyle=1;",
            "vertex": "1", "parent": "1"
        })
        ET.SubElement(title, "mxGeometry", attrib={"x": str(x+12), "y": str(y+8), "width": "120", "height": "18", "as": "geometry"})
        # swatches
        def swatch(lbl, color, ty):
            box = ET.SubElement(root, "mxCell", attrib={
                "id": uuid.uuid4().hex[:8], "value": "",
                "style": f"rounded=1;fillColor={color};strokeColor=#666666;",
                "vertex": "1", "parent": "1"
            })
            gy = {"L": 36, "E": 62, "B": 88}[ty]
            ET.SubElement(box, "mxGeometry", attrib={"x": str(x+12), "y": str(y+gy), "width": "18", "height": "12", "as": "geometry"})
            txt = ET.SubElement(root, "mxCell", attrib={
                "id": uuid.uuid4().hex[:8], "value": lbl,
                "style": "text;align=left;verticalAlign=middle;fontSize=12;",
                "vertex": "1", "parent": "1"
            })
            ET.SubElement(txt, "mxGeometry", attrib={"x": str(x+36), "y": str(y+gy-4), "width": "130", "height": "20", "as": "geometry"})

        swatch("Ledger",       "#FFE6E6", "L")
        swatch("Legal Entity", "#FFE2C2", "E")
        swatch("Business Unit","#FFF1B3", "B")

    add_legend()

    return ET.tostring(mxfile, encoding="utf-8", method="xml").decode("utf-8")


def _drawio_url_from_xml(xml: str) -> str:
    # draw.io expects raw DEFLATE then base64
    raw = zlib.compress(xml.encode("utf-8"), level=9)[2:-4]
    b64 = base64.b64encode(raw).decode("ascii")
    return f"https://app.diagrams.net/?title=EnterpriseStructure.drawio#R{b64}"


# Only run if df exists and has rows
if isinstance(df, pd.DataFrame) and not df.empty:
    _xml = _make_drawio_xml_bus(df)

    st.download_button(
        "⬇️ Download diagram (.drawio)",
        data=_xml.encode("utf-8"),
        file_name="EnterpriseStructure.drawio",
        mime="application/xml",
        use_container_width=True
    )
    st.markdown(f"[🔗 Open in draw.io (preview)]({_drawio_url_from_xml(_xml)})")
    st.caption("Opens in diagrams.net; File → Save to persist.")
# ======= END DRAW.IO BLOCK =======
