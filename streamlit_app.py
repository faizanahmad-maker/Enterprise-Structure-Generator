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
# ===================== DRAW.IO DIAGRAM (grouped by ledger, keeps unassigned, curved edges) =====================
if "df" in locals() and isinstance(df, pd.DataFrame) and not df.empty:
    import xml.etree.ElementTree as ET
    import zlib, base64, uuid

    def _make_drawio_xml(df: pd.DataFrame) -> str:
        # ---------- 1) Build sets + relationships from the dataframe ----------
        ledgers = [x for x in df["Ledger Name"].dropna().unique() if x]
        les     = [x for x in df["Legal Entity"].dropna().unique() if x]
        bus     = [x for x in df["Business Unit"].dropna().unique() if x]

        led_to_le = {}   # Ledger -> set(LEs)
        le_to_bus = {}   # LE -> set(BUs)
        for _, r in df.iterrows():
            L  = (str(r["Ledger Name"]).strip()   if r["Ledger Name"]   else "")
            E  = (str(r["Legal Entity"]).strip()  if r["Legal Entity"]  else "")
            BU = (str(r["Business Unit"]).strip() if r["Business Unit"] else "")
            if L and E:    led_to_le.setdefault(L, set()).add(E)
            if E and BU:   le_to_bus.setdefault(E, set()).add(BU)

        # ---------- 2) Grouped layout (by ledger), then "unassigned" tails ----------
        base_x, step, GAP, W, H = 40, 220, 160, 180, 60
        Y_LEDGER, Y_LE, Y_BU = 40, 240, 440

        led_x, le_x, bu_x = {}, {}, {}
        cur_x = base_x

        def _center_to_left(cx):  # convert center X to left X
            return int(cx - W/2)

        # Place each ledger as a block: its LEs centered beneath, each LE’s BUs beneath it
        for L in sorted(ledgers):
            les_in_group = sorted(led_to_le.get(L, [])) or [None]  # None = ledger with no LEs

            # For each LE, slots = max(1, number of BUs)
            meta = []
            for E in les_in_group:
                bu_list = sorted(le_to_bus.get(E, [])) if E else []
                meta.append((E, bu_list, max(1, len(bu_list))))

            total_slots = sum(s for _, _, s in meta)
            group_start = cur_x
            cursor = group_start

            # BUs bottom, LEs middle
            for E, bu_list, slots in meta:
                if bu_list:
                    centers = [cursor + (i + 0.5) * step for i in range(len(bu_list))]
                    for i, BU in enumerate(bu_list):
                        bu_x[BU] = _center_to_left(centers[i])
                    if E:
                        le_x[E] = _center_to_left(sum(centers)/len(centers))
                else:
                    if E:  # LE with no BUs gets its own slot
                        le_x[E] = _center_to_left(cursor + 0.5 * step)
                cursor += slots * step

            group_width = total_slots * step if total_slots > 0 else step
            led_center = group_start + group_width / 2.0
            led_x[L] = _center_to_left(led_center)
            cur_x += group_width + GAP

        # Unassigned LEs (present in df but not positioned above)
        unassigned_les = [E for E in sorted(les) if E not in le_x]
        if unassigned_les:
            start = cur_x
            for i, E in enumerate(unassigned_les):
                le_x[E] = _center_to_left(start + (i + 0.5) * step)
            cur_x = start + len(unassigned_les) * step + GAP

        # Unassigned BUs (present in df but not positioned above)
        unassigned_bus = [B for B in sorted(bus) if B not in bu_x]
        if unassigned_bus:
            start = cur_x
            for i, B in enumerate(unassigned_bus):
                bu_x[B] = _center_to_left(start + (i + 0.5) * step)
            cur_x = start + len(unassigned_bus) * step + GAP

        # ---------- 3) Build mxfile ----------
        mxfile  = ET.Element("mxfile", attrib={"host": "app.diagrams.net"})
        diagram = ET.SubElement(mxfile, "diagram", attrib={"id": str(uuid.uuid4()), "name": "Enterprise Structure"})
        model   = ET.SubElement(diagram, "mxGraphModel")
        root    = ET.SubElement(model, "root")
        ET.SubElement(root, "mxCell", attrib={"id": "0"})
        ET.SubElement(root, "mxCell", attrib={"id": "1", "parent": "0"})

        # Palette (match request): Ledger=Red, LE=Orange, BU=Yellow
        S_LEDGER = "rounded=1;whiteSpace=wrap;html=1;fillColor=#F8CECC;strokeColor=#B85450;fontSize=12;"
        S_LE     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FAD7AC;strokeColor=#D79B00;fontSize=12;"
        S_BU     = "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFF2CC;strokeColor=#D6B656;fontSize=12;"

        # Curved center-to-center edges; routing disabled to avoid “bus” lines
        S_EDGE   = (
            "endArrow=block;rounded=1;"
            "noEdgeStyle=1;orthogonal=0;curved=1;jettySize=0;"
            "exitX=0.5;exitY=0;entryX=0.5;entryY=1;"
        )

        def add_vertex(label, style, x, y, w=W, h=H):
            cid = uuid.uuid4().hex[:8]
            v = ET.SubElement(root, "mxCell",
                              attrib={"id": cid, "value": label, "style": style, "vertex": "1", "parent": "1"})
            ET.SubElement(v, "mxGeometry",
                          attrib={"x": str(int(x)), "y": str(int(y)), "width": str(w), "height": str(h), "as": "geometry"})
            return cid

        def add_edge(src_id, tgt_id):
            eid = uuid.uuid4().hex[:8]
            e = ET.SubElement(root, "mxCell",
                              attrib={"id": eid, "value": "", "style": S_EDGE, "edge": "1", "parent": "1",
                                      "source": src_id, "target": tgt_id})
            ET.SubElement(e, "mxGeometry", attrib={"relative": "1", "as": "geometry"})

        # Place vertices
        id_map = {}
        for B, x in bu_x.items(): id_map[("B", B)] = add_vertex(B, S_BU, x, Y_BU)
        for E, x in le_x.items(): id_map[("E", E)] = add_vertex(E, S_LE, x, Y_LE)
        for L, x in led_x.items(): id_map[("L", L)] = add_vertex(L, S_LEDGER, x, Y_LEDGER)

        # Draw edges (curved, upward)
        drawn = set()
        for _, r in df.iterrows():
            L  = (str(r["Ledger Name"]).strip()   if r["Ledger Name"]   else "")
            E  = (str(r["Legal Entity"]).strip()  if r["Legal Entity"]  else "")
            B  = (str(r["Business Unit"]).strip() if r["Business Unit"] else "")

            if B and E and ("B", B) in id_map and ("E", E) in id_map:
                k = ("B2E", B, E)
                if k not in drawn: add_edge(id_map[("B", B)], id_map[("E", E)]); drawn.add(k)

            if E and L and ("E", E) in id_map and ("L", L) in id_map:
                k = ("E2L", E, L)
                if k not in drawn: add_edge(id_map[("E", E)], id_map[("L", L)]); drawn.add(k)

        # ---------- 4) Legend ----------
        def add_text(text, x, y):
            tid = uuid.uuid4().hex[:8]
            t = ET.SubElement(root, "mxCell",
                              attrib={"id": tid, "value": text,
                                      "style": "text;html=1;align=left;verticalAlign=middle;resizable=0;autosize=1;",
                                      "vertex": "1", "parent": "1"})
            ET.SubElement(t, "mxGeometry", attrib={"x": str(x), "y": str(y),
                                                   "width": "90", "height": "20", "as": "geometry"})
            return tid

        add_vertex("Legend", "rounded=1;whiteSpace=wrap;html=1;fillColor=#FFFFFF;strokeColor=#666666;fontSize=12;",
                   10, 10, 180, 120)
        def legend_row(fill_hex, label, y):
            add_vertex("", f"rounded=1;whiteSpace=wrap;html=1;fillColor={fill_hex};strokeColor=#666666;",
                       20, y, 28, 18)
            add_text(label, 56, y - 1)

        legend_row("#F8CECC", "Ledger",        40)
        legend_row("#FAD7AC", "Legal Entity",  68)
        legend_row("#FFF2CC", "Business Unit", 96)

        return ET.tostring(mxfile, encoding="utf-8", method="xml").decode("utf-8")

    def _drawio_url_from_xml(xml: str) -> str:
        raw = zlib.compress(xml.encode("utf-8"), level=9)[2:-4]  # raw DEFLATE
        b64 = base64.b64encode(raw).decode("ascii")
        return f"https://app.diagrams.net/?title=EnterpriseStructure.drawio#R{b64}"

    _xml = _make_drawio_xml(df)
    st.download_button("⬇️ Download diagram (.drawio)", _xml.encode("utf-8"),
                       file_name="EnterpriseStructure.drawio", mime="application/xml")
    st.markdown(f"[🔗 Open in draw.io (preview)]({_drawio_url_from_xml(_xml)})")
    st.caption("Grouped by Ledger • curved center-to-center arrows • all unassigned LEs/BUs shown at right • Legend included.")
# =================== END DRAW.IO DIAGRAM ===================
