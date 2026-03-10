"""
WMS Analyzer – ren analyslogik utan GUI-beroenden.
Extraherad från wms_sök79.py.
"""
from __future__ import annotations

import os
from typing import Dict, List, Optional, Tuple

import pandas as pd


class WMSAnalyzer:
    """WMS-analysmotor: klassificerar pallar och spårar buffertuppdateringar."""

    def __init__(self, data_path: str) -> None:
        self.data_path = data_path

        try:
            self.receive_file = self._find_file_prefix("v_ask_receive_log")
        except FileNotFoundError:
            self.receive_file = None
        try:
            self.booking_file = self._find_file_prefix("v_ask_booking_putaway")
        except FileNotFoundError:
            self.booking_file = None
        try:
            self.buffert_file = self._find_file_prefix("v_ask_article_buffertpallet")
        except FileNotFoundError:
            self.buffert_file = None
        try:
            self.trans_file = self._find_file_prefix("v_ask_trans_log")
        except FileNotFoundError:
            self.trans_file = None
        try:
            self.pick_file = self._find_file_prefix("v_ask_pick_log_full")
        except FileNotFoundError:
            self.pick_file = None
        try:
            self.correct_file = self._find_file_prefix("v_ask_correct_log")
        except FileNotFoundError:
            self.correct_file = None

        self.receive_df = self._load_csv(self.receive_file)
        self.booking_df = self._load_csv(self.booking_file)
        self.buffert_df = self._load_csv(self.buffert_file)
        self.trans_df = self._load_csv(self.trans_file)
        self.pick_df = self._load_csv(self.pick_file)
        self.correct_df = self._load_csv(self.correct_file)

        self._prepare_dataframes()

    def _prepare_dataframes(self) -> None:
        for df, cols in [
            (self.receive_df, ["Inköpsnr", "Artikel", "Pallid", "Mottaget"]),
            (self.booking_df, ["Pall nr", "Inköpsnr"]),
            (self.buffert_df, ["Pallid", "Lagerplats"]),
            (self.trans_df, ["Pallid", "Till", "Timestamp", "Från"]),
            (self.pick_df, ["Pallid", "Artikelnr", "Plockat", "Ordernr"]),
            (self.correct_df, ["Pallid", "Antal", "Anledning", "Artikel"]),
        ]:
            for col in cols:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.strip()

        if "Mottaget" in self.receive_df.columns:
            self.receive_df["Mottaget_num"] = pd.to_numeric(
                self.receive_df["Mottaget"].str.replace(",", "."), errors="coerce"
            )
        if "Plockat" in self.pick_df.columns:
            self.pick_df["Plockat_num"] = pd.to_numeric(
                self.pick_df["Plockat"].str.replace(",", "."), errors="coerce"
            )
        if not self.correct_df.empty and "Antal" in self.correct_df.columns:
            self.correct_df["Antal_num"] = pd.to_numeric(
                self.correct_df["Antal"].str.replace(",", "."), errors="coerce"
            )

    def _find_file_prefix(self, prefix: str) -> str:
        for fname in os.listdir(self.data_path):
            if fname.startswith(prefix) and fname.endswith(".csv"):
                return os.path.join(self.data_path, fname)
        raise FileNotFoundError(f"Ingen fil börjar med {prefix!r} i {self.data_path}")

    def _load_csv(self, filename: Optional[str]) -> pd.DataFrame:
        if not filename or not os.path.exists(filename):
            return pd.DataFrame()
        try:
            return pd.read_csv(filename, sep="\t", encoding="utf-8", dtype=str)
        except Exception:
            try:
                return pd.read_csv(filename, dtype=str, sep=None, engine="python")
            except Exception:
                return pd.DataFrame()

    def _is_putaway(self, pallid: str) -> bool:
        if self.booking_df.empty or "Pall nr" not in self.booking_df.columns:
            return False
        try:
            return not self.booking_df[self.booking_df["Pall nr"] == pallid].empty
        except KeyError:
            return False

    def _get_buffert_location(self, pallid: str) -> Optional[str]:
        rows = self.buffert_df[self.buffert_df["Pallid"] == pallid]
        if not rows.empty:
            return rows.iloc[0]["Lagerplats"]
        return None

    def _get_latest_trans_destination(self, pallid: str) -> Optional[str]:
        rows = self.trans_df[self.trans_df["Pallid"] == pallid]
        if rows.empty:
            return None
        if "Timestamp" in rows.columns:
            try:
                tmp = pd.to_datetime(rows["Timestamp"], errors="coerce")
                rows = rows.assign(_ts=tmp)
                rows = rows.sort_values(by="_ts", ascending=False, na_position="last")
            except Exception:
                pass
        return rows.iloc[0]["Till"] if "Till" in rows.columns else None

    def _find_new_pallids_from_buffert_update(self, plock_loc: str, exclude: List[str]) -> List[str]:
        if "Från" not in self.trans_df.columns or "Till" not in self.trans_df.columns:
            return []
        loc_upper = plock_loc.upper() if plock_loc else ""
        rows = self.trans_df[
            (self.trans_df["Från"].str.upper() == loc_upper)
            & (self.trans_df["Till"].str.upper() == "CRANE")
            & (~self.trans_df["Pallid"].isin(exclude))
        ].copy()
        if rows.empty:
            return []
        if "Timestamp" in rows.columns:
            try:
                ts = pd.to_datetime(rows["Timestamp"], errors="coerce")
                rows = rows.assign(_ts=ts)
                rows = rows.sort_values("_ts", ascending=True, na_position="last")
            except Exception:
                pass
        return rows["Pallid"].dropna().unique().tolist()

    def _classify_location(self, pallid: str) -> Tuple[str, Optional[str]]:
        if self._is_putaway(pallid):
            return ("Ej inlagrad", None)

        loc = self._get_buffert_location(pallid)
        if loc:
            loc_upper = loc.upper()
            if loc_upper.startswith("UTE"):
                return ("Ej inlagrad", loc)
            if loc_upper.startswith("B") or loc_upper.startswith("L") or loc_upper.startswith("AA"):
                if not (loc_upper.startswith("AA75") or loc_upper.startswith("AA76")):
                    return ("Plockplats", loc)
            return ("Inlagrad", loc)

        latest_dest = self._get_latest_trans_destination(pallid)
        if latest_dest:
            dest_upper = latest_dest.upper()
            if dest_upper.startswith("TO") or dest_upper.startswith("PR"):
                return ("Skickad", latest_dest)
            if dest_upper.startswith("UTE"):
                return ("Ej inlagrad", latest_dest)
            if dest_upper.startswith("B") or dest_upper.startswith("L") or dest_upper.startswith("AA"):
                if not (dest_upper.startswith("AA75") or dest_upper.startswith("AA76")):
                    return ("Plockplats", latest_dest)
            return ("Inlagrad", latest_dest)
        return ("Okänt", None)

    def _get_event_timestamps(self, pallid: str) -> Dict[str, str]:
        ts: Dict[str, str] = {}
        if not self.receive_df.empty and "Ändrad" in self.receive_df.columns:
            rcv = self.receive_df[self.receive_df["Pallid"] == pallid]
            if not rcv.empty:
                ts["mottag"] = rcv["Ändrad"].iloc[0]
        if not self.booking_df.empty and "Ändrad" in self.booking_df.columns:
            put = self.booking_df[self.booking_df["Pall nr"] == pallid]
            if not put.empty:
                ts["putaway"] = put["Ändrad"].iloc[0]
        if not self.buffert_df.empty and "Datum/tid" in self.buffert_df.columns:
            buf = self.buffert_df[self.buffert_df["Pallid"] == pallid]
            if not buf.empty:
                ts["buffert"] = buf["Datum/tid"].iloc[0]
        if not self.pick_df.empty and "Datum" in self.pick_df.columns:
            pk = self.pick_df[self.pick_df["Pallid"] == pallid]
            if not pk.empty:
                ts["pick"] = pk["Datum"].iloc[0]
        if not self.correct_df.empty and "Ändrad" in self.correct_df.columns:
            co = self.correct_df[self.correct_df["Pallid"] == pallid]
            if not co.empty:
                ts["saldojustering"] = co["Ändrad"].iloc[0]
        if not self.trans_df.empty and "Timestamp" in self.trans_df.columns:
            tr = self.trans_df[self.trans_df["Pallid"] == pallid]
            if not tr.empty:
                ts["trans"] = tr["Timestamp"].iloc[0]
        return ts

    def analyze(self, purchase_number: str, article_number: str) -> str:
        purchase_number = str(purchase_number).strip()
        article_number = str(article_number).strip()

        recv_matches = self.receive_df[
            (self.receive_df["Inköpsnr"] == purchase_number)
            & (self.receive_df["Artikel"] == article_number)
        ]
        if recv_matches.empty:
            return (
                f"Inga mottagna pallar hittades för inköpsnr '{purchase_number}' "
                f"och artikelnr '{article_number}'."
            )

        initial_pall_ids = recv_matches["Pallid"].dropna().unique().tolist()
        pall_ids = list(initial_pall_ids)
        checked: set = set(initial_pall_ids)
        i = 0
        while i < len(pall_ids):
            pid = pall_ids[i]
            status, loc = self._classify_location(pid)
            if status == "Plockplats" and loc:
                new_ids = self._find_new_pallids_from_buffert_update(loc, list(checked))
                for nid in new_ids:
                    if nid not in checked:
                        pall_ids.append(nid)
                        checked.add(nid)
            i += 1

        total_received = recv_matches["Mottaget_num"].sum(skipna=True)
        pick_matches = self.pick_df[
            (self.pick_df["Pallid"].isin(pall_ids))
            & (self.pick_df["Artikelnr"] == article_number)
        ]
        shipment_totals = pick_matches.groupby("Ordernr")["Plockat_num"].sum(min_count=1)
        total_shipped = shipment_totals.sum()

        pall_corrections: Dict[str, float] = {}
        if not self.correct_df.empty:
            corr_pall = self.correct_df[
                (~self.correct_df["Pallid"].str.lower().isin(["nan", "", "none"]))
                & (self.correct_df["Pallid"].isin(pall_ids))
            ]
            if not corr_pall.empty:
                pall_corrections = corr_pall.groupby("Pallid")["Antal_num"].sum(min_count=1).to_dict()
            corr_general = self.correct_df[self.correct_df["Anledning"] == purchase_number]
            general_adjustment = corr_general["Antal_num"].sum(skipna=True) if not corr_general.empty else 0.0
        else:
            general_adjustment = 0.0

        saldo_lines: List[str] = []
        for pallid in pall_ids:
            recv_qty = recv_matches[recv_matches["Pallid"] == pallid]["Mottaget_num"].sum(skipna=True)
            picked_qty = pick_matches[pick_matches["Pallid"] == pallid]["Plockat_num"].sum(skipna=True)
            corr_qty = pall_corrections.get(pallid, 0.0)
            remaining = None
            if pd.notnull(recv_qty) and pd.notnull(picked_qty) and pd.notnull(corr_qty):
                remaining = recv_qty - picked_qty + corr_qty
            status, loc = self._classify_location(pallid)
            loc_str = loc if loc is not None else ""
            corr_info = f", justerat: {corr_qty}" if corr_qty != 0 else ""
            origin_label = "inköp" if pallid in initial_pall_ids else "buffert"

            if origin_label == "buffert" and remaining is not None and remaining < 0:
                ship_qty_val = None
                if pd.notnull(picked_qty) and pd.notnull(recv_qty) and pd.notnull(corr_qty):
                    ship_qty_val = float(picked_qty - recv_qty + corr_qty)
                buffer_loc = None
                try:
                    buf_rows = self.trans_df[
                        (self.trans_df["Pallid"] == pallid)
                        & (self.trans_df["Till"].str.upper() == "CRANE")
                    ]
                    if not buf_rows.empty and "Från" in buf_rows.columns:
                        buffer_loc = buf_rows.iloc[0]["Från"]
                except Exception:
                    buffer_loc = None
                buffer_loc_str = buffer_loc if buffer_loc is not None else ""
                if ship_qty_val is not None:
                    line1 = (
                        f"   - PallID {pallid} [{origin_label}] : Buffertuppdatering (plats: {buffer_loc_str})"
                        f" - antal: {ship_qty_val}{corr_info}"
                    )
                else:
                    line1 = (
                        f"   - PallID {pallid} [{origin_label}] : Buffertuppdatering (plats: {buffer_loc_str})"
                        f" - antal: {picked_qty}{corr_info}"
                    )
                line2 = (
                    f"   - PallID {pallid} [{origin_label}] : Skickad (plats: {loc_str})"
                    f" - kvar: 0.0{corr_info}"
                )
                ev = self._get_event_timestamps(pallid)
                if ev:
                    ev_parts = [f"{k}: {v}" for k, v in ev.items()]
                    line2 += "\n      • " + " | ".join(ev_parts)
                saldo_lines.append(line1)
                saldo_lines.append(line2)
            else:
                line = (
                    f"   - PallID {pallid} [{origin_label}] : {status} (plats: {loc_str})"
                    f" - kvar: {remaining}{corr_info}"
                )
                ev = self._get_event_timestamps(pallid)
                if ev:
                    ev_parts = [f"{k}: {v}" for k, v in ev.items()]
                    line += "\n      • " + " | ".join(ev_parts)
                saldo_lines.append(line)

        report_lines: List[str] = []
        report_lines.append(f"Analys för Inköpsnr: {purchase_number}, Artikelnr: {article_number}")
        report_lines.append(
            "\nFlödesöversikt:\n---------------\n"
            "1. **Mottagna pallar:**\n"
            f"   - Antal mottagningsrader: {len(recv_matches)}\n"
            f"   - Unika pallID: {len(pall_ids)}\n"
            f"   - Summa mottaget antal: {total_received}\n"
            "2. **Levererat till kund:**\n"
            f"   - Summa plockat antal: {total_shipped}\n"
            "   - Orderfördelning:\n"
        )

        if shipment_totals.empty:
            report_lines.append("   (Inga leveranser hittades i plockloggen för de matchande pallarna.)")
        else:
            for order, qty in shipment_totals.sort_index().items():
                report_lines.append(f"   - Order {order}: {qty}")

        total_pall_correction = sum(pall_corrections.values()) if pall_corrections else 0.0
        total_correction = (
            (total_pall_correction + general_adjustment)
            if (pall_corrections or general_adjustment)
            else 0.0
        )
        remaining_total = None
        if pd.notna(total_received) and pd.notna(total_shipped):
            try:
                remaining_total = total_received - total_shipped + total_correction
            except Exception:
                remaining_total = None
        if remaining_total is not None:
            report_lines.append(f"\n3. **Saldo kvar totalt:**\n   - Kvar att leverera: {remaining_total}")

        report_lines.append("\n4. **Saldo per pallID:**")
        report_lines.extend(saldo_lines)

        if total_correction != 0.0 or not pd.isna(general_adjustment):
            report_lines.append("\n5. **Saldojusteringar:**")
            if general_adjustment:
                report_lines.append(
                    f"   - Justeringar på inköpsnr {purchase_number}: {general_adjustment}"
                )
            for pid, adj in pall_corrections.items():
                report_lines.append(f"   - Justeringar på pallID {pid}: {adj}")

        return "\n".join(report_lines)
