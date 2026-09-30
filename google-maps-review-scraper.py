"""
Google Maps Review Scraper - VERSI DIPERBAIKI (v3)

Install:
    pip install selenium pandas openpyxl
    (webdriver-manager opsional; Selenium >= 4.6 sudah punya driver manager sendiri)

Contoh pemakaian:
    python scraping_data_ulasan_google_ulasan.py "https://maps.app.goo.gl/xxxx"
    python scraping_data_ulasan_google_ulasan.py --urls-file cabang.txt --headless
    python scraping_data_ulasan_google_ulasan.py URL --max-reviews 2000 --sort newest
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import (
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# --------------------------------------------------------------------------- #
# Konfigurasi
# --------------------------------------------------------------------------- #

# Isi di sini jika tidak ingin memakai argumen CLI (satu URL per cabang).
DEFAULT_URLS: list[str] = [
    "https://www.google.com/maps/place/C28+Social+Space+-+Malioboro/@-7.7913623,110.3614189,17z/data=!3m1!4b1!4m6!3m5!1s0x2e7a5926aec9cfa1:0xa10c2a7ead083ddb!8m2!3d-7.7913623!4d110.3614189!16s%2Fg%2F11vpmgmvrx?entry=ttu&g_ep=EgoyMDI2MDkyMy4wIKXMDSoASAFQAw%3D%3D",
]

REVIEW_CARD = "div.jftiEf"  # kartu satu ulasan
CHUNK_SIZE = 500            # jumlah ulasan yang diekstrak per panggilan JS

SORT_OPTIONS = {
    # nama: (index menu, regex label ID/EN)
    "relevant": (0, r"paling relevan|most relevant"),
    "newest": (1, r"terbaru|newest"),
    "highest": (2, r"tertinggi|highest"),
    "lowest": (3, r"terendah|lowest"),
}

log = logging.getLogger("gmaps")

# --------------------------------------------------------------------------- #
# JavaScript helper (dijalankan di dalam browser)
# --------------------------------------------------------------------------- #

# Cari ancestor yang bisa di-scroll dari kartu ulasan pertama.
FIND_PANEL_JS = """
let el = document.querySelector(arguments[0]);
while (el && el !== document.body) {
  const st = getComputedStyle(el);
  if (/(auto|scroll)/.test(st.overflowY) && el.scrollHeight > el.clientHeight) return el;
  el = el.parentElement;
}
return null;
"""

# Klik semua tombol "Selengkapnya" / "More" dan kembalikan jumlahnya.
EXPAND_JS = """
const btns = document.querySelectorAll(
  'button.w8nwRe, button[aria-label="Lihat ulasan lengkap"], button[aria-label="See more"]'
);
btns.forEach(b => { try { b.click(); } catch (e) {} });
return btns.length;
"""

# Ekstrak satu batch ulasan sekaligus (1 round-trip untuk ratusan ulasan).
EXTRACT_JS = """
const [start, size, sel] = arguments;
const cards = Array.from(document.querySelectorAll(sel)).slice(start, start + size);
const txt = (root, s) => { const n = root.querySelector(s); return n ? n.innerText.trim() : ''; };
return cards.map(c => {
  const ratingEl = c.querySelector('span.kvMYJc');
  const reviewEl = c.querySelector('.MyEned .wiI7pd') || c.querySelector('span.wiI7pd');
  const ownerEl  = c.querySelector('.CDe7pd .wiI7pd') || c.querySelector('.CDe7pd');
  return {
    review_id:    c.getAttribute('data-review-id') || '',
    reviewer:     txt(c, 'div.d4r55'),
    reviewer_meta: txt(c, 'div.RfnDt'),
    rating_label: ratingEl ? (ratingEl.getAttribute('aria-label') || '') : '',
    date_raw:     txt(c, 'span.rsqaWe'),
    text:         reviewEl ? reviewEl.innerText.trim() : '',
    owner_reply:  ownerEl ? ownerEl.innerText.trim() : '',
    photos:       c.querySelectorAll('button.Tya61d, button[jsaction*="photo"]').length,
    likes:        txt(c, 'span.pkWtMe'),
  };
});
"""

# --------------------------------------------------------------------------- #
# Fungsi utilitas
# --------------------------------------------------------------------------- #

_CTRL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

_UNITS = {
    "detik": 1 / 86400, "second": 1 / 86400,
    "menit": 1 / 1440, "minute": 1 / 1440,
    "jam": 1 / 24, "hour": 1 / 24,
    "hari": 1, "day": 1,
    "minggu": 7, "week": 7,
    "bulan": 30, "month": 30,
    "tahun": 365, "year": 365,
}
_UNIT_RE = "|".join(_UNITS)


def clean_text(value: str) -> str:
    """Hapus karakter kontrol yang menyebabkan error saat export ke Excel."""
    return _CTRL_CHARS.sub("", value or "").strip()


def parse_int(text: str) -> Optional[int]:
    digits = re.sub(r"\D", "", text or "")
    return int(digits) if digits else None


def parse_rating(label: str) -> Optional[float]:
    """'5 bintang' / 'Rated 4.0 out of 5' -> angka."""
    m = re.search(r"(\d+(?:[.,]\d+)?)", label or "")
    return float(m.group(1).replace(",", ".")) if m else None


def approx_date(relative: str, now: datetime) -> Optional[str]:
    """
    Konversi '2 bulan lalu' / 'seminggu lalu' / 'a year ago' menjadi perkiraan tanggal.
    Hasilnya hanya perkiraan (1 bulan = 30 hari, 1 tahun = 365 hari).
    """
    s = (relative or "").lower()
    if not s:
        return None
    m = re.search(rf"(\d+)\s*({_UNIT_RE})", s)
    if m:
        n, unit = int(m.group(1)), m.group(2)
    else:
        m = re.search(rf"\bse\s*({_UNIT_RE})", s) or re.search(rf"\b(?:a|an)\s+({_UNIT_RE})", s)
        if not m:
            return None
        n, unit = 1, m.group(1)
    return (now - timedelta(days=n * _UNITS[unit])).date().isoformat()


def slugify(text: str) -> str:
    slug = re.sub(r"[^A-Za-z0-9]+", "_", text or "").strip("_")
    return slug[:60] or "tempat"


# --------------------------------------------------------------------------- #
# Scraper
# --------------------------------------------------------------------------- #

class GoogleMapsReviewScraper:
    def __init__(self, headless: bool = False, block_images: bool = False):
        self.driver = self._build_driver(headless, block_images)
        self.wait = WebDriverWait(self.driver, 20)

    # ---- driver ----------------------------------------------------------- #
    @staticmethod
    def _build_driver(headless: bool, block_images: bool) -> webdriver.Chrome:
        options = webdriver.ChromeOptions()
        options.add_argument("--lang=id")
        options.add_argument("--window-size=1400,1000")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--no-sandbox")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        if headless:
            options.add_argument("--headless=new")
        if block_images:
            options.add_experimental_option(
                "prefs", {"profile.managed_default_content_settings.images": 2}
            )
        try:
            # Selenium >= 4.6 mengunduh driver yang cocok secara otomatis.
            return webdriver.Chrome(options=options)
        except WebDriverException as exc:
            log.warning("Selenium Manager gagal (%s). Mencoba webdriver-manager...", exc.msg)
            from webdriver_manager.chrome import ChromeDriverManager

            return webdriver.Chrome(
                service=Service(ChromeDriverManager().install()), options=options
            )

    def close(self) -> None:
        try:
            self.driver.quit()
        except Exception:
            pass

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()

    # ---- navigasi --------------------------------------------------------- #
    def _accept_consent(self) -> None:
        """Tutup halaman persetujuan cookie Google jika muncul."""
        if "consent.google" not in self.driver.current_url:
            return
        xpath = (
            "//button[.//span[contains(text(),'Terima semua') or contains(text(),'Accept all')]"
            " or contains(@aria-label,'Terima semua') or contains(@aria-label,'Accept all')]"
        )
        try:
            self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
            time.sleep(2)
        except TimeoutException:
            log.warning("Halaman consent terdeteksi tapi tombol tidak ditemukan.")

    def _safe_text(self, css: str) -> str:
        try:
            return self.driver.find_element(By.CSS_SELECTOR, css).text.strip()
        except Exception:
            return ""

    def get_place_info(self) -> dict:
        info = {"nama_tempat": "N/A", "rating_keseluruhan": "N/A",
                "total_ulasan": None, "kategori": "N/A"}
        info["nama_tempat"] = self._safe_text("h1.DUwDvf") or "N/A"
        info["kategori"] = self._safe_text("button.DkEaL") or "N/A"

        summary = self._safe_text("div.F7nice")          # contoh: "4,5\n(1.234)"
        m = re.search(r"(\d[.,]\d)", summary)
        if m:
            info["rating_keseluruhan"] = m.group(1)
        m = re.search(r"\(([\d.,]+)\)", summary)
        if m:
            info["total_ulasan"] = parse_int(m.group(1))
        return info

    def open_reviews_tab(self) -> None:
        xpaths = [
            '//button[@role="tab"][contains(@aria-label,"Ulasan") or contains(@aria-label,"Reviews")]',
            '//button[contains(@aria-label,"Ulasan") or contains(@aria-label,"Reviews")]',
        ]
        for xp in xpaths:
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, xp))
                ).click()
                log.info("Tab ulasan dibuka")
                return
            except TimeoutException:
                continue
        log.warning("Tab ulasan tidak ditemukan - mungkin panel ulasan sudah terbuka.")

    def sort_reviews(self, mode: str = "newest") -> None:
        index, label_re = SORT_OPTIONS[mode]
        try:
            btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    '//button[contains(@aria-label,"Urutkan") or contains(@aria-label,"Sort")]',
                ))
            )
            btn.click()
            time.sleep(1.5)
            items = self.driver.find_elements(By.CSS_SELECTOR, 'div[role="menuitemradio"]')
            target = next(
                (i for i in items if re.search(label_re, i.text, re.I)),
                items[index] if len(items) > index else None,
            )
            if target is None:
                raise RuntimeError("menu urutan kosong")
            target.click()
            time.sleep(2.5)
            log.info("Ulasan diurutkan: %s", mode)
        except Exception as exc:
            log.warning("Gagal mengurutkan (%s), lanjut dengan urutan default.", exc)

    # ---- scroll ----------------------------------------------------------- #
    def _count(self) -> int:
        return self.driver.execute_script(
            "return document.querySelectorAll(arguments[0]).length;", REVIEW_CARD
        )

    def _find_panel(self):
        WebDriverWait(self.driver, 20).until(
            lambda d: d.find_elements(By.CSS_SELECTOR, REVIEW_CARD)
        )
        panel = self.driver.execute_script(FIND_PANEL_JS, REVIEW_CARD)
        if panel is None:
            raise RuntimeError("Panel scroll ulasan tidak ditemukan.")
        return panel

    def _wait_for_growth(self, prev: int, timeout: float) -> int:
        """Tunggu sampai jumlah ulasan bertambah (maks `timeout` detik)."""
        end = time.time() + timeout
        while time.time() < end:
            time.sleep(0.4)
            n = self._count()
            if n > prev:
                return n
        return self._count()

    def scroll_reviews(self, max_reviews: Optional[int], total_reported: Optional[int],
                       stall_limit: int = 8, max_minutes: float = 90) -> int:
        """Scroll panel sampai semua ulasan termuat. Mengembalikan jumlah kartu ulasan."""
        log.info("Mulai scroll (target: %s, dilaporkan Google: %s)",
                 max_reviews or "semua", total_reported or "?")
        panel = self._find_panel()
        count = self._count()
        stalls = 0
        t0 = last_log_t = time.time()
        last_log_n = 0

        try:
            while True:
                if max_reviews and count >= max_reviews:
                    log.info("Target %d ulasan tercapai.", max_reviews)
                    break
                if time.time() - t0 > max_minutes * 60:
                    log.warning("Batas waktu %.0f menit tercapai.", max_minutes)
                    break

                try:
                    self.driver.execute_script(
                        "arguments[0].scrollTop = arguments[0].scrollHeight;", panel
                    )
                except StaleElementReferenceException:
                    panel = self._find_panel()
                    continue

                new = self._wait_for_growth(count, timeout=4 + stalls * 2)
                if new > count:
                    count, stalls = new, 0
                    if count - last_log_n >= 100 or time.time() - last_log_t > 15:
                        rate = (count - last_log_n) / max(time.time() - last_log_t, 1e-6)
                        pct = f" ({count / total_reported:.0%})" if total_reported else ""
                        log.info("Ulasan dimuat: %d%s | %.1f ulasan/detik", count, pct, rate)
                        last_log_n, last_log_t = count, time.time()
                    continue

                # Tidak ada penambahan
                stalls += 1
                # Jika sudah melewati jumlah yang dilaporkan Google, cukup beberapa kali cek.
                limit = 3 if total_reported and count >= total_reported else stall_limit
                log.info("Belum ada ulasan baru (%d/%d) - total %d", stalls, limit, count)
                if stalls >= limit:
                    log.info("Scroll selesai: tidak ada ulasan baru.")
                    break
                # Pancing loader: naik sedikit lalu turun lagi.
                self.driver.execute_script(
                    "arguments[0].scrollTop = Math.max(arguments[0].scrollTop - 800, 0);", panel
                )
                time.sleep(0.8)
        except KeyboardInterrupt:
            log.warning("Scroll dihentikan manual - memproses ulasan yang sudah dimuat.")

        return self._count()

    # ---- ekstraksi -------------------------------------------------------- #
    def expand_all_reviews(self, rounds: int = 3) -> None:
        total = 0
        for _ in range(rounds):
            clicked = self.driver.execute_script(EXPAND_JS)
            if not clicked:
                break
            total += clicked
            time.sleep(1.5)
        log.info("%d ulasan panjang di-expand", total)

    def extract_reviews(self, place_name: str) -> pd.DataFrame:
        total = self._count()
        if total == 0:
            Path("debug_page_source.html").write_text(self.driver.page_source, encoding="utf-8")
            self.driver.save_screenshot("debug_screenshot.png")
            log.error("Tidak ada ulasan ditemukan. File debug disimpan.")
            return pd.DataFrame()

        now = datetime.now()
        scraped_at = now.strftime("%Y-%m-%d %H:%M:%S")
        rows: list[dict] = []
        for start in range(0, total, CHUNK_SIZE):
            batch = self.driver.execute_script(EXTRACT_JS, start, CHUNK_SIZE, REVIEW_CARD)
            for r in batch:
                meta = r["reviewer_meta"]
                m = re.search(r"(\d[\d.,]*)\s*(?:ulasan|reviews?)", meta, re.I)
                rows.append({
                    "nama_tempat": place_name,
                    "review_id": r["review_id"],
                    "nama_reviewer": clean_text(r["reviewer"]),
                    "local_guide": bool(re.search(r"local guide", meta, re.I)),
                    "jumlah_ulasan_reviewer": parse_int(m.group(1)) if m else None,
                    "rating": parse_rating(r["rating_label"]),
                    "tanggal": clean_text(r["date_raw"]),
                    "perkiraan_tanggal": approx_date(r["date_raw"], now),
                    "ulasan": clean_text(r["text"]),
                    "balasan_pemilik": clean_text(r["owner_reply"]),
                    "jumlah_foto": r["photos"],
                    "foto": "Ya" if r["photos"] else "Tidak",
                    "jumlah_suka": parse_int(r["likes"]) or 0,
                    "waktu_scrape": scraped_at,
                })
            log.info("Diekstrak %d/%d", min(start + CHUNK_SIZE, total), total)

        df = pd.DataFrame(rows)

        # Deduplikasi: pakai review_id; kalau kosong pakai kombinasi kolom.
        before = len(df)
        has_id = df["review_id"] != ""
        df = pd.concat([
            df[has_id].drop_duplicates("review_id"),
            df[~has_id].drop_duplicates(["nama_reviewer", "tanggal", "ulasan"]),
        ]).sort_index()
        if before != len(df):
            log.info("%d ulasan duplikat dibuang", before - len(df))

        df.insert(0, "no", range(1, len(df) + 1))
        return df.reset_index(drop=True)

    # ---- alur utama per tempat -------------------------------------------- #
    def scrape_place(self, url: str, max_reviews: Optional[int] = None, sort: str = "newest",
                     stall_limit: int = 8, max_minutes: float = 90) -> pd.DataFrame:
        log.info("Membuka: %s", url)
        self.driver.get(url)
        self._accept_consent()
        try:
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1.DUwDvf")))
        except TimeoutException:
            log.warning("Judul tempat tidak muncul, tetap dilanjutkan.")

        info = self.get_place_info()
        log.info("Tempat: %s | Rating: %s | Total ulasan: %s | Kategori: %s",
                 info["nama_tempat"], info["rating_keseluruhan"],
                 info["total_ulasan"] or "?", info["kategori"])

        self.open_reviews_tab()
        self.sort_reviews(sort)
        self.scroll_reviews(max_reviews, info["total_ulasan"], stall_limit, max_minutes)
        self.expand_all_reviews()

        df = self.extract_reviews(info["nama_tempat"])
        if max_reviews and len(df) > max_reviews:
            df = df.head(max_reviews)
        return df


# --------------------------------------------------------------------------- #
# Penyimpanan & ringkasan
# --------------------------------------------------------------------------- #

def save_dataframe(df: pd.DataFrame, out_dir: Path, name: str, xlsx: bool = True) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = out_dir / f"{name}_{stamp}"
    df.to_csv(f"{base}.csv", index=False, encoding="utf-8-sig")
    log.info("Tersimpan: %s.csv", base)
    if xlsx:
        try:
            df.to_excel(f"{base}.xlsx", index=False, engine="openpyxl")
            log.info("Tersimpan: %s.xlsx", base)
        except Exception as exc:  # CSV sudah aman tersimpan
            log.warning("Export Excel gagal (%s). Gunakan file CSV.", exc)


def print_summary(df: pd.DataFrame) -> None:
    print(f"\nTotal ulasan : {len(df)}")
    print(f"Tempat       : {df['nama_tempat'].iloc[0]}")
    print(f"Dengan teks  : {(df['ulasan'] != '').sum()}")
    print(f"Rating rata2 : {df['rating'].mean():.2f}")
    print("Distribusi rating:")
    for rating, cnt in df["rating"].value_counts().sort_index(ascending=False).items():
        pct = cnt / len(df) * 100
        print(f"  {rating:.0f} bintang: {cnt:5d} ({pct:5.1f}%) {'█' * int(pct / 2)}")


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def load_urls(args: argparse.Namespace) -> list[str]:
    urls = list(args.urls)
    if args.urls_file:
        lines = Path(args.urls_file).read_text(encoding="utf-8").splitlines()
        urls += [ln.strip() for ln in lines if ln.strip() and not ln.startswith("#")]
    urls = urls or DEFAULT_URLS
    if not urls:
        entered = input("Tempel URL Google Maps lalu tekan Enter: ").strip().strip('"')
        if entered:
            urls = [entered]
    return urls


def main() -> int:
    ap = argparse.ArgumentParser(description="Scraper ulasan Google Maps")
    ap.add_argument("urls", nargs="*", help="URL Google Maps (boleh lebih dari satu cabang)")
    ap.add_argument("--urls-file", help="File txt berisi satu URL per baris")
    ap.add_argument("--max-reviews", type=int, default=None, help="Batasi jumlah ulasan per tempat")
    ap.add_argument("--sort", choices=list(SORT_OPTIONS), default="newest")
    ap.add_argument("--stall-limit", type=int, default=8,
                    help="Berapa kali tanpa ulasan baru sebelum scroll dihentikan")
    ap.add_argument("--max-minutes", type=float, default=90, help="Batas waktu scroll per tempat")
    ap.add_argument("--output-dir", default="output")
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--block-images", action="store_true", help="Blok gambar agar lebih cepat")
    ap.add_argument("--no-xlsx", action="store_true", help="Simpan CSV saja")
    args = ap.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(message)s", datefmt="%H:%M:%S")

    urls = load_urls(args)
    if not urls:
        ap.error("Tidak ada URL. Berikan URL, --urls-file, atau isi DEFAULT_URLS di dalam script.")

    out_dir = Path(args.output_dir)
    results: list[pd.DataFrame] = []

    with GoogleMapsReviewScraper(headless=args.headless, block_images=args.block_images) as scraper:
        for i, url in enumerate(urls, 1):
            log.info("=" * 60)
            log.info("CABANG %d/%d", i, len(urls))
            try:
                df = scraper.scrape_place(url, args.max_reviews, args.sort,
                                          args.stall_limit, args.max_minutes)
            except KeyboardInterrupt:
                log.warning("Dihentikan pengguna.")
                break
            except Exception:
                log.exception("Gagal scraping %s", url)
                continue

            if df.empty:
                continue
            print_summary(df)
            save_dataframe(df, out_dir, slugify(df["nama_tempat"].iloc[0]), not args.no_xlsx)
            results.append(df)

    if len(results) > 1:
        combined = pd.concat(results, ignore_index=True)
        combined["no"] = range(1, len(combined) + 1)
        save_dataframe(combined, out_dir, "gabungan_semua_cabang", not args.no_xlsx)

    return 0 if results else 1


if __name__ == "__main__":
    sys.exit(main())
