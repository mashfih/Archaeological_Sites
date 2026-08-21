import os
import sys
import sqlite3
import pandas as pd

try:
    from tabulate import tabulate
    HAS_TABULATE = True
except ImportError:
    HAS_TABULATE = False

# ------------------------------------------------------------
# CONFIGURATION
# ------------------------------------------------------------
EXCEL_FILE = r"C:\Users\Lenovo\Downloads\Telegram Desktop\all_values_corrected.xlsx"
DB_FILE = "bangladesh_archaeological_sites.db"
SHEET_NAME = "মার্জড"

# ------------------------------------------------------------
# AUTO-TRANSLATE ENGLISH KEYWORDS TO BENGALI
# (Includes common district, division, and category names)
# ------------------------------------------------------------
TRANSLATE = {
    # Religious / site types
    "mosque": "মসজিদ",
    "mosjid": "মসজিদ",
    "masjid": "মসজিদ",
    "temple": "মন্দির",
    "mandir": "মন্দির",
    "palace": "প্রাসাদ",
    "fort": "দুর্গ",
    "kella": "কেল্লা",
    "bihar": "বিহার",
    "vihara": "বিহার",
    "monument": "স্মৃতিস্তম্ভ",
    "mound": "ঢিবি",
    "dhobi": "ঢিবি",
    "dighi": "দিঘি",
    "pond": "পুকুর",
    "gate": "তোরণ",
    "arch": "তোরণ",
    "ghat": "ঘাট",
    "house": "ভবন",
    "building": "ভবন",
    "bari": "বাড়ি",
    "mazar": "মাজার",
    "dargah": "দরগাহ",
    "shrine": "মাজার",

    # Divisions
    "dhaka": "ঢাকা",
    "chittagong": "চট্টগ্রাম",
    "ctg": "চট্টগ্রাম",
    "barisal": "বরিশাল",
    "rajshahi": "রাজশাহী",
    "khulna": "খুলনা",
    "sylhet": "সিলেট",
    "rangpur": "রংপুর",
    "mymensingh": "ময়মনসিংহ",

    # Districts – adding common spellings
    "cumilla": "কুমিল্লা",
    "comilla": "কুমিল্লা",
    "gazipur": "গাজীপুর",
    "narsingdi": "নরসিংদী",
    "narayanganj": "নারায়ণগঞ্জ",
    "faridpur": "ফরিদপুর",
    "madaripur": "মাদারীপুর",
    "shariatpur": "শরীয়তপুর",
    "gopalganj": "গোপালগঞ্জ",
    "tangail": "টাঙ্গাইল",
    "kishoreganj": "কিশোরগঞ্জ",
    "manikganj": "মানিকগঞ্জ",
    "munshiganj": "মুন্সীগঞ্জ",
    "brahmanbaria": "ব্রাহ্মণবাড়িয়া",
    "noakhali": "নোয়াখালী",
    "feni": "ফেনী",
    "lakshmipur": "লক্ষ্মীপুর",
    "chandpur": "চাঁদপুর",
    "bogura": "বগুড়া",
    "bogra": "বগুড়া",
    "joypurhat": "জয়পুরহাট",
    "naogaon": "নওগাঁ",
    "natore": "নাটোর",
    "pabna": "পাবনা",
    "sirajganj": "সিরাজগঞ্জ",
    "jhenaidah": "ঝিনাইদহ",
    "jashore": "যশোর",
    "chuadanga": "চুয়াডাঙ্গা",
    "magura": "মাগুরা",
    "meherpur": "মেহেরপুর",
    "kushtia": "কুষ্টিয়া",
    "bagherhat": "বাগেরহাট",
    "satkhira": "সাতক্ষীরা",
    "kurigram": "কুড়িগ্রাম",
    "gaibandha": "গাইবান্ধা",
    "lalmonirhat": "লালমনিরহাট",
    "nilphamari": "নীলফামারী",
    "panchagarh": "পঞ্চগড়",
    "thakurgaon": "ঠাকুরগাঁও",
    "dinajpur": "দিনাজপুর",
    "sherpur": "শেরপুর",
    "netrokona": "নেত্রকোণা",
    "moulvibazar": "মৌলভীবাজার",
    "habiganj": "হবিগঞ্জ",
    "sunamganj": "সুনামগঞ্জ",
    "bandarban": "বান্দরবান",
    "khagrachhari": "খাগড়াছড়ি",
    "rangamati": "রাঙ্গামাটি",
    "cox": "কক্সবাজার",
    "cox's bazar": "কক্সবাজার",

    # Upazilas – common names
    "hatiazari": "হাটহাজারী",
    "hathazari": "হাটহাজারী",
    "raozan": "রাউজান",
    "fatehpur": "ফতেহপুর",
    "savar": "সাভার",
    "keraniganj": "কেরানীগঞ্জ",
    "sonargaon": "সোনারগাঁও",
    "pabna": "পাবনা",
    "bagmara": "বাগমারা",
    "puthia": "পুঠিয়া",
    "shahjadpur": "শাহজাদপুর",
    "boalia": "বোয়ালিয়া",
    "rulup": "রূপসা",  # etc.
    "patuakhali": "পটুয়াখালী",
    "jhalokathi": "ঝালকাঠি",
    "rajapur": "রাজাপুর",
    "barishal": "বরিশাল",
    "barguna": "বরগুনা",
    "pirojpur": "পিরোজপুর",
    "bhola": "ভোলা",
    "potuakhali": "পটুয়াখালী",
    "bagerhat": "বাগেরহাট",
    "kotalipara": "কোটালীপাড়া",
    "tongipara": "টুঙ্গিপাড়া",
    "maksudpur": "মকসুদপুর",
    "kalapara": "কলাপাড়া",
    "mirzaganj": "মির্জাগঞ্জ",
    "dumki": "দুমকী",
    "dasmina": "দশমিনা",
    "mathbaria": "মঠবাড়িয়া",
    "kawkhali": "কাউখালী",
    "bhandaria": "ভান্ডারিয়া",
    "rajpasha": "রাজপাশা",
    "betagi": "বেতাগী",
    "bakerganj": "বাকেরগঞ্জ",
    "gournadi": "গৌরনদী",
    "hizla": "হিজলা",
    "mehendiganj": "মেহেন্দীগঞ্জ",
    "banaripara": "বানারীপাড়া",
    "agailjhara": "আগৈলঝাড়া",
    "faridpur": "ফরিদপুর",
    "madhukhali": "মধুখালী",
    "banga": "ভাঙ্গা",
    "sadar": "সদর",
    "upazila": "",        # generic term – no translation
    "upazilla": "",
    "zilla": "",
    "district": "",
    "division": "",
}

def translate_keyword(term):
    term_lower = term.lower().strip()
    # Direct match
    if term_lower in TRANSLATE:
        return TRANSLATE[term_lower] if TRANSLATE[term_lower] else None
    # Partial match – for multi‑word terms like "old mosque"
    for eng, ben in TRANSLATE.items():
        if eng and ben and eng in term_lower:
            return ben
    # If it looks like English but not found, maybe it's a place name not in list – return as is
    # We'll return None only for generic terms, but for other English words we return the term itself.
    # But we want to avoid no results for "cumilla" – already handled above.
    return term

# ------------------------------------------------------------
# DISPLAY HELPER – paginated, compact listing
# ------------------------------------------------------------
def display_listing(df, title=None, page_size=20):
    """Display a compact listing: Division, District, Site Name (truncated), Status."""
    if df.empty:
        print("No data to display.")
        return None

    # Build compact DataFrame for display
    display_df = df[['Division', 'District', 'Site Name', 'Status']].copy()
    # Truncate Site Name to 35 characters
    display_df['Site Name'] = display_df['Site Name'].astype(str).str.slice(0, 35)
    display_df['Site Name'] = display_df['Site Name'].apply(lambda x: x + '…' if len(x) >= 35 else x)

    total = len(display_df)
    if title:
        print(f"\n{title} (Total: {total})")

    if total <= page_size:
        if HAS_TABULATE:
            print(tabulate(display_df, headers='keys', tablefmt='pipe', showindex=False))
        else:
            with pd.option_context('display.width', 1000, 'display.max_colwidth', 35):
                print(display_df.to_string(index=False))
        return df  # return original for details

    pages = (total + page_size - 1) // page_size
    for page in range(pages):
        start = page * page_size
        end = min(start + page_size, total)
        print(f"\n--- Page {page+1}/{pages} (rows {start+1}-{end}) ---")
        sub_df = display_df.iloc[start:end]
        if HAS_TABULATE:
            print(tabulate(sub_df, headers='keys', tablefmt='pipe', showindex=False))
        else:
            with pd.option_context('display.width', 1000, 'display.max_colwidth', 35):
                print(sub_df.to_string(index=False))
        if page < pages - 1:
            input("Press Enter to continue to next page...")
    return df

# ------------------------------------------------------------
# DATABASE CREATION (only if DB doesn't exist)
# ------------------------------------------------------------
def create_database():
    print("📁 Database not found. Creating from Excel file...")
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ Excel file not found at: {EXCEL_FILE}")
        print("Please update the EXCEL_FILE variable with the correct path.")
        sys.exit(1)

    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, dtype=str)
    df.columns = df.columns.str.strip()

    col_map = {
        'ক্রম': 'serial',
        'বিভাগ': 'division',
        'জেলা': 'district',
        'উপজেলা': 'upazila',
        'প্রাচীন নিদর্শন/পুরাকীর্তি': 'site_name',
        'সংক্ষিপ্ত বর্ণনা': 'description',
        'গ্রাম': 'village',
        'অক্ষাংশ': 'latitude_dms',
        'দ্রাঘিমাংশ': 'longitude_dms',
        'শ্রেণীবিভাগ': 'classification',
        'অবস্থা': 'status',
        'প্রজ্ঞাপন/গেজেট নাম': 'gazette_name',
        'গেজেট তারিখ': 'gazette_date',
        'ছবি': 'image_column',
        'সোর্স ইউআরএল': 'source_url',
        'পিডিএফ ইউআরএল': 'pdf_url',
        'ছবি ইউআরএল': 'image_url',
        'অক্ষাংশ (দশমিক)': 'latitude_dec',
        'দ্রাঘিমাংশ (দশমিক)': 'longitude_dec',
    }
    df.rename(columns=col_map, inplace=True)

    keep_cols = ['division', 'district', 'upazila', 'site_name', 'description',
                 'village', 'latitude_dms', 'longitude_dms', 'classification',
                 'status', 'gazette_name', 'gazette_date',
                 'source_url', 'pdf_url', 'image_url',
                 'latitude_dec', 'longitude_dec']
    df = df[keep_cols].copy()

    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("PRAGMA foreign_keys = ON;")

    cur.executescript("""
    DROP TABLE IF EXISTS archaeological_sites;
    DROP TABLE IF EXISTS upazilas;
    DROP TABLE IF EXISTS districts;
    DROP TABLE IF EXISTS divisions;
    DROP TABLE IF EXISTS gazettes;

    CREATE TABLE divisions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL
    );

    CREATE TABLE districts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        division_id INTEGER NOT NULL,
        name TEXT NOT NULL,
        UNIQUE(division_id, name),
        FOREIGN KEY (division_id) REFERENCES divisions(id)
    );

    CREATE TABLE upazilas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        district_id INTEGER NOT NULL,
        name TEXT NOT NULL,
        UNIQUE(district_id, name),
        FOREIGN KEY (district_id) REFERENCES districts(id)
    );

    CREATE TABLE gazettes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        date TEXT,
        UNIQUE(name, date)
    );

    CREATE TABLE archaeological_sites (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        upazila_id INTEGER NOT NULL,
        site_name TEXT NOT NULL,
        description TEXT,
        village TEXT,
        latitude_dms TEXT,
        longitude_dms TEXT,
        latitude_dec REAL,
        longitude_dec REAL,
        classification TEXT,
        status TEXT,
        gazette_id INTEGER,
        source_url TEXT,
        pdf_url TEXT,
        image_url TEXT,
        FOREIGN KEY (upazila_id) REFERENCES upazilas(id),
        FOREIGN KEY (gazette_id) REFERENCES gazettes(id)
    );

    CREATE INDEX idx_sites_upazila ON archaeological_sites(upazila_id);
    CREATE INDEX idx_sites_gazette ON archaeological_sites(gazette_id);
    """)

    def get_or_create_division(name):
        cur.execute("SELECT id FROM divisions WHERE name = ?", (name,))
        row = cur.fetchone()
        if row:
            return row[0]
        cur.execute("INSERT INTO divisions (name) VALUES (?)", (name,))
        return cur.lastrowid

    def get_or_create_district(division_id, name):
        cur.execute("SELECT id FROM districts WHERE division_id = ? AND name = ?", (division_id, name))
        row = cur.fetchone()
        if row:
            return row[0]
        cur.execute("INSERT INTO districts (division_id, name) VALUES (?, ?)", (division_id, name))
        return cur.lastrowid

    def get_or_create_upazila(district_id, name):
        cur.execute("SELECT id FROM upazilas WHERE district_id = ? AND name = ?", (district_id, name))
        row = cur.fetchone()
        if row:
            return row[0]
        cur.execute("INSERT INTO upazilas (district_id, name) VALUES (?, ?)", (district_id, name))
        return cur.lastrowid

    def get_or_create_gazette(name, date):
        if pd.isna(name) and pd.isna(date):
            return None
        name = name if pd.notna(name) else ""
        date = date if pd.notna(date) else ""
        cur.execute("SELECT id FROM gazettes WHERE name = ? AND date = ?", (name, date))
        row = cur.fetchone()
        if row:
            return row[0]
        cur.execute("INSERT INTO gazettes (name, date) VALUES (?, ?)", (name, date))
        return cur.lastrowid

    conn.execute("BEGIN TRANSACTION")
    insert_sql = """
    INSERT INTO archaeological_sites (
        upazila_id, site_name, description, village,
        latitude_dms, longitude_dms, latitude_dec, longitude_dec,
        classification, status, gazette_id,
        source_url, pdf_url, image_url
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """

    for _, row in df.iterrows():
        if pd.isna(row['site_name']) or pd.isna(row['division']) or pd.isna(row['district']) or pd.isna(row['upazila']):
            continue

        div_id = get_or_create_division(row['division'].strip())
        dist_id = get_or_create_district(div_id, row['district'].strip())
        upaz_id = get_or_create_upazila(dist_id, row['upazila'].strip())
        gazette_id = get_or_create_gazette(row.get('gazette_name'), row.get('gazette_date'))

        def safe_val(val):
            return None if pd.isna(val) else val

        values = (
            upaz_id,
            safe_val(row['site_name']),
            safe_val(row['description']),
            safe_val(row['village']),
            safe_val(row['latitude_dms']),
            safe_val(row['longitude_dms']),
            safe_val(row['latitude_dec']),
            safe_val(row['longitude_dec']),
            safe_val(row['classification']),
            safe_val(row['status']),
            gazette_id,
            safe_val(row['source_url']),
            safe_val(row['pdf_url']),
            safe_val(row['image_url'])
        )
        cur.execute(insert_sql, values)

    conn.commit()
    conn.close()
    print(f"✅ Database created: {DB_FILE}")
    print(f"   Total sites inserted: {len(df)}")

# ------------------------------------------------------------
# SEARCH FUNCTIONS
# ------------------------------------------------------------
def search_sites(search_term):
    conn = sqlite3.connect(DB_FILE)
    query = """
    SELECT 
        s.id,
        d.name AS Division,
        dist.name AS District,
        u.name AS Upazila,
        s.site_name AS 'Site Name',
        s.village AS Village,
        s.classification AS Classification,
        s.status AS Status,
        s.latitude_dec AS Latitude,
        s.longitude_dec AS Longitude,
        s.description AS Description,
        s.source_url AS Source_URL,
        s.pdf_url AS PDF_URL,
        s.image_url AS Image_URL
    FROM archaeological_sites s
    JOIN upazilas u ON s.upazila_id = u.id
    JOIN districts dist ON u.district_id = dist.id
    JOIN divisions d ON dist.division_id = d.id
    WHERE s.site_name LIKE ? 
       OR s.village LIKE ?
       OR dist.name LIKE ?
       OR u.name LIKE ?
       OR s.classification LIKE ?
       OR s.status LIKE ?
    ORDER BY d.name, dist.name, s.site_name
    """
    search_pattern = f"%{search_term}%"
    df = pd.read_sql_query(query, conn, params=(search_pattern,)*6)
    conn.close()
    return df

def show_statistics():
    conn = sqlite3.connect(DB_FILE)
    print("\n--- Statistics ---")
    div_df = pd.read_sql_query("""
        SELECT d.name AS Division, COUNT(*) AS count
        FROM archaeological_sites s
        JOIN upazilas u ON s.upazila_id = u.id
        JOIN districts dist ON u.district_id = dist.id
        JOIN divisions d ON dist.division_id = d.id
        GROUP BY d.name
        ORDER BY count DESC
    """, conn)
    print("\nSites by Division:")
    print(tabulate(div_df, headers='keys', tablefmt='pipe', showindex=False) if HAS_TABULATE else div_df.to_string(index=False))

    class_df = pd.read_sql_query("""
        SELECT classification, COUNT(*) AS count
        FROM archaeological_sites
        WHERE classification IS NOT NULL
        GROUP BY classification
        ORDER BY count DESC
        LIMIT 10
    """, conn)
    print("\nTop 10 Classifications:")
    print(tabulate(class_df, headers='keys', tablefmt='pipe', showindex=False) if HAS_TABULATE else class_df.to_string(index=False))
    conn.close()

def show_all_sites():
    conn = sqlite3.connect(DB_FILE)
    query = """
    SELECT 
        s.id,
        d.name AS Division,
        dist.name AS District,
        u.name AS Upazila,
        s.site_name AS 'Site Name',
        s.village AS Village,
        s.classification AS Classification,
        s.status AS Status
    FROM archaeological_sites s
    JOIN upazilas u ON s.upazila_id = u.id
    JOIN districts dist ON u.district_id = dist.id
    JOIN divisions d ON dist.division_id = d.id
    ORDER BY d.name, dist.name, s.site_name
    """
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

def export_to_csv(df, filename="search_results.csv"):
    df.to_csv(filename, index=False, encoding='utf-8-sig')
    print(f"✅ Exported {len(df)} rows to {filename}")

# ------------------------------------------------------------
# DETAIL VIEW
# ------------------------------------------------------------
def view_site_details(row):
    print("\n" + "="*60)
    print("SITE DETAILS")
    print("="*60)
    for col in ['Site Name', 'Division', 'District', 'Upazila', 'Village', 'Classification', 'Status']:
        if col in row:
            print(f"{col}: {row[col]}")
    print("\nDescription:")
    print(row.get('Description', 'No description') or 'No description')
    if 'Latitude' in row and 'Longitude' in row:
        print(f"\nCoordinates: {row['Latitude']}, {row['Longitude']}")
    if 'Source_URL' in row and pd.notna(row['Source_URL']):
        print(f"Source: {row['Source_URL']}")
    if 'PDF_URL' in row and pd.notna(row['PDF_URL']):
        print(f"PDF: {row['PDF_URL']}")
    if 'Image_URL' in row and pd.notna(row['Image_URL']):
        print(f"Image: {row['Image_URL']}")
    input("\nPress Enter to continue...")

# ------------------------------------------------------------
# MAIN MENU
# ------------------------------------------------------------
def main():
    if not os.path.exists(DB_FILE):
        create_database()
    else:
        print("✅ Database already exists.")

    print("\n" + "="*60)
    print("  BANGLADESH ARCHAEOLOGICAL SITES SEARCH ENGINE")
    print("="*60)

    while True:
        print("\nOptions:")
        print("  1. Search by keyword")
        print("  2. Show ALL sites (paginated)")
        print("  3. Show statistics")
        print("  4. Export search results")
        print("  5. Exit (or 'q')")
        choice = input("Enter choice (1-5) or type a search term directly: ").strip()

        if choice.lower() in ['5', 'q']:
            print("👋 Goodbye!")
            break

        if choice in ['1', '2', '3', '4']:
            if choice == '1':
                term = input("Search keyword: ").strip()
                if not term:
                    print("❌ No keyword entered.")
                    continue
                translated = translate_keyword(term)
                if translated is None:
                    print("ℹ️  Generic term – please provide a specific name.")
                    continue
                if translated != term:
                    print(f"🔁 Translated to: {translated}")
                df = search_sites(translated)
                if df.empty:
                    print("❌ No results found.")
                    # Suggest if the translated term is a place name, but we already did.
                    continue
                display_listing(df, title="🔍 Search Results")
                if input("\nView details of a site? (y/n): ").strip().lower() == 'y':
                    try:
                        row_num = int(input("Enter row number: ").strip())
                        if 1 <= row_num <= len(df):
                            view_site_details(df.iloc[row_num-1])
                        else:
                            print("Invalid row number.")
                    except ValueError:
                        print("Invalid input.")
            elif choice == '2':
                df = show_all_sites()
                display_listing(df, title="📋 All Sites")
                if input("\nView details of a site? (y/n): ").strip().lower() == 'y':
                    try:
                        row_num = int(input("Enter row number: ").strip())
                        if 1 <= row_num <= len(df):
                            # re‑query full details by id
                            conn = sqlite3.connect(DB_FILE)
                            full_df = pd.read_sql_query("""
                                SELECT 
                                    d.name AS Division,
                                    dist.name AS District,
                                    u.name AS Upazila,
                                    s.site_name AS 'Site Name',
                                    s.village AS Village,
                                    s.latitude_dec AS Latitude,
                                    s.longitude_dec AS Longitude,
                                    s.classification AS Classification,
                                    s.status AS Status,
                                    s.description AS Description,
                                    s.source_url AS Source_URL,
                                    s.pdf_url AS PDF_URL,
                                    s.image_url AS Image_URL
                                FROM archaeological_sites s
                                JOIN upazilas u ON s.upazila_id = u.id
                                JOIN districts dist ON u.district_id = dist.id
                                JOIN divisions d ON dist.division_id = d.id
                                WHERE s.id = ?
                            """, conn, params=(df.iloc[row_num-1]['id'],))
                            conn.close()
                            if not full_df.empty:
                                view_site_details(full_df.iloc[0])
                            else:
                                print("Could not retrieve details.")
                        else:
                            print("Invalid row number.")
                    except ValueError:
                        print("Invalid input.")
            elif choice == '3':
                show_statistics()
            elif choice == '4':
                term = input("Keyword to search and export: ").strip()
                if term:
                    translated = translate_keyword(term)
                    if translated is None:
                        print("Generic term – please provide a specific name.")
                        continue
                    if translated != term:
                        print(f"🔁 Translated to: {translated}")
                    df = search_sites(translated)
                    if df.empty:
                        print("❌ No results found.")
                    else:
                        export_to_csv(df)
                else:
                    print("No keyword given.")
        else:
            # Treat as direct search term
            term = choice
            translated = translate_keyword(term)
            if translated is None:
                print("ℹ️  Generic term – please provide a specific name.")
                continue
            if translated != term:
                print(f"🔁 Translated to: {translated}")
            df = search_sites(translated)
            if df.empty:
                print("❌ No results found.")
                # Suggest checking spelling or trying Bengali
                print("💡 Hint: Try typing the Bengali name directly (e.g., কুমিল্লা).")
                continue
            display_listing(df, title="🔍 Search Results")
            if input("\nView details of a site? (y/n): ").strip().lower() == 'y':
                try:
                    row_num = int(input("Enter row number: ").strip())
                    if 1 <= row_num <= len(df):
                        view_site_details(df.iloc[row_num-1])
                    else:
                        print("Invalid row number.")
                except ValueError:
                    print("Invalid input.")

if __name__ == "__main__":
    main()