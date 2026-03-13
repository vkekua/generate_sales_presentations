import os
import pandas as pd
from data import load_data, load_partners, prepare_partner_data
from ppt_builder import build_presentation
from config import OUTPUT_DIR


def main():
    print("📥 Loading data...")
    df       = load_data()
    partners = load_partners(df)

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print(f"🔁 Building presentations for {len(partners)} row(s)...\n")

    saved = 0
    skipped = 0

    for _, partner_row in partners.iterrows():
        partner = partner_row["Partner"]
        country = partner_row["Country"]

        month_raw   = pd.to_datetime(partner_row["Month"])
        month_label = month_raw.strftime("%B %Y")   # e.g. "March 2026" — used on slide
        month_file  = month_raw.strftime("%Y-%m")   # e.g. "2026-03"    — used in filename

        print(f"  ⏳ {partner} — {country} — {month_label}")

        data = prepare_partner_data(df, partner, country, month_raw)

        if data is None:
            print(f"  ⚠️  No data — skipping {partner} — {country} — {month_label}")
            skipped += 1
            continue

        safe_partner = partner.replace(" ", "_")
        safe_country = country.replace(" ", "_")
        output_path  = os.path.join(OUTPUT_DIR, f"{safe_partner}_{safe_country}_{month_file}.pptx")

        build_presentation(
            partner=partner,
            country=country,
            month=month_label,
            data=data,
            output_path=output_path,
        )

        print(f"  ✅ Saved: {output_path}") 
        saved += 1

    print(f"\n🎉 Done! {saved} file(s) saved to /{OUTPUT_DIR}/")
    print(f"⚠️  {skipped} partner(s) skipped due to missing data.")
    
if __name__ == "__main__":
    main()