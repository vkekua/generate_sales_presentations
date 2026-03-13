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

    print(f"🔁 Building presentations for {len(partners)} partner(s)...\n")

    for _, partner_row in partners.iterrows():
        partner = partner_row["Partner"]
        country = partner_row["Country"]

        # ✅ Format month cleanly — e.g. "2026-03" or "March 2026"
        month_raw   = pd.to_datetime(partner_row["Month"])
        month_label = month_raw.strftime("%B %Y")       # e.g. "March 2026" — for slide text
        month_file  = month_raw.strftime("%Y-%m")       # e.g. "2026-03"    — for filename

        print(f"  ⏳ {partner} — {country} — {month_label}")

        data        = prepare_partner_data(df, partner)
        safe_name   = partner.replace(" ", "_")
        output_path = os.path.join(OUTPUT_DIR, f"{safe_name}_{month_file}.pptx")

        build_presentation(
            partner=partner,
            country=country,
            month=month_label,   # ✅ clean label used on slide
            data=data,
            output_path=output_path,
        )

        print(f"  ✅ Saved: {output_path}")

    print(f"\n🎉 Done! {len(partners)} file(s) saved to /{OUTPUT_DIR}/")


if __name__ == "__main__":
    main()