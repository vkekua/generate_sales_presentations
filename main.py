import os
import pandas as pd
from data import load_data, load_partners, prepare_partner_data
from ppt_builder import build_presentation
from config import INPUT_FILE, OUTPUT_DIR


def generate_all(input_file: str, output_dir: str) -> list[str]:
    """Run the full pipeline. Returns list of generated .pptx file paths."""
    df       = load_data(input_file)
    partners = load_partners(df, input_file)

    os.makedirs(output_dir, exist_ok=True)

    generated = []

    for _, partner_row in partners.iterrows():
        partner = partner_row["Partner"]
        country = partner_row["Country"]

        month_raw   = pd.to_datetime(partner_row["Month"])
        month_label = month_raw.strftime("%B %Y")
        month_file  = month_raw.strftime("%Y-%m")

        data = prepare_partner_data(df, partner, country, month_raw)

        if data is None:
            continue

        safe_partner = partner.replace(" ", "_")
        safe_country = country.replace(" ", "_")
        output_path  = os.path.join(output_dir, f"{safe_partner}_{safe_country}_{month_file}.pptx")

        build_presentation(
            partner=partner,
            country=country,
            month=month_label,
            data=data,
            output_path=output_path,
        )

        generated.append(output_path)

    return generated


def main():
    print("Loading data...")
    generated = generate_all(INPUT_FILE, OUTPUT_DIR)
    print(f"\nDone! {len(generated)} file(s) saved to /{OUTPUT_DIR}/")


if __name__ == "__main__":
    main()