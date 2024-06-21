from perf_dash_oop import DealSplitter

def main():
    # welcome()
    deal_splitter = DealSplitter(
        template_file="template.xlsx",
        input_data_file="SellerSync_Data.xlsx",
        input_sheet_names=["GMS_AGG", "GMS_SKU"],
        output_folder="output",
        id_column="merchant_customer_id",
        filename_location={"col": "seller_name", "sheet": "GMS_AGG"}
    )
    deal_splitter.split_sellers()

main()