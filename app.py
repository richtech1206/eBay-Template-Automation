import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
import zipfile

class eBayTemplateGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("eBay 3D Seller Template Generator")
        
        # File paths
        self.vendor_items_file = ""
        self.competitor_files = []
        self.pies_file = ""
        self.category_id_file = ""
        self.image_list_file = ""
        self.brand_name = tk.StringVar()
        self.pies_file_type = tk.StringVar(value="single")

        # Labels to display file names
        self.vendor_items_label = tk.Label(self.root, text="")
        self.competitor_files_label = tk.Label(self.root, text="")
        self.pies_file_label = tk.Label(self.root, text="")
        self.category_id_label = tk.Label(self.root, text="")
        self.image_list_label = tk.Label(self.root, text="")

        # Set up UI
        self.setup_ui()
        
    def setup_ui(self):
        # Vendor items file
        tk.Label(self.root, text="Vendor Items File").grid(row=0, column=0, padx=10, pady=10)
        self.vendor_items_btn = tk.Button(self.root, text="Browse", command=self.load_vendor_items_file)
        self.vendor_items_btn.grid(row=0, column=1, padx=10, pady=10)
        self.vendor_items_label.grid(row=0, column=2, padx=10, pady=10, sticky='w')

        # Competitor files
        tk.Label(self.root, text="Competitor Output Files").grid(row=1, column=0, padx=10, pady=10)
        self.competitor_files_btn = tk.Button(self.root, text="Browse", command=self.load_competitor_files)
        self.competitor_files_btn.grid(row=1, column=1, padx=10, pady=10)
        self.competitor_files_label.grid(row=1, column=2, padx=10, pady=10, sticky='w')

        # PIES file
        tk.Label(self.root, text="PIES File").grid(row=2, column=0, padx=10, pady=10)
        self.pies_file_btn = tk.Button(self.root, text="Browse", command=self.load_pies_file)
        self.pies_file_btn.grid(row=2, column=1, padx=10, pady=10)
        self.pies_file_label.grid(row=2, column=2, padx=10, pady=10, sticky='w')
        
        # PIES file type radio buttons
        tk.Radiobutton(self.root, text="Single XLSX", variable=self.pies_file_type, value="single").grid(row=3, column=1, sticky='w')
        tk.Radiobutton(self.root, text="ZIP with multiple XLSX", variable=self.pies_file_type, value="zip").grid(row=4, column=1, sticky='w')

        # Category ID file
        tk.Label(self.root, text="Category ID File").grid(row=5, column=0, padx=10, pady=10)
        self.category_id_btn = tk.Button(self.root, text="Browse", command=self.load_category_id_file)
        self.category_id_btn.grid(row=5, column=1, padx=10, pady=10)
        self.category_id_label.grid(row=5, column=2, padx=10, pady=10, sticky='w')

        # Image list file
        tk.Label(self.root, text="Image List File").grid(row=6, column=0, padx=10, pady=10)
        self.image_list_btn = tk.Button(self.root, text="Browse", command=self.load_image_list_file)
        self.image_list_btn.grid(row=6, column=1, padx=10, pady=10)
        self.image_list_label.grid(row=6, column=2, padx=10, pady=10, sticky='w')

        # Brand name input
        tk.Label(self.root, text="Brand Name").grid(row=7, column=0, padx=10, pady=10)
        self.brand_name_entry = tk.Entry(self.root, textvariable=self.brand_name)
        self.brand_name_entry.grid(row=7, column=1, padx=10, pady=10)

        # Generate button
        self.generate_btn = tk.Button(self.root, text="Generate", command=self.generate_template)
        self.generate_btn.grid(row=8, column=0, columnspan=2, padx=10, pady=20)
        
    def load_vendor_items_file(self):
        self.vendor_items_file = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if self.vendor_items_file:
            self.vendor_items_label.config(text=self.vendor_items_file.split("/")[-1])
            messagebox.showinfo("File Loaded", "Vendor Items File Loaded Successfully")

    def load_competitor_files(self):
        self.competitor_files = filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")])
        if self.competitor_files:
            self.competitor_files_label.config(text=", ".join([file.split("/")[-1] for file in self.competitor_files]))
            messagebox.showinfo("Files Loaded", "Competitor Files Loaded Successfully")
        
    def load_pies_file(self):
        if self.pies_file_type.get() == "single":
            self.pies_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
            if self.pies_file:
                self.pies_file_label.config(text=self.pies_file.split("/")[-1])
                messagebox.showinfo("File Loaded", "PIES File Loaded Successfully")
        elif self.pies_file_type.get() == "zip":
            self.pies_file = filedialog.askopenfilename(filetypes=[("ZIP files", "*.zip")])
            if self.pies_file:
                self.pies_file_label.config(text=self.pies_file.split("/")[-1])
                messagebox.showinfo("File Loaded", "PIES ZIP File Loaded Successfully")

    def load_category_id_file(self):
        self.category_id_file = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if self.category_id_file:
            self.category_id_label.config(text=self.category_id_file.split("/")[-1])
            messagebox.showinfo("File Loaded", "Category ID File Loaded Successfully")

    def load_image_list_file(self):
        self.image_list_file = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if self.image_list_file:
            self.image_list_label.config(text=self.image_list_file.split("/")[-1])
            messagebox.showinfo("File Loaded", "Image List File Loaded Successfully")

    def get_max_attribute_count(self, attributes_segment):
        # Group by partnumber and count the occurrences of each
        attribute_counts = attributes_segment.groupby('partnumber').size().reset_index(name='count')
        # Find the maximum count of attributes for any part number
        max_count = attribute_counts['count'].max()
        return max_count

    def get_attributes_for_sku(self, sku, attributes_segment):
        sku_attributes = attributes_segment[attributes_segment['partnumber'] == sku]
        attributes_dict = {}
        for idx, (_, row) in enumerate(sku_attributes.iterrows()):
            attributes_dict[f'Attribute{idx+1}Name'] = row['attributename']
            attributes_dict[f'Attribute{idx+1}Value'] = row['productattribute']
        return attributes_dict

    def generate_template(self):
        try:
            # Read the source files
            vendor_items = pd.read_csv(self.vendor_items_file)
            vendor_items.columns = vendor_items.columns.str.lower().str.replace(' ', '')  # Normalize column names to lowercase and remove spaces

            competitor_data = pd.concat([pd.read_csv(file) for file in self.competitor_files])
            competitor_data.columns = competitor_data.columns.str.lower().str.replace(' ', '')  # Normalize column names to lowercase and remove spaces
            
            pies_data = {}
            if self.pies_file_type.get() == "single":
                pies_data = pd.read_excel(self.pies_file, sheet_name=None, dtype=str)
                pies_data = {sheet_name.lower().replace(' ', ''): pies_data[sheet_name].rename(columns=lambda x: x.lower().replace(' ', '')) for sheet_name in pies_data}  # Normalize sheet names and column names to lowercase and remove spaces
            elif self.pies_file_type.get() == "zip":
                with zipfile.ZipFile(self.pies_file, 'r') as z:
                    for filename in z.namelist():
                        if filename.endswith('.xlsx'):
                            with z.open(filename) as f:
                                sheet_data = pd.read_excel(f, sheet_name=None, engine='openpyxl', dtype=str)
                                for sheet_name in sheet_data:
                                    normalized_sheet_name = sheet_name.lower().replace(' ', '')
                                    if normalized_sheet_name not in pies_data:
                                        pies_data[normalized_sheet_name] = sheet_data[sheet_name].rename(columns=lambda x: x.lower().replace(' ', ''))
                                    else:
                                        pies_data[normalized_sheet_name] = pd.concat([pies_data[normalized_sheet_name], sheet_data[sheet_name].rename(columns=lambda x: x.lower().replace(' ', ''))])

            category_id_mapping = pd.read_csv(self.category_id_file)
            category_id_mapping.columns = category_id_mapping.columns.str.lower().str.replace(' ', '')  # Normalize column names to lowercase and remove spaces
            
            image_list = pd.read_csv(self.image_list_file)
            image_list.columns = image_list.columns.str.lower().str.replace(' ', '')  # Normalize column names to lowercase and remove spaces
            
            # Aggregate all attributes with their values
            description_segment = pies_data['descriptionsegment']
            pies_template = pies_data['piestemplate']
            report_segment = pies_data['report']
            attributes_segment = pies_data['productattributessegment']
            interchange_segment = pies_data['partinterchangesegment']
            
            # Get the maximum attribute count
            max_attribute_count = self.get_max_attribute_count(attributes_segment)
            
            # Initialize the output DataFrame with the correct column order
            template_columns = [
                'Item ID', 'SKU', 'Parent SKU', 'Title', 'Description', 'Tags', 'MetaKeywords', 'MetaDescription', 'MobileDescription',
                'CategoryID', 'CategoryID2', 'StoreCategory', 'StoreCategory2', 'eBayCatalogID', 'eBayCatalogSearch', 'PromoteCampaign', 
                'PromoteRate', 'CarMake', 'CarModel', 'CarYear', 'CarTrim', 'CarEngine', 'CopyCarCompatibility', 'CarCompatibility', 
                'CarCompatibilityNotes', 'DeleteCarCompatibility', 'KType (TecDoc)', 'PrivateListing', 'Quantity', 'WarehouseQuantity', 
                'InventoryControl', 'Price', 'WholesalePrice', 'OriginalRetailPrice', 'BestOffer', 'BestOfferAccept', 'BestOfferDecline', 
                'VATPercent', 'ListingDuration', 'AuctionStartPrice', 'AuctionReservePrice', 'AuctionBINPrice', 'Condition', 'ConditionNote', 
                'C:ASIN', 'C:UPC', 'C:EAN', 'C:MPN', 'C:Brand', 'C:Manufacturer', 'C:Size', 'C:Color', 'C:Material'
            ]

            # Add dynamic attribute columns based on the maximum attribute count
            for i in range(1, max_attribute_count + 1):
                template_columns.append(f'Attribute{i}Name')
                template_columns.append(f'Attribute{i}Value')

            # Continue with the rest of the columns
            template_columns += [
                'CountryCode', 'Location', 'PostalCode', 'PolicyPayment', 'PolicyShipping', 'PolicyReturn', 'PackageType', 'MeasurementSystem', 
                'PackageLength', 'PackageWidth', 'PackageDepth', 'WeightMajor', 'WeightMinor', 'ImageURLs', 'Image 1', 'Image 2', 
                'Image 3', 'Image 4', 'Image 5', 'Image 6', 'Image 7', 'Image 8', 'Image 9', 'Image 10', 'VariationImageOption', 
                'VariationImage 1', 'VariationImage 2', 'VariationImage 3', 'DeleteVariation'
            ]

            brand_name = self.brand_name.get()
            
            # Add a column for the count of each attribute name
            attributes_segment['attribute_count'] = attributes_segment.groupby('attributename')['attributename'].transform('count')

            # Sort by the count and then by the attribute name
            attributes_segment = attributes_segment.sort_values(by=['attribute_count', 'attributename'], ascending=[False, True])

            # Drop the temporary count column if no longer needed
            attributes_segment = attributes_segment.drop(columns=['attribute_count'])
            
            # Collect rows for the output DataFrame
            rows = []

            # Create dictionary mappings for faster lookup
            competitor_data_dict = {row['partnumber']: row for _, row in competitor_data.iterrows()}
            
            description_segment_dict = {}
            # Iterate through the rows and process descriptions
            for _, row in description_segment.iterrows():
                partnumber = row['partnumber']
                description = row['description']
                if partnumber not in description_segment_dict:
                    description_segment_dict[partnumber] = []
                if description not in description_segment_dict[partnumber]:
                    description_segment_dict[partnumber].append(description)

            # Convert sets to concatenated strings
            description_segment_dict = {k: ', '.join(v) for k, v in description_segment_dict.items()}
            
            # Initialize an empty dictionary to store interchange part numbers
            interchange_segment_dict = {}

            # Iterate through the rows and process interchange part numbers
            for _, row in interchange_segment.iterrows():
                item_partnumber = row['item_partnumber']
                partnumber = row['partnumber']
                if item_partnumber not in interchange_segment_dict:
                    interchange_segment_dict[item_partnumber] = []
                interchange_segment_dict[item_partnumber].append(partnumber)
                
            # Convert lists to concatenated strings with comma separation
            interchange_segment_dict = {k: ', '.join(v) for k, v in interchange_segment_dict.items()}
            
            report_segment_dict = {row['partnumber']: row for _, row in report_segment.iterrows()}
            pies_template_dict = {row['partnumber']: row for _, row in pies_template.iterrows()}
            image_list_dict = image_list.groupby('partnumber').apply(lambda x: x.sort_values('sortorder')['url'].tolist()).to_dict()
            category_id_mapping_dict = category_id_mapping.set_index('partterminologyname')['category_id'].to_dict()
            
            for idx, row in vendor_items.iterrows():
                sku = row['partnumber']
                
                # Initialize fields
                competitor_row = competitor_data_dict.get(sku, {})
                title = competitor_row.get('title', 'N/A')
                condition_code = competitor_row.get('conditioncode', 'N/A')
                quantity = competitor_row.get('quantity', 'N/A')
                price = competitor_row.get('price', 'N/A')
                copycarcompatibility = competitor_row.get('copycarcompatabilityid', 'N/A')
                
                # Description
                description = description_segment_dict.get(sku, 'N/A')
                
                # Interchange Numbers
                interchange_numbers = interchange_segment_dict.get(sku, 'N/A')
                
                description = f'{description}. Interchanges are: {interchange_numbers}'
                
                # Tags
                tags = f"{brand_name} - {datetime.now().strftime('%Y-%m-%d')}"
                
                # CategoryID
                part_terminology = pies_template_dict.get(sku, {}).get('partterminologyname', 'N/A')
                category_id = category_id_mapping_dict.get(part_terminology, '#N/A')
                    
                # Product Type
                product_type = report_segment_dict.get(sku, {}).get('partterminologyname', 'N/A')
                    
                # Ensure 'itemlevelgtin' column is treated as string
                upc = pies_template_dict.get(sku, {}).get('itemlevelgtin', 'N/A')
                if upc != 'N/A':
                    upc = upc[2:]
                
                # Brand
                brand = brand_name
                
                # Package Dimensions and Weight
                package_row = report_segment_dict.get(sku, {})
                package_length = package_row.get('length(in)', 'N/A')
                package_width = package_row.get('width(in)', 'N/A')
                package_depth = package_row.get('height(in)', 'N/A')
                weight_major = package_row.get('weight(lbs)', 'N/A')
                
                # C:MPN, C:Brand, C:Manufacturer
                c_mpn = sku
                c_brand = brand
                c_manufacturer = brand
                
                # Get image URLs
                image_urls = image_list_dict.get(sku.replace("-", ''), [])
                image_urls_dict = {f'Image {i+1}': url for i, url in enumerate(image_urls)}
                
                # Get all attribute names and values for the SKU
                attributes_for_sku = self.get_attributes_for_sku(sku, attributes_segment)
                
                # Prepare new row
                new_row = {
                    'Item ID': '',
                    'SKU': sku,
                    'Parent SKU': '',
                    'Title': title,
                    'Description': description,
                    'Tags': tags,
                    'MetaKeywords': '',
                    'MetaDescription': '',
                    'MobileDescription': '',
                    'CategoryID': category_id,
                    'CategoryID2': '',
                    'StoreCategory': '',
                    'StoreCategory2': '',
                    'eBayCatalogID': '',
                    'eBayCatalogSearch': '',
                    'PromoteCampaign': '',
                    'PromoteRate': '',
                    'CarMake': '',
                    'CarModel': '',
                    'CarYear': '',
                    'CarTrim': '',
                    'CarEngine': '',
                    'CopyCarCompatibility': copycarcompatibility,
                    'CarCompatibility': '',
                    'CarCompatibilityNotes': '',
                    'DeleteCarCompatibility': '',
                    'KType (TecDoc)': '',
                    'PrivateListing': '',
                    'Quantity': quantity,
                    'WarehouseQuantity': '',
                    'InventoryControl': '',
                    'Price': price,
                    'WholesalePrice': '',
                    'OriginalRetailPrice': '',
                    'BestOffer': '',
                    'BestOfferAccept': '',
                    'BestOfferDecline': '',
                    'VATPercent': '',
                    'ListingDuration': '',
                    'AuctionStartPrice': '',
                    'AuctionReservePrice': '',
                    'AuctionBINPrice': '',
                    'Condition': condition_code,
                    'ConditionNote': '',
                    'C:ASIN': '',
                    'C:UPC': upc,
                    'C:EAN': '',
                    'C:MPN': c_mpn,
                    'C:Brand': c_brand,
                    'C:Manufacturer': c_manufacturer,
                    'C:Size': '',
                    'C:Color': '',
                    'C:Material': '',
                    'C:Product Type': product_type,
                    'CountryCode': '',
                    'Location': '',
                    'PostalCode': '',
                    'PolicyPayment': '',
                    'PolicyShipping': '',
                    'PolicyReturn': '',
                    'PackageType': 'PackageThickEnvelope',
                    'MeasurementSystem': 'ENGLISH',
                    'PackageLength': package_length,
                    'PackageWidth': package_width,
                    'PackageDepth': package_depth,
                    'WeightMajor': weight_major,
                    'WeightMinor': '',
                    'ImageURLs': '',
                    **image_urls_dict,
                    **attributes_for_sku  # Add all attribute names and values
                }
                
                print(f'{new_row}\n')
                rows.append(new_row)

            # Create the DataFrame from rows
            template_df = pd.DataFrame(rows, columns=template_columns)
            
            # Save the result to an Excel file
            output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            template_df.to_excel(output_file, index=False)
            messagebox.showinfo("Success", "3D Seller Template Generated Successfully")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            
if __name__ == "__main__":
    root = tk.Tk()
    app = eBayTemplateGenerator(root)
    root.mainloop()
