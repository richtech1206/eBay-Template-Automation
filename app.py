import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
import zipfile
import io

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
            
    def generate_template(self):
        # try:
            # Read the source files
            vendor_items = pd.read_csv(self.vendor_items_file)
            vendor_items.columns = vendor_items.columns.str.lower().str.replace(' ', '')  # Normalize column names to lowercase and remove spaces

            competitor_data = pd.concat([pd.read_csv(file) for file in self.competitor_files])
            competitor_data.columns = competitor_data.columns.str.lower().str.replace(' ', '')  # Normalize column names to lowercase and remove spaces
            
            pies_data = {}
            if self.pies_file_type.get() == "single":
                pies_data = pd.read_excel(self.pies_file, sheet_name=None)
                pies_data = {sheet_name.lower().replace(' ', ''): pies_data[sheet_name].rename(columns=lambda x: x.lower().replace(' ', '')) for sheet_name in pies_data}  # Normalize sheet names and column names to lowercase and remove spaces
            elif self.pies_file_type.get() == "zip":
                with zipfile.ZipFile(self.pies_file, 'r') as z:
                    for filename in z.namelist():
                        if filename.endswith('.xlsx'):
                            with z.open(filename) as f:
                                sheet_data = pd.read_excel(f, sheet_name=None, engine='openpyxl')
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
            
            # Initialize the output DataFrame
            template_columns = [
                'Item ID', 'SKU', 'Parent SKU', 'Title', 'Description', 'Tags', 'MetaKeywords', 'MetaDescription', 'MobileDescription',
                'CategoryID', 'CategoryID2', 'StoreCategory', 'StoreCategory2', 'eBayCatalogID', 'eBayCatalogSearch', 'PromoteCampaign', 
                'PromoteRate', 'CarMake', 'CarModel', 'CarYear', 'CarTrim', 'CarEngine', 'CopyCarCompatibility', 'CarCompatibility', 
                'CarCompatibilityNotes', 'DeleteCarCompatibility', 'KType (TecDoc)', 'PrivateListing', 'Quantity', 'WarehouseQuantity', 
                'InventoryControl', 'Price', 'WholesalePrice', 'OriginalRetailPrice', 'BestOffer', 'BestOfferAccept', 'BestOfferDecline', 
                'VATPercent', 'ListingDuration', 'AuctionStartPrice', 'AuctionReservePrice', 'AuctionBINPrice', 'Condition', 'ConditionNote', 
                'C:ASIN', 'C:UPC', 'C:EAN', 'C:MPN', 'C:Brand', 'C:Manufacturer', 'C:Size', 'C:Color', 'C:Material', 'C:Product Type', 
                'Attribute1Name', 'Attribute1Value', 'Attribute2Name', 'Attribute2Value', 'Attribute3Name', 'Attribute3Value', 
                'Attribute4Name', 'Attribute4Value', 'Attribute5Name', 'Attribute5Value', 'Attribute6Name', 'Attribute6Value', 
                'Attribute7Name', 'Attribute7Value', 'Attribute8Name', 'Attribute8Value', 'Attribute9Name', 'Attribute9Value', 
                'Attribute10Name', 'Attribute10Value', 'Attribute11Name', 'Attribute11Value', 'Attribute12Name', 'Attribute12Value', 
                'Attribute13Name', 'Attribute13Value', 'Attribute14Name', 'Attribute14Value', 'Attribute15Name', 'Attribute15Value', 
                'Attribute16Name', 'Attribute16Value', 'Attribute17Name', 'Attribute17Value', 'Attribute18Name', 'Attribute18Value', 
                'Attribute19Name', 'Attribute19Value', 'Attribute20Name', 'Attribute20Value', 'Attribute21Name', 'Attribute21Value', 
                'Attribute22Name', 'Attribute22Value', 'Attribute23Name', 'Attribute23Value', 'Attribute24Name', 'Attribute24Value', 
                'Attribute25Name', 'Attribute25Value', 'Attribute26Name', 'Attribute26Value', 'Attribute27Name', 'Attribute27Value', 
                'Attribute28Name', 'Attribute28Value', 'Attribute29Name', 'Attribute29Value', 'Attribute30Name', 'Attribute30Value', 
                'CountryCode', 'Location', 'PostalCode', 'PolicyPayment', 'PolicyShipping', 'PolicyReturn', 'PackageType', 
                'MeasurementSystem', 'PackageLength', 'PackageWidth', 'PackageDepth', 'WeightMajor', 'WeightMinor', 'ImageURLs'
            ]
            # Add additional image columns up to Image10
            for i in range(1, 11):
                template_columns.append(f'Image {i}')
            
            brand_name = self.brand_name.get()
            
            # Aggregate all attributes with their values
            description_segment = pies_data['descriptionsegment']
            pies_template = pies_data['piestemplate']
            report_segment = pies_data['report']
            attributes_segment = pies_data['productattributessegment']
            
            # Get the list of all SKUs from vendor items
            all_skus = vendor_items['partnumber'].tolist()

            # Add attributes to the new_row dictionary
            attr_rows = attributes_segment[attributes_segment['partnumber'].isin(all_skus)]
            attributes = {}
            for _, attr_row in attr_rows.iterrows():
                attr_name = attr_row['attributename']
                attr_value = attr_row['productattribute']
                if attr_name in attributes:
                    attributes[attr_name].append(attr_value)
                else:
                    attributes[attr_name] = [attr_value]

            # Flatten the attributes dictionary to have a single list of key-value pairs
            flattened_attributes = []
            for attr_name, attr_values in attributes.items():
                flattened_attributes.extend([(attr_name, attr_value) for attr_value in attr_values])

            # Separate flattened_attributes into several parts with the same length as all_skus
            chunk_size = len(all_skus)
            separated_attributes = [flattened_attributes[i:i + chunk_size] for i in range(0, len(flattened_attributes), chunk_size)]

            # If the last chunk is smaller than chunk_size, pad it with empty tuples
            if separated_attributes[-1] and len(separated_attributes[-1]) < chunk_size:
                separated_attributes[-1].extend([('', '')] * (chunk_size - len(separated_attributes[-1])))

            # Collect rows for the output DataFrame
            rows = []

            for idx, row in vendor_items.iterrows():
                sku = row['partnumber']
                
                # Initialize fields
                title = condition_code = quantity = price = copycarcompatibility = None
                
                # Get data from competitor files
                competitor_row = competitor_data[competitor_data['partnumber'] == sku]
                if not competitor_row.empty:
                    title = competitor_row['title'].values[0]
                    condition_code = competitor_row['conditioncode'].values[0]
                    quantity = competitor_row['quantity'].values[0]
                    price = competitor_row['price'].values[0]
                    copycarcompatibility = competitor_row['copycarcompatabilityid'].values[0]
                
                # Set default values if none found
                title = title if title else 'N/A'
                condition_code = condition_code if condition_code else 'N/A'
                quantity = quantity if quantity else 'N/A'
                price = price if price else 'N/A'
                copycarcompatibility = copycarcompatibility if copycarcompatibility else 'N/A'
                
                # Description
                des_rows = description_segment[description_segment['partnumber'] == sku]
                descriptions = des_rows['description'].astype(str).tolist()
                description = ', '.join(descriptions)
                
                # Tags
                tags = f"{brand_name} - {datetime.now().strftime('%Y-%m-%d')}"
                
                # CategoryID
                part_terminology_rows = pies_template[pies_template['partnumber'] == sku]
                if not part_terminology_rows.empty:
                    part_terminology = part_terminology_rows['partterminologyname'].values[0]
                    category_id = category_id_mapping.loc[category_id_mapping['partterminologyname'] == part_terminology, 'category_id']
                    category_id = category_id.values[0] if not category_id.empty else '#N/A'
                else:
                    category_id = '#N/A'
                    
                # Product Type
                product_type_rows = report_segment[report_segment['partnumber'] == sku]
                if not product_type_rows.empty:
                    product_type = product_type_rows['partterminologyname'].values[0]
                else:
                    product_type = 'N/A'
                    
                # UPC
                gtin_rows = pies_template[pies_template['partnumber'] == sku]
                if not gtin_rows.empty:
                    gtin = gtin_rows['itemlevelgtin'].values[0]
                    upc = str(gtin)[2:] if pd.notna(gtin) else 'N/A'
                else:
                    upc = 'N/A'
                
                # Brand
                brand = brand_name
                
                # Package Dimensions and Weight
                package_rows = report_segment[report_segment['partnumber'] == sku]
                if not package_rows.empty:
                    package_length = package_rows['length(in)'].values[0]
                    package_width = package_rows['width(in)'].values[0]
                    package_depth = package_rows['height(in)'].values[0]
                else:
                    package_length = package_width = package_depth = 'N/A'
                
                weight_major_rows = report_segment[report_segment['partnumber'] == sku]
                if not weight_major_rows.empty:
                    weight_major = weight_major_rows['weight(lbs)'].values[0]
                else:
                    weight_major = 'N/A'                            
                
                # C:MPN, C:Brand, C:Manufacturer
                c_mpn = row['partnumber']
                c_brand = row['linecode']
                c_manufacturer = row['linecode']
                
                # Get image URLs
                image_rows = image_list[image_list['partnumber'] == sku]
                image_urls = image_rows.sort_values('sortorder')['url'].tolist()
                image_urls_dict = {f'Image {i+1}': url for i, url in enumerate(image_urls)}
                
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
                    **image_urls_dict
                }
                
                # Add attribute name and value pairs to the new_row dictionary
                for i, attr_chunk in enumerate(separated_attributes, start=1):
                    if i > 30:  # Limit to 30 attributes to match the column names in the template
                        break
                    attr_name_col = f'Attribute{i}Name'
                    attr_value_col = f'Attribute{i}Value'
                    if idx < len(attr_chunk):  # Ensure index is within the chunk's range
                        attr_name, attr_value = attr_chunk[idx]
                    else:
                        attr_name, attr_value = '', ''
                    new_row[attr_name_col] = attr_name
                    new_row[attr_value_col] = attr_value
                print(f'{new_row}\n')
                rows.append(new_row)
            
            # Create the DataFrame from rows
            template_df = pd.DataFrame(rows)
            
            # Save the result to an Excel file
            output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            template_df.to_excel(output_file, index=False)
            messagebox.showinfo("Success", "3D Seller Template Generated Successfully")
        # except Exception as e:
        #     messagebox.showerror("Error", f"An error occurred: {str(e)}")
            
if __name__ == "__main__":
    root = tk.Tk()
    app = eBayTemplateGenerator(root)
    root.mainloop()
