import os
import pandas as pd
import webbrowser

def process_auction_list():
    # Prompt user for the path to the CSV file
    file_path = input("Please enter the path to the CSV file: ")

    try:
        # Load the CSV file
        data = pd.read_csv(file_path)

        # Extract the relevant columns
        titles_and_skus = data[['Auction ID', 'Title', 'BBY SKU', 'MFG Name']]

        # Dictionary to keep track of item quantities
        item_counts = {}

        # Base URL for the links
        base_url = "https://www.bestbuy.ca/en-ca/product/"

        # Count occurrences of each unique item
        for item in titles_and_skus.to_dict(orient='records'):
            auction_id = item['Auction ID']
            title = item['Title']
            bby_sku = item['BBY SKU']
            mfg_name = item['MFG Name']

            # Skip items with missing SKUs
            if pd.isna(bby_sku):
                continue

            bby_sku = int(bby_sku)  # Ensure the SKU is an integer

            if (title, bby_sku, mfg_name) in item_counts:
                item_counts[(title, bby_sku, mfg_name)] += 1
            else:
                item_counts[(title, bby_sku, mfg_name)] = 1

        # Get the Auction ID for the filename
        auction_id = titles_and_skus['Auction ID'].iloc[0]

        # Convert item_counts to a list of tuples and sort by quantity in descending order
        sorted_items = sorted(item_counts.items(), key=lambda x: x[1], reverse=True)

        # Generate HTML content
        html_content = """
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Items List</title>
            <style>
                table {
                    width: 100%;
                    border-collapse: collapse;
                }
                table, th, td {
                    border: 1px solid black;
                }
                th, td {
                    padding: 10px;
                    text-align: left;
                }
            </style>
        </head>
        <body>
            <table>
                <tr>
                    <th>Item</th>
                    <th>Manufacturer</th>
                    <th>Quantity</th>
                    <th>Image</th>
                </tr>
        """

        # Append each item to the HTML content
        for (title, bby_sku, mfg_name), qty in sorted_items:
            link = f"{base_url}{bby_sku}"
            sku_str = str(bby_sku)
            image_url = f"https://multimedia.bbycastatic.ca/multimedia/products/500x500/{sku_str[:3]}/{sku_str[:5]}/{sku_str}.jpg"
            html_content += f"""
                <tr>
                    <td>{title}</td>
                    <td>{mfg_name}</td>
                    <td>{qty}</td>
                    <td><a href="{link}" target="_blank"><img src="{image_url}" alt="Image for {title}" width="100"></a></td>
                </tr>
            """

        # Close the HTML tags
        html_content += """
            </table>
        </body>
        </html>
        """

        # Create the "Auction Lists" directory if it doesn't exist
        output_dir = 'Auction Lists'
        os.makedirs(output_dir, exist_ok=True)

        # Save the HTML content to a file
        output_file = os.path.join(output_dir, f'auction_id_{auction_id}.html')
        with open(output_file, 'w') as file:
            file.write(html_content)

        print(f"'{output_file}' has been generated.")

        # Open the generated HTML file in the default web browser
        webbrowser.open(f'file://{os.path.abspath(output_file)}')
    except Exception as e:
        print(f"An error occurred: {e}")

def main():
    while True:
        print("Select an option:")
        print("[1] Get auction list")
        print("[2] Exit")
        choice = input("Enter your choice: ")

        if choice == '1':
            process_auction_list()
        elif choice == '2':
            print("Exiting the program.")
            break
        else:
            print("Invalid choice. Please try again.")

if __name__ == "__main__":
    main()
