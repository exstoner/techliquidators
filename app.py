import os
import requests
import pandas as pd
import webbrowser
import cloudconvert

# Set your CloudConvert API key here
api_key = ''

def download_auction_list(auction_id):
    url = f"https://techliquidators.ca/query2Excel.cfm?i={auction_id}"
    response = requests.get(url)
    directory = f"{auction_id}"
    os.makedirs(directory, exist_ok=True)
    file_path = os.path.join(directory, f"{auction_id}.xls")
    with open(file_path, 'wb') as file:
        file.write(response.content)
    return file_path

def convert_xls_to_csv(file_path):
    cloudconvert.configure(api_key=api_key)

    try:
        job = cloudconvert.Job.create(payload={
            "tasks": {
                "import-1": {
                    "operation": "import/upload"
                },
                "task-1": {
                    "operation": "convert",
                    "input": "import-1",
                    "input_format": "xls",
                    "output_format": "csv",
                    "engine": "libreoffice"
                },
                "export-1": {
                    "operation": "export/url",
                    "input": "task-1"
                }
            }
        })
    except Exception as e:
        raise Exception(f"Failed to create CloudConvert job: {e}")

    try:
        upload_task_id = job['tasks'][0]['id']
    except KeyError as e:
        raise KeyError(f"Failed to find 'tasks' in job response: {job}")

    upload_task = cloudconvert.Task.find(id=upload_task_id)

    upload_url = upload_task['result']['form']['url']
    upload_params = upload_task['result']['form']['parameters']

    with open(file_path, 'rb') as file:
        response = requests.post(upload_url, data=upload_params, files={'file': file})

    job = cloudconvert.Job.wait(id=job['id'])

    try:
        export_task = [task for task in job['tasks'] if task['name'] == 'export-1'][0]
    except KeyError as e:
        raise KeyError(f"Failed to find 'tasks' in job response: {job}")

    file_url = export_task['result']['files'][0]['url']
    csv_file_path = file_path.replace('.xls', '.csv')
    response = requests.get(file_url)
    with open(csv_file_path, 'wb') as file:
        file.write(response.content)
    
    return csv_file_path

def process_auction_list(auction_id):
    try:
        # Download the auction list
        xls_file = download_auction_list(auction_id)

        # Convert the .xls file to .csv using CloudConvert
        csv_file = convert_xls_to_csv(xls_file)

        # Load the .csv file with pandas
        data = pd.read_csv(csv_file)

        # Extract the relevant columns
        titles_and_skus = data[['Title', 'BBY SKU', 'MFG Name']]

        # Dictionary to keep track of item quantities
        item_counts = {}

        # Base URL for the links
        base_url = "https://www.bestbuy.ca/en-ca/product/"

        # Count occurrences of each unique item
        for item in titles_and_skus.to_dict(orient='records'):
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

        # Convert item_counts to a list of tuples and sort by quantity in descending order
        sorted_items = sorted(item_counts.items(), key=lambda x: x[1], reverse=True)

        # Calculate the total quantity of all items
        total_quantity = sum(item_counts.values())

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
            <h1>Items List</h1>
            <table>
                <tr>
                    <th>Manufacturer</th>
                    <th>Item</th>
                    <th>Image</th>
                    <th>Quantity</th>
                </tr>
        """

        # Append each item to the HTML content
        for (title, bby_sku, mfg_name), qty in sorted_items:
            link = f"{base_url}{bby_sku}"
            sku_str = str(bby_sku)
            image_url = f"https://multimedia.bbycastatic.ca/multimedia/products/500x500/{sku_str[:3]}/{sku_str[:5]}/{sku_str}.jpg"
            html_content += f"""
                <tr>
                    <td>{mfg_name}</td>
                    <td>{title}</td>
                    <td><a href="{link}" target="_blank"><img src="{image_url}" alt="Image for {title}" width="100"></a></td>
                    <td>{qty}</td>
                </tr>
            """

        # Append the total quantity to the HTML content
        html_content += f"""
                <tr>
                    <td colspan="3"><strong>Total Quantity</strong></td>
                    <td><strong>{total_quantity}</strong></td>
                </tr>
            </table>
        </body>
        </html>
        """

        # Save the HTML content to a file
        directory = f"{auction_id}"
        output_file = os.path.join(directory, f'{auction_id}.html')
        with open(output_file, 'w') as file:
            file.write(html_content)

        print(f"HTML file '{output_file}' has been generated.")

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
            auction_id = input("Please enter the auction ID: ")
            process_auction_list(auction_id)
        elif choice == '2':
            print("Exiting the program.")
            break
        else:
            print("Invalid choice. Please try again.")

if __name__ == "__main__":
    main()
