import os
import requests
import pandas as pd
import time
import xlwings as xw

def get_coin_data(api_url):
    try:
        headers = {
            'User-Agent': 'Your User Agent'  # Add your user agent here
        }
        response = requests.get(api_url, headers=headers)
        response.raise_for_status()  # Raise an exception for bad responses
        data = response.json()
        return data
    
    except Exception as e:
        print(f"Error: {e}")
        return None
    

def get_excel_app():
    try:
        # Try connecting to an existing Excel instance
        app = xw.apps.active
        if app is None:
            # If there's no active Excel instance, create a new one
            app = xw.App(visible=True)
        return app
    
    except Exception as e:
        print(f"Error getting Excel app: {e}")
        return None

def open_or_get_workbook(app, workbook_name):
    try:
        # Try getting the active workbook from the existing Excel instance
        wb = app.books[workbook_name]
        if wb is None:
            # If there's no active workbook, create a new one
            wb = app.books.add()
        return wb
    
    except Exception as e:
        print(f"Error getting workbook: {e}")
        return None

def update_excel_data(api_url_markets, workbook_name):
    coin_data = get_coin_data(api_url_markets)

    if coin_data is not None:
        df = pd.DataFrame(coin_data)
        df_selected = df[['id', 'symbol', 'name', 'current_price', 'market_cap', 'market_cap_rank', 'total_volume', 
                          'high_24h', 'low_24h', 'price_change_24h', 'price_change_percentage_24h', 
                          'market_cap_change_24h', 'market_cap_change_percentage_24h', 'circulating_supply', 
                          'total_supply', 'max_supply', 'ath', 'ath_change_percentage', 'ath_date', 
                          'atl', 'atl_change_percentage', 'atl_date']]

        app = get_excel_app()

        if app is not None:
            try:
                wb = open_or_get_workbook(app, workbook_name)
                sheet = wb.sheets[0]

                sheet.clear_contents()
                data_to_write = [df_selected.columns.tolist()] + df_selected.values.tolist()
                sheet.range('A1').value = data_to_write

                print("Data updated in Excel.")
            except Exception as e:
                print(f"Error updating Excel data: {e}")
            finally:
                # Save the workbook and do not close Excel if it was already open
                save_path = os.path.join(os.getcwd(), "coin_data.xlsm")
                wb.save(save_path)
        else:
            print("Failed to connect to Excel.")
    else:
        print("Failed to fetch coin data.")

# Function to repeatedly update data in Excel every 5 seconds
def update_data_periodically(api_url_markets, workbook_name):
    app = get_excel_app()
    
    if app is not None:
        try:
            while True:
                update_excel_data(api_url_markets, workbook_name)
                time.sleep(3600)
        except KeyboardInterrupt:
            # Close Excel only if it was opened by the script
            if app.books.count == 1:  # Check if there is only one workbook open
                app.quit()

if __name__ == "__main__":
    api_url_markets = "https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=market_cap_desc&per_page=100&page=1&sparkline=false&locale=en"
    workbook_name = "coin_data.xlsm"  # Change this to your desired workbook name

    update_data_periodically(api_url_markets, workbook_name)
