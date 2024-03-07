"""
Software prereqs:
- A version of python installed
- Anaconda prompt install with the command "xlwings addon install"
- Text editer to update the python script to whatever cryptos you 
  want and for the correct columns
- Making a copy of the excel sheet into one that runs macros



Instructions: 
- Install a version of python
- After that, type "pip install requests xlwings" into a terminal
- Install Anaconda prompt, and once it's opened type "xlwings addon install"
- Change the name of this file to whatever you're excel file is named
- Then, go down to line 52 in this file, and update the list to whatever cryptos 
  you want.
- Then go to line 53, and update it to whatever column the crypto names
  on your excel sheet are on.
- Go to line 54 and and update it to whatever column you what the prices to be on
- Change Capitalized to false if the names of the cryptos on your spreadsheet are all lowercased
- Open your excel file and go to file, save a copy, and select xlsm
- Open that copy and go the xlwings tab at the top, but the correct paths
  in for the interpter, python script, and then open VB
- In there just copy and paste the code I'll provide somewhere and change the names
  of the files if nessary.
- That's it, just close the file and open it up again to test it.
"""

"""
INSTRUCTIONS FOR THE VB PART

In VB just find the "This workbook" thing on the left and paste:
    Sub RunPythonScript()
        RunPython "import Investment; Investment.main()"
    End Sub

Save it, and then right click on the "This workbook" thing and click insert then module.
Then paste this into the new module:

    Sub RunPythonScript()
        RunPython "import Investment; Investment.main()"
    End Sub

But instead of Investment, type the name of the python script (which should match you excel file)
leave the .main()

Thats it (make sure to save it before exiting)
"""


crypto_names = ["Ethereum", "Vechain", "Cardano"]
cryto_column = 'A'
price_coulumn = 'J'
Capitalized = True

# Uses Coin Gecko API get method to fetch the prices of the cryptos in the list passed to it
# Full list of all supported cryptos can be found somewhere on their site
import requests

def fetch_crypto_prices(crypto_names):
    crypto_names_str = ",".join([name.lower() for name in crypto_names])
    url = f"https://api.coingecko.com/api/v3/simple/price?ids={crypto_names_str}&vs_currencies=usd"
    response = requests.get(url)
    data = response.json()
    prices = {name.capitalize(): data[name.lower()]['usd'] for name in crypto_names if name.lower() in data}
    return prices




# Uses xlwings to look in the A column and find the crypto names of the crypto's in the list
# Capitalization is important, in this it's set to find the cryptos capital name, switch to regular
# if not the case. 
import xlwings as xw

def update_excel_file(excel_file, prices):

    wb = xw.Book.caller()
    sht = wb.sheets('sheet1')

    for coin, price in prices.items():
        coin_capitalized = coin.capitalize()
        # Find the row (Assuming coin names are in column A)
        for i in range (1,21):
            cell = sht.range(f"{cryto_column}{i}") # Update this to be whatever column has the crypto name
            if Capitalized == True:
                if cell.value == coin_capitalized:
                    sht.range(f"{price_coulumn}{i}").value = price # Update this to be whatever column the "Todays Price" feild is
            else:
                if cell.value == coin:
                    sht.range(f"{price_coulumn}{i}").value = price # Update this to be whatever column the "Todays Price" feild is




# Contains the list of cryptos involved. The excel_file variable only exists for testing and
# developing  purposes, as I would use it to update the c
def main():
    
    excel_file = r'C:\Users\ryanc\OneDrive\Documents\CryptoInvestment\Investment.xlsx'  # Update to your actual file path
    
    prices = fetch_crypto_prices(crypto_names)

    update_excel_file(excel_file, prices)
    
    print("Excel file updated with the latest cryptocurrency prices.")


# Update would be that it would find where the "Todays Price" Feild is and update it like that
# Also it would only update the first cell of a crypto cluster.
# Leave the excel_file argument for later testing       
