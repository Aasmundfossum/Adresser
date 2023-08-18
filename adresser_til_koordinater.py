import requests
import pandas as pd
import numpy as np
import re

# ------------------------------------------------------------------------------------------------------------------------------------------------ #
# Her kan det lastes opp en Excel-fil som inneholder en liste med adresser, og programmet gir ut en ny Excel-fil ('Adresser_med_koordinater.xlsx') med adressene sammen med deres koordinater.
# Noen adresser fungerer ikke, og gir blank rute i output-filen. Koordinatene til disse må evt. finnes manuelt. 
# Kun disse tre linjene trenger å endres:

filnavn = 'Copy of 2023.08.07 FFAS - Fjernvarmekunder.xlsx'         # Navn på Excel-fil med adresser
arknavn = 'Ark1'                                                    # Navn på arket i Excel-filen med addresser
sted = ', Fredrikstad'                                              # Navn på sted/by som adressene er i. Gjør at det velges riktig sted dersom flere steder har samme gatenavn og nummer
                                                                    # (noe som ofte er tilfelle). Komma og mellomrom må inkluderes. Kan evt. være blank, med risiko for at feil koordinater velges.
# ------------------------------------------------------------------------------------------------------------------------------------------------ #



def adresse_til_koordinat(adresse_str):
    # Replace YOUR_API_KEY with your actual API key. Sign up and get an API key on https://www.geoapify.com/ 
    API_KEY = "400f888f4da9461387721ccbd1a0e0db"

    # Define the address to geocode
    #address = "1600 Amphitheatre Parkway, Mountain View, CA"
    address = adresse_str
    address = str(address)+str(sted)

    # Build the API URL
    url = f"https://api.geoapify.com/v1/geocode/search?text={address}&limit=1&apiKey={API_KEY}"

    # Send the API request and get the response
    response = requests.get(url)

    # Check the response status code
    if response.status_code == 200:
        # Parse the JSON data from the response
        data = response.json()

        # Extract the first result from the data
        try:
            result = data["features"][0]
            # Extract the latitude and longitude of the result
            latitude = result["geometry"]["coordinates"][1]
            longitude = result["geometry"]["coordinates"][0]
        except:
            try:
                if 'gate' in address:
                    address = address.replace('gate',' gate')
                elif 'vei' in address:
                    address = address.replace('vei',' vei')
                else:
                    pattern = r"_(\d+)_[a-zA-Z]+?$"
                    match = re.search(pattern, address)
    
                    if match:
                        substring = match.group(0)
                        address = address.replace(substring, "")
                

                url = f"https://api.geoapify.com/v1/geocode/search?text={address}&limit=1&apiKey={API_KEY}"
                response = requests.get(url)
                data = response.json()
                result = data["features"][0]
                latitude = result["geometry"]["coordinates"][1]
                longitude = result["geometry"]["coordinates"][0]
            except:
                latitude = float('nan')
                longitude = float('nan')

    return latitude, longitude

adresse_fil = pd.read_excel(filnavn,sheet_name=arknavn)

breddegrad = np.zeros(len(adresse_fil))
lengdegrad = np.zeros(len(adresse_fil))

for i in range(0,len(adresse_fil)):
    breddegrad[i],lengdegrad[i] = adresse_til_koordinat(str(adresse_fil.iloc[i,0]))
    print(i)

breddegrad = pd.DataFrame(breddegrad,columns = ['Breddegrad'])
lengdegrad = pd.DataFrame(lengdegrad, columns = ['Lengdegrad'])

ferdig = pd.concat([adresse_fil,breddegrad,lengdegrad,],axis=1)
ferdig.to_excel('Adresser_med_koordinater.xlsx')

print('Ferdig kjørt')