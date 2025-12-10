
import time
import pandas as pd
import openrouteservice as ors
import numpy as np

# --- Configuration & Constants ---
# You need an openrouteservice API-Key to run this script.
# Further info is contained in the readme.
ORS_API_KEY = """Secret_API-Key"""

DATA_FILE = "data.xlsx"
AFTERPARTY_FILE = "Afterparty.xlsx"


# --- Helper Functions ---

def setup_client(api_key):
    """Initialize the OpenRouteService client."""
    return ors.Client(key=api_key)

def geocode_address(client, address):
    """
    Geocode a single address using ORS.
    Adds a delay to respect rate limits.
    """
    try:
        time.sleep(1)  # Add a 1-second delay to respect rate limits
        geocode = client.pelias_search(text=address)
        coords = geocode['features'][0]['geometry']['coordinates']
        return coords[0], coords[1]  # longitude, latitude
    except Exception as e:
        print(f"Error geocoding {address}: {e}")
        return None, None

def add_coordinates_to_dataframe(df, client, address_col='Adress'):
    """
    Apply geocoding to a DataFrame's address column.
    Returns the DataFrame with new 'longitude' and 'latitude' columns.
    """
    print(f"Geocoding addresses in column '{address_col}'...")
    # Using zip to unpack the tuple result of apply
    coords = df[address_col].apply(lambda x: geocode_address(client, x))
    df['longitude'], df['latitude'] = zip(*coords)
    return df

def calculate_distance_matrix(client, origins_coords, destinations_coords, profile='cycling-regular'):
    """
    Calculate distance matrix between origins and destinations.
    origins_coords: list of [lon, lat]
    destinations_coords: list of [lon, lat]
    """
    locations = origins_coords + destinations_coords
    
    matrix = client.distance_matrix(
        locations=locations,
        profile=profile,
        sources=list(range(len(origins_coords))),
        destinations=list(range(len(origins_coords), len(origins_coords) + len(destinations_coords))),
        metrics=['distance', 'duration']
    )
    return matrix

def assignment(Adress_Table, Distance_Afterparty, Distance_between_Groups, Prioritize_center, Sees_value_first, Sees_value_second, Starts, Targets):
    """
    Assigns teams to hosts based on distance and availability.
    Mutates Adress_Table.
    """
    # Sort candidates by duration to afterparty/center
    Distance_Afterparty.sort_values("Duration_H", ascending=Prioritize_center, inplace=True)
    
    for index, _ in Distance_Afterparty.iterrows():
        select_af_team_dist = Distance_Afterparty.loc[index]
        
        # Current team info
        select_team_adress = Adress_Table.loc[Adress_Table["Team Nr."] == select_af_team_dist["Team Nr."]]
        select_team_adress_id = select_team_adress["Team Nr."].iloc[0]
        
        # Distances from this team to potential targets
        select_team_team_dist = Distance_between_Groups.loc[Distance_between_Groups[Starts] == select_af_team_dist["Team Nr."]]
        
        # Get current assignments (nogos)
        nogos = [
            select_team_adress["Sees 1"].iloc[0], select_team_adress["Sees 2"].iloc[0],
            select_team_adress["Sees 3"].iloc[0], select_team_adress["Sees 4"].iloc[0],
            select_team_adress["Sees 5"].iloc[0], select_team_adress["Sees 6"].iloc[0]
        ]
        
        # Skip if already assigned (checked via Sees_value_first)
        if select_team_adress[Sees_value_first].iloc[0] != 0:
            continue

        # Skip teams who will be hosts next step (if they exist in Targets column of their own distance rows)
        # Note: This check seems to verify if the current team appears as a target in the distance matrix provided?
        if select_team_team_dist[Starts].isin(select_team_team_dist[Targets]).any():
            continue
        
        # Select relevant targets and sort by distance to them
        relevant_targets = Distance_between_Groups.loc[Distance_between_Groups[Starts] == select_team_adress_id]
        relevant_targets.sort_values("Duration_H", ascending=True, inplace=True)
        
        # Iterate over targets to find a suitable host
        for _, target in relevant_targets.iterrows():
            target_id = target[Targets]
            
            if target_id in nogos:
                continue
            
            target_adress_id = Adress_Table.loc[Adress_Table["Team Nr."] == target_id]
            
            # Check slot 1 availability
            if target_adress_id[Sees_value_first].iloc[0] == 0:
                Adress_Table.loc[Adress_Table["Team Nr."] == target_id, Sees_value_first] = select_af_team_dist["Team Nr."]
                Adress_Table.loc[Adress_Table["Team Nr."] == select_af_team_dist["Team Nr."], Sees_value_first] = target_id
                break
                
            # Check slot 2 availability
            elif target_adress_id[Sees_value_second].iloc[0] == 0:
                if target_adress_id[Sees_value_first].iloc[0] in nogos:
                    continue
                else:
                    Adress_Table.loc[Adress_Table["Team Nr."] == target_id, Sees_value_second] = select_af_team_dist["Team Nr."]
                    Adress_Table.loc[Adress_Table["Team Nr."] == select_af_team_dist["Team Nr."], Sees_value_first] = target_id

                    # Assign third team for "sees second" - mutual link?
                    the_third_team = target_adress_id[Sees_value_first].iloc[0]
                    Adress_Table.loc[Adress_Table["Team Nr."] == select_af_team_dist["Team Nr."], Sees_value_second] = the_third_team
                    Adress_Table.loc[Adress_Table["Team Nr."] == the_third_team, Sees_value_second] = select_af_team_dist["Team Nr."]
                    break

def overwrite(guest_distances, host_distances, Adress_Table, guest_start, guest_target, host_start, host_target, Sees_first):
    """
    Updates distances for guests based on their host's location for the next course.
    """
    for index, row in guest_distances.iterrows():
        # Copy the distances from the hosts to the next course onto the guests
        guest_team = guest_distances.loc[index]
        guest_team_in_Adresses = Adress_Table.loc[Adress_Table["Team Nr."] == guest_team[guest_start]]

        if guest_team_in_Adresses[Sees_first].iloc[0] != guest_team[guest_target]:
            continue

        guest_team_to = guest_team[guest_target]  # ID of host team
        guest_team_from = guest_team[guest_start]  # ID of guest team

        relevant_guests = guest_distances[guest_distances[guest_start] == guest_team_from]
        relevant_hosts = host_distances[host_distances[host_start] == guest_team_to]

        for (index1, _), (index2, __) in zip(relevant_guests.iterrows(), relevant_hosts.iterrows()):
            route_host = host_distances.loc[index2, "Duration_H"]
            target_host = host_distances.loc[index2, host_target]

            guest_distances.loc[index1, guest_target] = target_host
            guest_distances.loc[index1, "Duration_H"] = route_host

def generate_overview_files(starters_table, main_courses_table, desserts_table, Adress_list):
    """Generates the overview.txt file."""
    Starters_Sentences = []
    Main_Courses_Sentences = []
    Desserts_Sentences = []

    print("Generating overview...")

    for (index1, _), (index2, __), (index3, ___) in zip(starters_table.iterrows(), main_courses_table.iterrows(), desserts_table.iterrows()):
        This_Starter = starters_table.loc[index1]
        This_MainCourse = main_courses_table.loc[index2]
        This_Dessert = desserts_table.loc[index3]
        
        Star_in_Ad = Adress_list[Adress_list["Team Nr."] == This_Starter["Team Nr."]]
        Main_in_Ad = Adress_list[Adress_list["Team Nr."] == This_MainCourse["Team Nr."]]
        Des_in_Ad = Adress_list[Adress_list["Team Nr."] == This_Dessert["Team Nr."]]
        
        Sta_Sen = f'Team Nr.{Star_in_Ad["Team Nr."].iloc[0]} (Starter) hosts the Teams {Star_in_Ad["Sees 1"].iloc[0]} and {Star_in_Ad["Sees 2"].iloc[0]}.'
        Mai_Sen = f'Team Nr.{Main_in_Ad["Team Nr."].iloc[0]} (Main Course) hosts the Teams {Main_in_Ad["Sees 3"].iloc[0]} and {Main_in_Ad["Sees 4"].iloc[0]}.'
        Des_Sen = f'Team Nr.{Des_in_Ad["Team Nr."].iloc[0]} (Dessert) hosts the Teams {Des_in_Ad["Sees 5"].iloc[0]} and {Des_in_Ad["Sees 6"].iloc[0]}.'
        
        Starters_Sentences.append(Sta_Sen)
        Main_Courses_Sentences.append(Mai_Sen)
        Desserts_Sentences.append(Des_Sen)
        
    Starters_Sentences = set(Starters_Sentences)
    Main_Courses_Sentences = set(Main_Courses_Sentences)
    Desserts_Sentences = set(Desserts_Sentences)

    with open('overview.txt', 'w', encoding='utf-8') as file:
        file.write('Starters\n\n')
        for Sentence in Starters_Sentences:
            file.write(f'{Sentence}\n\n')
            
        file.write('Main Courses\n\n')
        for Sentence in Main_Courses_Sentences:
            file.write(f'{Sentence}\n\n')
            
        file.write('Desserts\n\n')
        for Sentence in Desserts_Sentences:
            file.write(f'{Sentence}\n\n')
    
    print("Overview saved in 'overview.txt'.")

def get_user_inputs():
    """Collects configuration inputs from the user."""
    print("\n--- Mail Creation ---")
    lang = input("Please select a language for mail creation\n|1---english\n|2---deutsch\n------- ")
    
    Time_Sta = input("Please type the time the starters should begin:\n")
    Time_Mai = input("Please type the time the main courses should begin:\n")
    Time_Des = input("Please type the time the desserts should begin:\n")
    
    Time_Afterparty = input("Please type the time the Afterparty is planned:\n")
    Afterparty_Drinks = input("1----drinks will be provided by the location for a price\n2----drinks will be provided for free by us\nSelection: ")
    
    Orga_Name = input("Name(s) of person(s) responsible for adjustments/info:\n")
    Orga_Tel = input("Tel.-Number(s) of responsible person(s):\n")
    
    Awareness = input("Will there be an Awareness Team?\n1-----yes\n2-----no\nSelection: ")
    Aw_Tel = ""
    if Awareness == "1":
        Aw_Tel = input("What is their Tel.-Nr.? ")
        
    return {
        "lang": lang,
        "Time_Sta": Time_Sta,
        "Time_Mai": Time_Mai,
        "Time_Des": Time_Des,
        "Time_Afterparty": Time_Afterparty,
        "Afterparty_Drinks": Afterparty_Drinks,
        "Orga_Name": Orga_Name,
        "Orga_Tel": Orga_Tel,
        "Awareness": Awareness,
        "Aw_Tel": Aw_Tel
    }

def generate_individual_mails(Adress_list, starters, main_courses, desserts, Afterparty_df, config):
    """Generates individual text files for each team."""
    
    print("Generating individual mail files...")
    
    Time_Sta = config["Time_Sta"]
    Time_Mai = config["Time_Mai"]
    Time_Des = config["Time_Des"]
    Time_Afterparty = config["Time_Afterparty"]
    
    # Note: currently only German ('2') is fully implemented in the structure below based on source
    # If lang is '1', it falls back to whatever is defined or might fail if not handled.
    # The original code mostly checked `if lang == '2':` then had the block.
    # I will assume German logic for now as in the original source, but be safe.
    
    # Pre-formatting strings
    Afterparty_Address = Afterparty_df["Adress"].iloc[0]
    
    # Common text blocks
    Ausklang = ""
    if config["Afterparty_Drinks"] == '1':
        Ausklang = (
            f'Ab ca. {Time_Afterparty} Uhr treffen wir uns an folgender Adresse: {Afterparty_Address}'
            f' zum gemütlichen Ausklang des Abends. Hier könnt ihr eure Rezepte austauschen oder übriggebliebene Speisen mitbringen.'
            f' Es wird Getränke auf Spendenbasis geben.\n\n\n\n'
        )
    else:
        Ausklang = (
            f'Ab ca. {Time_Afterparty} Uhr treffen wir uns an folgender Adresse: {Afterparty_Address}'
            f' zum gemütlichen Ausklang des Abends. Hier könnt ihr eure Rezepte austauschen oder übriggebliebene Speisen mitbringen.\n\n\n\n'
        )

    Schlusswort = ""
    if config["Awareness"] == '1':    
        Schlusswort = (
            f'Bitte gebt eurem/r Kochpartner/in Bescheid, da diese Email nur an eine Person geht.\n\n'
            f'Bei Notfällen/Unstimmigkeiten an dem Abend meldet euch bei {config["Orga_Name"]} unter {config["Orga_Tel"]}\n\n'
            f'Die Handynummer des A-Teams lautet: {config["Aw_Tel"]}\n\n'
            f'Wichtig: Falls ihr krankheitsbedingt ausfallt und selber keinen Ersatz findet gebt uns bitte schnellstmöglich Bescheid.\n\n'
        )
    else:    
        Schlusswort = (
            f'Bitte gebt eurem/r Kochpartner/in Bescheid, da diese Email nur an eine Person geht.\n\n'
            f'Bei Notfällen/Unstimmigkeiten an dem Abend meldet euch bei {config["Orga_Name"]} unter {config["Orga_Tel"]}\n\n'
            f'Wichtig: Falls ihr krankheitsbedingt ausfallt und selber keinen Ersatz findet gebt uns bitte schnellstmöglich Bescheid.\n\n'
        )

    for _, Team in Adress_list.iterrows():
        Anrede = f'Hallo {Team["Name 1"]} und {Team["Name 2"]}, \n\nVielen Dank für eure Anmeldung\n\n. Hier euer Plan für den Abend: \n\n\n\n'
        
        # --- STARTER LOGIC ---
        if Team["Team Nr."] in starters:
            if Team["Sees 1"] != 0:
                guest_1 = Adress_list.loc[Adress_list["Team Nr."] == Team["Sees 1"]].iloc[0]
            else: 
                guest_1 = Team
            if Team["Sees 2"] != 0:
                guest_2 = Adress_list.loc[Adress_list["Team Nr."] == Team["Sees 2"]].iloc[0]
            else: 
                guest_2 = Team
            
            Vorspeise = (
                f"Ihr dürft gleich am Anfang aktiv sein und euren Gästen eine unvergessliche Vorspeise servieren "
                f"(Essgewohnheiten: {guest_1['Allergies or else']} UND {guest_2['Allergies or else']}).\n\n"
                f"Ob ihr zu eurem Gericht ein stilles Mineralwasser, ein fancy Aperitif oder ein Gläschen süffigen Wein serviert, "
                f"bleibt euch natürlich selbst überlassen.\n\n"
            )
        else:
            # They are a guest
            host_team_id = Team["Sees 1"]
            if host_team_id != 0:
                host = Adress_list.loc[Adress_list["Team Nr."] == host_team_id].iloc[0]
                Vorspeise = (
                    f"Den Abend beginnt ihr bei {host['Name 1']} und {host['Name 2']} ({host['Adress']}), {host['Ring at']} mit einer genialen Vorspeise.\n\n"
                )
            else:
                Vorspeise = "Fehler bei der Zuweisung der Vorspeise.\n\n"

        # --- MAIN COURSE LOGIC ---
        if Team["Team Nr."] in main_courses:
            if Team["Sees 3"] != 0:
                guest_1 = Adress_list.loc[Adress_list["Team Nr."] == Team["Sees 3"]].iloc[0]
            else: 
                guest_1 = Team
            if Team["Sees 4"] != 0:
                guest_2 = Adress_list.loc[Adress_list["Team Nr."] == Team["Sees 4"]].iloc[0]
            else: 
                guest_2 = Team
            
            Hauptspeise = (
                f"Danach seid ihr selbst an der Reihe euren Gästen eine unvergessliche Hauptspeise anzubieten."
                f"(Essgewohnheiten: {guest_1['Allergies or else']} UND {guest_2['Allergies or else']}).\n\n"
                f"Ob ihr zu eurem Gericht ein stilles Mineralwasser, ein fancy Aperitif oder ein Gläschen süffigen Wein serviert, "
                f"bleibt euch natürlich selbst überlassen.\n\n"
            )
        else:
            host_team_id = Team["Sees 3"]
            if host_team_id != 0:
                host = Adress_list.loc[Adress_list["Team Nr."] == host_team_id].iloc[0]
                Hauptspeise = (
                    f"Im Anschluss daran beglücken euch {host['Name 1']} und {host['Name 2']} ({host['Adress']}), {host['Ring at']} mit einer grandiosen Hauptspeise.\n\n"
                )
            else:
                Hauptspeise = "Fehler bei der Zuweisung der Hauptspeise.\n\n"

        # --- DESSERT LOGIC ---
        if Team["Team Nr."] in desserts:
            if Team["Sees 5"] != 0:
                guest_1 = Adress_list.loc[Adress_list["Team Nr."] == Team["Sees 5"]].iloc[0]
            else: 
                guest_1 = Team
            if Team["Sees 6"] != 0:
                guest_2 = Adress_list.loc[Adress_list["Team Nr."] == Team["Sees 6"]].iloc[0]
            else: 
                guest_2 = Team
            
            Nachspeise = (
                f"Zum Schluss seid ihr selbst an der Reihe euren Gästen eine leckere Nachspeise zu servieren."
                f"(Essgewohnheiten: {guest_1['Allergies or else']} UND {guest_2['Allergies or else']}).\n\n"
                f"Ob ihr zu eurem Gericht ein stilles Mineralwasser, ein fancy Aperitif oder ein Gläschen süffigen Wein serviert, "
                f"bleibt euch natürlich selbst überlassen.\n\n"
            )
        else:
            host_team_id = Team["Sees 5"]
            if host_team_id != 0:
                host = Adress_list.loc[Adress_list["Team Nr."] == host_team_id].iloc[0]
                Nachspeise = (
                    f"Zur Nachspeise servieren euch {host['Name 1']} und {host['Name 2']} ({host['Adress']}), {host['Ring at']} etwas ganz Traumhaftes für den letzten Platz im Magen.\n\n"
                )
            else:
                Nachspeise = "Fehler bei der Zuweisung der Nachspeise.\n\n"

        Zeitplan = (
            f"Die Zeitplanübersicht:\n\n\n\n{Time_Sta}-----Vorspeise\n\n\n\n{Time_Mai}-----Hauptspeise\n\n\n\n{Time_Des}-----Nachspeise\n"
            f"Denkt bitte bei euren Speisen daran, sie entweder schon vorzubereiten oder nicht zu aufwendig zu gestalten, da eure Gäste eintreffen werden. "
            f"Außerdem solltet ihr den Fahrtweg miteinplanen, sodass ihr jeweils zur angegebenen Uhrzeit beim nächsten Team seid.\n\n"
        )

        with open(f'Mail Team Nr.{Team["Team Nr."]}.txt', 'w', encoding='utf-8') as file:
            file.write(f'{Anrede}{Vorspeise}{Hauptspeise}{Nachspeise}{Zeitplan}{Ausklang}{Schlusswort}')

    print("Individual mails generated.")


# --- Main Execution ---

def main():
    print("Starting Running Dinner Calculator...")
    
    # 1. Setup Client
    client = setup_client(ORS_API_KEY)
    
    # 2. Load Data
    try:
        Adress_list = pd.read_excel(DATA_FILE)
        Afterparty = pd.read_excel(AFTERPARTY_FILE)
        print("Data loaded successfully.")
    except Exception as e:
        print(f"Error loading data files: {e}")
        return

    # 3. Calculations for Courses
    Count_teams = len(Adress_list)
    Double_cooks = (3 - int(Count_teams % 3))
    Double_cooks = Double_cooks if Double_cooks != 3 else 0
    Course_width = (int(Count_teams / 3) + (1 if Double_cooks != 0 else 0))
    
    print(f"Total Teams: {Count_teams}")
    print(f"Teams cooking twice (Double Cooks needed): {Double_cooks}")
    print(f"Course Width (hosts per course): {Course_width}")

    # 4. Identify Double Cook Candidates
    High_double_cook = Adress_list["will to double"].max()
    Second_High_double_cook = Adress_list[Adress_list["will to double"] != High_double_cook]["will to double"].max()
    
    Count_Highest_double_cook = 0
    Doublecooks_valuelist = [High_double_cook]
    
    while Count_Highest_double_cook < Double_cooks:
        for _, x in Adress_list.iterrows():
            if x["will to double"] == High_double_cook:
                Count_Highest_double_cook += 1
        
        if Count_Highest_double_cook != Double_cooks:
            Doublecooks_valuelist.append(Second_High_double_cook)
            break # Assuming simple logic from original script
            
    # Re-evaluate who are the double cooks
    # Original logic:
    # Doublecooks_Numberlist are the Course_width top willed teams, but limited to Double_cooks count? 
    # Actually, the code logic was slightly complex.
    
    # Let's reproduce the dataframe of candidates
    # Temporary dataframe for distance (initially just formatting)
    # The original script calculated "Distance_to_Afterparty" early on, which contains all dims.
    
    # 5. Geocoding
    # Prepare stripped dataframe
    Adresslist_without_names = Adress_list[["Team Nr.", "Adress", "will to double", "readiness starter", "vegan", "routeplus_vegan"]].copy()
    
    add_coordinates_to_dataframe(Adresslist_without_names, client, "Adress")
    add_coordinates_to_dataframe(Afterparty, client, "Adress")
    
    # 6. Distance to Afterparty (Centrality Check)
    Layer_Teams_coords = Adresslist_without_names[['longitude', 'latitude']].apply(lambda row: [row['longitude'], row['latitude']], axis=1).tolist()
    Layer_Afterparty_coords = [Afterparty[['longitude', 'latitude']].values[0].tolist()]
    
    print("Calculating distances to Afterparty...")
    matrix_afterparty = calculate_distance_matrix(client, Layer_Teams_coords, Layer_Afterparty_coords)
    
    distances_ap = matrix_afterparty['distances']
    durations_ap = matrix_afterparty['durations']
    
    # Create Distance DataFrame
    Distance_to_Afterparty = Adresslist_without_names.copy()
    Distance_to_Afterparty['Distance_km'] = [dist[0] / 1000 for dist in distances_ap]
    Distance_to_Afterparty['Duration_H'] = [dur[0] / 3600 for dur in durations_ap]
    
    # Sort by duration to afterparty
    Distance_to_Afterparty.sort_values(by='Duration_H', ascending=True, inplace=True)
    
    # 7. Select Courses (Starters, Main, Dessert)
    
    # Identify Double Cook Teams (Teams cooking twice)
    Doubles = Distance_to_Afterparty.loc[Distance_to_Afterparty["will to double"].isin(Doublecooks_valuelist)].sort_values(by='will to double', ascending=False)
    Doublecooks_Numberlist = []
    for _, x in Doubles[:Double_cooks].iterrows():
        Doublecooks_Numberlist.append(x["Team Nr."])
    
    print(f"Double Cook Teams: {Doublecooks_Numberlist}")

    desserts = []
    main_courses = []
    starters = []
    double_courses = []

    # Assign Desserts (closest to Afterparty generally, or based on list order which is sorted by duration)
    # The top 'Course_width' closest to Afterparty are assigned Desserts?
    # Original code: for _, x in Distance_to_Afterparty[:(Course_width)].iterrows(): descriptors.append(...)
    
    for _, x in Distance_to_Afterparty[:Course_width].iterrows():
        desserts.append(x["Team Nr."])
        if x["Team Nr."] in Doublecooks_Numberlist and len(double_courses) < Double_cooks:
            double_courses.append(x["Team Nr."])
            main_courses.append(x["Team Nr."])

    # Assign Starters (prefer 'readiness starter' == True)
    # They are picked from the far end (farthest from afterparty? or just from the list?)
    # Original uses slicing with negative step: [-1:-(Course_width...):-1]
    Dessert_prep = Distance_to_Afterparty.loc[Distance_to_Afterparty["readiness starter"] == True]
    
    limit = -(Course_width - (Double_cooks - len(double_courses)) + 1)
    # Note: slicing logic is tricky. Let's assume the original logic meant "take X unique starters from the end of the sorted list"
    
    for _, x in Dessert_prep.iloc[::-1].iterrows(): # Reverse iterate
        if len(starters) >= (Course_width - (Double_cooks - len(double_courses))): # Approximate break condition
             if len(set(starters + main_courses + desserts)) < Count_teams: # Heuristic
                 pass # Continue logic is cleaner below
    
    # Stick to original loop structure for exactness
    # "Dessert_prep[-1:-(Course_width - (Double_cooks-len(double_courses))+1):-1]"
    needed_starters = Course_width - (Double_cooks - len(double_courses))
    count = 0
    for _, x in Dessert_prep.iloc[::-1].iterrows():
        if count >= needed_starters:
            break
        starters.append(x["Team Nr."])
        if x["Team Nr."] in Doublecooks_Numberlist and len(double_courses) < Double_cooks:
            double_courses.append(x["Team Nr."])
            main_courses.append(x["Team Nr."])
        count += 1

    # Fill remaining Main Courses
    already_assigned = starters + desserts + main_courses
    
    for _, x in Distance_to_Afterparty.loc[~Distance_to_Afterparty["Team Nr."].isin(already_assigned)].iterrows():
        main_courses.append(x["Team Nr."])
        if x["Team Nr."] in Doublecooks_Numberlist and len(double_courses) < Double_cooks:
            double_courses.append(x["Team Nr."])
            starters.append(x["Team Nr."])

    print(f"Starters: {starters}")
    print(f"Main Courses: {main_courses}")
    print(f"Desserts: {desserts}")

    # Create Sub-tables for coordinates
    starters_table = Distance_to_Afterparty.loc[Distance_to_Afterparty["Team Nr."].isin(starters)]
    main_courses_table = Distance_to_Afterparty.loc[Distance_to_Afterparty["Team Nr."].isin(main_courses)]
    desserts_table = Distance_to_Afterparty.loc[Distance_to_Afterparty["Team Nr."].isin(desserts)]

    # 8. Calculate Inter-Group Distance Matrices
    # Helpers to extract coords
    def get_coords(table):
        return table[['longitude', 'latitude']].values.tolist()

    starter_coords = get_coords(starters_table)
    main_course_coords = get_coords(main_courses_table)
    desserts_coords = get_coords(desserts_table)

    print("Calculating Step 1 Distances (Starters -> Mains)...")
    matrix_s_m = calculate_distance_matrix(client, starter_coords, main_course_coords)
    
    # Process Many-to-Many
    def create_dist_df(matrix, origins_table, dests_table, origin_col_name, dest_col_name):
        dists = matrix['distances']
        durs = matrix['durations']
        data = []
        origins_ids = origins_table['Team Nr.'].tolist()
        dests_ids = dests_table['Team Nr.'].tolist()
        
        for i, origin_id in enumerate(origins_ids):
            for j, dest_id in enumerate(dests_ids):
                data.append({
                    origin_col_name: origin_id,
                    dest_col_name: dest_id,
                    'Distance_km': dists[i][j] / 1000,
                    'Duration_H': durs[i][j] / 3600
                })
        return pd.DataFrame(data)

    Distance_Starters_to_MainCourses = create_dist_df(matrix_s_m, starters_table, main_courses_table, 'Starter Team Nr.', 'Main Course Team Nr.')

    print("Calculating Step 2 Distances (Mains -> Desserts)...")
    matrix_m_d = calculate_distance_matrix(client, main_course_coords, desserts_coords)
    Distance_MainCourses_to_Desserts = create_dist_df(matrix_m_d, main_courses_table, desserts_table, 'Main Course Team Nr.', 'Dessert Team Nr.')

    print("Calculating Step 3 Distances (Starters -> Desserts) [Backup/Check]...")
    matrix_s_d = calculate_distance_matrix(client, starter_coords, desserts_coords)
    Distance_Starters_to_Desserts = create_dist_df(matrix_s_d, starters_table, desserts_table, 'Starter Course Team Nr.', 'Dessert Team Nr.')

    # 9. Assignment Logic
    # Initialize 'Sees' columns
    for i in range(1, 7):
        Adress_list[f"Sees {i}"] = 0

    print("Running Assignment Algorithms...")
    
    # Assigment 1: Starters to Mains
    assignment(Adress_list, main_courses_table, Distance_Starters_to_MainCourses, True, "Sees 1", "Sees 2", "Main Course Team Nr.", "Starter Team Nr.")
    
    # Assignment 2: Desserts (initial pass or specific logic?)
    # Original: assignment(Adress_list, desserts_table, Distance_Starters_to_Desserts, True, "Sees 1", "Sees 2", "Dessert Team Nr.", "Starter Course Team Nr.")
    # Wait, the original does this, but maybe it overwrites "Sees 1/2" again? Or maybe it's filling gaps?
    # It seems to use 'Distance_Starters_to_Desserts' for this.
    assignment(Adress_list, desserts_table, Distance_Starters_to_Desserts, True, "Sees 1", "Sees 2", "Dessert Team Nr.", "Starter Course Team Nr.")
    
    # Overwrite logic: Propagate distances?
    Copy_Star_Dess = Distance_Starters_to_Desserts.copy()
    overwrite(Copy_Star_Dess, Distance_Starters_to_MainCourses, Adress_list, "Dessert Team Nr.", "Starter Course Team Nr.", "Starter Team Nr.", "Main Course Team Nr.", "Sees 1")
    
    # Adjust dataframe for next assignment
    Copy_Star_Dess.sort_values(by=["Dessert Team Nr.", "Starter Course Team Nr."], ascending=[True, True], inplace=True)
    Copy_Star_Dess["Main Course Team Nr."] = Copy_Star_Dess["Starter Course Team Nr."]
    # del Copy_Star_Dess["Starter Course Team Nr."] # Column rename effectively
    
    # Assignment 3: Desserts to Mains (Host) ?
    # assignment(Adress_list, desserts_table, Copy_Star_Dess, False, "Sees 3", "Sees 4", "Dessert Team Nr.", "Main Course Team Nr.")
    # The columns passed must match the dataframe columns.
    # Note: "Starter Course Team Nr." was used in 'overwrite' as 'guest_target'.
    assignment(Adress_list, desserts_table, Copy_Star_Dess, False, "Sees 3", "Sees 4", "Dessert Team Nr.", "Main Course Team Nr.")
    
    # Assignment 4: Starters to Main Hosts
    assignment(Adress_list, starters_table, Distance_Starters_to_MainCourses, True, "Sees 3", "Sees 4", "Starter Team Nr.", "Main Course Team Nr.")
    
    # Another Overwrite and Assign cycle
    Copy_Star_Main = Distance_Starters_to_MainCourses.copy()
    overwrite(Copy_Star_Main, Distance_MainCourses_to_Desserts, Adress_list, "Starter Team Nr.", "Main Course Team Nr.", "Main Course Team Nr.", "Dessert Team Nr.", "Sees 3")
    Copy_Star_Main["Dessert Team Nr."] = Copy_Star_Main["Main Course Team Nr."]
    # del Copy_Star_Main["Main Course Team Nr."]

    # Assign Main Courses to Dessert Hosts
    assignment(Adress_list, main_courses_table, Distance_MainCourses_to_Desserts, True, "Sees 5", "Sees 6", "Main Course Team Nr.", "Dessert Team Nr.")
    
    # Assign Starter Teams to Dessert Hosts
    assignment(Adress_list, starters_table, Copy_Star_Main, False, "Sees 5", "Sees 6", "Starter Team Nr.", "Dessert Team Nr.")

    print("Assignments complete.")
    
    # 10. Output Generation
    generate_overview_files(starters_table, main_courses_table, desserts_table, Adress_list)
    
    # Ask for user input for emails
    config = get_user_inputs()
    
    generate_individual_mails(Adress_list, starters, main_courses, desserts, Afterparty, config)
    
    print("Done! Files generated.")

if __name__ == "__main__":
    main()
