import folium
from geopy.distance import distance, great_circle
import math
import webbrowser
import time
import os 
import pandas as pd
import requests
import tkinter as tk
from tkinter import messagebox

pillars = {}
def square_search(initial_coordinates, last_coordinates, radius):

    initial_latitude, initial_longitude = initial_coordinates[0], initial_coordinates[1]
    last_latitude, last_longitude = last_coordinates[0], last_coordinates[1]

    distance_between_circles = math.sqrt(3) * radius
    #start variables
    global pillars
    current_latitude, current_longitude = initial_latitude, initial_longitude
    count = 0
    is_offset = True

    #draw the pillar circles until past the last_latitude
    while current_latitude > last_latitude:
        #updates the number of circle drawn and ands the coordinates to the dictionary
        pillars[count] = (current_latitude, current_longitude)
        count += 1

        #assigns the bearing depending if the circle is offset
        if is_offset:
            angle = 150
        else: angle = 210

        #calculates the coordinates of the next circle
        next_circle = distance(kilometers=distance_between_circles / 1000).destination((current_latitude, current_longitude), bearing=angle)

        #updates current coordiantes to next circle
        current_latitude = next_circle.latitude
        current_longitude = next_circle.longitude

        #changes is_offset for next circle
        is_offset = not is_offset

    for i in range(len(pillars)):
        next_circle = distance(kilometers=distance_between_circles / 1000).destination((pillars[i][0], pillars[i][1]), bearing=90)
        current_latitude = next_circle.latitude
        current_longitude = next_circle.longitude
        while current_longitude < last_longitude:
            pillars[count] = (current_latitude, current_longitude)
            count += 1

            next_circle = distance(kilometers=distance_between_circles / 1000).destination((current_latitude, current_longitude), bearing=90)
            current_latitude = next_circle.latitude
            current_longitude = next_circle.longitude


    return pillars

circles_in_circle = {}
circles_in_circle_count = 0
def circle_search(big_coordinates, big_radius):
    circle_pillars = {}
    global circles_in_circle_count
    big_latitude, big_longitude = big_coordinates[0], big_coordinates[1]
    top_corner = distance(kilometers= math.sqrt(2*math.pow(big_radius, 2)) / 1000).destination((big_latitude, big_longitude), bearing=315)
    bottom_corner = distance(kilometers= math.sqrt(2*math.pow(big_radius, 2)) / 1000).destination((big_latitude, big_longitude), bearing=135)


    #initial coordinates
    initial_latitude, initial_longitude = top_corner.latitude, top_corner.longitude
    #last coordinates
    last_latitude, last_longitude = bottom_corner.latitude, bottom_corner.longitude
    #radius of cicles (meters)
    radius = big_radius/4
    #distance between circles centers
    distance_between_circles = math.sqrt(3) * radius 

    #start variables
    current_latitude, current_longitude = initial_latitude, initial_longitude
    circle_pillars = {}
    count = 0
    is_offset = True

    #draw the pillar circles until past the last_latitude
    while current_latitude > last_latitude:
        #updates the number of circle drawn and adds the coordinates to the dictionary
        circle_pillars[count] = (current_latitude, current_longitude)
        count += 1

        #assigns the bearing depending if the circle is offset
        if is_offset:
            angle = 150
        else: angle = 210

        #calculates the coordinates of the next circle
        next_circle = distance(kilometers=distance_between_circles / 1000).destination((current_latitude, current_longitude), bearing=angle)

        #updates current coordiantes to next circle
        current_latitude = next_circle.latitude
        current_longitude = next_circle.longitude
        #changes is_offset for next circle
        is_offset = not is_offset

    for i in range(len(circle_pillars)):
        next_circle = distance(kilometers=distance_between_circles / 1000).destination((circle_pillars[i][0], circle_pillars[i][1]), bearing=90)
        current_latitude = next_circle.latitude
        current_longitude = next_circle.longitude
        while current_longitude < last_longitude:

            dist = great_circle(big_coordinates, (current_latitude, current_longitude)).meters

            if dist <= big_radius:
                circles_in_circle[circles_in_circle_count] = (current_latitude, current_longitude)
                circles_in_circle_count += 1

            next_circle = distance(kilometers=distance_between_circles / 1000).destination((current_latitude, current_longitude), bearing=90)
            current_latitude = next_circle.latitude
            current_longitude = next_circle.longitude
    
    return circles_in_circle

def fetch_place_details(api_key, place_id):
    details_url = f"https://maps.googleapis.com/maps/api/place/details/json?place_id={place_id}&key={api_key}"
    try:
        response = requests.get(details_url)
        response.raise_for_status()  # Raise an error for bad HTTP status codes
    except requests.exceptions.RequestException as e:
        print(f"Error fetching details for place ID {place_id}: {e}")
        return None, None, None

    details = response.json().get('result', {})
    if not details:
        print(f"No details found for place ID {place_id}")
        return None, None, None

    phone_number = details.get('formatted_phone_number', 'N/A')
    formatted_address = details.get('formatted_address', 'N/A')
    website = details.get('website', 'N/A')
    return phone_number, formatted_address, website


overflown = []
no_result = []
def fetch_and_save_places(api_key, circle_coordinates, radius, searching_types, file_name):
    latitude = circle_coordinates[0]
    longitude = circle_coordinates[1]
    places_data = []  # Initialize this list at the start of the function
    next_page_token = None
    global overflown
    global no_result

    default_searching_types = ["store", "bakery", "bar", "beuty_salon", "book_store", "cafe", "car_wash", "clothing_store", "convenience_store", "dentist", "electronics_store", "food", "gym", "hair_care",  "home_goods_store", "jewelry_store", "liquor_store", "meal_delivery", "meal_takeaway", "movie_rental", "pet_store", "restaurant", "spa", "veterinary_care", "night_club", "shoe_store", "car_repair", "furniture_store"]

    searching_types = [type.strip() for type in searching_types.split(',')]
    searching_types += default_searching_types
    searching_types = set(searching_types)

    file_path = f"{file_name}.xlsx"

    place_count = 0
    while True:
        search_url = f"https://maps.googleapis.com/maps/api/place/nearbysearch/json?location={latitude},{longitude}&radius={radius}&key={api_key}"
        if next_page_token:
            search_url += f"&pagetoken={next_page_token}"

        response = requests.get(search_url)
        if response.status_code != 200:
            messagebox.showerror("Error", "Failed to fetch places. Please check the API key and parameters.")
            return
        
        results = response.json().get('results', [])
        if not results:
            print("No places found in the specific area.")
            no_result.append(circle_coordinates)
            break  # Break the loop if there are no results

        next_page_token = response.json().get('next_page_token')

        for place in results:
            # Place processing logic here
            place_count += 1
            name = place.get('name')
            place_id = place.get('place_id')
            types = place.get('types', [])
            all_types = ', '.join(types) if types else 'N/A'

            if all(searching_type not in types for searching_type in searching_types):
                continue

            # Fetch additional details
            phone_number, formatted_address, website = fetch_place_details(api_key, place_id)

            # Prepare the data to be added
            places_data.append({
                'Business Name': name,
                'Types': all_types,
                'Service Type': '',  # Leave this blank
                'Phone #': phone_number,
                'Website': website,
                'Address': formatted_address,
                'Owner/Manager Name': '',
                'Owner/Manager Email/#': '',
                'Time Called': '',
                'Day Called MM/DD': '',
                'Your Initials': '',
                'Business Answer': '',
                'Notes': '',
                'Place ID': place_id  # Include Place ID for reference
            })

        if not next_page_token:
            break

    if place_count >= 60:
        print(f"Circle Overflown ({place_count})")
        overflown.append(circle_coordinates)
               
    time.sleep(2)  # Wait for the API to prepare the next page

    # Check if the file already exists and write new data
    if os.path.exists(file_path):
        # Append data to the existing Excel file
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df = pd.DataFrame(places_data)
            startrow = writer.sheets["All Data"].max_row  # Start appending below the last row
            df.to_excel(writer, sheet_name="All Data", index=False, header=False, startrow=startrow)
    else:
        # Create a new Excel file with the data
        df = pd.DataFrame(places_data)
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name="All Data", index=False)

            # Write header titles only if creating a new file
            header_titles = [
                'Business Name', 'Types', 'Service Type', 'Phone #', 'Website',
                'Address', 'Owner/Manager Name', 'Owner/Manager Email/#',
                'Time Called', 'Day Called MM/DD', 'Your Initials', 
                'Business Answer', 'Notes', 'Place ID'
            ]
            for col_num, title in enumerate(header_titles, 1):
                writer.sheets["All Data"].cell(row=1, column=col_num, value=title)

    print(f"Data saved to {file_name}.xlsx")

def on_submit():
    api_key = api_key_entry.get()
    
    # Parse coordinates from entries
    try:
        initial_coordinates = tuple(map(float, initial_coordinates_entry.get().split(',')))
        last_coordinates = tuple(map(float, last_coordinates_entry.get().split(',')))
    except ValueError:
        messagebox.showerror("Error", "Invalid coordinate format. Use 'latitude,longitude'.")
        return
    
    try:
        radius = float(radius_entry.get())
    except ValueError:
        messagebox.showerror("Error", "Radius and max per sheet must be numbers.")
        return

    searching_types = searching_types_entry.get()
    file_name = file_name_entry.get()

    if not (api_key and initial_coordinates and last_coordinates and radius and file_name):
        messagebox.showerror("Error", "All fields must be filled out, and max per sheet must be greater than 0.")
        return
    
    # Perform the square search
    square_search(initial_coordinates, last_coordinates, radius)
    
    for coordinates in pillars.values():
        fetch_and_save_places(api_key, coordinates, radius, searching_types, file_name)

    if overflown:
        response = messagebox.askyesno("Confirm", "Do you want to search in the overflown circles?")
        if response:
            for coord in overflown:
                circle_search(coord, radius)
                print(circles_in_circle)
            for circle_coords in circles_in_circle.values():
                fetch_and_save_places(api_key, circle_coords, radius / 10, searching_types, file_name)
        else: 
            messagebox.showinfo("Canceled", "Search in overflown coordinates was canceled.")

    # Folium map generation using middle coordinates
    distance_vertex_square = great_circle(initial_coordinates, last_coordinates).meters
    middle_square = distance(kilometers=distance_vertex_square / 1000).destination(initial_coordinates, bearing=135)

    m = folium.Map(location=[middle_square.latitude, middle_square.longitude], zoom_start=13)

    for coordinates in pillars.values():

        if coordinates in overflown:
            color = "red"
        elif coordinates in no_result:
            color = "grey"
        else: 
            color = "green"

        folium.Circle(
            location=coordinates,
            radius=radius,
            color=color,
            fill=True,
            fill_color=color,
            fill_opacity=0.1,
        ).add_to(m)

    for circle_coords in circles_in_circle.values():
        folium.Circle(
            location=circle_coords,
            radius=radius / 4,
            color="white",
            fill=False,
            fill_opacity=0.1,
        ).add_to(m)
        
    m.save("visuals_search.html")
    log_maker(file_name)
    webbrowser.open("visuals_search.html")
    messagebox.showinfo("Success", f"Data saved to {file_name}.xlsx")


def log_maker(file_name):

    file_path = f"{file_name}.txt"

    with open(file_path, 'w') as file:
        file.write('\n' + repr(pillars))
        file.write('\n' + repr(overflown))
        file.write('\n' + repr(no_result))       

    print(f"Search log saved to {file_name}.txt")

# Create the main window
root = tk.Tk()
root.title("9Yaps Sales Software")

# Create input fields
tk.Label(root, text="API Key:").grid(row=0, column=0)
api_key_entry = tk.Entry(root)
api_key_entry.grid(row=0, column=1)

tk.Label(root, text="Initial Coords:").grid(row=1, column=0)
initial_coordinates_entry = tk.Entry(root)
initial_coordinates_entry.grid(row=1, column=1)

tk.Label(root, text="Last Coords:").grid(row=2, column=0)
last_coordinates_entry = tk.Entry(root)
last_coordinates_entry.grid(row=2, column=1)

tk.Label(root, text="Radius (meters):").grid(row=3, column=0)
radius_entry = tk.Entry(root)
radius_entry.grid(row=3, column=1)

tk.Label(root, text="Searching Types (comma separated):").grid(row=4, column=0)
searching_types_entry = tk.Entry(root)
searching_types_entry.grid(row=4, column=1)

tk.Label(root, text="File Name (without extension):").grid(row=5, column=0)
file_name_entry = tk.Entry(root)
file_name_entry.grid(row=5, column=1)

# Create a submit button
submit_button = tk.Button(root, text="Search", command=on_submit)
submit_button.grid(row=7, columnspan=2)

# Start the Tkinter event loop
root.mainloop()