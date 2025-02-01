import folium
import pandas as pd
import numpy as np
from geopy.distance import geodesic
from folium import Map, Marker, PolyLine, FeatureGroup, LayerControl
from datetime import datetime

# Load the data
file_path = 'SmartRoute Optimizer.xlsx'
shipments_df = pd.read_excel(file_path, sheet_name='Shipments_Data')
vehicle_df = pd.read_excel(file_path, sheet_name='Vehicle_Information')
store_location = pd.read_excel(file_path, sheet_name='Store Location')

# Store location
store_lat, store_long = store_location.iloc[0]['Latitute'], store_location.iloc[0]['Longitude']

# Assumptions
travel_time_per_km = 5  # minutes
delivery_time_per_shipment = 10  # minutes

# Define priority vehicle types and colors
vehicle_priority = ['3W', '4W-EV', '4W']
vehicle_colors = {'3W': 'blue', '4W-EV': 'green', '4W': 'red'}

# Vehicle limits
vehicle_limits = {'3W': 50, '4W-EV': 25, '4W': float('inf')}  # 4W has no limit

# Convert vehicle data to numeric
vehicle_df['Shipments_Capacity'] = pd.to_numeric(vehicle_df['Shipments_Capacity'], errors='coerce')
vehicle_df['Max Trip Radius (in KM)'] = pd.to_numeric(vehicle_df['Max Trip Radius (in KM)'], errors='coerce')

# Function to calculate MST distance using geodesic distance
def calculate_mst_distance(locations):
    total_distance = sum(geodesic(locations[i], locations[i+1]).km for i in range(len(locations)-1))
    return total_distance

# Function to calculate trip time
def calculate_trip_time(distance, num_shipments):
    return (distance * travel_time_per_km) + (num_shipments * delivery_time_per_shipment)

# Function to calculate time slot duration
def calculate_time_slot_duration(time_slot):
    start_time, end_time = map(lambda t: datetime.strptime(t.strip(), '%H:%M:%S'), time_slot.split('-'))
    return (end_time - start_time).total_seconds() / 60

# Function to assign priority vehicle type with limits
def assign_vehicle_type(num_shipments, mst_distance):
    for vehicle_type in vehicle_priority:
        if vehicle_limits[vehicle_type] > 0:  # Check if vehicle is available
            vehicle = vehicle_df[vehicle_df['Vehicle Type'] == vehicle_type].iloc[0]
            max_capacity, max_radius = vehicle['Shipments_Capacity'], vehicle['Max Trip Radius (in KM)']
            max_radius = max_radius if not pd.isna(max_radius) else float('inf')
            if num_shipments <= max_capacity and mst_distance <= max_radius:
                vehicle_limits[vehicle_type] -= 1  # Decrement vehicle count
                return vehicle_type
    return '4W'  # Default to 4W if no priority vehicle can accommodate

# Improved grouping function using nearest-neighbor logic
def group_shipments(shipments, max_capacity, max_radius):
    shipments = sorted(shipments, key=lambda x: geodesic((store_lat, store_long), (x['Latitude'], x['Longitude'])).km)
    grouped_shipments, remaining_shipments = [], shipments.copy()

    while remaining_shipments:
        group, group_distance = [], 0
        for shipment in remaining_shipments[:]:  # Iterate over copy of remaining shipments
            shipment_distance = geodesic((store_lat, store_long), (shipment['Latitude'], shipment['Longitude'])).km
            if len(group) < max_capacity and group_distance + shipment_distance <= max_radius:
                group.append(shipment)
                group_distance += shipment_distance
                remaining_shipments.remove(shipment)
        grouped_shipments.append(group)
    
    return grouped_shipments

# Optimized trip creation
def optimize_trips(shipments_df):
    trips = []
    grouped_shipments = shipments_df.groupby('Delivery Timeslot')

    for time_slot, group in grouped_shipments:
        shipments_in_slot = group.to_dict('records')
        time_slot_duration = calculate_time_slot_duration(time_slot)
        max_capacity, max_radius = vehicle_df['Shipments_Capacity'].max(), vehicle_df['Max Trip Radius (in KM)'].max()
        grouped_shipments_in_slot = group_shipments(shipments_in_slot, max_capacity, max_radius)

        for group in grouped_shipments_in_slot:
            locations = [(store_lat, store_long)] + [(s['Latitude'], s['Longitude']) for s in group] + [(store_lat, store_long)]
            mst_distance = calculate_mst_distance(locations)
            vehicle_type = assign_vehicle_type(len(group), mst_distance)

            trips.append({
                'TRIP ID': len(trips) + 1,
                'Shipments': group,
                'TIME SLOT': time_slot,
                'MST_DIST': mst_distance,
                'TRIP_TIME': calculate_trip_time(mst_distance, len(group)),
                'Vehical_Type': vehicle_type,
                'CAPACITY_UTI': len(group) / max_capacity,
                'TIME_UTI': calculate_trip_time(mst_distance, len(group)) / time_slot_duration,
                'COV_UTI': mst_distance / max_radius
            })
    
    return trips

# Optimize trips
trips = optimize_trips(shipments_df)

# Convert trips to DataFrame
trips_df = pd.DataFrame([{
    'TRIP ID': trip['TRIP ID'],
    'Shipment ID': ', '.join(str(s['Shipment ID']) for s in trip['Shipments']),
    'Latitude': ', '.join(str(s['Latitude']) for s in trip['Shipments']),
    'Longitude': ', '.join(str(s['Longitude']) for s in trip['Shipments']),
    'TIME SLOT': trip['TIME SLOT'],
    'Shipments': len(trip['Shipments']),
    'MST_DIST': trip['MST_DIST'],
    'TRIP_TIME': trip['TRIP_TIME'],
    'Vehical_Type': trip['Vehical_Type'],
    'CAPACITY_UTI': trip['CAPACITY_UTI'],
    'TIME_UTI': trip['TIME_UTI'],
    'COV_UTI': trip['COV_UTI']
} for trip in trips])

# Save trips to Excel
trips_df.to_excel('Sample Output Trip.xlsx', index=False)

# Visualize trips on a map
m = Map(location=[store_lat, store_long], zoom_start=12)
Marker([store_lat, store_long], tooltip='Store').add_to(m)

# Feature groups for each vehicle type
vehicle_groups = {vehicle: FeatureGroup(name=vehicle) for vehicle in vehicle_colors.keys()}

# Add shipments and routes
for trip in trips:
    vehicle_type, color = trip['Vehical_Type'], vehicle_colors.get(trip['Vehical_Type'], 'gray')
    locations = [(store_lat, store_long)] + [(s['Latitude'], s['Longitude']) for s in trip['Shipments']] + [(store_lat, store_long)]

    PolyLine(locations, color=color, weight=2.5, opacity=1).add_to(vehicle_groups[vehicle_type])
    for shipment in trip['Shipments']:
        Marker([shipment['Latitude'], shipment['Longitude']], tooltip=f"Shipment {shipment['Shipment ID']}").add_to(vehicle_groups[vehicle_type])

for group in vehicle_groups.values():
    group.add_to(m)

LayerControl().add_to(m)
m.save('route_map.html')

print("Optimized trips saved to 'Sample Output Trip.xlsx'")
print("Route map saved to 'route_map.html'")