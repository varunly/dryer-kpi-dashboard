"""
Lindner Dryer - Build Optimization Database
Run this script ONCE to analyze all historical data and create permanent optimization profiles
"""

import pandas as pd
import numpy as np
import json
from datetime import datetime
from dryer_kpi_monthly_final import parse_energy, parse_wagon, explode_intervals, allocate_energy, CONFIG

class OptimizationDatabaseBuilder:
    def __init__(self, energy_file, wagon_file):
        """Initialize with your data files"""
        self.energy_file = energy_file
        self.wagon_file = wagon_file
        self.product_profiles = {}
        
    def analyze_all_data(self):
        """Analyze complete historical data"""
        print("🔄 Loading and parsing data files...")
        
        # Load energy data
        e_raw = pd.read_excel(self.energy_file, sheet_name=CONFIG["energy_sheet"])
        e = parse_energy(e_raw)
        print(f"✅ Loaded {len(e)} energy records")
        
        # Load wagon data
        w_raw = pd.read_excel(
            self.wagon_file,
            sheet_name=CONFIG["wagon_sheet"],
            header=CONFIG["wagon_header_row"]
        )
        w = parse_wagon(w_raw)
        print(f"✅ Loaded {len(w)} wagon records")
        
        # Process intervals
        print("🔄 Processing zone intervals...")
        ivals = explode_intervals(w)
        print(f"✅ Created {len(ivals)} interval records")
        
        # Allocate energy
        print("🔄 Allocating energy to products...")
        alloc = allocate_energy(e, ivals)
        print(f"✅ Allocated energy to {len(alloc)} records")
        
        # Create comprehensive summary
        print("🔄 Creating product profiles...")
        
        # Overall product summary
        product_summary = alloc.groupby("Produkt").agg({
            "Energy_share_kWh": "sum",
            "m3": "sum",
            "Overlap_h": "sum"
        }).reset_index()
        
        product_summary["kWh_per_m3"] = (
            product_summary["Energy_share_kWh"] / 
            product_summary["m3"].replace(0, np.nan)
        )
        
        # Zone-specific data
        zone_summary = alloc.groupby(["Produkt", "Zone"]).agg({
            "Energy_share_kWh": ["sum", "mean", "std"],
            "m3": "sum",
            "Overlap_h": "sum"
        }).reset_index()
        
        zone_summary.columns = [
            'Produkt', 'Zone', 
            'Total_Energy_kWh', 'Avg_Energy_kWh', 'Std_Energy_kWh',
            'Total_Volume_m3', 'Total_Hours'
        ]
        
        zone_summary["kWh_per_m3"] = (
            zone_summary["Total_Energy_kWh"] / 
            zone_summary["Total_Volume_m3"].replace(0, np.nan)
        )
        
        # Build product profiles
        products = product_summary["Produkt"].unique()
        
        for product in products:
            print(f"  📊 Profiling {product}...")
            
            # Overall stats
            prod_data = product_summary[product_summary["Produkt"] == product].iloc[0]
            
            # Zone-specific stats
            zone_data = zone_summary[zone_summary["Produkt"] == product]
            
            # Count production runs
            wagon_count = len(w[w["Produkt"] == product])
            
            # Get product characteristics
            thickness = self._extract_thickness(product)
            product_type = self._extract_type(product)
            
            # Build comprehensive profile
            profile = {
                # Basic info
                "product": product,
                "thickness_mm": thickness,
                "type": product_type,
                
                # Production stats
                "total_wagons_produced": int(wagon_count),
                "total_volume_m3": float(prod_data["m3"]),
                "total_energy_kwh": float(prod_data["Energy_share_kWh"]),
                "total_production_hours": float(prod_data["Overlap_h"]),
                
                # Efficiency metrics
                "avg_kwh_per_m3": float(prod_data["kWh_per_m3"]),
                "kwh_per_wagon": float(prod_data["Energy_share_kWh"] / wagon_count) if wagon_count > 0 else 0,
                
                # Zone-specific data
                "zone_profiles": {},
                
                # Metadata
                "data_points": int(wagon_count),
                "confidence": self._calculate_confidence(wagon_count),
                "last_updated": datetime.now().isoformat()
            }
            
            # Add zone-specific profiles
            for _, zone_row in zone_data.iterrows():
                zone = zone_row['Zone']
                profile["zone_profiles"][zone] = {
                    "total_energy_kwh": float(zone_row['Total_Energy_kWh']),
                    "avg_energy_kwh": float(zone_row['Avg_Energy_kWh']),
                    "std_energy_kwh": float(zone_row['Std_Energy_kWh']) if pd.notna(zone_row['Std_Energy_kWh']) else 0,
                    "kwh_per_m3": float(zone_row['kWh_per_m3']) if pd.notna(zone_row['kWh_per_m3']) else 0,
                    "total_hours": float(zone_row['Total_Hours'])
                }
            
            self.product_profiles[product] = profile
        
        print(f"\n✅ Created profiles for {len(self.product_profiles)} products")
        
        return self.product_profiles
    
    def calculate_transition_matrix(self):
        """Calculate optimal transition costs between all product pairs"""
        print("\n🔄 Calculating transition cost matrix...")
        
        products = list(self.product_profiles.keys())
        transition_matrix = {}
        
        for prod1 in products:
            transition_matrix[prod1] = {}
            for prod2 in products:
                if prod1 == prod2:
                    transition_matrix[prod1][prod2] = 0
                else:
                    cost = self._calculate_transition_cost(prod1, prod2)
                    transition_matrix[prod1][prod2] = cost
        
        print(f"✅ Calculated {len(products)}x{len(products)} transition matrix")
        
        return transition_matrix
    
    def _calculate_transition_cost(self, prod1, prod2):
        """Calculate transition cost based on real data"""
        profile1 = self.product_profiles[prod1]
        profile2 = self.product_profiles[prod2]
        
        cost = 0
        
        # 1. Thickness difference cost (physical setup)
        thickness_diff = abs(profile1['thickness_mm'] - profile2['thickness_mm'])
        cost += thickness_diff * 3.0  # 3 kWh per mm
        
        # 2. Type change penalty (cleaning needed)
        if profile1['type'] != profile2['type']:
            cost += 50  # Fixed penalty for material type change
        
        # 3. Energy consumption difference (temperature adjustment)
        energy_diff = abs(profile1['avg_kwh_per_m3'] - profile2['avg_kwh_per_m3'])
        cost += energy_diff * 0.8
        
        # 4. Zone temperature differences
        for zone in ['Z2', 'Z3', 'Z4', 'Z5']:
            if zone in profile1['zone_profiles'] and zone in profile2['zone_profiles']:
                zone_energy_diff = abs(
                    profile1['zone_profiles'][zone]['kwh_per_m3'] - 
                    profile2['zone_profiles'][zone]['kwh_per_m3']
                )
                cost += zone_energy_diff * 0.2
        
        return round(cost, 2)
    
    def generate_optimization_rules(self):
        """Generate optimization rules based on data"""
        print("\n🔄 Generating optimization rules...")
        
        rules = {
            "product_grouping": {},
            "preferred_sequences": [],
            "avoid_transitions": [],
            "energy_intensive_products": [],
            "quick_changeover_groups": []
        }
        
        products = list(self.product_profiles.keys())
        
        # Group by type
        for product in products:
            ptype = self.product_profiles[product]['type']
            if ptype not in rules["product_grouping"]:
                rules["product_grouping"][ptype] = []
            rules["product_grouping"][ptype].append(product)
        
        # Sort within groups by thickness
        for ptype in rules["product_grouping"]:
            rules["product_grouping"][ptype].sort(
                key=lambda p: self.product_profiles[p]['thickness_mm']
            )
        
        # Identify energy-intensive products (top 30%)
        energy_sorted = sorted(
            products,
            key=lambda p: self.product_profiles[p]['avg_kwh_per_m3'],
            reverse=True
        )
        cutoff = int(len(energy_sorted) * 0.3)
        rules["energy_intensive_products"] = energy_sorted[:cutoff]
        
        # Find quick changeover pairs (thickness diff < 4mm, same type)
        for i, prod1 in enumerate(products):
            for prod2 in products[i+1:]:
                p1 = self.product_profiles[prod1]
                p2 = self.product_profiles[prod2]
                
                if (p1['type'] == p2['type'] and 
                    abs(p1['thickness_mm'] - p2['thickness_mm']) <= 4):
                    rules["quick_changeover_groups"].append([prod1, prod2])
        
        # Generate preferred sequences (ascending thickness within type)
        for ptype, prods in rules["product_grouping"].items():
            if len(prods) > 1:
                rules["preferred_sequences"].append({
                    "type": ptype,
                    "sequence": prods,
                    "reason": "Ascending thickness minimizes adjustments"
                })
        
        print(f"✅ Generated {len(rules['preferred_sequences'])} optimization rules")
        
        return rules
    
    def save_database(self, output_file="optimization_database.json"):
        """Save complete optimization database to JSON"""
        print(f"\n💾 Saving optimization database to {output_file}...")
        
        # Calculate transition matrix
        transition_matrix = self.calculate_transition_matrix()
        
        # Generate rules
        rules = self.generate_optimization_rules()
        
        # Compile complete database
        database = {
            "metadata": {
                "created": datetime.now().isoformat(),
                "source_energy_file": self.energy_file,
                "source_wagon_file": self.wagon_file,
                "total_products": len(self.product_profiles),
                "version": "1.0"
            },
            "product_profiles": self.product_profiles,
            "transition_matrix": transition_matrix,
            "optimization_rules": rules
        }
        
        # Save to JSON
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(database, f, indent=2, ensure_ascii=False)
        
        print(f"✅ Database saved successfully!")
        print(f"\n📊 Database Summary:")
        print(f"   Products profiled: {len(self.product_profiles)}")
        print(f"   Transition costs calculated: {len(self.product_profiles)**2}")
        print(f"   Optimization rules: {len(rules['preferred_sequences'])}")
        print(f"   Quick changeover pairs: {len(rules['quick_changeover_groups'])}")
        
        return database
    
    def save_excel_report(self, output_file="optimization_database.xlsx"):
        """Save human-readable Excel report"""
        print(f"\n📊 Creating Excel report...")
        
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            # Product profiles
            profiles_data = []
            for product, profile in self.product_profiles.items():
                profiles_data.append({
                    'Product': product,
                    'Type': profile['type'],
                    'Thickness (mm)': profile['thickness_mm'],
                    'Wagons Produced': profile['total_wagons_produced'],
                    'Total Volume (m³)': profile['total_volume_m3'],
                    'Total Energy (kWh)': profile['total_energy_kwh'],
                    'Avg kWh/m³': profile['avg_kwh_per_m3'],
                    'kWh/Wagon': profile['kwh_per_wagon'],
                    'Confidence': profile['confidence']
                })
            
            df_profiles = pd.DataFrame(profiles_data)
            df_profiles.to_excel(writer, sheet_name='Product_Profiles', index=False)
            
            # Transition matrix
            transition_matrix = self.calculate_transition_matrix()
            df_transitions = pd.DataFrame(transition_matrix)
            df_transitions.to_excel(writer, sheet_name='Transition_Matrix')
            
            # Zone details
            zone_data = []
            for product, profile in self.product_profiles.items():
                for zone, zone_prof in profile['zone_profiles'].items():
                    zone_data.append({
                        'Product': product,
                        'Zone': zone,
                        'Total Energy (kWh)': zone_prof['total_energy_kwh'],
                        'kWh/m³': zone_prof['kwh_per_m3'],
                        'Total Hours': zone_prof['total_hours']
                    })
            
            df_zones = pd.DataFrame(zone_data)
            df_zones.to_excel(writer, sheet_name='Zone_Details', index=False)
        
        print(f"✅ Excel report saved to {output_file}")
    
    def print_summary(self):
        """Print summary to console"""
        print("\n" + "="*60)
        print("OPTIMIZATION DATABASE SUMMARY")
        print("="*60)
        
        for product in sorted(self.product_profiles.keys()):
            profile = self.product_profiles[product]
            print(f"\n{product} ({profile['type']}-type, {profile['thickness_mm']}mm):")
            print(f"  • Wagons produced: {profile['total_wagons_produced']}")
            print(f"  • Avg energy: {profile['avg_kwh_per_m3']:.2f} kWh/m³")
            print(f"  • Energy/wagon: {profile['kwh_per_wagon']:.2f} kWh")
            print(f"  • Confidence: {profile['confidence']}")
    
    def _extract_thickness(self, product):
        """Extract thickness from product name"""
        import re
        match = re.search(r'\d+', product)
        return int(match.group()) if match else 36
    
    def _extract_type(self, product):
        """Extract type from product name"""
        return product[0] if product else 'L'
    
    def _calculate_confidence(self, data_points):
        """Calculate confidence score based on number of data points"""
        if data_points >= 50:
            return "High"
        elif data_points >= 20:
            return "Medium"
        elif data_points >= 5:
            return "Low"
        else:
            return "Very Low"


# Main execution
if __name__ == "__main__":
    print("="*60)
    print("LINDNER DRYER - OPTIMIZATION DATABASE BUILDER")
    print("="*60)
    print("\nThis script analyzes your complete historical data")
    print("and creates a permanent optimization database.\n")
    
    # File paths
    ENERGY_FILE = r"E:\Lindner\Python\Energieverbrauch Trockner 1, Stundenweise - Januar - September 2025.xlsx"
    WAGON_FILE = r"E:\Lindner\Python\Hordenwagenverfolgung_Stand 2025_10_12.xlsm"
    
    # Create builder
    builder = OptimizationDatabaseBuilder(ENERGY_FILE, WAGON_FILE)
    
    # Analyze all data
    try:
        profiles = builder.analyze_all_data()
        
        # Print summary
        builder.print_summary()
        
        # Save JSON database
        database = builder.save_database("optimization_database.json")
        
        # Save Excel report
        builder.save_excel_report("optimization_database.xlsx")
        
        print("\n" + "="*60)
        print("✅ OPTIMIZATION DATABASE CREATED SUCCESSFULLY!")
        print("="*60)
        print("\nGenerated files:")
        print("  1. optimization_database.json (for the app)")
        print("  2. optimization_database.xlsx (human-readable)")
        print("\nYou can now use this database for instant optimization!")
        
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
