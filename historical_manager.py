# historical_manager.py
# COPY ALL OF THIS INTO A NEW FILE

import pickle
import json
import pandas as pd
import numpy as np
from datetime import datetime
import os

class HistoricalDataManager:
    def __init__(self, storage_path="dryer_historical_data"):
        """Initialize historical data manager"""
        self.storage_path = storage_path
        if not os.path.exists(storage_path):
            os.makedirs(storage_path)
        
        self.kpi_file = os.path.join(storage_path, "kpi_history.pkl")
        self.optimization_file = os.path.join(storage_path, "optimization_history.pkl")
        self.consolidated_file = os.path.join(storage_path, "consolidated_yearly.pkl")
    
    def save_kpi_results(self, results, timestamp=None):
        """Save KPI analysis results with timestamp"""
        if timestamp is None:
            timestamp = datetime.now()
        
        history = self.load_kpi_history()
        
        entry = {
            'timestamp': timestamp,
            'summary': results['summary'],
            'yearly': results['yearly'],
            'products': results['yearly']['Produkt'].unique().tolist(),
            'total_energy': results['yearly']['Energy_kWh'].sum(),
            'avg_efficiency': results['yearly']['kWh_per_m3'].mean()
        }
        
        history.append(entry)
        
        if len(history) > 100:
            history = history[-100:]
        
        with open(self.kpi_file, 'wb') as f:
            pickle.dump(history, f)
        
        return True
    
    def load_kpi_history(self):
        """Load KPI history"""
        try:
            with open(self.kpi_file, 'rb') as f:
                return pickle.load(f)
        except FileNotFoundError:
            return []
    
    def get_consolidated_historical_data(self):
        """Get consolidated historical data for ALL products"""
        history = self.load_kpi_history()
        
        if not history:
            return None
        
        all_yearly_data = []
        for entry in history:
            yearly_df = entry['yearly'].copy()
            yearly_df['analysis_date'] = entry['timestamp']
            all_yearly_data.append(yearly_df)
        
        if all_yearly_data:
            combined = pd.concat(all_yearly_data, ignore_index=True)
            
            consolidated = combined.groupby(['Produkt', 'Zone']).apply(
                lambda x: pd.Series({
                    'Energy_kWh': x['Energy_kWh'].sum(),
                    'Volume_m3': x['Volume_m3'].sum(),
                    'kWh_per_m3': (x['Energy_kWh'].sum() / x['Volume_m3'].sum()) if x['Volume_m3'].sum() > 0 else 0,
                    'sample_count': len(x),
                    'confidence': min(len(x) / 10, 1.0)
                })
            ).reset_index()
            
            return consolidated
        return None
    
    def merge_with_current_data(self, current_yearly, weight_historical=0.3):
        """Merge historical data with current analysis"""
        historical = self.get_consolidated_historical_data()
        
        if historical is None or historical.empty:
            return current_yearly, "No historical data used"
        
        merged = current_yearly.merge(
            historical[['Produkt', 'Zone', 'kWh_per_m3', 'confidence']],
            on=['Produkt', 'Zone'],
            how='left',
            suffixes=('_current', '_historical')
        )
        
        merged['kWh_per_m3_combined'] = merged.apply(
            lambda row: (
                row['kWh_per_m3_current'] * (1 - weight_historical) +
                row['kWh_per_m3_historical'] * weight_historical * row.get('confidence', 1.0)
                if pd.notna(row.get('kWh_per_m3_historical')) 
                else row['kWh_per_m3_current']
            ),
            axis=1
        )
        
        merged['kWh_per_m3'] = merged['kWh_per_m3_combined']
        
        result = merged[['Produkt', 'Zone', 'Energy_kWh', 'Volume_m3', 'kWh_per_m3']].copy()
        
        products_with_history = merged['kWh_per_m3_historical'].notna().sum()
        status = f"Enhanced with historical data for {products_with_history} product-zone combinations"
        
        return result, status
    
    def save_optimization_result(self, products, optimal_order, metrics):
        """Save optimization results"""
        history = self.load_optimization_history()
        
        entry = {
            'timestamp': datetime.now(),
            'products': products,
            'optimal_order': optimal_order,
            'total_cost': metrics['best_cost'],
            'savings_vs_worst': metrics['savings_vs_worst'],
            'savings_vs_avg': metrics['savings_vs_avg']
        }
        
        history.append(entry)
        
        if len(history) > 50:
            history = history[-50:]
        
        with open(self.optimization_file, 'wb') as f:
            pickle.dump(history, f)
    
    def load_optimization_history(self):
        """Load optimization history"""
        try:
            with open(self.optimization_file, 'rb') as f:
                return pickle.load(f)
        except FileNotFoundError:
            return []
    
    def get_product_profile(self, product):
        """Get average historical profile for a product"""
        history = self.load_kpi_history()
        
        if not history:
            return None
        
        product_data = []
        for entry in history:
            yearly = entry['yearly']
            prod_rows = yearly[yearly['Produkt'] == product]
            if not prod_rows.empty:
                product_data.append({
                    'timestamp': entry['timestamp'],
                    'avg_kwh_per_m3': prod_rows['kWh_per_m3'].mean(),
                    'total_energy': prod_rows['Energy_kWh'].sum(),
                    'total_volume': prod_rows['Volume_m3'].sum()
                })
        
        if product_data:
            df = pd.DataFrame(product_data)
            return {
                'product': product,
                'avg_kwh_per_m3': df['avg_kwh_per_m3'].mean(),
                'std_kwh_per_m3': df['avg_kwh_per_m3'].std(),
                'total_runs': len(df),
                'last_run': df['timestamp'].max()
            }
        return None
