import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime
import requests
import json
import math

class OptionsAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("NSE Options Trading Analyzer")
        self.root.geometry("1400x800")
        self.root.configure(bg="#f0f0f0")
        
        self.analysis_data = []
        
        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('Header.TLabel', font=('Arial', 16, 'bold'), foreground="#667eea")
        self.style.configure('Config.TLabelframe', font=('Arial', 10, 'bold'))
        self.style.configure('Action.TButton', font=('Arial', 10, 'bold'), padding=10)
        
        self.create_widgets()
        
    def create_widgets(self):
        # Main container
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="📊 NSE Options Trading Analyzer", 
                               style='Header.TLabel')
        title_label.pack(pady=(0, 20))
        
        # Configuration Frame
        config_frame = ttk.LabelFrame(main_frame, text="Configuration", 
                                     style='Config.TLabelframe', padding="15")
        config_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Config Row 1
        config_row1 = ttk.Frame(config_frame)
        config_row1.pack(fill=tk.X, pady=5)
        
        ttk.Label(config_row1, text="Margin Percentage (%):").pack(side=tk.LEFT, padx=(0, 10))
        self.margin_var = tk.StringVar(value="15")
        margin_entry = ttk.Entry(config_row1, textvariable=self.margin_var, width=15)
        margin_entry.pack(side=tk.LEFT, padx=(0, 30))
        
        ttk.Label(config_row1, text="Lot Size Multiplier:").pack(side=tk.LEFT, padx=(0, 10))
        self.lot_multiplier_var = tk.StringVar(value="1")
        lot_entry = ttk.Entry(config_row1, textvariable=self.lot_multiplier_var, width=15)
        lot_entry.pack(side=tk.LEFT)
        
        # Buttons
        button_frame = ttk.Frame(config_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        fetch_btn = ttk.Button(button_frame, text="🔍 Fetch & Analyze Options", 
                              command=self.fetch_and_analyze, style='Action.TButton')
        fetch_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        export_btn = ttk.Button(button_frame, text="📥 Export to Excel", 
                               command=self.export_to_excel, style='Action.TButton')
        export_btn.pack(side=tk.LEFT)
        
        # Status Label
        self.status_label = ttk.Label(main_frame, text="", font=('Arial', 9))
        self.status_label.pack(pady=(0, 10))
        
        # Results Frame with Scrollbar
        results_frame = ttk.Frame(main_frame)
        results_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create Treeview
        columns = ('Symbol', 'Spot', '52W High', '52W Low', 'Percentile', 'Lot Size',
                  'CE Strike', 'CE Premium', 'CE IRR', 'PE Strike', 'PE Premium', 'PE IRR')
        
        self.tree = ttk.Treeview(results_frame, columns=columns, show='headings', height=20)
        
        # Configure columns
        column_widths = {
            'Symbol': 100, 'Spot': 100, '52W High': 100, '52W Low': 100,
            'Percentile': 90, 'Lot Size': 90, 'CE Strike': 90, 'CE Premium': 100,
            'CE IRR': 90, 'PE Strike': 90, 'PE Premium': 100, 'PE IRR': 90
        }
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths.get(col, 100), anchor=tk.CENTER)
        
        # Scrollbars
        vsb = ttk.Scrollbar(results_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(results_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        results_frame.grid_rowconfigure(0, weight=1)
        results_frame.grid_columnconfigure(0, weight=1)
        
        # Configure row colors
        self.tree.tag_configure('oddrow', background='#f9f9f9')
        self.tree.tag_configure('evenrow', background='#ffffff')
    
    def show_status(self, message, is_error=False):
        color = "red" if is_error else "green"
        self.status_label.config(text=message, foreground=color)
        self.root.update()
    
    def calculate_percentile(self, current, high, low):
        if high == low:
            return 50.0
        return ((current - low) / (high - low)) * 100
    
    def find_nearest_strike(self, price, strike_interval=50):
        return round(price / strike_interval) * strike_interval
    
    def calculate_strike_with_margin(self, spot_price, margin_percent, is_call=True):
        # First get the nearest strike to current spot price
        nearest_strike = self.find_nearest_strike(spot_price)
        
        # Then apply margin to this nearest strike
        if is_call:
            # For CE: 15% less means multiply by (1 - 0.15) = 0.85
            target_price = nearest_strike * (1 - margin_percent / 100)
        else:
            # For PE: 15% more means multiply by (1 + 0.15) = 1.15
            target_price = nearest_strike * (1 + margin_percent / 100)
        
        # Get nearest strike for the target price
        return self.find_nearest_strike(target_price)
    
    def generate_option_premium(self, spot_price, strike_price, option_type):
        """Simplified premium calculation for demonstration"""
        diff = abs(spot_price - strike_price)
        
        if option_type == 'CE':
            if spot_price > strike_price:
                base_premium = (spot_price - strike_price) + (spot_price * 0.02)
            else:
                base_premium = spot_price * 0.01
        else:  # PE
            if spot_price < strike_price:
                base_premium = (strike_price - spot_price) + (spot_price * 0.02)
            else:
                base_premium = spot_price * 0.01
        
        return round(base_premium, 2)
    
    def calculate_irr(self, premium, strike_price, days_to_expiry=30):
        """Calculate Internal Rate of Return"""
        margin_required = strike_price * 0.15  # 15% margin
        if margin_required == 0:
            return 0
        irr = (premium / margin_required) * (365 / days_to_expiry) * 100
        return round(irr, 2)
    
    def get_sample_data(self):
        """Sample stock data - Replace with actual NSE API calls"""
        return [
            {'symbol': 'NIFTY', 'spot_price': 19500, 'high_52w': 20000, 'low_52w': 18000, 'lot_size': 50},
            {'symbol': 'BANKNIFTY', 'spot_price': 44500, 'high_52w': 46000, 'low_52w': 42000, 'lot_size': 25},
            {'symbol': 'RELIANCE', 'spot_price': 2450, 'high_52w': 2650, 'low_52w': 2250, 'lot_size': 250},
            {'symbol': 'TCS', 'spot_price': 3600, 'high_52w': 3850, 'low_52w': 3200, 'lot_size': 125},
            {'symbol': 'INFY', 'spot_price': 1480, 'high_52w': 1600, 'low_52w': 1350, 'lot_size': 300},
            {'symbol': 'HDFCBANK', 'spot_price': 1650, 'high_52w': 1750, 'low_52w': 1450, 'lot_size': 550},
            {'symbol': 'ICICIBANK', 'spot_price': 950, 'high_52w': 1050, 'low_52w': 850, 'lot_size': 1375},
            {'symbol': 'SBIN', 'spot_price': 590, 'high_52w': 650, 'low_52w': 520, 'lot_size': 1500},
            {'symbol': 'BHARTIARTL', 'spot_price': 880, 'high_52w': 950, 'low_52w': 750, 'lot_size': 1220},
            {'symbol': 'HINDUNILVR', 'spot_price': 2580, 'high_52w': 2800, 'low_52w': 2350, 'lot_size': 300}
        ]
    
    def fetch_and_analyze(self):
        try:
            self.show_status("🔄 Analyzing options data...")
            
            # Clear existing data
            for item in self.tree.get_children():
                self.tree.delete(item)
            self.analysis_data = []
            
            # Get configuration
            margin_percent = float(self.margin_var.get())
            lot_multiplier = int(self.lot_multiplier_var.get())
            
            # Get stock data (replace with actual API call)
            stocks = self.get_sample_data()
            
            for stock in stocks:
                spot_price = stock['spot_price']
                high_52w = stock['high_52w']
                low_52w = stock['low_52w']
                
                # Calculate percentile
                percentile = self.calculate_percentile(spot_price, high_52w, low_52w)
                
                # Calculate strikes
                ce_strike = self.calculate_strike_with_margin(spot_price, margin_percent, is_call=True)
                pe_strike = self.calculate_strike_with_margin(spot_price, margin_percent, is_call=False)
                
                # Generate premiums
                ce_premium = self.generate_option_premium(spot_price, ce_strike, 'CE')
                pe_premium = self.generate_option_premium(spot_price, pe_strike, 'PE')
                
                # Calculate IRR
                ce_irr = self.calculate_irr(ce_premium, ce_strike)
                pe_irr = self.calculate_irr(pe_premium, pe_strike)
                
                # Adjusted lot size
                adjusted_lot_size = stock['lot_size'] * lot_multiplier
                
                # Store data
                row_data = {
                    'Symbol': stock['symbol'],
                    'Spot Price': spot_price,
                    '52W High': high_52w,
                    '52W Low': low_52w,
                    'Percentile': f"{percentile:.2f}",
                    'Lot Size': adjusted_lot_size,
                    'CE Strike': ce_strike,
                    'CE Premium': ce_premium,
                    'CE IRR': f"{ce_irr:.2f}",
                    'PE Strike': pe_strike,
                    'PE Premium': pe_premium,
                    'PE IRR': f"{pe_irr:.2f}",
                    'Margin Used': margin_percent
                }
                
                self.analysis_data.append(row_data)
                
                # Insert into treeview
                values = (
                    stock['symbol'],
                    f"₹{spot_price:.2f}",
                    f"₹{high_52w:.2f}",
                    f"₹{low_52w:.2f}",
                    f"{percentile:.2f}%",
                    adjusted_lot_size,
                    f"₹{ce_strike}",
                    f"₹{ce_premium}",
                    f"{ce_irr:.2f}%",
                    f"₹{pe_strike}",
                    f"₹{pe_premium}",
                    f"{pe_irr:.2f}%"
                )
                
                tag = 'evenrow' if len(self.analysis_data) % 2 == 0 else 'oddrow'
                self.tree.insert('', tk.END, values=values, tags=(tag,))
            
            self.show_status("✅ Analysis completed successfully!")
            
        except ValueError as e:
            self.show_status(f"⚠️ Invalid input: {str(e)}", is_error=True)
            messagebox.showerror("Input Error", f"Please check your input values: {str(e)}")
        except Exception as e:
            self.show_status(f"❌ Error: {str(e)}", is_error=True)
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
    
    def export_to_excel(self):
        try:
            if not self.analysis_data:
                messagebox.showwarning("No Data", "Please analyze data before exporting!")
                return
            
            # Create DataFrame
            df = pd.DataFrame(self.analysis_data)
            
            # Generate filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"Options_Analysis_{timestamp}.xlsx"
            
            # Export to Excel
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Options Analysis', index=False)
                
                # Auto-adjust column widths
                worksheet = writer.sheets['Options Analysis']
                for idx, col in enumerate(df.columns):
                    max_length = max(df[col].astype(str).apply(len).max(), len(col)) + 2
                    worksheet.column_dimensions[chr(65 + idx)].width = max_length
            
            self.show_status(f"✅ Excel file exported: {filename}")
            messagebox.showinfo("Success", f"Data exported successfully to:\n{filename}")
            
        except Exception as e:
            self.show_status(f"❌ Export failed: {str(e)}", is_error=True)
            messagebox.showerror("Export Error", f"Failed to export: {str(e)}")

def main():
    root = tk.Tk()
    app = OptionsAnalyzer(root)
    root.mainloop()

if __name__ == "__main__":
    main()