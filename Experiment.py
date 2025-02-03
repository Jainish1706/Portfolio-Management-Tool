import tkinter as tk
from tkinter import ttk
import pandas as pd
from main import process_trade_data, buy_and_sell_q, load_worksheet

def process_inputs():
    # Here you would collect inputs and pass them through your backend logic
    isin = entry_isin.get()
    symbol = entry_symbol.get()
    bs = entry_bs.get()
    quantity = int(entry_quantity.get())
    date = entry_date.get()
    unit_price = float(entry_price.get())
    
    # Create DataFrame or pass inputs to existing functions in main.py
    df = pd.DataFrame({
        'ISIN': [isin],
        'symbol': [symbol],
        'BS': [bs],
        'Quant': [quantity],
        'Date': [date],
        'Unit_Price': [unit_price]
    })

    # Process the data using existing functions
    q_buy, q_sell = buy_and_sell_q(df)
    
    # Display processed holdings (output) in the output section
    output_text.delete(1.0, tk.END)
    output_text.insert(tk.END, f"Buy Queue:\n{q_buy}\nSell Queue:\n{q_sell}")

# Initialize the GUI window
root = tk.Tk()
root.title("Buy-Sell Holdings")

# Input fields for user input
tk.Label(root, text="ISIN").grid(row=0, column=0)
entry_isin = tk.Entry(root)
entry_isin.grid(row=0, column=1)

tk.Label(root, text="Symbol").grid(row=1, column=0)
entry_symbol = tk.Entry(root)
entry_symbol.grid(row=1, column=1)

tk.Label(root, text="Buy/Sell (B/S)").grid(row=2, column=0)
entry_bs = tk.Entry(root)
entry_bs.grid(row=2, column=1)

tk.Label(root, text="Quantity").grid(row=3, column=0)
entry_quantity = tk.Entry(root)
entry_quantity.grid(row=3, column=1)

tk.Label(root, text="Date (dd-mm-yyyy)").grid(row=4, column=0)
entry_date = tk.Entry(root)
entry_date.grid(row=4, column=1)

tk.Label(root, text="Unit Price").grid(row=5, column=0)
entry_price = tk.Entry(root)
entry_price.grid(row=5, column=1)

# Button to process input and display the result
process_button = tk.Button(root, text="Process", command=process_inputs)
process_button.grid(row=6, column=0, columnspan=2)

# Output field to display the result
output_text = tk.Text(root, height=10, width=50)
output_text.grid(row=7, column=0, columnspan=2)

# Run the GUI
root.mainloop()

# import tkinter as tk
# from tkinter import messagebox

# # Function to handle user input
# def take_input():
#     user_input = input_field.get()
#     if user_input.lower() == "stop":
#         root.quit()  # Close the GUI
#     else:
#         inputs.append(user_input)
#         input_field.delete(0, tk.END)  # Clear the input field for new entry
#         print(f"Inputs so far: {inputs}")  # For debugging, prints input list to console

# # Main window setup
# root = tk.Tk()
# root.title("Input Collector")

# inputs = []  # To store user inputs

# # GUI elements
# label = tk.Label(root, text="Enter input (type 'stop' to end):")
# label.pack(pady=10)

# input_field = tk.Entry(root, width=30)
# input_field.pack(pady=10)

# submit_button = tk.Button(root, text="Submit", command=take_input)
# submit_button.pack(pady=10)

# root.mainloop()
