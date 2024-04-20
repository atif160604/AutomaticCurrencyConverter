import tkinter
from tkinter import *
from tkinter import messagebox
import model
from Vat import Vat


def convert_to_bool(text: str):
    if text.upper() == "TRUE":
        return True
    elif text.upper() == "FALSE":
        return False
    else:
        messagebox.showinfo(title="Error", message="Field is left blank")


# --------------------------- BUTTON FUNCTIONALITY -----------------------

def currency_add():
    currency_data = currency_entry.get().upper()
    if currency_data in model.list_of_all_currencies and currency_data not in model.list_of_currency:
        model.list_of_currency.append(currency_data)
        model.add_record_fta(currency_data)
        messagebox.showinfo(title="Currency Data", message=f"The Currency {currency_data} has been Added")
        currency_entry.delete(0, END)
    elif currency_data in model.list_of_all_currencies and currency_data in model.list_of_currency:
        messagebox.showinfo(title="Currency Data", message=f"The Currency {currency_data} is already in list")
    else:
        messagebox.showinfo(title="Currency Data", message=f"The Currency {currency_data} is not in list")
    print(model.list_of_currency)
    model.column_to_start_rates_vat += 1
    print("success")


def currency_delete():
    currency_data = currency_entry.get().upper()
    if currency_data in model.list_of_currency:
        messagebox.showinfo(title="Currency Data", message=f"The Currency {currency_data} has been deleted")
        model.list_of_currency.remove(currency_data)
        model.delete_record_fta(currency_data)
        model.column_to_start_rates_vat -= 1
        print(model.column_to_start_rates_vat)
        currency_entry.delete(0, END)
    else:
        messagebox.showinfo(title="Currency Data", message=f"The Currency {currency_data} does not exist")


def currency_find():
    currency_data = currency_entry.get().upper()
    if currency_data in model.list_of_currency:
        messagebox.showinfo(title="Currency Data", message=f"The Currency {currency_data} is found")
    else:
        messagebox.showinfo(title="Currency Data", message=f"The Currency {currency_data} is not found")


def currency_list():
    templist = []
    for currency_rate in model.list_of_currency:
        templist.append(currency_rate)
    messagebox.showinfo(title="Currency Data", message=f"All the Avaailable currencies: {templist}")


def currency_vat_add():
    currency_from_data = currency_conversion_from_entry.get().upper()
    currency_to_data = currency_conversion_to_entry.get().upper()
    is_vat_data = convert_to_bool(is_vat_entry.get())
    is_corp_data = convert_to_bool(is_corp_entry.get())
    is_red = convert_to_bool(is_corp_red.get())
    is_highlighted = convert_to_bool(is_highlighted_entry.get())

    if isinstance(is_vat_data, bool) and isinstance(is_corp_data, bool) and is_vat_data != is_corp_data:

        if is_vat_data:
            if currency_from_data in model.list_of_all_currencies and currency_to_data in model.list_of_all_currencies:
                model.new_conversion(currency_from_data, currency_to_data, is_vat_data, is_corp_data, is_red,
                                     is_highlighted)

                model.add_record_corp(currency_from_data, currency_to_data, str(is_vat_data), str(is_corp_data), str(is_red),
                                     str(is_highlighted))
                messagebox.showinfo(title="Currency data", message="Conversion has been added")

                currency_conversion_from_entry.delete(0, END)
                currency_conversion_to_entry.delete(0, END)
                is_vat_entry.delete(0, END)
                is_corp_entry.delete(0, END)
                is_corp_red.delete(0, END)
                is_highlighted_entry.delete(0, END)

            else:
                messagebox.showinfo(title="Not Found", message="Currency does not exist")


        elif is_corp_data:
            if currency_from_data in model.list_of_all_currencies_Vat and currency_to_data in model.list_of_all_currencies_Vat:
                model.new_conversion(currency_from_data, currency_to_data, is_vat_data, is_corp_data, is_red,
                                     is_highlighted)
                messagebox.showinfo(title="Currency data", message="Conversion has been added")

                currency_conversion_from_entry.delete(0, END)
                currency_conversion_to_entry.delete(0, END)
                is_vat_entry.delete(0, END)
                is_corp_entry.delete(0, END)
                is_corp_red.delete(0, END)
                is_highlighted_entry.delete(0, END)

            else:
                messagebox.showinfo(title="Not Found", message="Currency does not exist")
    else:
        messagebox.showinfo(title="Currency data", message="VAT and Corp both cannot be True/False")
    for each in model.list_of_vat_currencies:
        print(each)


def currency_vat_delete():
    currency_from_data = currency_conversion_from_entry.get().upper()
    currency_to_data = currency_conversion_to_entry.get().upper()
    is_vat_data = convert_to_bool(is_vat_entry.get())
    is_corp_data = convert_to_bool(is_corp_entry.get())
    is_red = convert_to_bool(is_corp_red.get())
    is_highlighted = convert_to_bool(is_highlighted_entry.get())

    temp_currency = Vat(currency_from_data, currency_to_data, is_vat_data, is_corp_data, is_red, is_highlighted)
    if model.find_conversion(temp_currency):
        model.delete_record_corp(currency_from_data, currency_to_data, str(is_vat_data), str(is_corp_data), str(is_red),
                                     str(is_highlighted))

        model.delete_conversion(currency_from_data, currency_to_data, is_vat_data, is_corp_data)
        messagebox.showinfo(title="Currency Data", message="Conversion has been deleted")

        currency_conversion_from_entry.delete(0, END)
        currency_conversion_to_entry.delete(0, END)
        is_vat_entry.delete(0, END)
        is_corp_entry.delete(0, END)
        is_corp_red.delete(0, END)
        is_highlighted_entry.delete(0, END)

    else:
        messagebox.showinfo(title="Currency Data", message="Conversion cannot be found")


def currency_vat_find():
    currency_from_data = currency_conversion_from_entry.get().upper()
    print(type(currency_from_data))
    currency_to_data = currency_conversion_to_entry.get().upper()
    is_vat_data = convert_to_bool(is_vat_entry.get())
    is_corp_data = convert_to_bool(is_corp_entry.get())
    is_red = convert_to_bool(is_corp_red.get())
    is_highlighted = convert_to_bool(is_highlighted_entry.get())
    print(is_highlighted)

    temp_currency = Vat(currency_from_data, currency_to_data, is_vat_data, is_corp_data, is_red, is_highlighted)
    if model.find_conversion(temp_currency):
        messagebox.showinfo(title="Currency Data", message="Conversion exists in list")
    else:
        messagebox.showinfo(title="Currency Data", message="Conversion does not exist")


def currency_vat_list():
    templist = []
    for currency_rate in model.list_of_vat_currencies:
        templist.append((currency_rate.init_currency, currency_rate.currency_to_convert))
    messagebox.showinfo(title="Currency Data", message=f"All the Available currencies: {templist}")


def refresh():
    model.refresh_excel()
    messagebox.showinfo(title="Success", message="Excel Sheet has been updated")


def clear():
    currency_conversion_from_entry.delete(0, END)
    currency_conversion_to_entry.delete(0, END)
    is_vat_entry.delete(0, END)
    is_corp_entry.delete(0, END)
    is_corp_red.delete(0, END)
    is_highlighted_entry.delete(0, END)

    currency_entry.delete(0, END)

    messagebox.showinfo(title="Clear", message="All Clear")
# Window setup
window = Tk()
window.title("Currency Conversion")
window.config(padx=20, pady=20)

# Canvas with logo
canvas = Canvas(width=200, height=200)
lock_img = PhotoImage(file=model.persistent_image_path)  # Ensure the 'logo.png' file exists in your project directory
canvas.create_image(100, 100, image=lock_img)
canvas.grid(row=0, column=0, columnspan=3)

# Frame for currency conversion controls
currency_frame = Frame(window, padx=10, pady=10)
currency_frame.grid(row=1, column=0, sticky="ew")

Label(currency_frame, text="Currency:").grid(row=0, column=0, sticky="w")
currency_entry = Entry(currency_frame, width=25)
currency_entry.grid(row=0, column=1, sticky="ew")
currency_entry.focus()

# Buttons for currency actions
Button(currency_frame, text="Add", width=10, command=currency_add).grid(row=0, column=2, padx=5)
Button(currency_frame, text="Delete", width=10, command=currency_delete).grid(row=0, column=3, padx=5)
Button(currency_frame, text="Find", width=10, command=currency_find).grid(row=0, column=4, padx=5)
Button(currency_frame, text="List", width=10, command=currency_list).grid(row=0, column=5, padx=5)

# Frame for VAT & Corp
vat_corp_frame = Frame(window, padx=10, pady=10)
vat_corp_frame.grid(row=2, column=0, sticky="ew", columnspan=3)

Label(vat_corp_frame, text="VAT & Corp:").grid(row=0, column=0, columnspan=6)
Label(vat_corp_frame, text="Currency conversion from:").grid(row=1, column=0, sticky="w")
currency_conversion_from_entry = Entry(vat_corp_frame, width=15)
currency_conversion_from_entry.grid(row=1, column=1, sticky="ew")

Label(vat_corp_frame, text="to").grid(row=1, column=2, sticky="ew")
currency_conversion_to_entry = Entry(vat_corp_frame, width=15)
currency_conversion_to_entry.grid(row=1, column=3, sticky="ew")

Label(vat_corp_frame, text="is VAT?:").grid(row=2, column=0, sticky="w")
is_vat_entry = Entry(vat_corp_frame, width=15)
is_vat_entry.grid(row=2, column=1, sticky="ew")

Label(vat_corp_frame, text="is Corp?:").grid(row=2, column=2, sticky="ew")
is_corp_entry = Entry(vat_corp_frame, width=15)
is_corp_entry.grid(row=2, column=3, sticky="ew")

Label(vat_corp_frame, text="is Red?:").grid(row=2, column=4, sticky="w")
is_corp_red = Entry(vat_corp_frame, width=15)
is_corp_red.grid(row=2, column=5, sticky="ew")

# Adding 'is Highlighted?' label and entry
Label(vat_corp_frame, text="is Highlighted?:").grid(row=2, column=6, sticky="w")
is_highlighted_entry = Entry(vat_corp_frame, width=15)
is_highlighted_entry.grid(row=2, column=7, sticky="ew")

# Frame for additional VAT & Corp controls
vat_corp_control_frame = Frame(window, padx=10, pady=10)
vat_corp_control_frame.grid(row=3, column=0, sticky="ew", columnspan=3)

Button(vat_corp_control_frame, text="Add", width=10, command=currency_vat_add).grid(row=0, column=0, padx=5)
Button(vat_corp_control_frame, text="Delete", width=10, command=currency_vat_delete).grid(row=0, column=1, padx=5)
Button(vat_corp_control_frame, text="Find", width=10, command=currency_vat_find).grid(row=0, column=2, padx=5)
Button(vat_corp_control_frame, text="List", width=10, command=currency_vat_list).grid(row=0, column=3, padx=5)

# Frame for refresh and clear controls
refresh_clear_frame = Frame(window, padx=10, pady=10)
refresh_clear_frame.grid(row=4, column=0, sticky="ew", columnspan=3)

Button(refresh_clear_frame, text="Refresh", width=20, command=refresh).grid(row=0, column=0, padx=5)
Button(refresh_clear_frame, text="Clear", width=20, command=clear).grid(row=0, column=1, padx=5)

# Label for 'Created by Atif Agboatwala'
creator_label = Label(window, text="Created by Atif Agboatwala")
creator_label.grid(row=5, column=2, sticky="e")

window.mainloop()
