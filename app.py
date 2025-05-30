import streamlit as st
import pandas as pd
import xlwings as xw

file_path = "Qarabag Economic Model_2032_oil.xlsx"
result_sheet = "Result"
price_sheet = "Intro"
price_cell = "B14"
capex_sheet = "Model Inputs"
capex_items_range = "B134:B146"
capex_years_range = "E133:AK133"

def update_oil_price_and_load_data(oil_price):
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    sht_price = wb.sheets[price_sheet]
    sht_result = wb.sheets[result_sheet]
    sht_price.range(price_cell).value = oil_price
    wb.app.calculate()
    indicators = [str(i).strip() for i in sht_result.range("B3:B8").value]
    entities = [str(e).strip() for e in sht_result.range("C2:G2").value]
    values = sht_result.range("C3:G8").value
    wb.save()
    wb.close()
    app.quit()
    data = pd.DataFrame(values, index=indicators, columns=entities)
    return data

def read_current_oil_price():
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    sht_price = wb.sheets[price_sheet]
    price = sht_price.range(price_cell).value
    wb.close()
    app.quit()
    return price

def read_capex_breakdown(mode="total"):
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    sht = wb.sheets[capex_sheet]
    items = [str(i).strip() for i in sht.range(capex_items_range).value]
    years = sht.range(capex_years_range).value
    values = sht.range("E134:AK146").value
    wb.close()
    app.quit()
    df = pd.DataFrame(values, index=items, columns=years)
    if mode == "total":
        return df.sum(axis=1)
    elif mode == "year":
        return df.sum(axis=0)
    else:
        return None

entity_map = {
    "socar karabakh": "SOCAR Karabakh",
    "soa": "SOCAR Karabakh",
    "rsa project": "RSA Project",
    "project": "RSA Project",
    "total project": "RSA Project",
    "contractor party 2": "Contractor Party 2",
    "foreign contractor": "Contractor Party 2",
    "bp": "Contractor Party 2",
    "other contractor": "Contractor Party 2",
    "state share": "State Share",
    "state/sofaz incl. ag": "State/SOFAZ incl. AG"
}

indicator_map = {
    "capex": "CAPEX, MM USD",
    "cash flow": "Cash Flow, MM USD",
    "npv": "NPV10, MM USD",
    "npv10": "NPV10, MM USD",
    "irr": "IRR, %",
    "non-discounted payback": "Non-Discounted Payback Period",
    "discounted payback": "Discounted Payback Period"
}

st.title("Qarabag Economic Model Assistant")

user_input = st.text_input("Ask your question here:")

if user_input:
    lowered = user_input.lower()

    if "hello" in lowered and "ocean" in lowered:
        st.write("Hello SUMI User, how can I help you with?")

    elif "what oil price" in lowered and "used" in lowered:
        price = read_current_oil_price()
        st.write(f"The current oil price used in the model is {price} USD/bbl.")

    elif "capex" in lowered and "breakdown" in lowered:
        if "life of field" in lowered or "total" in lowered:
            capex_totals = read_capex_breakdown(mode="total")
            st.write("CAPEX Breakdown - Life of Field (Total):")
            st.dataframe(capex_totals)
        elif "year by year" in lowered:
            capex_years = read_capex_breakdown(mode="year")
            st.write("CAPEX Breakdown - Year by Year:")
            st.dataframe(capex_years)
        else:
            st.write("Would you like the CAPEX breakdown by 'life of field' (total) or 'year by year'?")

    elif ("oil price" in lowered) and any(cmd in lowered for cmd in ["update", "set", "change"]):
        words = user_input.split()
        new_price = None
        for w in words:
            try:
                val = float(w)
                if val > 0:
                    new_price = val
                    break
            except:
                continue
        if new_price is not None:
            try:
                data = update_oil_price_and_load_data(new_price)
                st.write(f"Oil price updated to {new_price} USD/bbl and data refreshed.")
            except Exception as e:
                st.error(f"Failed to update oil price: {e}")
        else:
            st.write("Sorry, I couldn't find a valid oil price in your request.")

    else:
        found_entity = next((entity_map[e] for e in entity_map if e in lowered), None)
        found_indicator = next((indicator_map[i] for i in indicator_map if i in lowered), None)
        if not found_entity and "capex" not in lowered:
            st.write("Sorry, I couldn't find which party you're asking about.")
        elif not found_indicator and "capex" not in lowered:
            st.write("Sorry, I couldn't understand which metric you're referring to.")
        else:
            try:
                data = update_oil_price_and_load_data(read_current_oil_price())
                value = data.loc[found_indicator, found_entity]
                st.write(f"{found_entity}'s {found_indicator} is {value}")
            except:
                st.write("Data not found for the given query.")
