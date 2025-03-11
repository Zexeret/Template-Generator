##### FCN Utility Functions START #####

# Placeholder using this function
# [[Count i]]
def numBBGValues(sheet_data):
    bbg_values = [sheet_data[0].get("BBG Code 1", ''), sheet_data[0].get("BBG Code 2", ''), sheet_data[0].get("BBG Code 3", ''), sheet_data[0].get("BBG Code 4", '')]
    return str(sum(1 for bbg in bbg_values if bbg != ''))

# Placeholder using this function
# [[ShareBasketTab]]
# [[ObservationDatesTab]]
def getTableData(sheet_data, start_row=1, end_row=None, start_col=1, end_col=None):
    if not sheet_data:
        return ValueError("Empty sheet data")

    start_row = max(start_row, 1) - 1
    end_row = end_row - 1 if end_row is not None else len(sheet_data) - 1
    start_col = max(start_col, 1) - 1
    end_col = end_col - 1 if end_col is not None else len(sheet_data[0]) - 1
    
    headers = list(sheet_data[0].keys())
    headers = headers[start_col:end_col + 1]
    table_data = []
    for row in sheet_data[start_row:end_row + 1]:
        row_value = {}
        for i, header in enumerate(headers):
            row_value[header] = list(row.values())[start_col + i]   
        table_data.append(row_value)
    return table_data


##### FCN Utility Functions END #####


# TESTING PURPOSE ONLY
# To test the working of your own custom functions
# Run testCustomFunction in jupyter
# Modify data object as per your usecase
def sumofBBG(sheet_data):
    if not sheet_data:
        return ValueError("Empty sheet data")

    fieldMap = sheet_data[0]
    counter_value = float(fieldMap["Counter"])
    bbg_value = float(fieldMap["BBG 3.0"])

    result = counter_value + bbg_value

    # Round the value to 3 places max
    result = round(counter_value + bbg_value, 3)
    return result


def sampleDemoFunction(fieldMap):
    # fieldMap will give you every data as string.
    counter_value = float(fieldMap[0]["Age"])
    bbg_value = float(fieldMap[0]["Phone"])

    result = counter_value + bbg_value

    # Round the value to 3 places max
    result = round(counter_value + bbg_value, 3)
    return result


def testCustomFunction():
    data = [{"Age" : "10",
            "Phone" : "12345678",
            "Address": "XYZ"}]
    
    result = sumofBBG(data)
    print(result)