def sumofBBG(fieldMap):
    counter_value = float(fieldMap["Counter"])
    bbg_value = float(fieldMap["BBG 3.0"])

    result = counter_value + bbg_value

    # Round the value to 3 places max
    result = round(counter_value + bbg_value, 3)
    return result




# TESTING PURPOSE ONLY
# To test the working of your own custom functions
# Run testCustomFunction in jupyter
# Modify data object as per your usecase
def sampleDemoFunction(fieldMap):
    # fieldMap will give you every data as string.
    counter_value = float(fieldMap["Age"])
    bbg_value = float(fieldMap["Phone"])

    result = counter_value + bbg_value

    # Round the value to 3 places max
    result = round(counter_value + bbg_value, 3)
    return result


def testCustomFunction():
    data = {"Age" : "10",
            "Phone" : "12345678",
            "Address": "XYZ"}
    
    result = sumofBBG(data)
    print(result)