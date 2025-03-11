from datetime import datetime

def dateFormatter(value):
    if isinstance(value, datetime):
        return value.strftime("%d %b %y")  # Convert datetime object to "10 Oct 24"

    if isinstance(value, str):  # If it's a string like "10 October 2024"
        try:
            # Try parsing different possible formats
            parsed_date = datetime.strptime(value, "%d %B %Y")  # Full month name
            return parsed_date.strftime("%d %b %y")
        except ValueError:
            pass  # Continue if parsing fails
    
    return str(value)  # Return as-is if it doesn't match

    

def percentageFormatter(cell):
    value = cell.value
    if isinstance(value, (int, float)):
        return f"{cell * 100:.0f}%"   
    


#### TESTING PURPOSE
def demoFormatter(value):
    return value + " Formatted"




    
