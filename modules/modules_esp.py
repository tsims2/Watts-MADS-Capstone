from .stylesheet import stylesheet
from .settlement_api_dispatch_data import DispatchDataProcessor
from .settlement_cs_performance_calculations import CSPerformanceCalculator
from .settlement_cs_report_generation  import CSReportGenerator

color_palette = [('#FF5400'),  # orange
                    ('#FFC907'),  # yellow
                    ('#072B60'),  # royal blue
                    ('#00C9AZ'),  # green
                    ('#59C8E3'), # sky blue
                    ('#020244'), # navy
                    ('#EBF2F5'), # light blue
                    ('#363C49'), # gray
                    ('#0078A6'), #Teal
                    ('#12B34C'), #Green
                    ]  

def get_float_value(self, value_str):
    try:
        return float(value_str)
    except ValueError:
        return 0.0

def sanitize_filename(filename):
    # Replace invalid characters with underscores
    return "".join(c if c.isalnum() or c in (' ', '.', '_') else '_' for c in filename)

def convert_to_string_without_decimal(value):
    try:
        # First, try to convert the value to a float, then to an int to remove any decimals
        int_value = int(float(value))
        return str(int_value)
    except (ValueError, TypeError):
        # If conversion fails, return the value as a string directly
        return str(value)
