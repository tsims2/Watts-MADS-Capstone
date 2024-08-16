from .modules_external import *

class CSPerformanceCalculator:
    def __init__(self, asset_id, uan, customer, season, event, event_type, customer_asset, curtailment_strategy, dispatch_list, customer_dispatch_list):
        self.asset_id = asset_id
        self.uan = uan
        self.customer = customer
        self.season = season
        self.event = event
        self.event_type = event_type
        self.customer_asset = customer_asset
        self.curtailment_strategy = curtailment_strategy
        self.dispatches = dispatch_list
        self.customer_dispatches = customer_dispatch_list

        # Read and convert asset_event_data to EST
        file_pattern = f"Settlements/CS/{self.season}/Raw Meter Data/*{self.asset_id}*"
        matching_files = glob.glob(file_pattern)
        if not matching_files:
            raise FileNotFoundError(f"No files found matching pattern: {file_pattern}")
        file_path = matching_files[0]

        # Read the CSV file
        with open(file_path, 'r') as f:
            lines = f.readlines()

        # Extract metadata
        self.metadata = {
            'Asset ID': lines[0].strip().split(',')[1].split('\n')[0],
            'Utility Account Number': self.uan,
            'Customer': lines[2].strip().split(',')[1],
            'Date Range Collected': lines[3].strip().split(',')[1],
            'Date Retrieved': lines[4].strip().split(',')[1],
            'Meter Tags': ','.join(lines[6].strip().split(',')[1:])
        }

        # Extract meter tags
        meter_tags = lines[6].strip().split(',')[1:]

        # Read the actual data
        self.asset_event_data = pd.read_csv(file_path, skiprows=7, names=['Time Intervals'] + meter_tags)

        # Convert 'Time Intervals' to datetime and set as index
        self.asset_event_data['Time Intervals'] = pd.to_datetime(self.asset_event_data['Time Intervals']).dt.tz_localize(None)
        self.asset_event_data.set_index('Time Intervals', inplace=True)

        # If there are multiple meter columns, sum them up
        if len(self.asset_event_data.columns) > 1:
            self.asset_event_data = self.asset_event_data.sum(axis=1)
        else:
            self.asset_event_data = self.asset_event_data.iloc[:, 0]
        
 
    def create_cs_excel(self):
        # Define the event load and the one hour preceding it
        self.event_load = self.get_event_load()

        # Define the baseline dates and data
        self.baseline_dates, self.baseline_data, self.baseline = self.get_baseline()

        # Calculate Baseline Adjustment
        self.baseline_adjustment = self.get_baseline_adjustment()

        # Calculate Adjusted Baseline
        self.adjusted_baseline = self.baseline + self.baseline_adjustment

        # Calculate Performance
        self.performance_df = self.calculate_performance()

        # Create Excel workbook
        event_date = pd.to_datetime(self.event).date()
        excel_path = f"Settlements/CS/{self.season}/CPOWER Results/{self.event_type} Event {event_date}/{self.asset_id} {self.event_type} {event_date} {self.customer} Performance Calculations.xlsx"
        
        # Create directory structure if it doesn't exist
        os.makedirs(os.path.dirname(excel_path), exist_ok=True)

        # Prepare data for "Calculations" sheet
        self.calculations = pd.DataFrame(self.baseline_data)
        self.calculations.columns = ['Baseline Data']
        self.calculations['Date'] = self.calculations.index.date
        self.calculations['Time'] = self.calculations.index.time
        self.calculations = self.calculations.pivot(index='Time', columns='Date', values='Baseline Data')
        
        # Add additional columns
        self.calculations["Baseline"] = self.baseline
        self.calculations["Baseline Adjustment"] = self.baseline_adjustment
        self.calculations["Adjusted Baseline"] = self.adjusted_baseline
        
        event_column_name = f"Event {event_date}"
        event_data = self.asset_event_data[self.asset_event_data.index.date == event_date]
        event_data.index = event_data.index.time
        self.calculations[event_column_name] = event_data.reindex(self.calculations.index)
        
        self.calculations["Performance"] = self.calculations["Adjusted Baseline"] - self.calculations[event_column_name]

        # Create and save Excel file
        wb = Workbook()
        
        # Create "Summary" sheet
        summary_sheet = wb.active
        summary_sheet.title = "Summary"
        
        # Add metadata
        for key, value in self.metadata.items():
            summary_sheet.append([key, value])
        
        summary_sheet.append([])  # Add an empty row for spacing
            
        # Add performance dataframe
        summary_sheet.append(["Performance Data"])
        performance_start_row = summary_sheet.max_row + 1
        for r in dataframe_to_rows(self.performance_df, index=True, header=True):
            summary_sheet.append(r)
        performance_end_row = summary_sheet.max_row
        
        # Add total performance
        total_performance = self.performance_df['Performance'].mean()
        summary_sheet.append([])
        summary_sheet.append(["Total Performance", total_performance])

        # Create line chart
        chart = LineChart()
        chart.title = "Performance Summary"
        chart.x_axis.title = "Time"
        chart.y_axis.title = "Value"

        # Add data series to chart for Adjusted Baseline and Event Load
        col_indices = {
            'Adjusted Baseline': 4,  # Assuming 'Adjusted Baseline' is the third column
            event_column_name: 5     # Assuming 'Event Load' is the fourth column
        }

        adjusted_baseline_data = Reference(summary_sheet, min_col=col_indices['Adjusted Baseline'], 
                                            min_row=performance_start_row, max_row=performance_end_row)
        chart.add_data(adjusted_baseline_data, titles_from_data=True)

        event_load_data = Reference(summary_sheet, min_col=col_indices[event_column_name], 
                                    min_row=performance_start_row, max_row=performance_end_row)
        chart.add_data(event_load_data, titles_from_data=True)

        # Add chart to summary sheet
        summary_sheet.add_chart(chart, "A" + str(summary_sheet.max_row + 2))
        
        # Create "Calculations" sheet
        calc_sheet = wb.create_sheet("Calculations")
        for r in dataframe_to_rows(self.calculations, index=True, header=True):
            calc_sheet.append(r)
        
        # Save the workbook
        try:
            wb.save(excel_path)
        except PermissionError:
            return None

        return total_performance

    def get_event_load(self):
        # Create a temporary column combining Event Date and Start Time
        self.dispatches['Event Start'] = pd.to_datetime(self.dispatches['Event Date'].astype(str) + ' ' + self.dispatches['Start Time'].astype(str))

        # Filter dispatches for the event date and start time
        self.event_dispatch = self.dispatches[self.dispatches['Event Date'] == self.event]
                
        # Combine date and time
        self.event_start = self.event_dispatch['Event Start'].iloc[0]
        self.event_end = pd.to_datetime(f"{self.event_start.date()} {self.event_dispatch['End Time'].iloc[0]}")
        
        # Include 180 Minutes
        event_load = self.asset_event_data[(self.asset_event_data.index >= self.event_start - pd.Timedelta(minutes=180)) & 
                                            (self.asset_event_data.index <= self.event_end)]
        
        return event_load

    def get_baseline(self):
        event_date = pd.to_datetime(self.event)
        
        # Get list of dates to drop
        drop_dates = self.get_baseline_drop_dates(self.event, self.dispatches)

        # Remove any dates after the event and in the drop dates
        baseline_data = self.asset_event_data[self.asset_event_data.index < event_date]
        baseline_data = baseline_data[~baseline_data.index.floor('D').isin([d.date() for d in drop_dates])]

        # Get the latest 10 unique days
        unique_days = baseline_data.index.floor('D').unique()[-10:]
        baseline_data = baseline_data[baseline_data.index.floor('D').isin(unique_days)]

        # Calculate the average baseline
        baseline = baseline_data.groupby(baseline_data.index.time).mean()

        return unique_days, baseline_data, baseline

    def get_baseline_adjustment(self):
        # Calculate adjustment window (30 - 15 minutes before event start)
        adjustment_start = self.event_start - pd.Timedelta(minutes=120)
        adjustment_end = self.event_start - pd.Timedelta(minutes=60)
            
        # Get average load for adjustment period on event day
        event_day_adj = self.event_load[(self.event_load.index.time >= adjustment_start.time()) & 
                                        (self.event_load.index.time < adjustment_end.time())].mean()
        
        # Get average load for adjustment period in baseline
        baseline_adj = self.baseline_data[(self.baseline_data.index.time >= adjustment_start.time()) & 
                                            (self.baseline_data.index.time < adjustment_end.time())].mean()
        
        # Calculate adjustment factor
        adjustment_factor = max((event_day_adj - baseline_adj),0)

        # Make a copy of the baseline and replace all its values with the adjustment factor
        baseline_adjustment = self.baseline.copy()
        baseline_adjustment.loc[:] = adjustment_factor

        return baseline_adjustment

    def calculate_performance(self):        
        # Create a copy of the event_load and remove the date
        actual_load = self.event_load.copy()
        actual_load.index = actual_load.index.time

        self.full_performance_df = pd.DataFrame({
            'Baseline': self.baseline,
            'Adjustment': self.baseline_adjustment,
            'Adjusted Baseline': self.adjusted_baseline,
            'Event Load': actual_load
        })

        # Ensure self.event_start and self.event_end are datetime objects
        self.event_start = pd.to_datetime(self.event_start)
        self.event_end = pd.to_datetime(self.event_end)
        
        self.full_performance_df['Performance'] = self.full_performance_df['Adjusted Baseline'] - self.full_performance_df['Event Load']

        performance_df = self.full_performance_df[
            (self.full_performance_df.index > self.event_start.time()) & 
            (self.full_performance_df.index <= self.event_end.time())
        ]
        
        # Calculate average performance
        self.average_performance = performance_df['Performance'].mean()

        return performance_df

    def get_baseline_drop_dates(self, event, dispatches):
        def convert_to_datetime(ts):
            if isinstance(ts, pd.Timestamp):
                return ts.to_pydatetime()
            elif isinstance(ts, date):
                return datetime.combine(ts, datetime.min.time())
            return ts

        years = range(2021, 2036)  # This covers 2021 to 2035

        # All event days are dropped
        filtered_dispatches = self.dispatches[
            (self.dispatches["Program Type"] == self.event_type)
        ]

        date_drop_list = [
            convert_to_datetime(pd.to_datetime(event_date).date())
            for event_date in filtered_dispatches["Event Start"]
        ]

        # Add holidays to the list
        for year in years:
            date_drop_list.extend([
                datetime(year, 1, 1),  # New Year's Day
                datetime(year, 7, 4),  # Independence Day
                datetime(year, 12, 25),  # Christmas Day
                # Martin Luther King Jr. Day (third Monday in January)
                datetime(year, 1, 1) + timedelta(days=((7 - datetime(year, 1, 1).weekday()) + 1) % 7 + (3 - 1) * 7),
                # Presidents Day (third Monday in February)
                datetime(year, 2, 1) + timedelta(days=((7 - datetime(year, 2, 1).weekday()) + 1) % 7 + (3 - 1) * 7),
                # Memorial Day (last Monday in May)
                datetime(year, 5, 31) - timedelta(days=(datetime(year, 5, 31).weekday() + 1) % 7),
                # Labor Day (first Monday in September)
                datetime(year, 9, 1) + timedelta(days=(7 - datetime(year, 9, 1).weekday()) % 7) if datetime(year, 9, 1).weekday() != 0 else datetime(year, 9, 1),
                # Thanksgiving Day (fourth Thursday in November)
                datetime(year, 11, 1) + timedelta(days=((7 - datetime(year, 11, 1).weekday()) + 1) % 7 + (4 - 1) * 7),
            ])

        # Determine if the event is on a weekday or weekend
        is_weekend = self.event.weekday() >= 5  # 5 and 6 are Saturday and Sunday

        # Add weekends or weekdays to drop list based on the event day
        for year in years:
            start_date = datetime(year, 1, 1)
            end_date = datetime(year, 12, 31)

            current_date = start_date
            while current_date <= end_date:
                if is_weekend:
                    if current_date.weekday() < 5:  # Drop weekdays for weekend events
                        date_drop_list.append(current_date)
                else:
                    if current_date.weekday() >= 5:  # Drop weekends for weekday events
                        date_drop_list.append(current_date)
                current_date += timedelta(days=1)

        # Ensure no duplicates and return the list
        return list(set(date_drop_list))
    
