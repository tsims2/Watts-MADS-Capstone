from .modules_external import *

class DispatchDataProcessor:
    def __init__(self, file_path, year=None):
        self.file_path = file_path
        self.year = year
        self.dispatches_df = None
        self.isone_dispatches_df = None

        self.cs_daily_programs = [
            'NGRID-Connected Solutions-Daily Dispatch',
            'Rhode Island Connected Solutions-Daily Dispatch',
            'Eversource-Connected Solutions-Daily Dispatch',
            'Unitil-Connected Solutions-Daily Dispatch'
        ]
        self.cs_targeted_programs = [
            'NGRID-Connected Solutions-Targeted Dispatch',
            'Rhode Island Connected Solutions-Targeted Dispatch',
            'Cape Light Compact Program - Targeted Dispatch',
            'EVERSOURCE-Connected Solutions-Targeted Dispatch',
            'UNITIL-Connected Solutions-Targeted Dispatch',
            'Efficiency Maine-Demand Response Initiative',
            'Liberty-Connected Solutions-Targeted Dispatch'
        ]

    def read_data(self):
        self.dispatches_df = pd.read_excel(self.file_path)
        self.isone_dispatches_df = self.dispatches_df[self.dispatches_df['Region'] == 'ISONE']
        self.isone_dispatches_df.loc[:, 'Start Date & Time (Local)'] = pd.to_datetime(self.isone_dispatches_df['Start Date & Time (Local)'])
                
        if self.year:
            self.isone_dispatches_df = self.isone_dispatches_df[self.isone_dispatches_df['Start Date & Time (Local)'].dt.year == self.year]

    def get_dispatch_data(self):
        if self.isone_dispatches_df is None:
            self.read_data()

        def create_event_summary(programs, event_type=None):
            df = self.isone_dispatches_df[self.isone_dispatches_df['Program'].isin(programs)]
            if event_type:
                df = df[df['Event Type'] == event_type]
            
            summary = df[['Start Date & Time (Local)', 'End Date & Time (Local)', 'Program']].drop_duplicates()
            summary['Event Date'] = summary['Start Date & Time (Local)'].dt.date
            summary['Start Time'] = summary['Start Date & Time (Local)'].dt.time
            summary['End Time'] = summary['End Date & Time (Local)'].dt.time
            
            def get_program_type(program):
                if "Daily" in program:
                    return "Daily"
                elif "Targeted" in program or "Efficiency Maine" in program:
                    return "Targeted"
                elif "Active Demand" in program:
                    return "ADCR"
                else:
                    return "Unknown"
            
            summary['Program Type'] = summary['Program'].apply(get_program_type)
            summary['Program Name'] = summary['Program']  # Add this line to include Program Name
            
            return summary[['Event Date', 'Start Time', 'End Time', 'Program Type', 'Program Name']].sort_values('Event Date')

        def create_customer_df(programs, event_type=None):
            df = self.isone_dispatches_df[self.isone_dispatches_df['Program'].isin(programs)].copy()
            if event_type:
                df = df[df['Event Type'] == event_type]

            if 'Accnt #' in df.columns:
                df = df.rename(columns={'Accnt #': 'Utility Account Number'})
            
            # Ensure 'Program' column is included
            if 'Program' not in df.columns:
                df['Program'] = df['Program Name']
            
            return df

        cs_daily_dispatch_events = create_event_summary(self.cs_daily_programs)
        cs_targeted_dispatch_events = create_event_summary(self.cs_targeted_programs)

        cs_daily_customer_dispatches = create_customer_df(self.cs_daily_programs)
        cs_targeted_customer_dispatches = create_customer_df(self.cs_targeted_programs)

        return (
            cs_daily_dispatch_events,
            cs_daily_customer_dispatches,
            cs_targeted_dispatch_events,
            cs_targeted_customer_dispatches,
        )