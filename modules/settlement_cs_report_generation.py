from .modules_external import *

class CSReportGenerator:
    def __init__(self, asset_id, customer, season, event, event_type, customer_asset, utility_account_number, address, customer_share, dispatches, customer_dispatches):
        self.asset_id = asset_id
        self.customer = customer
        self.season = season
        self.event = event
        self.event_type = event_type
        self.customer_asset = customer_asset
        self.dispatches = dispatches
        self.customer_dispatches = customer_dispatches
        self.utility_account_number = utility_account_number
        self.customer_share = customer_share
        self.address = address

        os.makedirs("Data/Temp Assets", exist_ok=True)

    def aggregate_reports(self, comp_card, report_type, report_format, month=None, selected_events=None):
        self.program = comp_card['Program'].iloc[0]
        self.report_format = report_format  # Store report_format as an instance variable

        pdf_writer = PyPDF2.PdfWriter()
        
        if report_format == "Full":
            # Create cover page
            cover_page = self.create_cover_page(comp_card, report_type, month, selected_events)
            
            pdf_writer.add_page(cover_page)
            
            # Create table of contents
            toc_buffer = self.create_table_of_contents(comp_card, report_type, month, selected_events)
            toc_reader = PdfReader(toc_buffer)
            for page in toc_reader.pages:
                pdf_writer.add_page(page)

            # Update self.toclength
            self.toclength = len(toc_reader.pages)

            # At the start of the aggregate_reports method:
            base_page_number = self.toclength + 2  # ToC and summary
        else:
            self.toclength = 0
            base_page_number = 1

        # Create summary page
        if report_type == "Event":
            summary_page = self.create_event_summary_page(comp_card)
        elif report_type == "Monthly":
            summary_page = self.create_monthly_summary_page(comp_card, month)
        elif report_type == "Custom":
            summary_page = self.create_custom_summary_page(comp_card, selected_events)
        
        pdf_writer.add_page(summary_page)

        if report_format == "Full":
            # Group by company
            for company, company_data in comp_card.groupby('Customer'):
                company_page_number = base_page_number

                if report_type == "Event":
                    # Add individual asset reports
                    for asset, data in company_data.iterrows():
                        asset_page_number = company_page_number
                        self.asset_id = data['Asset ID']
                        self.utility_account_number = data['Utility Account Number']
                        self.address = data['Facility Address']
                        self.customer_asset = data['Customer Asset']
                        program = data['Program']
                        curtailment_strategy = data['Curtailment Strategy']

                        asset_report = self.create_asset_report(program, asset_page_number)

                        pdf_writer.add_page(asset_report)
                        company_page_number += 1

                    # Update base_page_number for the next company
                    base_page_number = company_page_number

                elif report_type == "Custom":
                    base_page_number = self.toclength + 2  # ToC and summary

                    for asset, asset_data in company_data.groupby('Asset ID'):
                        # Create asset divider page
                        asset_divider = self.create_asset_divider_page(asset_data['Customer Asset'].iloc[0], asset_data['Utility Account Number'].iloc[0])
                        pdf_writer.add_page(asset_divider)
                        base_page_number += 1

                        for event_date in selected_events:
                            self.event = event_date
                            self.asset_id = asset
                            self.utility_account_number = asset_data['Utility Account Number'].iloc[0]
                            self.address = asset_data['Facility Address'].iloc[0]
                            self.customer_asset = asset_data['Customer Asset'].iloc[0]
                            program = asset_data['Program'].iloc[0]
                            curtailment_strategy = asset_data['Curtailment Strategy'].iloc[0]

                            asset_report = self.create_asset_report(program, base_page_number)

                            pdf_writer.add_page(asset_report)
                            base_page_number += 1
                            
                elif report_type == "Monthly":
                    events = self.get_monthly_events(month, self.program)
                    
                    # Add individual asset reports for each event
                    for asset_index, (_, data) in enumerate(company_data.iterrows()):
                        # Create asset divider page
                        divider_page = self.create_asset_divider_page(data['Customer Asset'], data['Utility Account Number'])
                        pdf_writer.add_page(divider_page)
                        company_page_number += 1

                        for event_index, event in enumerate(events):
                            asset_page_number = company_page_number + event_index
                            
                            self.asset_id = data['Asset ID']
                            self.utility_account_number = data['Utility Account Number']
                            self.address = data['Facility Address']
                            self.customer_asset = data['Customer Asset']
                            self.event = event.date()
                            program = data['Program']
                            curtailment_strategy = data['Curtailment Strategy']

                            asset_report = self.create_asset_report(program, asset_page_number)
                                
                            pdf_writer.add_page(asset_report)

                        company_page_number += len(events)

                    # Update base_page_number for the next company
                    base_page_number = company_page_number

        # Save the complete report
        output_path = self.get_output_path(comp_card['Customer'].iloc[0], report_type, month, report_format)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with open(output_path, 'wb') as output_file:
            pdf_writer.write(output_file)

    def create_asset_divider_page(self, customer_asset, uan):
        buffer = BytesIO()
        c = canvas.Canvas(buffer, pagesize=letter)
        width, height = letter

        # Set background color
        c.setFillColor(colors.Color(246/255, 246/255, 234/255))  # Creme color
        c.rect(0, 0, width, height, fill=True)

        # Set text color and font
        c.setFillColor(colors.Color(2/255, 2/255, 68/255))  # Navy color
        c.setFont("Public Sans-Bold", 14)

        # Draw centered text
        text = f"{customer_asset}\n({uan})"
        textobject = c.beginText()
        textobject.setTextOrigin(50, height / 2)
        for line in text.split('\n'):
            textobject.textLine(line)
        c.drawText(textobject)

        c.showPage()
        c.save()
        buffer.seek(0)
        return PdfReader(buffer).pages[0]

    def create_cover_page(self, comp_card, report_type, month=None, selected_events=None):
        if report_type == "Monthly":
            # Convert month name to number
            month_number = list(month_abbr).index(month[:3].title())
            report_start = datetime.strptime(f'2024-{month_number:02d}-01 00:00:00', '%Y-%m-%d %H:%M:%S')
            report_end = report_start + relativedelta(day=31)
            first_day = datetime.strftime(report_start, '%m/%d/%Y')
            last_day = datetime.strftime(report_end, '%m/%d/%Y')
            report_time_string = f"{month} Results | {first_day} - {last_day}"
        elif report_type == "Custom":
            start_date = min(selected_events).strftime('%m/%d/%Y')
            end_date = max(selected_events).strftime('%m/%d/%Y')
            report_time_string = f"Custom Report | {start_date} - {end_date}"
        else:  # Event
            first_day = datetime.strftime(self.event, '%m/%d/%Y')
            report_time_string = f"Event Results | {first_day}"

        # Create a BytesIO object to hold the PDF content
        buffer = BytesIO()
        c = canvas.Canvas(buffer, pagesize=letter)

        page_width, page_height = letter

        image_path = r"C:\Users\Tyler.Sims\OneDrive - CPower\Documents\CPower\ISONE\ISONEdb\Branding\CPower_Icon_Stacked_Full_Color.png"
        background_image_path = "C:/Users/Tyler.Sims/OneDrive - CPower/Documents/CPower/ISONE/ISONEdb/Branding/New England Stock Image.jpg"

        # Draw background image
        with Image.open(background_image_path) as img:
            img_width, img_height = img.size
            img_width = img_width * 0.95
            img_height = img_height * 0.95
        c.drawImage(background_image_path, -450, 50, width=img_width, height=img_height, mask='auto')

        # Draw white box
        box_x, box_y, box_width, box_height = 0, 0, 612, 264
        c.setStrokeColor('#F6F6EA')
        c.setLineWidth(0)
        c.setFillColor('#F6F6EA')
        c.rect(box_x, box_y, box_width, box_height, fill=True)

        # Draw CPower logo
        with Image.open(image_path) as img:
            img_width, img_height = img.size
            img_width = img_width * 0.08
            img_height = img_height * 0.08
        c.drawImage(image_path, page_width/2 - img_width/2, 60, width=img_width, height=img_height, mask='auto')

        reversed_program_tab_name_mapping = {
            "CLT": "Cape Light Compact Program - Targeted Dispatch",
            "EMT": "Efficiency Maine-Demand Response Initiative",
            "EVD": "Eversource-Connected Solutions-Daily Dispatch",
            "EVT": "EVERSOURCE-Connected Solutions-Targeted Dispatch",
            "LBT": "Liberty-Connected Solutions-Targeted Dispatch",
            "NGD": "NGRID-Connected Solutions-Daily Dispatch",
            "NGT": "NGRID-Connected Solutions-Targeted Dispatch",
            "RID": "Rhode Island Connected Solutions-Daily Dispatch",
            "RIT": "Rhode Island Connected Solutions-Targeted Dispatch",
            "UND": "Unitil-Connected Solutions-Daily Dispatch",
            "UNT": "UNITIL-Connected Solutions-Targeted Dispatch",
        }

        utility_name = reversed_program_tab_name_mapping[comp_card['Program'].iloc[0]]

        # Draw text
        c.setFont("Public Sans-Bold", 20)
        c.setFillColor(color_palette[5])
        c.drawCentredString(page_width/2, 220, f"{comp_card['Customer'].iloc[0]}")

        c.setFont("Public Sans", 14)
        c.setFillColor('black')
        c.drawCentredString(page_width/2, 190, "Preliminary Report")
        c.setFont("Public Sans", 14)
        c.drawCentredString(page_width/2, 165, f"{utility_name}")

        c.setFont("Public Sans", 14)
        c.drawCentredString(page_width/2, 140, f"{report_time_string}")

        c.setFont("Public Sans", 10)
        c.setFillColor(color_palette[5])
        c.drawCentredString(page_width/2, 40, "a CPowered report")

        c.showPage()
        c.save()

        # Move buffer position to the beginning
        buffer.seek(0)
        return PdfReader(buffer).pages[0]
    
    def create_table_of_contents(self, comp_card, report_type, month=None, selected_events=None):
        buffer = BytesIO()
        c = canvas.Canvas(buffer, pagesize=letter)
        width, height = letter

        def calculate_toc_length():
            total_lines = 1  # Start with 1 for the Summary line
            for _ in comp_card.iterrows():
                total_lines += 1  # Asset name
                if report_type == "Monthly":
                    total_lines += 1  # Divider page
                total_lines += len(events)  # Events for this asset
                total_lines += 1  # Spacing between assets
            return (total_lines // 45) + 1  # 45 lines per page, round up

        # Get events based on report type
        if report_type == "Event":
            events = [self.event]
        elif report_type == "Monthly":
            events = self.get_monthly_events(month, self.program)
        elif report_type == "Custom":
            events = selected_events
        else:
            events = []

        # Calculate ToC length
        self.toclength = calculate_toc_length()

        def new_page():
            nonlocal y_position
            c.showPage()
            c.setFont("Public Sans-Bold", 16)
            c.drawString(50, height - 50, "Table of Contents (continued)")
            y_position = height - 80

        # Set up the first page
        c.setFont("Public Sans-Bold", 16)
        c.drawString(50, height - 50, "Table of Contents")

        y_position = height - 80
        content_page_number = self.toclength + 1  # Start numbering after ToC

        # Add Summary entry (in bold)
        c.setFont("Public Sans-Bold", 12)
        c.drawString(50, y_position, "Summary")
        c.setFont("Public Sans", 11)
        c.drawString(500, y_position, str(content_page_number))
        y_position -= 20
        content_page_number += 1

        # List each asset and its events
        for _, asset in comp_card.iterrows():
            if y_position < 50:
                new_page()

            c.setFont("Public Sans-Bold", 12)
            asset_line = f"{asset['Customer Asset']} ({asset['Asset ID']})"
            c.drawString(50, y_position, asset_line)

            if report_type == "Monthly":
                c.drawString(500, y_position, str(content_page_number))
                content_page_number += 1  # Account for divider page

            y_position -= 20

            c.setFont("Public Sans", 11)
            for event in events:
                if y_position < 50:
                    new_page()
                    c.setFont("Public Sans", 11)  # Reset font after new page

                event_line = f"Event - {event.strftime('%m/%d/%Y')}"
                c.drawString(70, y_position, event_line)
                c.drawString(500, y_position, str(content_page_number))
                y_position -= 15
                content_page_number += 1

            y_position -= 10  # Add some space between assets

        # Save the last page
        c.showPage()
        c.save()
        buffer.seek(0)

        return buffer
    
    def get_monthly_events(self, month, program):
        # Convert month name to number
        month_names = ['January', 'February', 'March', 'April', 'May', 'June',
                        'July', 'August', 'September', 'October', 'November', 'December']
        month_number = month_names.index(month) + 1

        # Get the year from self.season (assuming it's in format "Summer YYYY")
        year = int(self.season.split()[-1])

        # Get the start and end date of the month
        _, last_day = monthrange(year, month_number)
        start_date = datetime(year, month_number, 1)
        end_date = datetime(year, month_number, last_day)

        # Filter events for the specified month
        monthly_events = []
        for _, event_data in self.dispatches.iterrows():
            event_date = datetime.strptime(str(event_data['Event Date']), '%Y-%m-%d')
            if start_date <= event_date <= end_date:
                if "D" in program and "Daily" in event_data['Program Type']:
                    monthly_events.append(event_date)
                elif "D" not in program and "Daily" not in event_data['Program Type']:
                    monthly_events.append(event_date)

        return sorted(monthly_events)

    def create_event_summary_page(self, comp_card):
        buffer = BytesIO()
        c = canvas.Canvas(buffer, pagesize=letter)
        width, height = letter

        company = comp_card['Customer'].iloc[0]
        program = comp_card['Program'].iloc[0]
        utility_name = program.split("-")[0]

        # Add the text to the canvas
        textobject = c.beginText(35, 700)
        textobject.setFont("Public Sans-Bold", 12)
        textobject.setFillColor(color_palette[5]) 
        textobject.textLines(f'''{company}''')
        c.drawText(textobject)

        def ordinal(n):
            if 10 <= n % 100 <= 20:
                suffix = 'th'
            else:
                suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
            return f"{n}{suffix}"

        period = f"Event on {self.event.strftime('%B %d, %Y')}"

        # Filter dispatches for the current season and sort from latest to earliest
        season_start = pd.to_datetime(f"{self.season.split()[-1]}-06-01").date()  # Convert to date
        season_end = pd.to_datetime(f"{int(self.season.split()[-1]) + 1}-05-31").date()  # Convert to date

        # Convert 'Event Date' to date objects for comparison
        self.dispatches['Event Date'] = pd.to_datetime(self.dispatches['Event Date']).dt.date

        season_dispatches = self.dispatches[
            (self.dispatches['Event Date'] >= season_start) & 
            (self.dispatches['Event Date'] <= season_end)
        ].sort_values('Event Date', ascending=False)

        # Separate daily and targeted dispatches
        daily_dispatches = season_dispatches[season_dispatches['Program Type'] == 'Daily']
        targeted_dispatches = season_dispatches[season_dispatches['Program Type'] == 'Targeted']

        # Convert self.event to date for comparison
        event_date = pd.to_datetime(self.event).date()

        # Find the index of the current event
        if self.event_type == 'Daily':
            event_index = daily_dispatches[daily_dispatches['Event Date'] == event_date].index[0]
            true_perf_len = len(daily_dispatches.loc[:event_index])
        else:  # Targeted
            event_index = targeted_dispatches[targeted_dispatches['Event Date'] == event_date].index[0]
            true_perf_len = len(targeted_dispatches.loc[:event_index])

        true_perf_len_ordinal = ordinal(true_perf_len)

        # Add the text to the canvas
        textobject = c.beginText(35, 680)
        textobject.setFont("Libre Franklin", 11)
        textobject.setFillColor('black')  
        if utility_name == 'Efficiency Maine':
            textobject.textLines(f'''
                Thank you for your recent performance in the Demand Response Initiative Program. This is your 
                {true_perf_len_ordinal} event of the year. As a reminder, your seasonal performance and final payout are the average 
                performance of your top three events.
                ''')
        else:
            textobject.textLines(f'''
                Thank you for your recent performance in the Connected Solutions {self.event_type} program. This is your 
                {true_perf_len_ordinal} event of the year. As a reminder, your seasonal performance and final payout are the average 
                performance of all your events.
                ''')
        c.drawText(textobject)

        # Add the text to the canvas
        textobject = c.beginText(35, 620)
        textobject.setFont("Public Sans-Bold", 11)
        textobject.setFillColor('black')  
        textobject.textLines(f'''Aggregated Performance Summary for Season''')
        c.drawText(textobject)

        # Calculate performance metrics
        total_performance = comp_card['Average Performance'].dropna().loc[lambda x: x >= 0].sum()

 
        customer_share = float(self.customer_share)
        program = comp_card['Program'].iloc[0]
        if "D" in program and "RID" in program:
            program_price = 300
        elif "D" in program and "RID" not in program:
            program_price = 200
        elif "T" in program and "RIT" in program:
            program_price = 40
        elif "T" in program and "RIT" not in program:
            program_price = 35

        def format_number(num):
            if num < 0:
                return f"({abs(int(round(num))):,})"
            else:
                return f"{int(round(num)):,}"

        def format_currency(num):
            if num < 0:
                return f"(${abs(num):,.2f})"
            else:
                return f"${num:,.2f}"

        if utility_name == 'Efficiency Maine':
            lower_perf = total_performance * 0.9
            upper_perf = total_performance
            lower_rev = max(total_performance * program_price * 0.9 * customer_share, 0)
            upper_rev = max(total_performance * program_price * customer_share, 0)
        else:
            lower_perf = total_performance * 0.9
            upper_perf = total_performance
            lower_rev = max(total_performance * program_price * 0.9 * customer_share, 0)
            upper_rev = max(total_performance * program_price * customer_share, 0)

        if lower_perf < upper_perf:
            Rolling_Performance = f"{format_number(lower_perf)} kW - {format_number(upper_perf)} kW"
            Rolling_Revenue = f"{format_currency(lower_rev)} - {format_currency(upper_rev)}"
        else:
            Rolling_Performance = f"{format_number(upper_perf)} kW - {format_number(lower_perf)} kW"
            Rolling_Revenue = f"{format_currency(upper_rev)} - {format_currency(lower_rev)}"

        # Add the text to the canvas
        textobject = c.beginText(55, 600)
        textobject.setFont("Libre Franklin", 11)
        textobject.setFillColor('black')  
        textobject.textLines(f'''
            Assets Dispatched:
            Events This Season: 
            Rolling Seasonal Projected Performance:
            Rolling Seasonal Projected Revenues:
            ''')
        c.drawText(textobject)

        # Add the text to the canvas
        textobject = c.beginText(280, 600)
        textobject.setFont("Libre Franklin", 11)
        textobject.setFillColor('black')  
        textobject.textLines(f'''
            {len(comp_card)}
            {true_perf_len}
            {Rolling_Performance}
            {Rolling_Revenue}
            ''')
        c.drawText(textobject)

        if len(comp_card['Customer Asset']) > 1:
            # Add the text to the canvas
            textobject = c.beginText(35, 520)
            textobject.setFont("Public Sans-Bold", 11)
            textobject.setFillColor('black')  
            textobject.textLines(f'''Performance by Asset for Event {period}''')
            c.drawText(textobject)

            plt.rcParams['font.family'] = 'Public Sans'
            plt.rcParams['font.style'] = 'normal'
            plt.rcParams['font.size'] = 12
            plt.rcParams['font.weight'] = 'normal'

            fig, ax = plt.subplots(figsize=(12, 4.5))

            # Calculate bar positions
            x = np.arange(len(comp_card['Utility Account Number']))
            width = 0.8  # Width of the bars

            # Plot the bar chart with green color
            bars = ax.bar(x, comp_card['Average Performance'], width, color='green')

            max_performance = comp_card['Average Performance'].max()
            if max_performance > 0.0001:  # Check if max performance is greater than a small threshold
                ax.set_ylim(0, max(max_performance * 1.1, 0.1))  # Add 10% padding at the top, with a minimum of 0.1
            else:
                ax.set_ylim(0, 1)  # Set y-axis from 0 to 1 when max performance is 0 or very small

            # Add performance values above the bars
            for bar, performance in zip(bars, comp_card['Average Performance']):
                value = bar.get_height()
                height = max(bar.get_height(), 0)
                error = height * 0.1
                ax.errorbar(bar.get_x() + bar.get_width() / 2, height, yerr=error, fmt='none', color=color_palette[0], capsize=0)
                '''color = 'green' if height > 0 else 'red'
                ax.text(bar.get_x() + bar.get_width() / 2, height, f'{round(value)}', 
                        ha='center', va='bottom', color=color, fontsize=12, fontweight='bold')'''

            # Add labels and formatting
            ax.set_xlabel('Assets')
            ax.set_ylabel('kW')

            # Improved format_label function
            def format_label(label, max_chars=10):
                if len(label) > max_chars:
                    return label[:max_chars-3] + '...'
                return label
            
            fig_width_inches = fig.get_size_inches()[0]
            char_width_inches = 0.15  # Approximate width of a character
            max_chars = int((fig_width_inches / len(comp_card['Utility Account Number'])) / char_width_inches)
            max_chars = max(5, min(15, max_chars))  # Ensure max_chars is between 5 and 15

            # Set x-ticks and labels
            ax.set_xticks(x)
            # Create two-level labels
            labels = []
            for asset, account in zip(comp_card['Customer Asset'], comp_card['Utility Account Number']):
                asset_label = format_label(asset, max_chars)
                account_label = format_label(account, max_chars)
                labels.append(f"{asset_label}\n{account_label}")

            # Set x-labels with two levels
            ax.set_xticklabels(labels, rotation=0, ha='center', va='top', fontsize=10)

            # Adjust bottom margin to accommodate two-line labels
            plt.subplots_adjust(bottom=0.2)

            # Remove borders
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)

            # Add horizontal grid lines
            ax.yaxis.grid(True, linestyle='--', which='major', color='grey', alpha=.25)

            # Adjust layout to prevent cutting off labels
            #plt.tight_layout()

            # Save the plot
            plt.savefig("Data/Temp Assets/performance_by_asset.png", dpi=290, bbox_inches='tight')

            # Get the dimensions of the image
            img = Image.open("Data/Temp Assets/performance_by_asset.png")
            img_width, img_height = img.size
            mag = .17

            # Add the centered image to the canvas
            c.drawImage("Data/Temp Assets/performance_by_asset.png", 35, 370, width=img_width*mag, height=img_height*mag)

            # Close the figure
            plt.close(fig)

            chart_space = 0

        else:
            chart_space = 230

        # Add the text to the canvas
        textobject = c.beginText(35, 290 + chart_space)
        textobject.setFont("Libre Franklin-Medium", 11)
        textobject.setFillColor(color_palette[0]) 
        textobject.textLines(f'''
            Individual preliminary performance reports are available below for each asset.
        ''')
        c.drawText(textobject)

        # Add the text to the canvas
        textobject = c.beginText(35, 263+ chart_space)
        textobject.setFont("Libre Franklin", 11)
        textobject.setFillColor('black') 
        textobject.textLines(f'''
            For more detailed load data please visit the

            Best,

            The CPower Team
        ''')
        c.drawText(textobject)

        hyperlink = "CPower Portal (portal.cpowercorp.com)."
        textobject = c.beginText(253, 263+ chart_space)  # Adjust x position according to the length of the preceding text
        textobject.setFont("Libre Franklin-Medium", 11)
        textobject.setFillColor(color_palette[8])  # Set color to blue
        textobject.textLine(hyperlink)
        c.drawText(textobject)

        link = "https://nam02.safelinks.protection.outlook.com/?url=https%3A%2F%2Fportal.cpowercorp.com%2FAuthentication&data=05%7C01%7CTyler.Sims%40CPowerEnergyManagement.com%7C40328f4a152944bd80a108db7f286350%7C36d907c0d528449a85027ed7ff4b32dc%7C1%7C0%7C638243583688071448%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C3000%7C%7C%7C&sdata=jpbysi1AUhcZLUFdofrrkLILDy4e8hbmAwxTidJ3B3c%3D&reserved=0"
        c.linkURL(link, (253, 263 + chart_space, 253 + 200, 263 + 15 + chart_space), relative=1)

        textobject = c.beginText(35, 40)
        textobject.setFont("Libre Franklin", 10)
        textobject.setFillColor(color_palette[7])  
        textobject.textLines(f'''
            *Preliminary reports are for discussion purposes only. Final results are pending until settlement with utility.
            *Seasonal rolling projections use a 10% margin of error.
            ''')
        c.drawText(textobject)

        if self.report_format == "Full":
            page_number = 1 + self.toclength

            c.saveState()
            c.setFont("Helvetica", 8)
            c.drawRightString(570, 25, f"Page {page_number}")
            c.restoreState()

        c.showPage()
        c.save()
        buffer.seek(0)

        return PdfReader(buffer).pages[0]

    def create_monthly_summary_page(self, comp_card, month=None):
        # Convert month name to number
        month_num = pd.to_datetime(month, format='%B').month

        buffer = BytesIO()
        c = canvas.Canvas(buffer, pagesize=letter)
        width, height = letter

        company = comp_card['Customer'].iloc[0]
        program = comp_card['Program'].iloc[0]
        utility_name = program.split("-")[0]

        # Add the text to the canvas
        textobject = c.beginText(35, 710)
        textobject.setFont("Public Sans-Bold", 12)
        textobject.setFillColor(color_palette[5]) 
        textobject.textLines(f'''{company}''')
        c.drawText(textobject)
        
        # Filter events for the specified month from the dispatch list
        month_events = [event for event in self.dispatches['Event Date'] 
                        if pd.to_datetime(event).month == month_num]
        true_event_len = len(self.dispatches[self.dispatches['Program Name'] == comp_card['Full Program Name'].iloc[0]])

        # Create a summary table
        summary_data = []
        for event in month_events:
            event_date = pd.to_datetime(event).strftime('%Y-%m-%d')
            event_column = f'Event {event_date}'
            if event_column in comp_card.columns:
                total_assets = comp_card[event_column].count()
                total_performance = comp_card[event_column].sum()
                avg_performance = comp_card[event_column].mean()
                summary_data.append([event_date, total_assets, total_performance, avg_performance])

        summary_df = pd.DataFrame(summary_data, columns=['Event Date', 'Total Assets', 'Total Performance (kW)', 'Average Performance (kW)'])

        true_perf_len = len(month_events)

        # Add the text to the canvas
        textobject = c.beginText(35, 680)
        textobject.setFont("Libre Franklin", 11)
        textobject.setFillColor('black')  
        if utility_name == 'Efficiency Maine':
            textobject.textLines(f'''
                Thank you for your recent performance in the Demand Response Initiative Program. This month has had 
                {true_perf_len} events. As a reminder, your seasonal performance and final payout are the average performance of 
                your top three events this year.
                ''')
        else:
            textobject.textLines(f'''
                Thank you for your recent performance in the Connected Solutions Program! 

                There as been {true_event_len} events this season. As a reminder, your seasonal incentive is calculated using the average
                performance of all events. Negative performances impact your seasonal performance average.
                ''')
        c.drawText(textobject)

        # Add the text to the canvas
        textobject = c.beginText(35, 610)
        textobject.setFont("Public Sans-Bold", 11)
        textobject.setFillColor('black')  
        textobject.textLines(f'''Aggregated Performance Summary for {month}''')
        c.drawText(textobject)

        # Calculate performance metrics
        total_performance = comp_card['Average Performance'].dropna().sum()

        customer_share = float(self.customer_share)
        program = comp_card['Program'].iloc[0]
        if "D" in program and "RID" in program:
            program_price = 300
        elif "D" in program and "RID" not in program:
            program_price = 200
        elif "T" in program and "RIT" in program:
            program_price = 40
        elif "T" in program and "RIT" not in program:
            program_price = 35

        def format_number(num):
            if num < 0:
                return f"({abs(int(round(num))):,})"
            else:
                return f"{int(round(num)):,}"

        def format_currency(num):
            if num < 0:
                return f"(${abs(num):,.2f})"
            else:
                return f"${num:,.2f}"

        if utility_name == 'Efficiency Maine':
            lower_perf = total_performance * 0.9
            upper_perf = total_performance
            lower_rev = max(total_performance * program_price * 0.9 * customer_share, 0)
            upper_rev = max(total_performance * program_price * customer_share, 0)
        else:
            lower_perf = total_performance * 0.9
            upper_perf = total_performance
            lower_rev = max(total_performance * program_price * 0.9 * customer_share, 0)
            upper_rev = max(total_performance * program_price * customer_share, 0)

        if lower_perf < upper_perf:
            Rolling_Performance = f"{format_number(lower_perf)} kW - {format_number(upper_perf)} kW"
            Rolling_Revenue = f"{format_currency(lower_rev)} - {format_currency(upper_rev)}"
        else:
            Rolling_Performance = f"{format_number(upper_perf)} kW - {format_number(lower_perf)} kW"
            Rolling_Revenue = f"{format_currency(upper_rev)} - {format_currency(lower_rev)}"

        # Add the text to the canvas
        textobject = c.beginText(55, 590)
        textobject.setFont("Libre Franklin", 11)
        textobject.setFillColor('black')  
        textobject.textLines(f'''
            Assets Dispatched:
            Events This Month: 
            Rolling Seasonal Projected Performance:
            Rolling Seasonal Projected Revenues:
            ''')
        c.drawText(textobject)

        # Add the text to the canvas
        textobject = c.beginText(275, 590)
        textobject.setFont("Libre Franklin", 11)
        textobject.setFillColor('black')  
        textobject.textLines(f'''
            {len(comp_card)}
            {true_perf_len}
            {Rolling_Performance}
            {Rolling_Revenue}
            ''')
        c.drawText(textobject)

        # Add the text to the canvas
        textobject = c.beginText(35, 520)
        textobject.setFont("Public Sans-Bold", 11)
        textobject.setFillColor('black')  
        textobject.textLines(f'''Performance by Event''')
        c.drawText(textobject)

        plt.rcParams['font.family'] = 'Public Sans'
        plt.rcParams['font.style'] = 'normal'
        plt.rcParams['font.size'] = 12
        plt.rcParams['font.weight'] = 'normal'

        fig, ax = plt.subplots(figsize=(12,3))

        # Ensure bar heights don't go below 0
        bar_heights = [max(0, perf) for perf in summary_df['Total Performance (kW)']]

        bars = ax.bar(summary_df['Event Date'], bar_heights, color=color_palette[9])        

        ax.set_xlabel(f'{self.season} Event Date')
        ax.set_ylabel('Total Performance (kW)')

        # Remove borders
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        plt.xticks(rotation=45, ha='right')

        # Color weekend dates in gray
        for label in ax.get_xticklabels():
            date = label.get_text()  # Get the x-coordinate of the label
            if pd.to_datetime(date).weekday() >= 5:  # 5 and 6 are Saturday and Sunday
                label.set_color('gray')

        # Get current x-tick labels and positions
        ticks = ax.get_xticks()
        labels = [item.get_text() for item in ax.get_xticklabels()]

        new_labels = []
        for date in labels:
            try:
                date_parsed = pd.to_datetime(date)
                # Reformat the date to remove the year
                new_text = date_parsed.strftime('%b %d')
                new_labels.append(new_text)
            except ValueError:
                # Handle any parsing errors for dates that might not be formatted correctly
                new_labels.append(date)

        # Set new x-tick labels
        ax.set_xticks(ticks)
        ax.set_xticklabels(new_labels)

        # Set y-axis limits
        y_max = max(bar_heights)
        ax.set_ylim(0, y_max * 1.1)  # Add 10% padding at the top, starting from 0

        # Add horizontal grid lines
        ax.yaxis.grid(True, linestyle='--', which='major', color='grey', alpha=.25)

        # Add performance values above the bars
        for bar, performance in zip(bars, summary_df['Total Performance (kW)']):
            height = bar.get_height()
            error = height * 0.1
            ax.errorbar(bar.get_x() + bar.get_width() / 2, height, yerr=error, fmt='none', color=color_palette[0], capsize=0)
            label = f'{int(round(abs(performance)))}'
        # Save the plot
        plt.savefig("Data/Temp Assets/performance_by_asset.png", dpi=290, bbox_inches='tight')

        # Add the centered image to the canvas
        img = Image.open("Data/Temp Assets/performance_by_asset.png")
        img_width, img_height = img.size
        mag = .17

        c.drawImage("Data/Temp Assets/performance_by_asset.png", 35, 330, width=img_width*mag, height=img_height*mag)

        # Close the figure
        plt.close(fig)

        # Add the text to the canvas
        textobject = c.beginText(35, 300)
        textobject.setFont("Libre Franklin-Medium", 11)
        textobject.setFillColor(color_palette[0]) 
        if self.report_format == "Full":
            textobject.textLines(f'''
                Individual preliminary performance reports are available below for each asset.
            ''')
        else:
            textobject.textLines(f'''
                Individual preliminary performance reports are available upon request.
            ''')

        c.drawText(textobject)

        # Add the text to the canvas
        textobject = c.beginText(35, 275)
        textobject.setFont("Libre Franklin", 11)
        textobject.setFillColor('black') 
        textobject.textLines(f'''
            For more detailed load data please visit the

            If you have questions please reach out to Northeast@CPowerEnergyManagement.com.
            
            Best,

            The Northeast CPower Team
        ''')
        c.drawText(textobject)

        hyperlink = "CPower Portal (portal.cpowercorp.com)."
        textobject = c.beginText(253, 275)  # Adjust x position according to the length of the preceding text
        textobject.setFont("Libre Franklin-Medium", 11)
        textobject.setFillColor(color_palette[8])  # Set color to blue
        textobject.textLine(hyperlink)
        c.drawText(textobject)

        link = "https://nam02.safelinks.protection.outlook.com/?url=https%3A%2F%2Fportal.cpowercorp.com%2FAuthentication&data=05%7C01%7CTyler.Sims%40CPowerEnergyManagement.com%7C40328f4a152944bd80a108db7f286350%7C36d907c0d528449a85027ed7ff4b32dc%7C1%7C0%7C638243583688071448%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C3000%7C%7C%7C&sdata=jpbysi1AUhcZLUFdofrrkLILDy4e8hbmAwxTidJ3B3c%3D&reserved=0"
        c.linkURL(link, (253, 275, 253 + 200, 275 + 15), relative=1)

        textobject = c.beginText(35, 40)
        textobject.setFont("Libre Franklin", 10)
        textobject.setFillColor(color_palette[7])  
        textobject.textLines(f'''
            *Preliminary reports are for discussion purposes only. Final results are pending until settlement with utility.
            *Seasonal rolling projections use a 10% margin of error.
            ''')
        c.drawText(textobject)

        if self.report_format == "Full":
            c.saveState()
            c.setFont("Helvetica", 8)
            c.drawRightString(570, 25, f"Page {1 + self.toclength}")
            c.restoreState()

        c.showPage()
        c.save()
        buffer.seek(0)
        return PdfReader(buffer).pages[0]

    def create_custom_summary_page(self, comp_card, selected_events):
        buffer = BytesIO()
        c = canvas.Canvas(buffer, pagesize=letter)
        width, height = letter

        company = comp_card['Customer'].iloc[0]
        program = comp_card['Program'].iloc[0]
        utility_name = program.split("-")[0]

        # Add the text to the canvas
        textobject = c.beginText(35, 700)
        textobject.setFont("Public Sans-Bold", 12)
        textobject.setFillColor(color_palette[5]) 
        textobject.textLines(f'''{company}''')
        c.drawText(textobject)

        period = f"Custom Report for {len(selected_events)} events"
        start_date = min(selected_events).strftime('%B %d, %Y')
        end_date = max(selected_events).strftime('%B %d, %Y')

        # Add the text to the canvas
        textobject = c.beginText(35, 680)
        textobject.setFont("Libre Franklin", 11)
        textobject.setFillColor('black')  
        if utility_name == 'Efficiency Maine':
            textobject.textLines(f'''
                Thank you for your recent performance in the Demand Response Initiative Program. This report covers 
                {len(selected_events)} events from {start_date} to {end_date}. As a reminder, your seasonal performance and final 
                payout are the average performance of your top three events.
                ''')
        else:
            textobject.textLines(f'''
                Thank you for your recent performance in the Connected Solutions {self.event_type} program. This report covers 
                {len(selected_events)} events from {start_date} to {end_date}. As a reminder, your seasonal performance and final 
                payout are the average performance of all your events.
                ''')
        c.drawText(textobject)

        # Add the text to the canvas
        textobject = c.beginText(35, 620)
        textobject.setFont("Public Sans-Bold", 11)
        textobject.setFillColor('black')  
        textobject.textLines(f'''Aggregated Performance Summary for Selected Events''')
        c.drawText(textobject)

        # Calculate performance metrics
        event_columns = [f"Event {event}" for event in selected_events]

        total_performance = comp_card[event_columns].mean(axis=1).sum()

        customer_share = float(self.customer_share)
        if "D" in program and "RID" in program:
            program_price = 300
        elif "D" in program and "RID" not in program:
            program_price = 200
        elif "T" in program and "RIT" in program:
            program_price = 40
        elif "T" in program and "RIT" not in program:
            program_price = 35

        def format_number(num):
            if num < 0:
                return f"({abs(int(round(num))):,})"
            else:
                return f"{int(round(num)):,}"

        def format_currency(num):
            if num < 0:
                return f"(${abs(num):,.2f})"
            else:
                return f"${num:,.2f}"

        if utility_name == 'Efficiency Maine':
            lower_perf = total_performance * 0.9
            upper_perf = total_performance
            lower_rev = max(total_performance * program_price * 0.9 * customer_share, 0)
            upper_rev = max(total_performance * program_price * customer_share, 0)
        else:
            lower_perf = total_performance * 0.9
            upper_perf = total_performance
            lower_rev = max(total_performance * program_price * 0.9 * customer_share, 0)
            upper_rev = max(total_performance * program_price * customer_share, 0)

        if lower_perf < upper_perf:
            Rolling_Performance = f"{format_number(lower_perf)} kW - {format_number(upper_perf)} kW"
            Rolling_Revenue = f"{format_currency(lower_rev)} - {format_currency(upper_rev)}"
        else:
            Rolling_Performance = f"{format_number(upper_perf)} kW - {format_number(lower_perf)} kW"
            Rolling_Revenue = f"{format_currency(upper_rev)} - {format_currency(lower_rev)}"

        # Add the text to the canvas
        textobject = c.beginText(55, 600)
        textobject.setFont("Libre Franklin", 11)
        textobject.setFillColor('black')  
        textobject.textLines(f'''
            Assets Dispatched:
            Events in Report: 
            Average Performance Based Selected Events:
            Average Revenue Based on Selected Events:
            ''')
        c.drawText(textobject)

        # Add the text to the canvas
        textobject = c.beginText(330, 600)
        textobject.setFont("Libre Franklin", 11)
        textobject.setFillColor('black')  
        textobject.textLines(f'''
            {len(comp_card)}
            {len(selected_events)}
            {Rolling_Performance}
            {Rolling_Revenue}
            ''')
        c.drawText(textobject)

        if len(comp_card['Customer Asset']) > 1:
            # Add the text to the canvas
            textobject = c.beginText(35, 530)
            textobject.setFont("Public Sans-Bold", 11)
            textobject.setFillColor('black')  
            textobject.textLines(f'''Average Performance by Asset for Selected Events''')
            c.drawText(textobject)

            plt.rcParams['font.family'] = 'Public Sans'
            plt.rcParams['font.style'] = 'normal'
            plt.rcParams['font.size'] = 12
            plt.rcParams['font.weight'] = 'normal'

            fig, ax = plt.subplots(figsize=(12, 4.5))

            # Calculate average performance for each asset across selected events
            avg_performances = comp_card.groupby('Customer Asset')[event_columns].mean().mean(axis=1)

            # Calculate bar positions
            x = np.arange(len(avg_performances))
            width = 0.8  # Width of the bars

            # Plot the bar chart with green color
            bars = ax.bar(x, avg_performances, width, color='green')

            max_performance = avg_performances.max()
            if max_performance > 0.0001:  # Check if max performance is greater than a small threshold
                ax.set_ylim(0, max(max_performance * 1.1, 0.1))  # Add 10% padding at the top, with a minimum of 0.1
            else:
                ax.set_ylim(0, 1)  # Set y-axis from 0 to 1 when max performance is 0 or very small

            # Add performance values above the bars
            for bar, performance in zip(bars, avg_performances):
                value = bar.get_height()
                height = max(bar.get_height(), 0)
                error = height * 0.1
                ax.errorbar(bar.get_x() + bar.get_width() / 2, height, yerr=error, fmt='none', color=color_palette[0], capsize=0, elinewidth=3)
                '''color = 'green' if height > 0 else 'red'
                ax.text(bar.get_x() + bar.get_width() / 2, height, f'{round(value)}', 
                        ha='center', va='bottom', color=color, fontsize=12, fontweight='bold')'''

            # Add labels and formatting
            ax.set_xlabel('Assets')
            ax.set_ylabel('Average kW')

            # Improved format_label function
            def format_label(label, max_chars=10):
                if len(label) > max_chars:
                    return label[:max_chars-3] + '...'
                return label
            
            fig_width_inches = fig.get_size_inches()[0]
            char_width_inches = 0.15  # Approximate width of a character
            max_chars = int((fig_width_inches / len(avg_performances)) / char_width_inches)
            max_chars = max(5, min(15, max_chars))  # Ensure max_chars is between 5 and 15

            # Set x-ticks and labels
            ax.set_xticks(x)
            # Create two-level labels
            labels = []
            for asset, account in zip(avg_performances.index, comp_card['Utility Account Number'].unique()):
                asset_label = format_label(asset, max_chars)
                account_label = format_label(account, max_chars)
                labels.append(f"{asset_label}\n{account_label}")

            # Set x-labels with two levels
            ax.set_xticklabels(labels, rotation=0, ha='center', va='top', fontsize=10)

            # Adjust bottom margin to accommodate two-line labels
            plt.subplots_adjust(bottom=0.2)

            # Remove borders
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)

            # Add horizontal grid lines
            ax.yaxis.grid(True, linestyle='--', which='major', color='grey', alpha=.25)

            # Save the plot
            plt.savefig("Data/Temp Assets/performance_by_asset.png", dpi=290, bbox_inches='tight')

            # Add the centered image to the canvas
            img = Image.open("Data/Temp Assets/performance_by_asset.png")
            img_width, img_height = img.size
            mag = .17

            c.drawImage("Data/Temp Assets/performance_by_asset.png", 35, 330, width=img_width*mag, height=img_height*mag)

            # Close the figure
            plt.close(fig)

            chart_space = 0

        else:
            chart_space = 230

        # Add the text to the canvas
        textobject = c.beginText(35, 300 + chart_space)
        textobject.setFont("Libre Franklin-Medium", 11)
        textobject.setFillColor(color_palette[0]) 
        textobject.textLines(f'''
            Individual preliminary performance reports are available below for each asset.
        ''')
        c.drawText(textobject)

        # Add the text to the canvas
        textobject = c.beginText(35, 283+ chart_space)
        textobject.setFont("Libre Franklin", 11)
        textobject.setFillColor('black') 
        textobject.textLines(f'''
            For more detailed load data please visit the

            Best,

            The CPower Team
        ''')
        c.drawText(textobject)

        hyperlink = "CPower Portal (portal.cpowercorp.com)."
        textobject = c.beginText(253, 283+ chart_space)  # Adjust x position according to the length of the preceding text
        textobject.setFont("Libre Franklin-Medium", 11)
        textobject.setFillColor(color_palette[8])  # Set color to blue
        textobject.textLine(hyperlink)
        c.drawText(textobject)

        link = "https://nam02.safelinks.protection.outlook.com/?url=https%3A%2F%2Fportal.cpowercorp.com%2FAuthentication&data=05%7C01%7CTyler.Sims%40CPowerEnergyManagement.com%7C40328f4a152944bd80a108db7f286350%7C36d907c0d528449a85027ed7ff4b32dc%7C1%7C0%7C638243583688071448%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C3000%7C%7C%7C&sdata=jpbysi1AUhcZLUFdofrrkLILDy4e8hbmAwxTidJ3B3c%3D&reserved=0"
        c.linkURL(link, (253, 293 + chart_space, 253 + 200, 263 + 15 + chart_space), relative=1)

        textobject = c.beginText(35, 40)
        textobject.setFont("Libre Franklin", 10)
        textobject.setFillColor(color_palette[7])  
        textobject.textLines(f'''
            *Preliminary reports are for discussion purposes only. Final results are pending until settlement with utility.
            *Performance and revenue projections use a 10% margin of error.
            ''')
        c.drawText(textobject)

        if self.report_format == "Full":
            page_number = 1 + self.toclength

            c.saveState()
            c.setFont("Helvetica", 8)
            c.drawRightString(570, 25, f"Page {page_number}")
            c.restoreState()

        c.showPage()
        c.save()
        buffer.seek(0)

        return PdfReader(buffer).pages[0]

    def create_asset_report(self, program, page_number=None):
        # Load the data from the Excel file
        try:
            event_date = pd.to_datetime(self.event).date()

            # Determine the file path based on the input parameters
            file_path = fr"Settlements\CS\{self.season}\CPOWER Results\{self.event_type} Event {event_date}\{self.asset_id} {self.event_type} {event_date} {self.customer} Performance Calculations.xlsx"

            df = pd.read_excel(file_path, sheet_name='Calculations')
            # Drop the first row and reset the index
            df = df.iloc[1:].reset_index(drop=True)
            # Set the first column as the index
            df.set_index(df.columns[0], inplace=True)

            df.index = pd.to_datetime([event_date.strftime('%Y-%m-%d') + ' ' + str(time) for time in df.index])

            # Now extract the event day load
            event_day_load = df[f'Event {event_date}']
            baseline = df['Baseline']
            adjusted_baseline = df['Adjusted Baseline']

            # Use .loc for explicit indexing:
            self.dispatches.loc[:, 'Event Start'] = pd.to_datetime(
                self.dispatches['Event Date'].astype(str) + ' ' + 
                self.dispatches['Start Time'].astype(str)
            )

            # Filter dispatches for the event date and start time
            event_dispatch = self.dispatches[self.dispatches['Event Date'] == self.event]
            
            if event_dispatch.empty:
                raise ValueError(f"No dispatch found for event date: {event_date}")
            
            # Combine date and time
            self.event_start = event_dispatch['Event Start'].iloc[0]
            self.event_end = pd.to_datetime(f"{self.event_start.date()} {event_dispatch['End Time'].iloc[0]}")

            # Calculate averages over the event period
            event_period = (event_day_load.index > self.event_start) & (event_day_load.index <= self.event_end)
            avg_baseline = baseline[event_period].mean()
            avg_adjusted_baseline = adjusted_baseline[event_period].mean()
            avg_demand = event_day_load[event_period].mean()
            performance = avg_adjusted_baseline - avg_demand
            performance_percentage = (performance / avg_adjusted_baseline) * 100 if avg_adjusted_baseline != 0 else 0

            baseline.index = pd.to_datetime(baseline.index.astype(str))
            adjusted_baseline.index = pd.to_datetime(adjusted_baseline.index.astype(str))
            event_day_load.index = pd.to_datetime(event_day_load.index.astype(str))

        except Exception as e:
            event_date = None
            event_day_load = None
            baseline = None
            adjusted_baseline = None
            # Create a temporary column combining Event Date and Start Time
            self.dispatches['Event Start'] = pd.to_datetime(self.dispatches['Event Date'].astype(str) + ' ' + self.dispatches['Start Time'].astype(str))

            # Filter dispatches for the event date and start time
            event_dispatch = self.dispatches[self.dispatches['Event Date'] == self.event]
            
            if event_dispatch.empty:
                raise ValueError(f"No dispatch found for event date: {event_date}")

            self.event_start = event_dispatch['Event Start'].iloc[0]
            self.event_end = pd.to_datetime(f"{self.event_start.date()} {event_dispatch['End Time'].iloc[0]}")
            avg_baseline = np.nan
            avg_adjusted_baseline = np.nan
            avg_demand = np.nan
            performance = np.nan
            performance_percentage = np.nan
        
        if program == 'Efficiency Maine-Connected Solutions-Targeted Dispatch-Demand Response Initiative':
            sub_program = "Demand Response Initiative Program"
        else:
            sub_program = "Connected Solutions " +  self.event_type

        # Determine the file path for the PDF report
        full_path = fr"Settlements\CS\{self.season}\Reports\Event\{self.event_type}\{self.event.strftime('%Y.%m.%d')}\Individual"

        # Create the directory if it does not exist
        if not os.path.exists(full_path):
            os.makedirs(full_path)

        if page_number is not None:
            # Create a BytesIO object to hold the PDF content
            buffer = BytesIO()
            c = canvas.Canvas(buffer, pagesize=letter)

        else:
            pdf_filename = os.path.join(full_path, f"{self.customer} {self.asset_id} {sub_program} Event {self.event.strftime('%Y.%m.%d')}.pdf")
            c = canvas.Canvas(pdf_filename, pagesize = letter)
        
        # Use this x-coordinate when drawing the image.
        c.drawImage("Data\Image Assets\CPower_Logo.jpg", 470, 740, width=100, height=20)  # adjust the values as necessary
        
        # Add the text to the canvas
        textobject = c.beginText(35, 20)
        textobject.setFont("Libre Franklin", 10)
        textobject.setFillColor(color_palette[7])  
        textobject.textLines(f'''
            *Preliminary report is for discussion purposes only. Our data, your data, and the utilities data may vary.
            ''')
        
        c.drawText(textobject)
        
        # Add the text to the canvas
        textobject = c.beginText(35, 750)
        textobject.setFont("Public Sans-Bold", 12)
        textobject.setFillColor(color_palette[5]) 
        textobject.textLines(f'''
            Preliminary Report 
            Connected Solutions {self.event_type} Event - {self.event.strftime('%m/%d/%Y')}
            ''')

        c.drawText(textobject)
        
        # Add the text to the canvas
        textobject = c.beginText(35, 710)
        textobject.setFont("Public Sans-Bold", 11)
        textobject.setFillColor('black')  
        textobject.textLines(f'''Asset Information''')
        
        c.drawText(textobject)
        
        # Add the text to the canvas
        textobject = c.beginText(55, 690)
        textobject.setFont("Libre Franklin", 11)
        textobject.setFillColor('black') 
        textobject.textLines(f'''
            Name:            
            Owner:                 
            Address:               
            Utility Account Number:
            ''')

        # Add Asset ID if different from Utility Account Number
        if self.asset_id != self.utility_account_number:
            textobject.textLine("Program ID:")
        
        c.drawText(textobject)   
        
        # Add the text to the canvas
        textobject = c.beginText(220, 690)
        textobject.setFont("Libre Franklin", 11)
        textobject.setFillColor('black') 
        textobject.textLines(f'''
            {self.customer_asset}
            {self.customer}
            {self.address}
            {self.utility_account_number}
            ''')

        # Add Asset ID if different from Utility Account Number
        if self.asset_id != self.utility_account_number:
            textobject.textLine(self.asset_id)
        
        c.drawText(textobject) 
        
        # Adjust y-coordinate based on whether Asset ID was added
        y_coord = 630 if self.asset_id == self.utility_account_number else 610
                
        # Add the text to the canvas
        textobject = c.beginText(35, y_coord)
        textobject.setFont("Public Sans-Bold", 11)
        textobject.setFillColor('black')  
        textobject.textLines(f'''Event Day Load''')
        
        c.drawText(textobject)
        
        # Add the text to the canvas
        textobject = c.beginText(55, y_coord - 20)
        textobject.setFont("Libre Franklin", 11)
        textobject.setFillColor('black')  
        textobject.textLines(f'''
            Average Baseline:
            Average Adjusted Baseline:  
            Average Demand: 
            ''')
        
        c.drawText(textobject)

        # Add the text to the canvas
        textobject = c.beginText(43, y_coord - 54)
        textobject.setFont("Helvetica", 10)
        textobject.setFillColor('black')  
        textobject.textLines(f'''
            __________________________________________________________________
            ''')
        
        c.drawText(textobject)
        
        # Add the text to the canvas
        textobject = c.beginText(55, y_coord - 68)
        textobject.setFont("Libre Franklin-Medium", 11)
        textobject.setFillColor(color_palette[9] if performance > 0 else 'red')
        textobject.textLines(f'''
            Performance:    
            ''')
        
        c.drawText(textobject)
        
        textobject = c.beginText(220, y_coord - 20)
        textobject.setFont("Libre Franklin", 11)
        textobject.setFillColor('black')  
        textobject.textLines(f'''
            {avg_baseline:.2f} kW
            {avg_adjusted_baseline:.2f} kW  
            {avg_demand:.2f} kW 
            ''')
        c.drawText(textobject)

        textobject = c.beginText(220, y_coord - 68)
        textobject.setFont("Libre Franklin-Medium", 11)
        textobject.setFillColor(color_palette[9] if performance > 0 else 'red')  
        textobject.textLines(f'''{performance:.2f} kW''') #[{performance_percentage:.2f}%]  
            
        c.drawText(textobject)
                
        plt.rcParams['font.family'] = 'Public Sans'
        plt.rcParams['font.style'] = 'normal'
        plt.rcParams['font.size'] = 10
        plt.rcParams['font.weight'] = 'normal'

        if event_day_load is not None:
            # Generate the plot as an image
            fig, ax = plt.subplots(figsize=(12, 3))  # Create a new figure with a single Axes

            # Set x-axis limits to show 3 hours before the event and 1 hour after
            start_time = self.event_start - pd.Timedelta(hours=3)
            end_time = self.event_end + pd.Timedelta(hours=1)
            ax.set_xlim(start_time, end_time)

            ax.plot(baseline.index, baseline.values, color=color_palette[0], label='Baseline (kW)', linewidth=0.9, linestyle='--') 
            
            # Plot adjusted baseline with 5% error lines
            ax.plot(adjusted_baseline.index, adjusted_baseline.values, color=color_palette[0], label='Adjusted Baseline (kW)', linewidth=1.2)
            ax.fill_between(adjusted_baseline.index, 
                            adjusted_baseline.values * 0.9, 
                            adjusted_baseline.values * 1.1, 
                            color=color_palette[0], alpha=0.2)
            
            # Plot event day load with 5% error lines
            ax.plot(event_day_load.index, event_day_load.values, color=color_palette[5], label='Demand (kW)', linewidth=0.9)
            ax.fill_between(event_day_load.index, 
                            event_day_load.values * 0.9, 
                            event_day_load.values * 1.1, 
                            color=color_palette[5], alpha=0.2)

            filtered_event_day_load = event_day_load.loc[event_day_load.index.isin(adjusted_baseline.index)]

            # Highlight the specified time span with a yellow color
            ax.axvspan(self.event_start, self.event_end, color=color_palette[5], alpha=0.3, label='Event')

            # Only shade performance during the event window
            event_mask = (filtered_event_day_load.index >= self.event_start) & (filtered_event_day_load.index <= self.event_end)
            ax.fill_between(filtered_event_day_load.index[event_mask], 
                            filtered_event_day_load.values[event_mask], 
                            adjusted_baseline.values[event_mask], 
                            where=[f <= a for f, a in zip(filtered_event_day_load.values[event_mask], adjusted_baseline.values[event_mask])], 
                            color=color_palette[9], label='Performance', interpolate=True, alpha=0.3)

            ax.fill_between(filtered_event_day_load.index[event_mask], 
                            adjusted_baseline.values[event_mask], 
                            filtered_event_day_load.values[event_mask], 
                            where=[f > a for f, a in zip(filtered_event_day_load.values[event_mask], adjusted_baseline.values[event_mask])], 
                            color='red', label='No Performance', interpolate=True, alpha=0.3)

            # Format x-axis to show times
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
            ax.xaxis.set_major_locator(mdates.HourLocator(interval=1))

            plt.xlabel('Time')
            plt.ylabel('kW Demand')

            plt.legend()

            # Add horizontal grid lines
            ax.yaxis.grid(True, linestyle='--', which='major', color='grey', alpha=.25)

            # Remove borders
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)

            plt.xlabel('Time')
            plt.ylabel('kW Demand')
            
            plt.legend()
            
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
            
            # get the current y-axis limits
            ymin, ymax = ax.get_ylim()
            
            # set the y-axis lower limit to 0 or maintain ymin if it's less than 0
            # and upper limit as per your existing maximum limit
            ax.set_ylim(min(0, ymin), ymax)
            
            # Add horizontal grid lines
            ax.yaxis.grid(True, linestyle='--', which='major', color='grey', alpha=.25)

            # Remove borders
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            
            # Save the plot as an image
            plt.savefig("Data/Temp Assets/plot.png", dpi=300, bbox_inches='tight')
            
            # Close the figure
            plt.close(fig)

            # Get the dimensions of the image
            img = Image.open("Data/Temp Assets/plot.png")
            img_width, img_height = img.size
            mag = .17

            # Calculate the x position to center the image
            x = (595.27 - img_width*mag) / 2  # Replace 595.27 with your PDF width if it's not A4

            # Add the centered image to the canvas
            c.drawImage("Data/Temp Assets/plot.png", x, 370, width=img_width*mag, height=img_height*mag)

        if avg_demand is np.nan:
            # Define the size and position of the block
            block_width = 530
            block_height = 150

            x = (595.27 - block_width) / 2

            # Draw the block
            c.setFillColor('white')
            c.rect(x, 370, block_width, block_height, fill=True)
            
            message = "Performance not available."
            text_width = c.stringWidth(message, "Helvetica", 12)

            # Calculate the position to center the text
            text_x = x + (block_width - text_width) / 2
            text_y = 370 + block_height / 2

            c.setFillColor('black')
            c.setFont("Helvetica", 12)
            c.drawString(text_x, text_y, message)
            
        # Add the first part of the text
        textobject = c.beginText(30, 340)
        textobject.setFont("Libre Franklin", 11)
        textobject.setFillColor('black')  
        textobject.textLines('''
            For more information please visit the '''
        )
        c.drawText(textobject)  

        # Add the hyperlink text
        hyperlink = "CPower Portal (portal.cpowercorp.com)."
        textobject = c.beginText(218, 340)  # Adjust x position according to the length of the preceding text
        textobject.setFont("Libre Franklin-Medium", 11)
        textobject.setFillColor(color_palette[8])  # Set color to blue
        textobject.textLine(hyperlink)
        c.drawText(textobject)

        # Add a hyperlink to the text
        link = "https://nam02.safelinks.protection.outlook.com/?url=https%3A%2F%2Fportal.cpowercorp.com%2FAuthentication&data=05%7C01%7CTyler.Sims%40CPowerEnergyManagement.com%7C40328f4a152944bd80a108db7f286350%7C36d907c0d528449a85027ed7ff4b32dc%7C1%7C0%7C638243583688071448%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C3000%7C%7C%7C&sdata=jpbysi1AUhcZLUFdofrrkLILDy4e8hbmAwxTidJ3B3c%3D&reserved=0"
        c.linkURL(link, (200, 340, 200 + 200, 340 + 15), relative=1)

        # Add the second part of the text
        textobject = c.beginText(30, 320)  # Adjust x position according to the length of the preceding text
        textobject.setFont("Libre Franklin-Medium", 11)
        textobject.setFillColor('black')
        textobject.textLines('''
            If you have questions regarding your report, contact us at northeast@cpowerenergymanagement.com.
            ''')
        c.drawText(textobject)

        # Add page number if provided
        if page_number is not None:
            c.saveState()
            c.setFont("Helvetica", 8)
            c.drawRightString(570, 25, f"Page {page_number}")
            c.restoreState()

            # Save the canvas to the buffer
            c.save()

            # Move buffer position to the beginning
            buffer.seek(0)

            # Create a PDF page object from the buffer
            pdf_page = PdfReader(buffer).pages[0]

            return pdf_page

        # Save the canvas as a pdf file
        c.save()

    def get_output_path(self, company, report_type, month=None, report_format="Full"):
        format_suffix = "One-Pager" if report_format == "OnePager" else "Full"
        if report_type == "Monthly":
            return f"Settlements/CS/{self.season}/Reports/Monthly/{month}/{self.event_type}/Company/{company} {month} {self.program} Report {format_suffix}.pdf"
        elif report_type == "Custom":
            return f"Settlements/CS/{self.season}/Reports/Custom/{self.event_type}/Company/{company} {month} {self.program} Report {format_suffix}.pdf"
        else:  # Event
            return f"Settlements/CS/{self.season}/Reports/Event/{self.event_type}/{self.event.strftime('%Y.%m.%d')}/Company/{company} {self.event.strftime('%Y.%m.%d')} {self.program} Report {format_suffix}.pdf"
