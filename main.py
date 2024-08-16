from modules import *

class Watts(QMainWindow):
    def __init__(self):
        super(Watts, self).__init__()
        self.app = app
        self.setWindowTitle("Watts")
        self.setMinimumWidth(1000)  # Set the minimum width of the app window
        self.setGeometry(100, 100, 1000, 700)

        self.threadpool = QThreadPool()

        current_year = datetime.now().year

        self.get_dispatches(year=current_year)

        self.create_widgets()
        self.load_settle_data()
        self.set_columns_for_settle_treeview()
        self.apply_daily_connected_solution_filter()
        self.event_summary_settle_cs.click()

        self.toggle_treeview_button.setChecked(True)

    def create_widgets(self):
        central_widget = QtWidgets.QWidget(self)
        self.setCentralWidget(central_widget)
        self.main_layout = QtWidgets.QHBoxLayout(central_widget)

        # region Left Layout

        # region Work Stations
        self.work_window_layout = QtWidgets.QVBoxLayout()

        # Create a widget to hold the main content
        self.main_content_widget = QWidget()
        self.main_content_layout = QVBoxLayout(self.main_content_widget)

        # Create button layout for work window
        work_station_button_layout = QtWidgets.QHBoxLayout()
        self.settle_button = QtWidgets.QPushButton("Settle", central_widget)
        work_station_button_layout.addWidget(self.settle_button)
        self.work_window_layout.addLayout(work_station_button_layout)
        
        # Add horizontal line
        self.work_window_layout.addWidget(QtWidgets.QLabel("............................................................................................"), 0, QtCore.Qt.AlignCenter)
        # endregion

        # region Program Filters Buttons
        program_filter_layout = QtWidgets.QVBoxLayout()
        program_filter_row = QtWidgets.QHBoxLayout()
        self.connected_solutions_button = QtWidgets.QPushButton("Connected Solutions", central_widget)
        program_filter_row.addWidget(self.connected_solutions_button)
        program_filter_layout.addLayout(program_filter_row)
        self.work_window_layout.addLayout(program_filter_layout)

        self.program_type_checkbox_layout = QtWidgets.QHBoxLayout()
        self.daily_button = QtWidgets.QPushButton("Daily", central_widget)
        self.daily_button.setCheckable(True)
        self.daily_button.setChecked(True)
        self.daily_button.clicked.connect(self.apply_daily_connected_solution_filter)
        self.program_type_checkbox_layout.addWidget(self.daily_button)

        self.targeted_button = QtWidgets.QPushButton("Targeted", central_widget)
        self.targeted_button.setCheckable(True)
        self.targeted_button.setChecked(False)
        self.targeted_button.clicked.connect(self.apply_targeted_connected_solution_filter)
        self.program_type_checkbox_layout.addWidget(self.targeted_button)

        self.program_type_checkbox_widget = QtWidgets.QWidget(central_widget)
        self.program_type_checkbox_widget.setLayout(self.program_type_checkbox_layout)
        self.work_window_layout.addWidget(self.program_type_checkbox_widget)

        # Create utility checkboxes (initially hidden)
        self.utility_checkbox_layout = QtWidgets.QVBoxLayout()

        utility_row1 = QtWidgets.QHBoxLayout()
        self.cape_light_checkbox = QCheckBox("Cape Light", central_widget)
        self.cape_light_checkbox.setChecked(True)  # Set initial state to checked
        self.cape_light_checkbox.stateChanged.connect(self.apply_connected_solution_filters)
        utility_row1.addWidget(self.cape_light_checkbox)
        self.efficiency_maine_checkbox = QCheckBox("Efficiency Maine", central_widget)
        self.efficiency_maine_checkbox.setChecked(True)
        self.efficiency_maine_checkbox.stateChanged.connect(self.apply_connected_solution_filters)
        utility_row1.addWidget(self.efficiency_maine_checkbox)
        self.eversource_checkbox = QCheckBox("Eversource", central_widget)
        self.eversource_checkbox.setChecked(True)
        self.eversource_checkbox.stateChanged.connect(self.apply_connected_solution_filters)
        utility_row1.addWidget(self.eversource_checkbox)

        self.utility_checkbox_layout.addLayout(utility_row1)

        utility_row2 = QtWidgets.QHBoxLayout()
        self.liberty_checkbox = QCheckBox("Liberty", central_widget)
        self.liberty_checkbox.setChecked(True)
        self.liberty_checkbox.stateChanged.connect(self.apply_connected_solution_filters)
        utility_row2.addWidget(self.liberty_checkbox)
        self.ngrid_checkbox = QCheckBox("NGRID", central_widget)
        self.ngrid_checkbox.setChecked(True)
        self.ngrid_checkbox.stateChanged.connect(self.apply_connected_solution_filters)
        utility_row2.addWidget(self.ngrid_checkbox)
        self.ri_energy_checkbox = QCheckBox("Rhode Island", central_widget)
        self.ri_energy_checkbox.setChecked(True)
        self.ri_energy_checkbox.stateChanged.connect(self.apply_connected_solution_filters)
        utility_row2.addWidget(self.ri_energy_checkbox)
        self.unitil_checkbox = QCheckBox("Unitil", central_widget)
        self.unitil_checkbox.setChecked(True)
        self.unitil_checkbox.stateChanged.connect(self.apply_connected_solution_filters)
        utility_row2.addWidget(self.unitil_checkbox)
        self.utility_checkbox_layout.addLayout(utility_row2)

        self.utility_checkbox_widget = QtWidgets.QWidget(central_widget)
        self.utility_checkbox_widget.setLayout(self.utility_checkbox_layout)
        self.work_window_layout.addWidget(self.utility_checkbox_widget)

        # Add horizontal line
        self.work_window_layout.addWidget(QtWidgets.QLabel("............................................................................................"), 0, QtCore.Qt.AlignCenter)
        # endregion

        # region GUI Layouts for Settle, and Pay widgets
        self.widget_layout = QtWidgets.QVBoxLayout()
        self.title_layout = QtWidgets.QVBoxLayout()

        # region Settle Connected Solutions
        self.settlement_period_label_settle_cs = None

        settle_cs_row1_layout = QtWidgets.QHBoxLayout()
        self.summer_label_settle_cs = QtWidgets.QLabel("Summer", central_widget)
        self.summer_year_combo_settle_cs = QtWidgets.QComboBox(central_widget)
        summers = [str(year) for year in range(2019, 2031)]
        self.summer_year_combo_settle_cs.addItems(summers)
        self.select_summer_settle_cs_button = QtWidgets.QPushButton("Select", central_widget)
        self.select_summer_settle_cs_button.clicked.connect(self.handle_select_summer_settle_cs)
        settle_cs_row1_layout.addWidget(self.summer_label_settle_cs)
        settle_cs_row1_layout.addWidget(self.summer_year_combo_settle_cs)
        settle_cs_row1_layout.addWidget(self.select_summer_settle_cs_button)
        self.widget_layout.addLayout(settle_cs_row1_layout)

        # Set default month and year to the month before the app is opened
        current_date = datetime.now()
        default_year_cs = current_date.year

        self.summer_year_combo_settle_cs.setCurrentText(str(default_year_cs))

        # Title for Event Performance Calculation
        self.event_performance_title_settle_cs = QtWidgets.QLabel("Event Performance Calculation", central_widget)
        self.widget_layout.addWidget(self.event_performance_title_settle_cs)

        settle_cs_row4_layout = QtWidgets.QHBoxLayout()
        self.download_settle_cs_button = QtWidgets.QPushButton("Pull Data", central_widget)
        settle_cs_row4_layout.addWidget(self.download_settle_cs_button)
        self.widget_layout.addLayout(settle_cs_row4_layout)

        self.calculate_settle_cs_button = QtWidgets.QPushButton("Calculate", central_widget)
        self.calculate_settle_cs_button.clicked.connect(self.handle_calculate_settle_cs)
        settle_cs_row4_layout.addWidget(self.calculate_settle_cs_button)
        self.widget_layout.addLayout(settle_cs_row4_layout)

        # Title for Report Creation
        self.report_generation_title_settle_cs = QtWidgets.QLabel("Report Generation", central_widget)
        self.widget_layout.addWidget(self.report_generation_title_settle_cs)

        settle_cs_row2_layout = QtWidgets.QHBoxLayout()
        self.calculate_event_performances_settle_cs_checkbox = QtWidgets.QCheckBox("Event", central_widget)
        settle_cs_row2_layout.addWidget(self.calculate_event_performances_settle_cs_checkbox)

        self.calculate_month_performances_settle_cs_checkbox = QtWidgets.QCheckBox("Month", central_widget)
        settle_cs_row2_layout.addWidget(self.calculate_month_performances_settle_cs_checkbox)

        self.calculate_custom_performances_settle_cs_checkbox = QtWidgets.QCheckBox("Custom", central_widget)
        settle_cs_row2_layout.addWidget(self.calculate_custom_performances_settle_cs_checkbox)
        self.widget_layout.addLayout(settle_cs_row2_layout)

        settle_cs_row3_layout = QtWidgets.QHBoxLayout()
        self.event_selection_label_settle_cs = QtWidgets.QLabel("Event Selection:", central_widget)
        self.event_selection_combo_settle_cs = QtWidgets.QComboBox(central_widget)
        individual_event_options = [f"Event {i}" for i in range(1, 31)]
        self.event_selection_combo_settle_cs.addItems(individual_event_options)
        settle_cs_row3_layout.addWidget(self.event_selection_label_settle_cs)
        settle_cs_row3_layout.addWidget(self.event_selection_combo_settle_cs)
        self.widget_layout.addLayout(settle_cs_row3_layout)
        self.event_selection_label_settle_cs.setVisible(False)
        self.event_selection_combo_settle_cs.setVisible(False)

        settle_cs_row5_layout = QtWidgets.QHBoxLayout()
        self.month_report_selection_label_settle_cs = QtWidgets.QLabel("Month Selection:", central_widget)
        self.month_report_selection_combo_settle_cs = QtWidgets.QComboBox(central_widget)
        report_options = ["June", "July", "August", "September"]
        self.month_report_selection_combo_settle_cs.addItems(report_options)
        settle_cs_row5_layout.addWidget(self.month_report_selection_label_settle_cs)
        settle_cs_row5_layout.addWidget(self.month_report_selection_combo_settle_cs)
        self.widget_layout.addLayout(settle_cs_row5_layout)
        self.month_report_selection_label_settle_cs.setVisible(False)
        self.month_report_selection_combo_settle_cs.setVisible(False)

        settle_cs_row3_layout = QtWidgets.QVBoxLayout()  # Changed to QVBoxLayout for better organization
        self.custom_report_selection_label_settle_cs = QtWidgets.QLabel("Custom Event Selection:", central_widget)
        self.custom_report_selection_list_settle_cs = QtWidgets.QListWidget(central_widget)
        self.custom_report_selection_list_settle_cs.setSelectionMode(QtWidgets.QAbstractItemView.MultiSelection)

        settle_cs_row3_layout.addWidget(self.custom_report_selection_label_settle_cs)
        settle_cs_row3_layout.addWidget(self.custom_report_selection_list_settle_cs)
        self.widget_layout.addLayout(settle_cs_row3_layout)

        self.custom_report_selection_label_settle_cs.setVisible(False)
        self.custom_report_selection_list_settle_cs.setVisible(False)

        settle_cs_row7_layout = QtWidgets.QHBoxLayout()
        self.select_recommended_settle_cs_button = QtWidgets.QPushButton("Select Recommended", central_widget)
        self.select_recommended_settle_cs_button.clicked.connect(self.handle_select_recommended_settle_cs)
        self.generate_report_settle_cs_button = QtWidgets.QPushButton("Generate Asset Report", central_widget)
        self.generate_report_settle_cs_button.clicked.connect(self.handle_generate_report_settle_cs)
        settle_cs_row7_layout.addWidget(self.select_recommended_settle_cs_button)
        settle_cs_row7_layout.addWidget(self.generate_report_settle_cs_button)
        self.widget_layout.addLayout(settle_cs_row7_layout)

        settle_cs_row8_layout = QtWidgets.QHBoxLayout()
        self.generate_customer_report_settle_cs_button = QtWidgets.QPushButton("Generate Customer Report", central_widget)
        self.generate_customer_report_settle_cs_button.clicked.connect(self.handle_generate_customer_report)
        settle_cs_row8_layout.addWidget(self.generate_customer_report_settle_cs_button)
        self.widget_layout.addLayout(settle_cs_row8_layout)

        self.calculate_event_performances_settle_cs_checkbox.stateChanged.connect(
            lambda state: self.update_custom_event_selection_visibility_settle_cs(state, "Event")
        )
        self.calculate_month_performances_settle_cs_checkbox.stateChanged.connect(
            lambda state: self.update_custom_event_selection_visibility_settle_cs(state, "Month")
        )
        self.calculate_custom_performances_settle_cs_checkbox.stateChanged.connect(
            lambda state: self.update_custom_event_selection_visibility_settle_cs(state, "Custom")
        )

        # endregion

        # Create group widgets and set their layouts
        self.gui_widget = QtWidgets.QWidget()
        self.gui_widget.setLayout(self.widget_layout)

        # Add group widgets to work_window_layout
        self.work_window_layout.addLayout(self.title_layout)
        self.work_window_layout.addWidget(self.gui_widget)
        
        # Add a vertical spacer to push the widgets to the top
        vertical_spacer = QtWidgets.QSpacerItem(0, 0, QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Expanding)
        self.work_window_layout.addItem(vertical_spacer)

        # Add the main content widget to the work window layout
        self.work_window_layout.addWidget(self.main_content_widget)

        # Create a label for status messages
        self.status_message_label = QLabel(self)
        self.status_message_label.setStyleSheet("color: #4A4A4A; font-size: 12px; padding: 5px;")
        self.status_message_label.setWordWrap(True)
        self.status_message_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.status_message_label.setVisible(False)  # Initially hidden

        # Add the status message label to the bottom of the work window layout
        self.work_window_layout.addWidget(self.status_message_label)

        # Create a widget to hold the work window layout
        self.work_window_widget = QWidget()
        self.work_window_widget.setLayout(self.work_window_layout)

        # Add the work window widget to your main layout
        self.main_layout.addWidget(self.work_window_widget)

        # endregion
        # endregion

        # region Right Layout 
        right_layout = QVBoxLayout()

        # Create a stacked layout for the main content area
        self.stacked_layout = QStackedLayout()
        stacked_widget = QWidget()
        stacked_widget.setLayout(self.stacked_layout)
        right_layout.addWidget(stacked_widget)

        # region Create the Treeview
        self.treeview = QtWidgets.QTreeWidget()
        self.treeview.setSortingEnabled(True)

        # Create a container widget for the treeview
        self.treeview_container = QWidget()
        self.treeview_layout = QVBoxLayout(self.treeview_container)

        # Add the treeview to the layout
        self.treeview_layout.addWidget(self.treeview)

        # Create filter layout
        filter_layout = QHBoxLayout()

        # Asset ID filter
        self.asset_id_filter = QLineEdit()
        self.asset_id_filter.setPlaceholderText("Filter Asset ID")
        self.asset_id_filter.textChanged.connect(self.apply_filters)
        filter_layout.addWidget(QLabel("Asset ID:"))
        filter_layout.addWidget(self.asset_id_filter)

        # Company filter
        self.company_filter = QLineEdit()
        self.company_filter.setPlaceholderText("Filter Company")
        self.company_filter.textChanged.connect(self.apply_filters)
        filter_layout.addWidget(QLabel("Company:"))
        filter_layout.addWidget(self.company_filter)

        # Facility filter
        self.facility_filter = QLineEdit()
        self.facility_filter.setPlaceholderText("Filter Facility")
        self.facility_filter.textChanged.connect(self.apply_filters)
        filter_layout.addWidget(QLabel("Facility:"))
        filter_layout.addWidget(self.facility_filter)

        # Curtailment Strategy filter
        self.curtailment_strategy_filter = QComboBox()
        self.curtailment_strategy_filter.addItem("All")
        self.curtailment_strategy_filter.currentTextChanged.connect(self.apply_filters)
        filter_layout.addWidget(QLabel("Curtailment Strategy:"))
        filter_layout.addWidget(self.curtailment_strategy_filter)

        # Add the filter layout to the main layout above the treeview
        self.treeview_layout.insertLayout(0, filter_layout)

        # Create controls for Treeview
        treeview_controls_layout = QHBoxLayout()

        self.toggle_all_button = QPushButton(" Toggle ")
        self.toggle_all_button.clicked.connect(self.toggle_all_items)
        treeview_controls_layout.addWidget(self.toggle_all_button)

        # Create a horizontal layout for the labels
        labels_layout = QHBoxLayout()
        labels_layout.setAlignment(Qt.AlignVCenter)

        self.total_label = QLabel("Total { Assets: 0, Customers: 0 }")
        self.selected_label = QLabel("Selected { Assets: 0, Customers: 0 }")
        self.total_label.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Preferred)
        self.selected_label.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Preferred)
        labels_layout.addWidget(self.total_label)
        labels_layout.addWidget(self.selected_label)

        # Add the labels layout to the main controls layout
        treeview_controls_layout.addLayout(labels_layout)

        # Add a stretch to push the save button to the right
        treeview_controls_layout.addStretch(5)

        # Add the controls layout to the main treeview layout
        self.treeview_layout.addLayout(treeview_controls_layout)

        # Add the container widget to the stacked layout
        self.stacked_layout.addWidget(self.treeview_container)
        # endregion

        # region Create Data View Widget
        self.data_viewer_widget = QWidget()
        self.stacked_layout.addWidget(self.data_viewer_widget)

        # Create the data viewer layout
        main_layout = QVBoxLayout(self.data_viewer_widget)
        self.data_viewer_layout = QVBoxLayout()
        main_layout.addLayout(self.data_viewer_layout)

        # region Data View Button Section
        self.data_view_button_layout = QHBoxLayout()
        main_layout.addLayout(self.data_view_button_layout)

        self.data_view_button_group = QButtonGroup(self)
        self.data_view_button_group.setExclusive(True)
        
        # region Settle
        # region Connected Solutions
        self.season=default_year_cs
        self.event_summary_settle_cs = self.create_data_view_button("Event Summary", self.create_heatmap_settle_cs)
        self.event_summary_settle_cs.setObjectName("event_summary_button")
        self.event_summary_settle_cs.setVisible(True)                

        self.geo_analysis_button_settle_cs = self.create_data_view_button("Geo Analysis", self.create_geo_analysis_settle_cs)
        self.geo_analysis_button_settle_cs.setObjectName("geo_analysis_button")
        self.geo_analysis_button_settle_cs.setVisible(True)
        # endregion
        # endregion
        # endregion
        # endregion

        # region Create always-visible bottom controls
        bottom_controls = QWidget()
        bottom_controls_layout = QHBoxLayout(bottom_controls)

        # Create a button group for radio button behavior
        self.view_button_group = QButtonGroup(self)
        self.view_button_group.setExclusive(True)

        self.toggle_treeview_button = QPushButton("Data Table")
        self.toggle_treeview_button.setCheckable(True)
        self.toggle_treeview_button.clicked.connect(self.toggle_treeview)
        self.toggle_treeview_button.setObjectName("toggle_treeview_button")
        self.view_button_group.addButton(self.toggle_treeview_button)
        bottom_controls_layout.addWidget(self.toggle_treeview_button)

        self.toggle_data_viewer_button = QPushButton("Data Visuals")
        self.toggle_data_viewer_button.setCheckable(True)
        self.toggle_data_viewer_button.clicked.connect(self.toggle_data_viewer)
        self.toggle_data_viewer_button.setObjectName("toggle_data_viewer_button")
        self.view_button_group.addButton(self.toggle_data_viewer_button)
        bottom_controls_layout.addWidget(self.toggle_data_viewer_button)

        right_layout.addWidget(bottom_controls)

        right_column = QtWidgets.QWidget()
        right_column.setLayout(right_layout)
        self.main_layout.addWidget(right_column, 6) 
        # endregion
        
        # endregion

    # region Data Viewer Buttons
    def create_data_view_button(self, text, connection_function):
        button = QPushButton(text, self)
        button.setCheckable(True)
        button.clicked.connect(connection_function)
        self.data_view_button_group.addButton(button)
        self.data_view_button_layout.addWidget(button)
        button.setVisible(False)
        button.setProperty("data-view-button", True)
        return button

    # region Connected Solutions
    def create_heatmap_settle_cs(self):
        # Clear current layout
        self.clear_data_viewer()

        # Create main layout
        main_layout = QHBoxLayout()

        # Create left panel for heatmap and event selection
        left_panel = QVBoxLayout()

        # Add title
        title = QLabel("Performance Cross Analysis")
        title.setStyleSheet("font-size: 18px; font-weight: bold; padding: 10px;")
        title.setAlignment(Qt.AlignCenter)
        left_panel.addWidget(title)

        # Event selection
        event_layout = QHBoxLayout()
        event_label = QLabel("Select Event:")
        self.event_combo = QComboBox()
        try:
            events = self.get_event_list()
            self.event_combo.addItems(events)
            self.event_combo.currentTextChanged.connect(self.update_heatmap_settle_cs)
            event_layout.addWidget(event_label)
            event_layout.addWidget(self.event_combo)
        except FileNotFoundError:
            self.receive_message("No events found. The CPOWER Results folder might not exist yet.")
            return

        left_panel.addLayout(event_layout)

        # Create a scroll area for the heatmap
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scroll_area.setStyleSheet("QScrollArea { border: none; }")

        # Create a widget to hold the heatmap
        heatmap_widget = QWidget()
        self.heatmap_layout = QVBoxLayout(heatmap_widget)
        self.heatmap_layout.setContentsMargins(0, 0, 0, 0)
        
        # Set the heatmap widget as the scroll area's widget
        self.scroll_area.setWidget(heatmap_widget)

        # Add the scroll area to the left panel
        left_panel.addWidget(self.scroll_area)

        # Add left panel to main layout
        main_layout.addLayout(left_panel, 1)  # 50% of horizontal space

        # Create right panel for line graph and additional info
        right_panel = QVBoxLayout()

        # Create a widget for the detailed performance analysis
        self.detailed_analysis_widget = QWidget()
        self.detailed_analysis_layout = QVBoxLayout(self.detailed_analysis_widget)
        right_panel.addWidget(self.detailed_analysis_widget)

        # Add right panel to main layout
        main_layout.addLayout(right_panel, 1)  # 50% of horizontal space

        # Set the main layout
        main_widget = QWidget()
        main_widget.setLayout(main_layout)
        self.data_viewer_layout.addWidget(main_widget)

        # Initial heatmap generation
        self.update_heatmap_settle_cs()

    def update_heatmap_settle_cs(self):
        selected_event = self.event_combo.currentText()
        data = self.prepare_heatmap_data(selected_event)

        # Sort data by over/under performance
        data['OverUnder'] = data.sum(axis=1) - data.index.map(self.get_expected_value)
        data = data.sort_values('OverUnder')

        for i in reversed(range(self.heatmap_layout.count())): 
            self.heatmap_layout.itemAt(i).widget().setParent(None)

        scene = QGraphicsScene()
        view = QGraphicsView(scene)
        view.setRenderHint(QPainter.Antialiasing)
        view.setFrameStyle(QFrame.NoFrame)
        view.setStyleSheet("border: none; background: transparent;")

        cell_size = 20
        x_offset = 450  # Increase x_offset to add more space between columns
        y_offset = 60
        row_height = 30

        # Filter out non-datetime columns
        datetime_columns = [col for col in data.columns if not col == 'OverUnder']
        
        event_start = pd.to_datetime(selected_event.split()[-1] + ' ' + datetime_columns[0])
        event_end = pd.to_datetime(selected_event.split()[-1] + ' ' + datetime_columns[-1])

        headers = ["Asset", "Actual/\nForecasted", "Over-Under/\nPercent", f"Performance from {event_start.strftime('%H:%M')} to {event_end.strftime('%H:%M')}"]
        for i, header in enumerate(headers):
            text = WrappedTextItem(header, x_offset/3 if i < 3 else x_offset)
            text.setFont(QFont("Public Sans", 10, QFont.Bold))
            if i == 1:  # Right align 'Actual/Forecasted'
                text.setPos(x_offset / 3 + (x_offset / 3) - text.boundingRect().width(), 0)
            elif i == 2:  # Left align 'Percent/Over-Under'
                text.setPos(2 * x_offset / 3, 0)
            elif i == 3:  # Position 'Performance from...' header
                text.setPos(x_offset, 0)
            else:
                text.setPos(i * (x_offset / 3), 0)
            scene.addItem(text)

        self.highlighted_row = None

        for row, (account, values) in enumerate(data.iterrows()):
            customer_name = str(self.get_customer_name(account))
            customer_asset = str(self.get_customer_asset(account))

            customer_name = (customer_name[:30] + '...') if len(customer_name) > 30 else customer_name
            customer_asset = (customer_asset[:30] + '...') if len(customer_asset) > 30 else customer_asset

            # Add a background rectangle for the row
            row_rect = QGraphicsRectItem(0, y_offset + row * row_height, scene.width(), row_height)
            row_rect.setBrush(QBrush(Qt.transparent))
            row_rect.setPen(QPen(Qt.NoPen))
            row_rect.setData(0, account)
            scene.addItem(row_rect)

            customer_name_text = scene.addText(customer_name)
            font = QFont()
            font.setBold(True)  # Set the font to bold
            customer_name_text.setFont(font)  # Apply the bold font to the QGraphicsTextItem
            customer_name_text.setPos(0, y_offset + row * row_height)

            customer_asset_text = scene.addText(customer_asset)
            customer_asset_text.setPos(0, y_offset + row * row_height + 12)

            expected = self.get_expected_value(account)
            total_performance = sum(filter(lambda x: not pd.isna(x), values[datetime_columns]))
            percent_captured = int((total_performance / expected * 100) if expected != 0 else 0)
            over_under = total_performance - expected

            def format_number(num):
                return f"({abs(num):.0f})" if num < 0 else f"{num:.0f}"

            def format_percent(num):
                return f"({abs(num)}%)" if num < 0 else f"{num}%"

            # Create actual/forecasted text items separately
            actual_text_color = QColor("green") if total_performance >= 0 else QColor("red")
            actual_text = QGraphicsTextItem(format_number(total_performance))
            actual_text.setDefaultTextColor(QColor(actual_text_color))
            font = QFont()
            font.setBold(True)  # Set the font to bold
            actual_text.setFont(font)  # Apply the bold font to the QGraphicsTextItem
            actual_text.setPos(x_offset / 3 + (x_offset / 3) - actual_text.boundingRect().width(), 
                            y_offset + row * row_height)
            scene.addItem(actual_text)

            forecasted_text_color = QColor("green") if total_performance >= 0 else QColor("red")
            forecasted_text = QGraphicsTextItem(format_number(expected))
            forecasted_text.setDefaultTextColor(QColor(forecasted_text_color))
            font = QFont()
            font.setItalic(True)  # Set the font to bold
            forecasted_text.setFont(font)  # Apply the bold font to the QGraphicsTextItem
            forecasted_text.setPos(x_offset / 3 + (x_offset / 3) - forecasted_text.boundingRect().width(), 
                                y_offset + row * row_height + 12)
            scene.addItem(forecasted_text)

            # Create over-under/ percent text items separately
            over_under_color = QColor("green") if over_under >= 0 else QColor("red")
            over_under_text = QGraphicsTextItem(format_number(over_under))
            over_under_text.setDefaultTextColor(over_under_color)
            over_under_text.setOpacity(0.5)  # Make it slightly opaque
            over_under_text.setPos(2 * x_offset / 3, y_offset + row * row_height)
            scene.addItem(over_under_text)

            percent_color = QColor("green") if percent_captured >= 0 else QColor("red")
            percent_text = QGraphicsTextItem(format_percent(percent_captured))
            percent_text.setDefaultTextColor(percent_color)
            percent_text.setOpacity(0.5)  # Make it slightly opaque
            percent_text.setPos(2 * x_offset / 3, y_offset + row * row_height + 12)
            scene.addItem(percent_text)

            for col, value in enumerate(values[datetime_columns]):
                if pd.notna(value):
                    norm_value = value / expected if expected != 0 else 0
                    color = self.get_color(norm_value)
                    tooltip = f"{value:.2f}"
                    
                    y_center = y_offset + row * row_height + row_height / 2 - cell_size / 2
                    
                    item = CircularItem(x_offset + col * cell_size, y_center, 
                                        cell_size * 0.95, cell_size * 0.95, color, tooltip)
                    item.setData(0, account)
                    item.setData(1, col)
                    scene.addItem(item)

        scene.setSceneRect(scene.itemsBoundingRect())
        scene.mouseReleaseEvent = lambda event: self.on_heatmap_click(event, scene, selected_event)

        view.setScene(scene)
        self.heatmap_layout.addWidget(view)

        # Add sorting information
        sort_info = QLabel("Assets sorted by over/under performance (least to greatest)")
        sort_info.setStyleSheet("font-style: italic; padding: 5px;")
        self.heatmap_layout.addWidget(sort_info)

    def prepare_heatmap_data(self, selected_event):
        # Construct the path to the selected event folder
        results_folder = f"Settlements/CS/{self.season}/CPOWER Results/{selected_event}"
        
        if not os.path.exists(results_folder):
            self.receive_message(f"Folder not found for event: {selected_event}")
            return pd.DataFrame()

        # Initialize the DataFrame
        df_heatmap = pd.DataFrame()

        # Get visible items from treeview
        visible_items = self.get_visible_items_from_treeview()

        for excel_file in os.listdir(results_folder):
            if excel_file.endswith('.xlsx'):
                file_path = os.path.join(results_folder, excel_file)
                try:
                    # Extract asset ID from the filename
                    asset_id = excel_file.split()[0]
                    
                    # Skip if asset is not visible in treeview
                    if asset_id not in visible_items:
                        continue

                    # Load the Excel file
                    df = pd.read_excel(file_path, sheet_name='Summary', header=None, skiprows=10)

                    # Process the data
                    time_col = df.iloc[:, 0]

                    # Check for valid (non-NA) rows
                    valid_rows = time_col.notna()

                    # Filter out rows that contain non-date text
                    def is_valid_date(value):
                        try:
                            pd.to_datetime(value, format='%H:%M:%S', errors='raise')
                            return True
                        except (ValueError, TypeError):
                            return False

                    valid_rows = valid_rows & time_col.apply(is_valid_date)

                    # Filter the time column to only include valid rows
                    valid_times = time_col[valid_rows]

                    # Convert the valid times to datetime
                    times = pd.to_datetime(valid_times, format='%H:%M:%S', errors='raise')

                    gross_performance = df.iloc[valid_rows.index[valid_rows], 5]

                    data = pd.Series(gross_performance.values, index=times)
                    
                    # Add to the heatmap DataFrame
                    df_heatmap[asset_id] = data

                except Exception as e:
                    self.receive_message(f"Error processing file {excel_file}: {str(e)}")

        # Transpose the DataFrame so assets are rows and time intervals are columns
        df_heatmap = df_heatmap.T

        # After processing, ensure column labels are in HH:MM format
        df_heatmap.columns = pd.to_datetime(df_heatmap.columns).strftime('%H:%M')
        return df_heatmap
    
    def get_visible_items_from_treeview(self):
        visible_items = []
        for i in range(self.treeview.topLevelItemCount()):
            item = self.treeview.topLevelItem(i)
            if not item.isHidden():
                asset_id = item.text(3)  # Assuming Asset ID is in column 3
                visible_items.append(asset_id)
        return visible_items

    def on_heatmap_click(self, event, scene, selected_event):
        item = scene.itemAt(event.scenePos(), QTransform())
        if isinstance(item, (CircularItem, QGraphicsRectItem)):
            account = item.data(0)
            
            # Remove previous highlight
            if self.highlighted_row is not None:
                self.highlighted_row.setBrush(QBrush(Qt.transparent))
            
            # Highlight the selected row
            row_items = [i for i in scene.items() if isinstance(i, QGraphicsRectItem) and i.data(0) == account]
            if row_items:
                row_items[0].setBrush(QBrush(QColor(200, 200, 200, 100)))  # Light gray, semi-transparent
                self.highlighted_row = row_items[0]
            
            self.show_detailed_performance_analysis(account, selected_event)

    def show_detailed_performance_analysis(self, account, selected_event):
        # Clear previous analysis
        for i in reversed(range(self.detailed_analysis_layout.count())): 
            self.detailed_analysis_layout.itemAt(i).widget().setParent(None)

        # Add title
        customer_name = self.get_customer_name(account)
        customer_asset = self.get_customer_asset(account)
        title = QLabel(f"{customer_name} - {customer_asset} Detailed Performance Analysis")
        title.setStyleSheet("font-size: 16px; font-weight: bold; padding: 10px;")
        title.setAlignment(Qt.AlignCenter)
        self.detailed_analysis_layout.addWidget(title)

        # Add performance text
        performance_text = self.get_performance_text(account, selected_event)
        performance_label = QLabel("Actual Performance")
        performance_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #072B60; padding: 5px 0;")
        self.detailed_analysis_layout.addWidget(performance_label)
        
        performance_value = QLabel(performance_text)
        performance_value.setStyleSheet("font-family: monospace; background-color: #f0f0f0; padding: 10px; border-radius: 5px;")
        self.detailed_analysis_layout.addWidget(performance_value)

        # Create line graph and baseline chart
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(4, 1), sharex=True)  # Adjusted height

        ax1 = self.create_line_graph(account, selected_event, ax1)
        ax2 = self.create_baseline_chart(account, selected_event, ax2)

        # Set common x-label closer to the plot
        fig.text(0.5, 0.2, 'Time', ha='center', va='center')  # Moved up slightly

        # Set common y-label
        fig.text(0.04, 0.6, 'kW Demand', ha='center', va='center', rotation='vertical')

        # Add titles to subplots
        ax1.set_title("Event Analysis", fontsize=12, fontweight='bold', pad=5)
        ax2.set_title("Baseline Analysis", fontsize=12, fontweight='bold', pad=5)

        # Create a single legend for both charts inside the plot area
        handles1, labels1 = ax1.get_legend_handles_labels()
        handles2, labels2 = ax2.get_legend_handles_labels()
        handles = handles1 + handles2
        labels = labels1 + labels2
        
        # Remove duplicate labels
        handles_labels = dict(zip(labels, handles))
        handles = list(handles_labels.values())
        labels = list(handles_labels.keys())
        
        # Rename 'Average Baseline' to 'Baseline kW'
        labels = ['Baseline kW' if label == 'Average Baseline' else label for label in labels]
        
        # Place legend inside the first subplot (ax1) in the upper right corner
        fig.legend(handles, labels, loc='lower center', bbox_to_anchor=(0.5, 0.05), ncol=3, frameon=False)  # Adjusted bbox_to_anchor

        # Adjust layout to ensure everything fits
        plt.tight_layout(pad=1.0)
        plt.subplots_adjust(bottom=0.3, left=0.15, right=0.95, hspace=0.4)  # Adjusted bottom, increased top

        # Remove white background and figure border
        fig.patch.set_facecolor('none')
        fig.patch.set_alpha(0)
        ax1.set_facecolor('none')
        ax2.set_facecolor('none')

        # Create a FigureCanvas
        canvas = FigureCanvas(fig)
        
        # Set a minimum size for the canvas
        canvas.setMinimumSize(600, 400)

        # Create a scroll area
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(canvas)
        scroll_area.setStyleSheet("border: none;")  # Remove scroll area border

        # Create a QWidget to hold the scroll area
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.addWidget(scroll_area)

        self.detailed_analysis_layout.addWidget(widget)

    def create_line_graph(self, account, selected_event, ax):
        # Construct the path to the selected event folder
        results_folder = f"Settlements/CS/{self.season}/CPOWER Results/{selected_event}"
        
        if not os.path.exists(results_folder):
            self.receive_message(f"Folder not found for event: {selected_event}")
            return ax

        # Find the Excel file for the selected account
        excel_file = next((f for f in os.listdir(results_folder) if f.startswith(account) and f.endswith('.xlsx')), None)
        
        if excel_file is None:
            self.receive_message(f"No data file found for account: {account}")
            return ax

        file_path = os.path.join(results_folder, excel_file)

        try:
            # Load the Excel file
            df = pd.read_excel(file_path, sheet_name='Calculations')
            
            # Drop the first row and reset the index
            df = df.iloc[1:].reset_index(drop=True)
            
            # Set the first column as the index
            df.set_index(df.columns[0], inplace=True)

            # Convert event date to datetime
            event_date = pd.to_datetime(selected_event.split()[-1]).date()

            # Combine event date with time index
            df.index = pd.to_datetime(event_date.strftime('%Y-%m-%d') + ' ' + df.index.astype(str))

            # Convert event date column to proper format
            event_col = f'Event {event_date}'
                
            # Determine which type of events to show based on the filter
            is_daily = self.daily_button.isChecked()
            is_targeted = self.targeted_button.isChecked()

            if is_daily:
                dispatch_list = self.cs_daily_dispatch_events
            elif is_targeted:
                dispatch_list = self.cs_targeted_dispatch_events

            matching_event = dispatch_list[dispatch_list['Event Date'] == event_date].iloc[0]
            event_start = datetime.combine(event_date, matching_event['Start Time'])
            event_end = datetime.combine(event_date, matching_event['End Time'])

            # Subtracting 3 hours using Timedelta
            start_time = event_start - pd.Timedelta(hours=3)
            end_time = event_end + pd.Timedelta(hours=1)
            df_filtered = df.loc[start_time:end_time]

            # Plot the data
            ax.plot(df_filtered.index, df_filtered['Baseline'], color=color_palette[0], label='Baseline kW', linewidth=0.9, linestyle='--')
            
            # Plot adjusted baseline with 5% error lines
            ax.plot(df_filtered.index, df_filtered['Adjusted Baseline'], color=color_palette[0], label='Adjusted Baseline', linewidth=1.2)
            ax.fill_between(df_filtered.index, 
                            df_filtered['Adjusted Baseline'] * 0.9, 
                            df_filtered['Adjusted Baseline'] * 1.1, 
                            color=color_palette[0], alpha=0.2)
            
            # Plot event day load with 5% error lines
            ax.plot(df_filtered.index, df_filtered[event_col], color=color_palette[5], label='Demand (kW)', linewidth=0.9)
            ax.fill_between(df_filtered.index, 
                            df_filtered[event_col] * 0.9, 
                            df_filtered[event_col] * 1.1, 
                            color=color_palette[5], alpha=0.2)

            # Highlight the event period
            ax.axvspan(event_start, event_end, color=color_palette[5], alpha=0.3, label='Event')

            # Highlight the baseline adjustment window
            adjustment_end = event_start - pd.Timedelta(hours=1)
            adjustment_start = adjustment_end - pd.Timedelta(hours=2)
            ax.axvspan(adjustment_start, adjustment_end, color='yellow', alpha=0.3, label='Adjustment Window')

            # Shade performance
            event_mask = (df_filtered.index >= event_start) & (df_filtered.index <= event_end)
            ax.fill_between(df_filtered.index[event_mask], 
                            df_filtered[event_col][event_mask], 
                            df_filtered['Adjusted Baseline'][event_mask], 
                            where=df_filtered[event_col][event_mask] <= df_filtered['Adjusted Baseline'][event_mask], 
                            color=color_palette[9], label='Performance', interpolate=True, alpha=0.3)

            ax.fill_between(df_filtered.index[event_mask], 
                            df_filtered['Adjusted Baseline'][event_mask], 
                            df_filtered[event_col][event_mask], 
                            where=df_filtered[event_col][event_mask] > df_filtered['Adjusted Baseline'][event_mask], 
                            color='red', label='No Performance', interpolate=True, alpha=0.3)

            # Format x-axis to show times
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
            ax.xaxis.set_major_locator(mdates.HourLocator(interval=1))

            # Remove borders
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False) 

            # Add horizontal grid lines
            ax.yaxis.grid(True, linestyle='--', which='major', color='grey', alpha=.25)

            # Set y-axis limits
            y_min = df_filtered[['Baseline', 'Adjusted Baseline', event_col]].min().min()
            y_max = df_filtered[['Baseline', 'Adjusted Baseline', event_col]].max().max()
            ax.set_ylim(y_min - (y_max - y_min) * 0.1, y_max + (y_max - y_min) * 0.1)

            return ax

        except Exception as e:
            self.receive_message(f"Error processing file for account {account}: {str(e)}")
            # Print full traceback for debugging
            traceback.print_exc()
            return ax

    def create_baseline_chart(self, account, selected_event, ax):
        df, start_time, end_time = self.get_baseline_data(account, selected_event)
        
        if df.empty:
            ax.text(0.5, 0.5, "No baseline data available", ha='center', va='center', transform=ax.transAxes)
            return ax

        baseline_dates = df.iloc[:, :10].columns.tolist()
        baseline_times = pd.date_range(start=start_time, end=end_time, freq='15T')

        # Plot individual baseline days in gray
        for date in baseline_dates:
            ax.plot(baseline_times, df[date], color='gray', alpha=0.5, linewidth=1)

        # Plot the average baseline in dotted orange with 5% error lines
        average_baseline = df.iloc[:, :10].mean(axis=1)
        ax.plot(baseline_times, average_baseline, color=color_palette[0], linestyle='--', linewidth=1, label='Baseline kW')
        ax.fill_between(baseline_times, 
                        average_baseline * 0.9, 
                        average_baseline * 1.1, 
                        color=color_palette[0], alpha=0.2)

        # Highlight the event period
        event_date = pd.to_datetime(selected_event.split()[-1]).date()
        is_daily = self.daily_button.isChecked()
        dispatch_list = self.cs_daily_dispatch_events if is_daily else self.cs_targeted_dispatch_events
        matching_event = dispatch_list[dispatch_list['Event Date'] == event_date].iloc[0]
        event_start = datetime.combine(event_date, matching_event['Start Time'])
        event_end = datetime.combine(event_date, matching_event['End Time'])
        ax.axvspan(event_start, event_end, color=color_palette[5], alpha=0.3, label='Event')

        # Highlight the baseline adjustment window
        adjustment_end = event_start - pd.Timedelta(hours=1)
        adjustment_start = adjustment_end - pd.Timedelta(hours=2)
        ax.axvspan(adjustment_start, adjustment_end, color='yellow', alpha=0.3, label='Adjustment Window')

        # Format x-axis to show times
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=1))

        # Remove borders
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        # Add horizontal grid lines
        ax.yaxis.grid(True, linestyle='--', which='major', color='grey', alpha=.25)

        # Add hover annotations
        cursor = mplcursors.cursor(ax, hover=True)

        @cursor.connect("add")
        def on_add(sel):
            line = sel.artist
            date = line.get_gid()
            x, y = sel.target
            sel.annotation.set(text=f"{date}\n{x.strftime('%H:%M')}: {y:.2f} kW", position=(0, 20))
            sel.annotation.get_bbox_patch().set(alpha=0.8)

        return ax
    
    def get_baseline_data(self, account, selected_event):
        # Construct the path to the selected event folder
        results_folder = f"Settlements/CS/{self.season}/CPOWER Results/{selected_event}"
        
        if not os.path.exists(results_folder):
            self.receive_message(f"Folder not found for event: {selected_event}")
            return pd.DataFrame(), None, None

        # Find the Excel file for the selected account
        excel_file = next((f for f in os.listdir(results_folder) if f.startswith(account) and f.endswith('.xlsx')), None)
        
        if excel_file is None:
            self.receive_message(f"No data file found for account: {account}")
            return pd.DataFrame(), None, None

        file_path = os.path.join(results_folder, excel_file)

        try:
            # Load the Excel file
            df = pd.read_excel(file_path, sheet_name='Calculations')
            
            # Drop the first row and reset the index
            df = df.iloc[1:].reset_index(drop=True)
            
            # Set the first column as the index
            df.set_index(df.columns[0], inplace=True)

            # Convert event date to datetime
            event_date = pd.to_datetime(selected_event.split()[-1]).date()

            # Combine event date with time index
            df.index = pd.to_datetime(event_date.strftime('%Y-%m-%d') + ' ' + df.index.astype(str))

            # Determine which type of events to show based on the filter
            is_daily = self.daily_button.isChecked()
            is_targeted = self.targeted_button.isChecked()

            if is_daily:
                dispatch_list = self.cs_daily_dispatch_events
            elif is_targeted:
                dispatch_list = self.cs_targeted_dispatch_events

            matching_event = dispatch_list[dispatch_list['Event Date'] == event_date].iloc[0]
            event_start = datetime.combine(event_date, matching_event['Start Time'])
            event_end = datetime.combine(event_date, matching_event['End Time'])

            # Subtracting 3 hours using Timedelta
            start_time = event_start - pd.Timedelta(hours=3)
            end_time = event_end + pd.Timedelta(hours=1)
            df_filtered = df.loc[start_time:end_time]

            return df_filtered, start_time, end_time

        except Exception as e:
            self.receive_message(f"Error processing file for account {account}: {str(e)}")
            return pd.DataFrame(), None, None

    def get_performance_text(self, account, selected_event):
        # Get performance data
        baseline, adjustment, adjusted_baseline, load, performance = self.get_performance_data(account, selected_event)

        # Create text in a format similar to performance reports
        performance_text = f"""
        {baseline:>10.2f} kW (Baseline)
        {adjustment:>10.2f} kW (Adjustment)
        {adjusted_baseline:>10.2f} kW (Adjusted Baseline)
        {load:>10.2f} kW (Load)
        =======================
        {performance:>10.2f} kW (Performance)
        """

        return performance_text
    
    def get_performance_data(self, account, selected_event):
        # Get the baseline data
        df, start_time, end_time = self.get_baseline_data(account, selected_event)
        
        if df.empty:
            return 0, 0, 0, 0, 0  # Return zeros if no data is found

        # Convert event date to datetime
        event_date = pd.to_datetime(selected_event.split()[-1]).date()

        # Determine which type of events to show based on the filter
        is_daily = self.daily_button.isChecked()
        is_targeted = self.targeted_button.isChecked()

        if is_daily:
            dispatch_list = self.cs_daily_dispatch_events
        elif is_targeted:
            dispatch_list = self.cs_targeted_dispatch_events

        matching_event = dispatch_list[dispatch_list['Event Date'] == event_date].iloc[0]
        event_start = datetime.combine(event_date, matching_event['Start Time'])
        event_end = datetime.combine(event_date, matching_event['End Time'])

        # Get the event column name
        event_col = f'Event {event_date}'

        # Calculate average baseline during the event
        baseline = df.loc[event_start:event_end, 'Baseline'].mean()

        # Calculate adjustment
        adjustment_start = event_start - pd.Timedelta(hours=3)
        adjustment_end = event_start
        adjustment = df.loc[adjustment_start:adjustment_end, 'Adjusted Baseline'].mean() - df.loc[adjustment_start:adjustment_end, 'Baseline'].mean()

        # Calculate adjusted baseline
        adjusted_baseline = baseline + adjustment

        # Calculate average load during the event
        load = df.loc[event_start:event_end, event_col].mean()

        # Calculate performance
        performance = adjusted_baseline - load

        return baseline, adjustment, adjusted_baseline, load, performance

    def get_color(self, value):
        # Color scale from red (low) to green (high) with less vibrancy
        r = int(max(0, min(128, 128 * (1 - value) * 2)))  # Red decreases as value increases
        g = int(max(0, min(128, 128 * value * 2)))  # Green increases as value increases
        return QColor(r, g, 0, 200)  # Less vibrant colors with slightly opaque

    def get_event_list(self):
        base_folder = f"Settlements/CS/{self.season}/CPOWER Results"
        if not os.path.exists(base_folder):
            self.receive_message(f"The folder {base_folder} does not exist.")
            return []
        
        event_type = "Daily" if self.daily_button.isChecked() else "Targeted"
        events = [f for f in os.listdir(base_folder) if os.path.isdir(os.path.join(base_folder, f)) and f.startswith(f"{event_type} Event")]
        return sorted(events)

    def get_customer_name(self, account):
        account_str = str(account)
        for df in self.cs_enroll_dfs.values():
            df['Asset ID'] = df['Asset ID'].astype(str)
            customer = df.loc[df['Asset ID'] == account_str, 'Customer']
            if not customer.empty:
                return customer.iloc[0]
        return "Unknown Customer"

    def get_customer_asset(self, account):
        account_str = str(account)
        for df in self.cs_enroll_dfs.values():
            df['Asset ID'] = df['Asset ID'].astype(str)
            asset = df.loc[df['Asset ID'] == account_str, 'Customer Asset']
            if not asset.empty:
                return str(asset.iloc[0])  # Convert to string here
        return "Unknown Asset"

    def get_expected_value(self, account):
        account_str = str(account)
        for df in self.cs_enroll_dfs.values():
            df['Asset ID'] = df['Asset ID'].astype(str)
            expected = df.loc[df['Asset ID'] == account_str, 'Forecasted KW']
            if not expected.empty:
                try:
                    return float(expected.iloc[0])
                except ValueError:
                    continue
        return 0

    def create_geo_analysis_settle_cs(self):
        # Clear current layout
        self.clear_data_viewer()

        # Create main layout
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)  # Remove margins
        layout.setSpacing(0)  # Remove spacing between widgets

        # Add title
        title = QLabel("Geo Analysis - Connected Solutions")
        title.setStyleSheet("font-size: 18px; font-weight: bold; padding: 10px;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # Add event selection dropdown
        event_layout = QHBoxLayout()
        event_layout.setContentsMargins(10, 0, 10, 10)  # Add some padding
        event_label = QLabel("Select Event:")
        self.event_combo_geo = QComboBox()
        self.populate_event_selection_combo_geo()
        self.event_combo_geo.currentIndexChanged.connect(self.update_geo_map)
        event_layout.addWidget(event_label)
        event_layout.addWidget(self.event_combo_geo)
        layout.addLayout(event_layout)

        # Create QWebEngineView for the map
        self.map_view = QWebEngineView()
        self.map_view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        layout.addWidget(self.map_view)

        # Create a widget to hold the layout
        widget = QWidget()
        widget.setLayout(layout)

        # Add the widget to the data viewer layout
        self.data_viewer_layout.addWidget(widget)

        # Initial map update
        self.update_geo_map()

    def update_geo_map(self):
        selected_event = self.event_combo_geo.currentText()
            
        # Determine which type of events to show based on the filter
        is_daily = self.daily_button.isChecked()
        is_targeted = self.targeted_button.isChecked()

        # Get the selected utilities
        selected_utilities = [checkbox.text() for checkbox in self.utility_checkbox_widget.findChildren(QCheckBox) if checkbox.isChecked()]
        
        # Combine all Connected Solutions enrollment DataFrames
        cs_df = pd.concat([df for df in self.cs_enroll_dfs.values()], ignore_index=True)

        # Merge cs_df with self.df to get the coordinates
        cs_df = cs_df.merge(self.df[['ACCPRG', 'Coordinates']], on='ACCPRG', how='left')

        # Apply filters
        if is_daily:
            cs_df = cs_df[cs_df['Program'].str.contains('Daily', case=False, na=False)]
        elif is_targeted:
            cs_df = cs_df[cs_df['Program'].str.contains('Targeted', case=False, na=False)]

        if selected_utilities:
            cs_df = cs_df[cs_df['Program'].apply(lambda x: any(utility.lower() in x.lower() for utility in selected_utilities))]

        # Create a folium map centered on New England with a grayscale style
        m = folium.Map(location=[43.8, -71.8], zoom_start=7, tiles='CartoDB positron')

        # Calculate max potential money for scaling
        def calculate_potential(row):
            performance = self.get_performance(row['ACCPRG'], selected_event)
            if performance is None or performance[0] is None:
                return 0
            forecasted = float(row.get('Forecasted KW', 0))
            return max(forecasted - performance[0], 0)

        max_potential_money = cs_df.apply(calculate_potential, axis=1).max()

        # Add markers for each asset
        for _, row in cs_df.iterrows():
            coordinates = row['Coordinates']
            if isinstance(coordinates, tuple) and len(coordinates) == 2:
                #try:
                lat, lon = map(float, coordinates)
                print((lat, lon))
                if not (math.isnan(lat) or math.isnan(lon)):
                    actual, forecasted = self.get_performance(row['ACCPRG'], selected_event)
                    print(("actual, forecasted"),(actual, forecasted))
                    if actual is not None and forecasted is not None:
                        performance_percentage = (actual / forecasted) * 100 if forecasted != 0 else 0
                        color = self.get_color_for_performance(performance_percentage)
                        
                        # Calculate the potential money on the table
                        potential_money = max(forecasted - actual, 0)
                        circle_size = 3 + (potential_money / max_potential_money) * 30 if max_potential_money > 0 else 3
                        
                        folium.CircleMarker(
                            location=[lat, lon],
                            radius=circle_size,
                            popup=f"<strong>{row['Customer']} - {row['Customer Asset']}</strong><br>"
                                f"Actual: {actual:.2f} kW<br>"
                                f"Forecasted: {forecasted:.2f} kW<br>"
                                f"Performance: {performance_percentage:.2f}%<br>"
                                f"Potential Improvement: {potential_money:.2f} kW",
                            tooltip=row['Asset ID'],
                            color=color,
                            fill=True,
                            fillColor=color,
                            fillOpacity=0.7
                        ).add_to(m)
                    else:
                        self.receive_message(f"Invalid coordinates (NaN) for ACCPRG {row['ACCPRG']}")
                #except ValueError:
                #    self.receive_message(f"Invalid coordinates for ACCPRG {row['ACCPRG']}: {coordinates}")
            else:
                self.receive_message(f"Missing or invalid coordinates for ACCPRG {row['ACCPRG']}")

        # Add a legend
        legend_html = '''
        <div style="position: fixed; 
                    bottom: 50px; 
                    left: 50px; 
                    width: 180px; 
                    background-color: white; 
                    border: 2px solid grey; 
                    z-index:9999; 
                    font-size:14px;
                    padding: 10px;">
            <strong>Performance</strong><br>
            <i class="fa fa-circle" style="color:#1a9850"></i> 90%+ <br>
            <i class="fa fa-circle" style="color:#91cf60"></i> 75-90% <br>
            <i class="fa fa-circle" style="color:#d9ef8b"></i> 50-75% <br>
            <i class="fa fa-circle" style="color:#fee08b"></i> 25-50% <br>
            <i class="fa fa-circle" style="color:#d73027"></i> <25% <br>
            <br>
            <strong>Circle Size</strong><br>
            Larger circles indicate more potential for improvement
        </div>
        '''
        m.get_root().html.add_child(folium.Element(legend_html))

        # Save map to data
        data = io.BytesIO()
        m.save(data, close_file=False)

        # Load map in QWebEngineView
        self.map_view.setHtml(data.getvalue().decode())

    def get_performance(self, accprg, event):
        for i in range(self.treeview.topLevelItemCount()):
            item = self.treeview.topLevelItem(i)
            if item.text(4) == accprg:  # Assuming ACCPRG is in column 4
                try:
                    forecasted = float(item.text(13))  # Forecasted kW in column 13
                    for column in range(self.treeview.columnCount()):
                        if self.treeview.headerItem().text(column) == f"{event}":
                            actual = float(item.text(column))
                            return actual, forecasted
                except ValueError:
                    self.receive_message(f"Invalid performance or forecast value for ACCPRG {accprg} on event {event}")
                    return None, None
        self.receive_message(f"No performance data found for ACCPRG {accprg} on event {event}")
        return None, None

    def get_color_for_performance(self, performance_percentage):
        if performance_percentage >= 90:
            return '#1a9850'  # Dark green
        elif performance_percentage >= 75:
            return '#91cf60'  # Light green
        elif performance_percentage >= 50:
            return '#d9ef8b'  # Yellowish green
        elif performance_percentage >= 25:
            return '#fee08b'  # Light orange
        else:
            return '#d73027'  # Red

    def populate_event_selection_combo_geo(self):
        self.event_combo_geo.clear()
        events = self.get_event_list()
        
        for event in events:
            self.event_combo_geo.addItem(event.split[2])
        
    # endregion
    # endregion

    # region GUI Settle Connected Solutions Settle Buttons
    def handle_select_summer_settle_cs(self):
        self.get_dispatches(year=int(self.summer_year_combo_settle_cs.currentText()))
        self.set_columns_for_settle_treeview()

    def handle_calculate_settle_cs(self):
        self.calculate_settle_cs_button.setEnabled(False)
        self.calculate_settle_cs_button.setText("Calculating...")
        
        self.calculate_spinner_timer = QTimer(self)
        self.calculate_spinner_timer.timeout.connect(self.update_calculate_spinner)
        self.calculate_spinner_timer.start(100)
        
        selected_items = self.get_selected_items_for_calculation()
        
        worker = Worker(self.perform_cs_calculations, selected_items)
        worker.signals.finished.connect(self.on_calculate_finished)
        worker.signals.result.connect(self.update_calculation_results)
        worker.signals.error.connect(self.handle_worker_error)
        
        QThreadPool.globalInstance().start(worker)

    def handle_worker_error(self, error_tuple):
        exctype, value, traceback_str = error_tuple
        self.receive_message(f"An error occurred: {value}")

    def update_calculate_spinner(self):
        current_text = self.calculate_settle_cs_button.text()
        if current_text.endswith("..."):
            self.calculate_settle_cs_button.setText("Calculating")
        else:
            self.calculate_settle_cs_button.setText(current_text + ".")

    def on_calculate_finished(self):
        self.calculate_spinner_timer.stop()
        self.calculate_settle_cs_button.setEnabled(True)
        self.calculate_settle_cs_button.setText("Calculate")

    def get_selected_items_for_calculation(self):
        selected_items = []
        for item in range(self.treeview.topLevelItemCount()):
            if not self.treeview.topLevelItem(item).isHidden() and self.treeview.topLevelItem(item).checkState(0) == QtCore.Qt.Checked:
                item_data = {
                    'item_index': item,
                    'asset_id': self.treeview.topLevelItem(item).text(3),
                    'accprg': self.treeview.topLevelItem(item).text(4),
                    'customer': self.treeview.topLevelItem(item).text(5),
                    'customer_asset': self.treeview.topLevelItem(item).text(6),
                    'uan': self.treeview.topLevelItem(item).text(7),
                    'program': self.treeview.topLevelItem(item).text(2),
                    'curtailment_strategy': self.treeview.topLevelItem(item).text(9),
                    'data_source': self.treeview.topLevelItem(item).text(10),
                }
                selected_items.append(item_data)
        return selected_items
    
    def perform_cs_calculations(self, selected_items):
        results = []
        year = self.summer_year_combo_settle_cs.currentText()
        season = "Summer " + year
        
        is_daily = self.daily_button.isChecked()
        event_type = "Daily" if is_daily else "Targeted"
        dispatch_list = self.cs_daily_dispatch_events if is_daily else self.cs_targeted_dispatch_events
        customer_dispatch_list = self.cs_daily_customer_dispatches if is_daily else self.cs_targeted_customer_dispatches

        for item in selected_items:
            try:
                tab = self.program_tab_name_mapping[item['program']]

                if item['accprg'] not in self.cs_settle_dfs[tab]['ACCPRG'].values:
                    continue

                events_to_process = [col for col in self.cs_settle_dfs[tab].columns 
                                    if re.match(r'.*\d{4}-\d{2}-\d{2}', col) and col != "Seasonal Average"]
                
                for event_column_name in events_to_process:
                    event_date = pd.to_datetime(event_column_name).date()
                    matching_event = dispatch_list[dispatch_list['Event Date'] == event_date]
                    matching_event = matching_event[matching_event['Program Name'] == item['program']]

                    if not matching_event.empty:
                        cs_calculator = CSPerformanceCalculator(
                            item['asset_id'], item['uan'], item['customer'], season, event_date, event_type,
                            item['customer_asset'], item['curtailment_strategy'],
                            dispatch_list, customer_dispatch_list,
                        )
                        event_performance_value = cs_calculator.create_cs_excel()
                        results.append((tab, item['accprg'], event_column_name, event_performance_value))
                    else:
                        self.receive_message(f"No matching event found for {event_column_name}")

                cpower_values = []
                for result in results:
                    tab, accprg, event_column_name, event_performance_value = result
                    if re.match(r'.*\d{4}-\d{2}-\d{2}', event_column_name):
                        cpower_values.append(event_performance_value)

                # Convert to numeric, replacing non-numeric values with NaN
                cpower_values = pd.to_numeric(cpower_values, errors='coerce')

                # Calculate average, ignoring NaN values
                if len(cpower_values) > 0:
                    cpower_avg = np.nanmean(cpower_values)
                else:
                    cpower_avg = np.nan  # or any other default value you prefer

                results.append((tab, item['accprg'], "Seasonal Average", cpower_avg))

            except Exception as e:
                self.receive_message(f"Error processing asset {item['asset_id']}: {str(e)}")

        return results

    def update_calculation_results(self, results):
        for tab, accprg, column_name, value in results:
            self.cs_settle_dfs[tab].loc[self.cs_settle_dfs[tab]['ACCPRG'] == accprg, column_name] = value

        # Update the treeview
        self.update_treeview_with_results(results)

        # Save the updated DataFrame
        try:
            file_path = f"Settlements/CS/{self.season}/Connected Solutions {self.season} Settlement Tracker.xlsx"
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
                for tab, df in self.cs_settle_dfs.items():
                    df.to_excel(writer, sheet_name=tab, index=False)
            self.receive_message(f"Updated and saved Connected Solutions Settlement Tracker for {self.season}")
        except Exception as e:
            self.receive_message(f"Failed to save updated CS Settlement Tracker: {e}")

    def update_treeview_with_results(self, results):
        for tab, accprg, column_name, value in results:
            for i in range(self.treeview.topLevelItemCount()):
                item = self.treeview.topLevelItem(i)
                if item.text(4) == accprg:  # Assuming ACCPRG is in column 4
                    for j in range(self.treeview.columnCount()):
                        if self.treeview.headerItem().text(j) == column_name:
                            item.setText(j, str(round(value, 1) if isinstance(value, float) else value))
                            break
                    break

    def populate_event_selection_combo_settle_cs(self):
        self.event_selection_combo_settle_cs.clear()
        self.custom_report_selection_list_settle_cs.clear()
        if self.daily_button.isChecked():
            events = self.cs_daily_dispatch_events.drop_duplicates(subset=['Event Date'])['Event Date'].tolist()
        else:  # targeted
            events = self.cs_targeted_dispatch_events.drop_duplicates(subset=['Event Date'])['Event Date'].tolist()
        
        for event in events:
            item = event.strftime('%Y-%m-%d')
            self.event_selection_combo_settle_cs.addItem(item)
            item = QtWidgets.QListWidgetItem(event.strftime('%Y-%m-%d'))
            self.custom_report_selection_list_settle_cs.addItem(item)

    def handle_select_recommended_settle_cs(self):
        company_performances = {}
        company_items = {}

        # Determine if it's daily or targeted
        event_type = "Daily" if self.daily_button.isChecked() else "Targeted"

        # Find the correct average column
        average_column = -1
        for column in range(self.treeview.columnCount()):
            if self.treeview.headerItem().text(column) == f"CPOWER Average":
                average_column = column
                break

        if average_column == -1:
            self.receive_message(f"Error: CPOWER Average column not found.")
            return

        # First pass: collect all items and their performances by company
        for item in range(self.treeview.topLevelItemCount()):
            tree_item = self.treeview.topLevelItem(item)
            if not tree_item.isHidden():
                company_name = tree_item.text(5)  # Assuming Company Name is in column 5
                average_value = tree_item.text(average_column)
                
                if company_name not in company_performances:
                    company_performances[company_name] = []
                    company_items[company_name] = []
                
                try:
                    average_float = float(average_value)
                    company_performances[company_name].append(average_float)
                    company_items[company_name].append(tree_item)
                except ValueError:
                    # If conversion to float fails, append 0
                    company_performances[company_name].append(0)
                    company_items[company_name].append(tree_item)

        # Second pass: select items based on company performance
        selected_items = []
        for company, performances in company_performances.items():
            if any(perf > 0 for perf in performances):
                for item in company_items[company]:
                    item.setCheckState(0, QtCore.Qt.Checked)
                    asset_id = item.text(3)  # Assuming Asset ID is in column 3
                    selected_items.append(asset_id)
            else:
                for item in company_items[company]:
                    item.setCheckState(0, QtCore.Qt.Unchecked)

        if selected_items:
            self.receive_message(f"Selected {len(selected_items)} items from companies with positive average performance.")
        else:
            self.receive_message("No items found from companies with positive average performance.")
            
    def handle_generate_report_settle_cs(self):
        selected_items = []
        for item in range(self.treeview.topLevelItemCount()):
            if not self.treeview.topLevelItem(item).isHidden() and self.treeview.topLevelItem(item).checkState(0) == QtCore.Qt.Checked:
                accprg = self.treeview.topLevelItem(item).text(4)
                asset_id = self.treeview.topLevelItem(item).text(3)
                customer = self.treeview.topLevelItem(item).text(5)
                customer_asset = self.treeview.topLevelItem(item).text(6)
                utility_account_number = self.treeview.topLevelItem(item).text(7)
                program = self.treeview.topLevelItem(item).text(2)
                curtailment_strategy = self.treeview.topLevelItem(item).text(10)
                facility_address = self.df.loc[self.df['ACCPRG'] == accprg, 'Facility Address'].iloc[0] if not self.df[self.df['ACCPRG'] == accprg].empty else ""
                
                # Get the correct tab name using the mapping
                tab_name = next((v for k, v in self.program_tab_name_mapping.items() if k in program), None)

                customer_share = self.cs_enroll_dfs[tab_name].loc[self.cs_enroll_dfs[tab_name]['ACCPRG'] == accprg, 'Customer Share'].iloc[0] if not self.cs_enroll_dfs[tab_name][self.cs_enroll_dfs[tab_name]['ACCPRG'] == accprg].empty else ""

                # If customer_share is still empty, fall back to self.df
                if not customer_share:
                    customer_share = self.df.loc[self.df['ACCPRG'] == accprg, 'Customer Share'].iloc[0] if not self.df[self.df['ACCPRG'] == accprg].empty else ""

                selected_items.append((accprg, asset_id, customer, customer_asset, utility_account_number, program, curtailment_strategy, facility_address, customer_share))

        season = "Summer " + self.summer_year_combo_settle_cs.currentText()
        event_date = pd.to_datetime(self.event_selection_combo_settle_cs.currentText()).date()
        event_type = "Daily" if self.daily_button.isChecked() else "Targeted"

        for accprg, asset_id, customer, customer_asset, utility_account_number, program, curtailment_strategy, facility_address, customer_share in selected_items:
            #try:
            # Determine which type of events to show based on the filter
            is_daily = self.daily_button.isChecked()
            is_targeted = self.targeted_button.isChecked()

            if is_daily:
                dispatch_list = self.cs_daily_dispatch_events
                customer_dispatch_list = self.cs_daily_customer_dispatches
            elif is_targeted:
                dispatch_list = self.cs_targeted_dispatch_events
                customer_dispatch_list = self.cs_targeted_customer_dispatches
                
            report_generator = CSReportGenerator(
                asset_id, customer, season, event_date, event_type, customer_asset, 
                utility_account_number, facility_address, customer_share,
                dispatch_list, customer_dispatch_list,
            )

            #if curtailment_strategy == 'Battery':
            #    report_generator.create_non_baseline_asset_report(program)
            #else:
            report_generator.create_asset_report(program)

            self.receive_message(f"Generated report for {asset_id}")
            #except Exception as e:
            #    self.receive_message(f"Error generating report for {asset_id}: {str(e)}")

        self.receive_message(f"Report generation complete for {len(selected_items)} assets.")

    def handle_generate_customer_report(self):
        if self.calculate_event_performances_settle_cs_checkbox.isChecked():
            report_type = "Event"
        elif self.calculate_month_performances_settle_cs_checkbox.isChecked():
            report_type = "Monthly"
        elif self.calculate_custom_performances_settle_cs_checkbox.isChecked():
            report_type = "Custom"
        else:
            report_type = "Event"  # Default to Event if none is checked

        # Ask user if they want a full report or one-pager
        reply = QMessageBox.question(self, 'Report Type',
                                    'Do you want to generate a full detail report? If no, this will generate a one-page summary.',
                                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                    QMessageBox.StandardButton.Yes)
        
        if reply == QMessageBox.StandardButton.Yes:
            report_format = "Full"
        else:
            report_format = "OnePager"

        if report_type == "Custom":
            selected_events = [item.text() for item in self.custom_report_selection_list_settle_cs.selectedItems()]
            if not selected_events:
                QMessageBox.warning(self, "No Events Selected", "Please select at least one event for the custom report.")
                return
        else:
            selected_events = None

        self.generate_customer_report(report_type, report_format, selected_events)

    def generate_customer_report(self, report_type, report_format, selected_events=None):
        customer_dfs = self.get_customer_dfs()
        season = "Summer " + self.summer_year_combo_settle_cs.currentText()
        event_type = "Daily" if self.daily_button.isChecked() else "Targeted"

        for customer, df in customer_dfs.items():
            if report_type == "Event":
                event_date = pd.to_datetime(self.event_selection_combo_settle_cs.currentText()).date()
                self.generate_event_customer_report(df, season, event_date, event_type, report_format)
            elif report_type == "Monthly":
                month = self.month_report_selection_combo_settle_cs.currentText()
                self.generate_monthly_customer_report(df, season, month, event_type, report_format)
            elif report_type == "Season":
                self.generate_season_customer_report(df, season, event_type, report_format)
            elif report_type == "Custom":
                selected_events = [pd.to_datetime(item.text()).date() for item in self.custom_report_selection_list_settle_cs.selectedItems()]
                if selected_events:
                    self.generate_custom_customer_report(df, season, selected_events, event_type, report_format)
                else:
                    self.receive_message("No events selected for custom report.")

    def generate_custom_customer_report(self, customer_df, season, selected_events, event_type, report_format):
        # Determine which type of events to show based on the filter
        is_daily = self.daily_button.isChecked()
        is_targeted = self.targeted_button.isChecked()

        if is_daily:
            dispatch_list = self.cs_daily_dispatch_events
            customer_dispatch_list = self.cs_daily_customer_dispatches
        elif is_targeted:
            dispatch_list = self.cs_targeted_dispatch_events
            customer_dispatch_list = self.cs_targeted_customer_dispatches

        cs_calculator = CSReportGenerator(
            asset_id=customer_df['Asset ID'].iloc[0],
            customer=customer_df['Customer'].iloc[0],
            season=season,
            event=None,  # We'll handle multiple events in the aggregate_reports method
            event_type=event_type,
            customer_asset=customer_df['Customer Asset'].iloc[0],
            utility_account_number=customer_df['Utility Account Number'].iloc[0],
            address=customer_df['Facility Address'].iloc[0],
            customer_share=customer_df['Customer Share'].iloc[0],
            dispatches=dispatch_list,
            customer_dispatches=customer_dispatch_list,
        )
        cs_calculator.aggregate_reports(customer_df, "Custom", report_format, selected_events=selected_events)

    def get_customer_dfs(self):
        selected_items = []

        # Get all column headers
        headers = [self.treeview.headerItem().text(i) for i in range(self.treeview.columnCount())]
        
        # Filter for CPOWER event columns
        cpower_event_columns = [col for col in headers if re.match(r'.*\d{4}-\d{2}-\d{2}', col) and col != "Seasonal Average"]
        
        for item in range(self.treeview.topLevelItemCount()):
            if not self.treeview.topLevelItem(item).isHidden() and self.treeview.topLevelItem(item).checkState(0) == Qt.Checked:
                asset_id = self.treeview.topLevelItem(item).text(3)
                accprg = self.treeview.topLevelItem(item).text(4)
                customer = self.treeview.topLevelItem(item).text(5)
                customer_asset = self.treeview.topLevelItem(item).text(6)
                program = self.program_tab_name_mapping[self.treeview.topLevelItem(item).text(2)]
                full_program_name = self.treeview.topLevelItem(item).text(2)
                utility_account_number = self.treeview.topLevelItem(item).text(7)
                curtailment_strategy = self.treeview.topLevelItem(item).text(9)
                facility_address = self.df.loc[self.df['ACCPRG'] == accprg, 'Facility Address'].iloc[0] if not self.df[self.df['ACCPRG'] == accprg].empty else ""

                # Get the correct tab name using the mapping
                tab_name = next((v for k, v in self.program_tab_name_mapping.items() if v in program), None)
                if tab_name and tab_name in self.cs_enroll_dfs:
                    customer_share = self.cs_enroll_dfs[tab_name].loc[self.cs_enroll_dfs[tab_name]['ACCPRG'] == accprg, 'Customer Share'].iloc[0] if not self.cs_enroll_dfs[tab_name][self.cs_enroll_dfs[tab_name]['ACCPRG'] == accprg].empty else ""
                    performance_contact = self.cs_enroll_dfs[tab_name].loc[self.cs_enroll_dfs[tab_name]['ACCPRG'] == accprg, 'Performance Report Contact'].iloc[0] if not self.cs_enroll_dfs[tab_name][self.cs_enroll_dfs[tab_name]['ACCPRG'] == accprg].empty else ""
                else:
                    customer_share = ""
                    performance_contact = ""

                # If customer_share is still empty, fall back to self.df
                if not customer_share:
                    customer_share = self.df.loc[self.df['ACCPRG'] == accprg, 'Customer Share'].iloc[0] if not self.df[self.df['ACCPRG'] == accprg].empty else ""
                
                # Get CPOWER event performances
                performances = {}
                for column in cpower_event_columns:
                    column_index = headers.index(column)
                    performance = self.treeview.topLevelItem(item).text(column_index)
                    if performance == "Pending" or performance == "Data Downloaded":
                        performance = 0
                    else:
                        try:
                            performance = float(performance)
                        except ValueError:
                            performance = 0
                    performances[pd.to_datetime(column.split()[-1]).date()] = performance
                
                # Calculate average performance
                avg_performance = np.nanmean(list(performances.values())) if performances else np.nan
                
                try:
                    expected_kw = float(self.treeview.topLevelItem(item).text(12))
                except ValueError:
                    expected_kw = float('nan')
                
                selected_items.append((
                    asset_id, accprg, customer, customer_asset, program, full_program_name,
                    utility_account_number, curtailment_strategy, 
                    facility_address, customer_share, performance_contact, performances, avg_performance, expected_kw
                ))

        # Create DataFrame
        df = pd.DataFrame(selected_items, columns=[
            'Asset ID', 'ACCPRG', 'Customer', 'Customer Asset', 'Program', 'Full Program Name',
            'Utility Account Number', 'Curtailment Strategy', 
            "Facility Address", "Customer Share", "Performance Contact", 'Performances', 'Average Performance', 'Expected kW'
        ])

        # Create individual event columns with dates
        for date in performances.keys():
            df[f'Event {date}'] = df['Performances'].apply(lambda x: x.get(date, np.nan))

        # Drop the Performances column as we now have individual event columns
        df = df.drop('Performances', axis=1)

        return {customer: group for customer, group in df.groupby(['Customer', 'Program'])}

    def generate_event_customer_report(self, customer_df, season, event_date, event_type, report_format):
        # Determine which type of events to show based on the filter
        is_daily = self.daily_button.isChecked()
        is_targeted = self.targeted_button.isChecked()

        if is_daily:
            dispatch_list = self.cs_daily_dispatch_events
            customer_dispatch_list = self.cs_daily_customer_dispatches
        elif is_targeted:
            dispatch_list = self.cs_targeted_dispatch_events
            customer_dispatch_list = self.cs_targeted_customer_dispatches

        cs_calculator = CSReportGenerator(
            asset_id=customer_df['Asset ID'].iloc[0],
            customer=customer_df['Customer'].iloc[0],
            season=season,
            event=event_date,
            event_type=event_type,
            customer_asset=customer_df['Customer Asset'].iloc[0],
            utility_account_number=customer_df['Utility Account Number'].iloc[0],
            address=customer_df['Facility Address'].iloc[0],
            customer_share=customer_df['Customer Share'].iloc[0],
            dispatches=dispatch_list,
            customer_dispatches=customer_dispatch_list,
        )
        cs_calculator.aggregate_reports(customer_df, "Event", report_format)

    def generate_monthly_customer_report(self, customer_df, season, month, event_type, report_format):
        # Determine which type of events to show based on the filter
        is_daily = self.daily_button.isChecked()
        is_targeted = self.targeted_button.isChecked()

        if is_daily:
            dispatch_list = self.cs_daily_dispatch_events[self.cs_daily_dispatch_events['Program Name'] == customer_df['Full Program Name'].iloc[0]]
            customer_dispatch_list = self.cs_daily_customer_dispatches
        elif is_targeted:
            dispatch_list = self.cs_targeted_dispatch_events[self.cs_targeted_dispatch_events['Program Name'] == customer_df['Full Program Name'].iloc[0]]
            customer_dispatch_list = self.cs_targeted_customer_dispatches

        cs_calculator = CSReportGenerator(
            asset_id=customer_df['Asset ID'].iloc[0],
            customer=customer_df['Customer'].iloc[0],
            season=season,
            event=None,
            event_type=event_type,
            customer_asset=customer_df['Customer Asset'].iloc[0],
            utility_account_number=customer_df['Utility Account Number'].iloc[0],
            address=customer_df['Facility Address'].iloc[0],
            customer_share=customer_df['Customer Share'].iloc[0],
            dispatches=dispatch_list,
            customer_dispatches=customer_dispatch_list,
        )
        cs_calculator.aggregate_reports(customer_df, "Monthly", report_format, month=month)

    def handle_request_data_settle_cs(self):
        selected_items = []
        for item in range(self.treeview.topLevelItemCount()):
            if not self.treeview.topLevelItem(item).isHidden() and self.treeview.topLevelItem(item).checkState(0) == QtCore.Qt.Checked:
                accprg = self.treeview.topLevelItem(item).text(4)
                selected_items.append(accprg)
        self.receive_message(f"Requesting data for the following:<br>{'<br>'.join(selected_items)}")

    def handle_send_data_settle_cs(self):
        selected_items = []
        for item in range(self.treeview.topLevelItemCount()):
            if not self.treeview.topLevelItem(item).isHidden() and self.treeview.topLevelItem(item).checkState(0) == QtCore.Qt.Checked:
                accprg = self.treeview.topLevelItem(item).text(4)
                selected_items.append(accprg)
        self.receive_message(f"Sending data for the following:<br>{'<br>'.join(selected_items)}")
        
    def handle_report_selected_settle_cs(self):
        selected_items = []
        for item in range(self.treeview.topLevelItemCount()):
            if not self.treeview.topLevelItem(item).isHidden() and self.treeview.topLevelItem(item).checkState(0) == QtCore.Qt.Checked:
                accprg = self.treeview.topLevelItem(item).text(4)
                selected_items.append(accprg)
        self.receive_message(f"Generating sales report for selected items:<br>{'<br>'.join(selected_items)}")

    def handle_report_all_settle_cs(self):
        pass

    def handle_pull_data_settle_cs(self):
        self.download_settle_cs_button.setEnabled(False)
        self.download_settle_cs_button.setText("Pulling Data...")
        
        # Create a QTimer to update the button text with a spinning animation
        self.spinner_timer = QTimer(self)
        self.spinner_timer.timeout.connect(self.update_spinner_pull_meter_cs)
        self.spinner_timer.start(100)
        
        # Use QThreadPool to run the data retrieval in the background
        worker = Worker(self.pull_data_cs)
        worker.signals.finished.connect(self.on_pull_data_finished)
        QThreadPool.globalInstance().start(worker)

    def update_spinner_pull_meter_cs(self):
        current_text = self.download_settle_cs_button.text()
        if current_text.endswith("..."):
            self.download_settle_cs_button.setText("Pulling Data")
        else:
            self.download_settle_cs_button.setText(current_text + ".")

    def on_pull_data_finished(self):
        self.spinner_timer.stop()
        self.download_settle_cs_button.setEnabled(True)
        self.download_settle_cs_button.setText("Pull Data")

    # endregion

    # region Notification Utils
    def receive_message(self, response):
        # Update the text of the status message label
        current_time = datetime.now().strftime("%H:%M:%S")
        self.status_message_label.setText(f"{current_time} - {response}")

        # Ensure the label is visible
        self.status_message_label.setVisible(True)

        # Optionally, you can add a timer to clear the message after a certain period
        QTimer.singleShot(30000, lambda: self.status_message_label.setVisible(False))  # Hide after 30 seconds
    # endregion

    # region GUI and Data Table Utils
    def clean_company_address(self, address):
        address = address.replace('\n', ', ')
        address = re.sub(r',\s*,', ',', address)
        address = address.strip(', ')
        return address

    def get_dispatches(self, year=2024):
        processor = DispatchDataProcessor(r"Data\Dispatch Events.xlsx", year=year)
        (
            self.cs_daily_dispatch_events,
            self.cs_daily_customer_dispatches,
            self.cs_targeted_dispatch_events,
            self.cs_targeted_customer_dispatches,
        ) = processor.get_dispatch_data()

    def load_primary_data(self):
        try:
            dynamics_df_raw = pd.read_excel("Data/Dynamics Source Data - Utility Accounts.xlsx")

            # Fill NaN values with an empty string to avoid issues during concatenation
            dynamics_df_raw = dynamics_df_raw.astype(object).fillna('')
            
            # Rename columns to match the specified format
            column_renames = {
                "Enrollment Status": "Status",
                "Program": "Program",
                "Name": "ACCPRG",
                "Account": "ACC",
                "Vendor ID (Account) (Account)": "Vendor ID",
                "Company": "Company",
                "Facility Name/Store # (Account) (Account)": "Facility Name",
                "Address 1 (Company) (Company)": "Company Address",
                "Service Address Line 1 (Account) (Account)": "Facility Address 1",
                "Service Address Line 2 (Account) (Account)": "Facility Address 2",
                "Service Address City (Account) (Account)": "Facility Address 3",
                "Service Address State (Account) (Account)": "Facility Address 4",
                "Service Address Zip 1 (Account) (Account)": "Facility Address 5",
                "Service Address Latitude (Account) (Account)": "Latitude",
                "Service Address Longitude (Account) (Account)": "Longitude",
                "Earliest Start Date": "Earliest Start",
                "End Date": "End Date",
                "Utility (Account) (Account)": "Utility",
                "Utility Account Number (Account) (Account)": "Utility Account Number",
                "Asset ID": "Asset ID",
                "Secondary ID (Account) (Account)": "Secondary ID",
                "Aggregation ID": "Aggregation ID",
                "Resource ID": "Resource ID",
                "Resource Name": "Resource Name",
            }
            dynamics_df_raw.rename(columns=column_renames, inplace=True)

            dynamics_df_raw['Company Address'] = dynamics_df_raw['Company Address'].apply(self.clean_company_address)

            # Fill NaN values with an empty string to avoid issues during concatenation
            dynamics_df_raw.fillna('', inplace=True)

            # Ensure all values are converted to strings before joining them
            dynamics_df_raw['Facility Address'] = dynamics_df_raw[
                ['Facility Address 1', 'Facility Address 2', 'Facility Address 3', 'Facility Address 4', 'Facility Address 5']
            ].apply(lambda x: ', '.join(x.astype(str)).strip(), axis=1)

            dynamics_df_raw['Coordinates'] = list(zip(dynamics_df_raw['Latitude'], dynamics_df_raw['Longitude']))
            
        except Exception as e:
            self.receive_message(f'Failed to read in Dynamics Enrollment Data: {e}')
            dynamics_df_raw = pd.DataFrame()

        self.df = dynamics_df_raw

    def load_enroll_data(self):
        try:
            file_path = f"Enrollments/CS/Connected Solutions Enrollment Tracker.xlsx"
            
            # List of all tabs to read
            self.cs_settle_tabs = ['CLT', 'EMT', 'EVD', 'EVT', 'LBT', 'NGD', 'NGT', 'RID', 'RIT', 'UND', 'UNT']
            
            # Read each tab into a separate DataFrame
            self.cs_enroll_dfs = {}
            for tab in self.cs_settle_tabs:
                try:
                    self.cs_enroll_dfs[tab] = pd.read_excel(file_path, sheet_name=tab)
                except Exception as e:
                    self.receive_message(f'Failed to read tab {tab} from Connected Solutions Enrollment Tracker: {e}')
                    self.cs_enroll_dfs[tab] = pd.DataFrame(columns=['ACCPRG', 'Customer', 'Customer Asset'])
            
            self.receive_message(f"Loaded Connected Solutions Enrollment Tracker.")

        except Exception as e:
                self.receive_message(f'Failed to read Connected Solutions Enrollment Tracker: {e}')
                self.cs_enroll_dfs = {tab: pd.DataFrame(columns=['ACCPRG', 'Customer', 'Customer Asset']) for tab in self.cs_settle_tabs}

    def load_settle_data(self):
        self.load_primary_data()
        self.load_enroll_data()

        connected_solutions_summer_season = "Summer " + self.summer_year_combo_settle_cs.currentText()

        try:
            file_path = f"Settlements/CS/{connected_solutions_summer_season}/Connected Solutions {connected_solutions_summer_season} Settlement Tracker.xlsx"
            
            # Ensure the directory exists
            os.makedirs(os.path.dirname(file_path), exist_ok=True)

            # Read each tab into a separate DataFrame
            self.cs_settle_tabs = ['CLT', 'EMT', 'EVD', 'EVT', 'LBT', 'NGD', 'NGT', 'RID', 'RIT', 'UND', 'UNT']
            self.cs_settle_dfs = {}

            for tab in self.cs_settle_tabs:
                try:
                    df = pd.read_excel(file_path, sheet_name=tab)
                except Exception:
                    # If the sheet doesn't exist, create a new DataFrame
                    df = pd.DataFrame(columns=['ACCPRG', 'Customer', 'Customer Asset', 'Expected KW'])

                # Ensure required columns exist
                required_columns = ["ACCPRG", "Customer", "Customer Asset", "Expected KW"]
                for col in required_columns:
                    if col not in df.columns:
                        df[col] = ""

                # Get the enrollment data for this tab
                enroll_df = self.cs_enroll_dfs.get(tab, pd.DataFrame())

                if not enroll_df.empty:
                    # Identify missing ACCPRGs
                    missing_accprgs = set(enroll_df['ACCPRG']) - set(df['ACCPRG'])

                    if missing_accprgs:
                        new_rows = []
                        for accprg in missing_accprgs:
                            new_row = enroll_df[enroll_df['ACCPRG'] == accprg]
                            customer_name = new_row['Customer'].values[0] if len(new_row['Customer'].values) > 0 else ""
                            customer_asset = new_row['Customer Asset'].values[0] if len(new_row['Customer Asset'].values) > 0 else ""
                            expected = new_row['Forecasted KW'].values[0] if len(new_row['Forecasted KW'].values) > 0 else 0
                            new_data = {
                                'ACCPRG': accprg,
                                'Customer': customer_name,
                                'Customer Asset': customer_asset,
                                'Expected KW': expected,
                            }
                            new_rows.append(new_data)
                        
                        new_df = pd.DataFrame(new_rows)
                        df = pd.concat([df, new_df], ignore_index=True)

                # Determine if it's a daily or targeted tab
                is_targeted = tab.endswith('T')
                is_daily = tab.endswith('D')

                # Get the appropriate event dates
                if is_daily:
                    events = self.cs_daily_dispatch_events['Event Date'].tolist()
                elif is_targeted:
                    events = self.cs_targeted_dispatch_events['Event Date'].tolist()
                else:
                    events = []

                # Add missing columns for each event
                for event in events:
                    cpower_col = f"{event}"
                    if cpower_col not in df.columns:
                        df[cpower_col] = ""

                # Add or update average columns
                cpower_cols = [col for col in df.columns 
                                    if re.match(r'.*\d{4}-\d{2}-\d{2}', col) and col != "Seasonal Average"]


                df[f"Seasonal Average"] = df[cpower_cols].mean(axis=1, numeric_only=True)

                # Ensure the average columns are at the end
                cols = df.columns.tolist()
                cols = [col for col in cols if col not in [f"Seasonal Average"]] + [f"Seasonal Average"]
                df = df[cols]

                self.cs_settle_dfs[tab] = df

            # Save the updated tracker
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
                for tab, df in self.cs_settle_dfs.items():
                    df.to_excel(writer, sheet_name=tab, index=False)
            
            self.receive_message(f"Loaded and updated Connected Solutions Settlement Tracker for {connected_solutions_summer_season}")

        except Exception as e:
            self.receive_message(f'Failed to read or update Connected Solutions Settlement Tracker: {e}')
            self.cs_settle_dfs = {tab: pd.DataFrame(columns=['ACCPRG', 'Customer', 'Customer Asset']) for tab in self.cs_settle_tabs}

        # Keep the existing dispatch event DataFrames unchanged
        if not hasattr(self, 'cs_daily_dispatch_events'):
            self.cs_daily_dispatch_events = pd.DataFrame()
        if not hasattr(self, 'cs_targeted_dispatch_events'):
            self.cs_targeted_dispatch_events = pd.DataFrame()

    def resize_columns_to_fit(self):
        max_width = 200  # Set the maximum width for columns (adjust as needed)
        for column in range(self.treeview.columnCount()):
            self.treeview.resizeColumnToContents(column)
            current_width = self.treeview.columnWidth(column)
            if current_width > max_width:
                self.treeview.setColumnWidth(column, max_width)
            else:
                # Add extra space to the column width
                new_width = min(current_width + 20, max_width)  # Adjust the value as needed
                self.treeview.setColumnWidth(column, new_width)

    def clear_data_viewer(self):
        for i in reversed(range(self.data_viewer_layout.count())):
            widget_to_remove = self.data_viewer_layout.itemAt(i).widget()
            self.data_viewer_layout.removeWidget(widget_to_remove)
            widget_to_remove.setParent(None)

    def update_custom_event_selection_visibility_settle_cs(self, state, report_type):
        self.event_selection_label_settle_cs.setVisible(False)
        self.event_selection_combo_settle_cs.setVisible(False)
        self.month_report_selection_label_settle_cs.setVisible(False)
        self.month_report_selection_combo_settle_cs.setVisible(False)
        self.custom_report_selection_label_settle_cs.setVisible(False)
        self.custom_report_selection_list_settle_cs.setVisible(False)

        if state == QtCore.Qt.Checked:
            if report_type == "Event":
                self.calculate_month_performances_settle_cs_checkbox.setChecked(False)
                self.calculate_custom_performances_settle_cs_checkbox.setChecked(False)
                self.event_selection_label_settle_cs.setVisible(True)
                self.event_selection_combo_settle_cs.setVisible(True)
                self.populate_event_selection_combo_settle_cs()
            elif report_type == "Month":
                self.calculate_event_performances_settle_cs_checkbox.setChecked(False)
                self.calculate_custom_performances_settle_cs_checkbox.setChecked(False)
                self.month_report_selection_label_settle_cs.setVisible(True)
                self.month_report_selection_combo_settle_cs.setVisible(True)
            elif report_type == "Season":
                self.calculate_event_performances_settle_cs_checkbox.setChecked(False)
                self.calculate_month_performances_settle_cs_checkbox.setChecked(False)
                self.calculate_custom_performances_settle_cs_checkbox.setChecked(False)
            elif report_type == "Custom":
                self.calculate_event_performances_settle_cs_checkbox.setChecked(False)
                self.calculate_month_performances_settle_cs_checkbox.setChecked(False)
                self.custom_report_selection_label_settle_cs.setVisible(True)
                self.custom_report_selection_list_settle_cs.setVisible(True)
                self.populate_event_selection_combo_settle_cs()

    def toggle_all_items(self):
        # Get the current state of the first item (if it exists)
        if self.treeview.topLevelItemCount() > 0:
            current_state = self.treeview.topLevelItem(0).checkState(0)
            # Toggle to the opposite state
            new_state = Qt.Unchecked if current_state == Qt.Checked else Qt.Checked
        else:
            new_state = Qt.Checked  # Default to checked if no items

        for i in range(self.treeview.topLevelItemCount()):
            item = self.treeview.topLevelItem(i)
            if not item.isHidden():
                item.setCheckState(0, new_state)
        self.update_selected_count()

    def update_selected_count(self):
        self.selected_count = 0
        self.total_count = 0
        self.selected_customers = set()
        self.total_customers = set()
        
        for i in range(self.treeview.topLevelItemCount()):
            item = self.treeview.topLevelItem(i)
            if not item.isHidden():
                self.total_count += 1
                customer = item.text(5)  # Assuming the customer name is in column 5
                self.total_customers.add(customer)
                
                if item.checkState(0) == Qt.Checked:
                    self.selected_count += 1
                    self.selected_customers.add(customer)
        
        self.update_status_bar()
        
    def update_status_bar(self):
        total_text = f"Total {{ Assets: {self.total_count}, Customers: {len(self.total_customers)} }}"
        selected_text = f"Selected {{ Assets: {self.selected_count}, Customers: {len(self.selected_customers)} }}"
        
        self.total_label.setText(total_text)
        self.selected_label.setText(selected_text)

    def update_settlement_period_label(self, program_type):
        if self.settlement_period_label_settle_cs is not None:
            self.title_layout.removeWidget(self.settlement_period_label_settle_cs)
            self.settlement_period_label_settle_cs.deleteLater()
            self.settlement_period_label_settle_cs = None

        self.settlement_period_label_settle_cs = QtWidgets.QLabel(self)
        self.settlement_period_label_settle_cs.setText(f"Summer {self.summer_year_combo_settle_cs.currentText()} {program_type}")
        self.settlement_period_label_settle_cs.setObjectName("settlement_period_label")
        self.title_layout.insertWidget(0, self.settlement_period_label_settle_cs)

    def toggle_treeview(self):
        self.stacked_layout.setCurrentIndex(0)
        self.toggle_treeview_button.setChecked(True)

    def toggle_data_viewer(self):
        self.stacked_layout.setCurrentIndex(1)
        self.toggle_data_viewer_button.setChecked(True)

    # endregion

    # region Data Table Controls
    def apply_daily_connected_solution_filter(self):
        self.daily_button.setChecked(True)
        self.targeted_button.setChecked(False)
        self.set_columns_for_settle_treeview()
        self.apply_connected_solution_filters()
        self.update_settlement_period_label("Daily")

    def apply_targeted_connected_solution_filter(self):
        self.daily_button.setChecked(False)
        self.targeted_button.setChecked(True)
        self.set_columns_for_settle_treeview()
        self.apply_connected_solution_filters()
        self.update_settlement_period_label("Targeted")

    def apply_connected_solution_filters(self):
        selected_utilities = [checkbox.text() for checkbox in self.utility_checkbox_widget.findChildren(QCheckBox) if checkbox.isChecked()]
        
        program_type = "Daily" if self.daily_button.isChecked() else "Targeted"

        for item in range(self.treeview.topLevelItemCount()):
            asset_item = self.treeview.topLevelItem(item)
            program = asset_item.text(2).lower()
            enrollment_status = asset_item.text(1).lower()

            if "cape light"  in program:
                program = "cape light connected solutions targeted"
            if "efficiency maine" in program:
                program = "efficiency maine connected solutions targeted"

            # Hide the item by default
            asset_item.setHidden(True)

            # Check if it's a Connected Solutions program
            if "connected solutions" not in program:
                continue

            # Check if it matches the current program type (Daily or Targeted)
            if program_type.lower() not in program:
                continue

            # Check utility match
            utility_match = not selected_utilities or any(utility.lower() in program for utility in selected_utilities)
            if not utility_match:
                continue

            # If we've made it this far, show the item
            asset_item.setHidden(False)

        self.update_selected_count()
        self.resize_columns_to_fit()

    def apply_super_filter(self, filter_text="Connected Solutions", partial=False):
        selected_utilities = [checkbox.text() for checkbox in self.utility_checkbox_widget.findChildren(QCheckBox) if checkbox.isChecked()]
        selected_program_types = [checkbox.text() for checkbox in self.program_type_checkbox_widget.findChildren(QCheckBox) if checkbox.isChecked()]
        
        asset_id_filter = self.asset_id_filter.text().lower()
        company_filter = self.company_filter.text().lower()
        facility_filter = self.facility_filter.text().lower()
        curtailment_strategy_filter = self.curtailment_strategy_filter.currentText()

        for item in range(self.treeview.topLevelItemCount()):
            asset_item = self.treeview.topLevelItem(item)
            program = asset_item.text(2).lower()
            enrollment_status = asset_item.text(1).lower()
            asset_id = asset_item.text(3).lower()
            company = asset_item.text(5).lower()
            facility = asset_item.text(6).lower()
            curtailment_strategy = asset_item.text(9)

            # Apply all filters
            if ((partial and filter_text.lower() in program) or (not partial and filter_text.lower() == program)) and \
            asset_id_filter in asset_id and \
            company_filter in company and \
            facility_filter in facility and \
            (curtailment_strategy_filter == "All" or curtailment_strategy_filter == curtailment_strategy):

                if "connected solutions" in program:
                    utility_match = not selected_utilities or any(utility.lower() in program for utility in selected_utilities)
                    program_type_match = not selected_program_types or any(program_type.lower() in program for program_type in selected_program_types)
                else:
                    utility_match = True
                    program_type_match = True

                if (
                    (partial and filter_text.lower() in program) or
                    (not partial and filter_text.lower() == program)
                ) and utility_match and program_type_match:
                    if (self.settle_button.isChecked()) and enrollment_status == "pending enrollment" or enrollment_status == "new" or enrollment_status == "not enrolled":
                        asset_item.setHidden(True)
                    else:
                        asset_item.setHidden(False)

                else:
                    asset_item.setHidden(True)
            else:
                asset_item.setHidden(True)
        
        self.total_count = self.treeview.topLevelItemCount()

        # Resize columns to fit contents
        self.resize_columns_to_fit()

    def populate_curtailment_strategies(self):
        strategies = set()
        for i in range(self.treeview.topLevelItemCount()):
            item = self.treeview.topLevelItem(i)
            strategy = item.text(9)  # Assuming Curtailment Strategy is in column 9
            if strategy:
                strategies.add(strategy)
        
        self.curtailment_strategy_filter.clear()
        self.curtailment_strategy_filter.addItem("All")
        self.curtailment_strategy_filter.addItems(sorted(strategies))

    def apply_filters(self):
        asset_id_filter = self.asset_id_filter.text().lower()
        company_filter = self.company_filter.text().lower()
        facility_filter = self.facility_filter.text().lower()
        curtailment_strategy_filter = self.curtailment_strategy_filter.currentText()

        for i in range(self.treeview.topLevelItemCount()):
            item = self.treeview.topLevelItem(i)
            asset_id = item.text(3).lower()  # Assuming Asset ID is in column 3
            company = item.text(5).lower()  # Assuming Company is in column 5
            facility = item.text(6).lower()  # Assuming Facility is in column 6
            curtailment_strategy = item.text(9)  # Assuming Curtailment Strategy is in column 9

            asset_id_match = asset_id_filter in asset_id
            company_match = company_filter in company
            facility_match = facility_filter in facility
            strategy_match = curtailment_strategy_filter == "All" or curtailment_strategy_filter == curtailment_strategy

            item.setHidden(not (asset_id_match and company_match and facility_match and strategy_match))

        self.update_selected_count()

    def set_columns_for_settle_treeview(self):
        self.treeview.clear()  # Clear existing data
        self.season = "Summer " + self.summer_year_combo_settle_cs.currentText()

        # Determine which type of events to show based on the filter
        is_daily = self.daily_button.isChecked()
        is_targeted = self.targeted_button.isChecked()

        if is_daily:
            self.events = self.cs_daily_dispatch_events['Event Date'].drop_duplicates().tolist()
            self.event_type = "Daily"
        elif is_targeted:
            self.events = self.cs_targeted_dispatch_events['Event Date'].drop_duplicates().tolist()
            self.event_type = "Targeted"
        else:
            self.events = []
            self.event_type = ""
            
        # Create header labels
        header_labels = ["", "Status", "Program", "Asset ID", "ACCPRG", "Company",
                        "Facility Name", "Utility Account Number", "Asset Type",
                        "Curtailment Strategy", "Data Source", "Meter Data Tags", "Contracted kW", "Forecasted kW"]

        # Add event columns
        for event_date in self.events:
            header_labels.append(f"{event_date}")
        header_labels.append(f"Seasonal Average")

        self.treeview.setColumnCount(len(header_labels))
        self.treeview.setHeaderLabels(header_labels)

        self.program_tab_name_mapping = {
            "Cape Light Compact Program - Targeted Dispatch": "CLT",
            "Efficiency Maine-Demand Response Initiative": "EMT",
            "Eversource-Connected Solutions-Daily Dispatch": "EVD",
            "EVERSOURCE-Connected Solutions-Targeted Dispatch": "EVT",
            "Liberty-Connected Solutions-Targeted Dispatch": "LBT",
            "NGRID-Connected Solutions-Daily Dispatch": "NGD",
            "NGRID-Connected Solutions-Targeted Dispatch": "NGT",
            "Rhode Island Connected Solutions-Daily Dispatch": "RID",
            "Rhode Island Connected Solutions-Targeted Dispatch": "RIT",
            "Unitil-Connected Solutions-Daily Dispatch": "UND",
            "UNITIL-Connected Solutions-Targeted Dispatch": "UNT",
        }

        def add_performance_data(row):
            cpower_performances = []
            for event_date in self.events:
                cpower_performance = row.get(f"{event_date}", "Pending")
                if isinstance(cpower_performance, pd.Series):
                    cpower_performance = cpower_performance.iloc[0] if not cpower_performance.empty else np.nan
                    if pd.isna(cpower_performance):
                        item_data.append("Pending")
                    elif cpower_performance != "Pending":
                        try:
                            cpower_performances.append(float(cpower_performance))
                            item_data.append(format_performance(cpower_performance))
                        except ValueError:
                            item_data.append("Pending")

                elif cpower_performance != "Pending":
                    try:
                        cpower_performances.append(float(cpower_performance))
                        item_data.append(format_performance(cpower_performance))
                    except ValueError:
                        item_data.append("Pending")
                else:
                    item_data.append("Pending")

            cpower_average = row.get(f"Seasonal Average", "Pending")

            if isinstance(cpower_average, pd.Series):
                cpower_average = str(cpower_average.iloc[0]) if not cpower_average.empty else np.nan

            if pd.isna(cpower_average):
                if cpower_performances:
                    cpower_average = str(sum(cpower_performances) / len(cpower_performances))
                else:
                    cpower_average = "Pending"  

            item_data.append(format_performance(cpower_average))

        def format_performance(value):
            if isinstance(value, pd.Series):
                if value.empty:
                    return "Pending"
                value = value.iloc[0]
            
            try:
                float_value = float(value)
                if float_value == 0:
                    return "-"
                return f"{float_value:.1f}"
            except ValueError:
                return value
            
        self.cs_settle_tabs_callable = lambda: self.cs_settle_tabs

        for df_label in self.cs_settle_tabs_callable():
            for _, row in self.cs_enroll_dfs[df_label].iterrows():
                enrollment_status = "Enrolled"
                meter_data_tag = str(row.get("Meter Data Tags", "Meter Data Tags"))
                accprg = str(row.get("ACCPRG", "ACCPRG"))
                asset_id = str(row.get("Asset ID", "Asset ID"))
                contracted_kw = str(row.get("Contracted KW", "Contracted KW"))
                program_name = str(row.get("Program", "Program"))
                company_name = str(row.get("Customer", "Customer"))
                facility_name = str(row.get("Customer Asset", "Customer Asset"))
                utility_account_number = str(row.get("Utility Account Number", "Utility Account Number"))
                asset_type = str(row.get("Asset Type", "Asset Type"))
                curtailment_strategy = str(row.get("Curtailment Strategy", "Curtailment Strategy"))
                data_source = str(row.get("Data Source", "Data Source"))
                forecasted_kw = str(row.get("Forecasted KW", "Forecasted kW"))

                matching_settlement_row = self.cs_settle_dfs[df_label][self.cs_settle_dfs[df_label]['ACCPRG'] == accprg].copy()

                item_data = ["", enrollment_status, program_name, asset_id, accprg, company_name,
                            facility_name, utility_account_number, asset_type,
                            curtailment_strategy, data_source, meter_data_tag, contracted_kw, forecasted_kw]

                # Add performance data
                add_performance_data(matching_settlement_row)

                item = QTreeWidgetItem(item_data)
                
                # Set the data for sorting
                for col in range(len(item_data)):
                    try:
                        item.setData(col, Qt.UserRole, float(item_data[col]))
                    except ValueError:
                        item.setData(col, Qt.UserRole, item_data[col])

                self.treeview.addTopLevelItem(item)
                
                item.setFlags(item.flags() | QtCore.Qt.ItemIsEditable)
                item.setCheckState(0, QtCore.Qt.Unchecked)
                self.treeview.addTopLevelItem(item)
            
            delegate = PerformanceDelegate(self.treeview)
            for col in range(14, self.treeview.columnCount()): 
                self.treeview.setItemDelegateForColumn(col, delegate)


        # After populating the treeview
        self.populate_curtailment_strategies()

    # endregion

class WorkerSignals(QObject):
    finished = pyqtSignal()
    error = pyqtSignal(tuple)
    result = pyqtSignal(object)

class Worker(QRunnable):
    def __init__(self, fn, *args, **kwargs):
        super(Worker, self).__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

    @pyqtSlot()
    def run(self):
        try:
            result = self.fn(*self.args, **self.kwargs)
        except:
            traceback.print_exc()
            exctype, value = sys.exc_info()[:2]
            self.signals.error.emit((exctype, value, traceback.format_exc()))
        else:
            self.signals.result.emit(result)
        finally:
            self.signals.finished.emit()

class CircularItem(QGraphicsEllipseItem):
    def __init__(self, x, y, w, h, color, tooltip):
        super().__init__(x, y, w, h)
        self.setBrush(QBrush(color))
        self.setPen(QPen(Qt.NoPen))
        self.setToolTip(tooltip)

class WrappedTextItem(QGraphicsTextItem):
    def __init__(self, text, max_width):
        super().__init__(text)
        self.max_width = max_width

    def paint(self, painter, option, widget=None):
        metrics = QFontMetrics(self.font())
        lines = []
        for line in self.toPlainText().split('\n'):
            if metrics.width(line) > self.max_width:
                words = line.split()
                new_line = words[0]
                for word in words[1:]:
                    if metrics.width(new_line + " " + word) <= self.max_width:
                        new_line += " " + word
                    else:
                        lines.append(new_line)
                        new_line = word
                lines.append(new_line)
            else:
                lines.append(line)
        self.setPlainText('\n'.join(lines))
        super().paint(painter, option, widget)

class PerformanceDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        value = index.data(Qt.DisplayRole)
        option.palette.setColor(option.palette.Text, QColor('black'))
        if isinstance(value, pd.Series):
            value = value.iloc[0] if not value.empty else "Pending"
        
        # Center align the text
        option.displayAlignment = Qt.AlignCenter

        font = QFont(option.font)
        font.setBold(False)
        option.font = font
    
        # Make text bold only if it's not "Pending" or "-"
        if value not in ["Pending", "-"]:
            font.setBold(True)
            option.font = font
        
        if value != "Pending" and value != "-":
            try:
                float_value = float(value)
                if float_value > 0:
                    option.palette.setColor(option.palette.Text, QColor(color_palette[10]))
                elif float_value < 0: 
                    option.palette.setColor(option.palette.Text, QColor(color_palette[11]))
            except ValueError:
                pass
        
        QStyledItemDelegate.paint(self, painter, option, index)

    def initStyleOption(self, option, index):
        super().initStyleOption(option, index)
        option.displayAlignment = Qt.AlignCenter

    def createEditor(self, parent, option, index):
        editor = super().createEditor(parent, option, index)
        if isinstance(editor, QLineEdit):
            editor.setAlignment(Qt.AlignCenter)
        return editor
            
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(stylesheet) 
    window = Watts()
    window.show()

    sys.exit(app.exec_())
