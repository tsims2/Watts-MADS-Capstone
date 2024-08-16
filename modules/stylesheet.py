stylesheet = """
QWidget {
    background-color: #F0F0F0;
    font-family: 'Public Sans';
    color: #333333;
}

QPushButton {
    background-color: #D3D3D3;
    color: #333333;
    border: none;
    padding: 4px;
    border-radius: 5px;
    font-family: 'Public Sans';
}

QPushButton:hover {
    background-color: #00C9A2;
}

QPushButton:pressed {
    background-color: #00C9A2;
}

/* Settle button: navy with yellow text */
QPushButton[text="Settle"] {
    background-color: #020244;  /* Navy */
    color: #FFC907;  /* Yellow */
}

/* Connected Solutions button: remains orange */
QPushButton[text="Connected Solutions"] {
    background-color: #FF5400;
    color: #FFFFFF;
}

/* Data Table and Data Visuals buttons: navy */
QPushButton[objectName="toggle_treeview_button"], QPushButton[objectName="toggle_data_viewer_button"] {
    background-color: #020244;  /* Navy */
    color: #FFFFFF;
}

QPushButton[objectName="toggle_treeview_button"]:checked, QPushButton[objectName="toggle_data_viewer_button"]:checked {
    background-color: #020244;  /* Navy */
    color: #FFFFFF;
}

/* Data view button stays green when clicked */
QPushButton[objectName="toggle_data_viewer_button"]:checked {
    background-color: #00C9A2;
}

QPushButton[objectName="toggle_treeview_button"]:checked, QPushButton[objectName="toggle_data_viewer_button"]:checked {
    background-color: #020244;
    color: #FFFFFF;
}

/* Data view button stays green when clicked */
QPushButton[objectName="toggle_data_viewer_button"]:checked {
    background-color: #00C9A2;
}

QCheckBox {
    border: none;
    font-family: 'Public Sans';
}

QCheckBox::indicator {
    border: 2px solid #FF8500;
    width: 16px;
    height: 16px;
}

QCheckBox::indicator:checked {
    background-color: #FF8500;
}

QComboBox {
    background-color: #FFFFFF;
    color: #333333;
    border: none;
    padding: 2px;
    border-radius: 5px;
    font-family: 'Public Sans';
}

QComboBox::drop-down {
    border: none;
}

QComboBox QAbstractItemView {
    background-color: #FFFFFF;
    color: #333333;
}

/* Workstation buttons */
QPushButton[text="View"], QPushButton[text="Enroll"], QPushButton[text="Pay"] {
    background-color: #FFFFFF;
    color: #333333;
}

QPushButton[text="View"]:checked, QPushButton[text="Enroll"]:checked, QPushButton[text="Pay"]:checked {
    background-color: #020244;
    color: #FFC907;
    font-weight: bold;
}

QPushButton[text="View"]:hover, QPushButton[text="Enroll"]:hover, QPushButton[text="Pay"]:hover {
    background-color: #020244;
}

/* Program filter buttons */
QPushButton[text="ADCR"], QPushButton[text="On Peak"], QPushButton[text="Clean Peak"], QPushButton[text="Daily"], QPushButton[text="Targeted"] {
    background-color: #FFFFFF;
    color: #333333;
}

QPushButton[text="ADCR"]:checked, QPushButton[text="On Peak"]:checked, QPushButton[text="Clean Peak"]:checked, QPushButton[text="Daily"]:checked, QPushButton[text="Targeted"]:checked {
    background-color: #FF5400;
    color: #FFFFFF;
    font-weight: bold;
}

QPushButton[text="ADCR"]:hover, QPushButton[text="On Peak"]:hover, QPushButton[text="Clean Peak"]:hover, QPushButton[text="Daily"]:hover, QPushButton[text="Targeted"]:hover {
    background-color: #FF5400;
}

/* Table view buttons */
QPushButton[text="Unify"], QPushButton[text="Save"], QPushButton[text="Download"] {
    background-color: #020244;
    color: #FFFFFF;
}

QPushButton[text="Unify"]:hover, QPushButton[text="Save"]:hover, QPushButton[text="Download"]:hover {
    background-color: #00C9A2;
    color: #FFFFFF;
}

/* Chat send button */
QPushButton[text="Send"] {
    background-color: #072B60;
    color: #FFFFFF;
}

QPushButton[text="Send"]:hover {
    background-color: #00C9A2;
}

QScrollBar:horizontal, QScrollBar:vertical {
    background: #E0E0E0;
    border: none;
    border-radius: 5px;
}

QScrollBar::handle:horizontal, QScrollBar::handle:vertical {
    background: #020244;
    border-radius: 5px;
}

QScrollBar::handle:horizontal {
    min-width: 20px;
}

QScrollBar::handle:vertical {
    min-height: 20px;
}

QScrollBar::add-line, QScrollBar::sub-line {
    border: none;
    background: none;
}

QTreeWidget {
    background-color: #FFFFFF;
    color: #333333;
    border: none;
    border-radius: 10px;
    font-family: 'Public Sans';
}

QHeaderView::section {
    background-color: #020244;
    color: #FFC907;
    font-weight: bold;
    padding: 2px;
    border: none;
}

QLineEdit, QTextEdit, QPlainTextEdit, QSpinBox, QDoubleSpinBox, QDateEdit, QTimeEdit, QDateTimeEdit {
    background-color: #FFFFFF;
    color: #333333;
    border: none;
    border-radius: 5px;
    padding: 2px;
    font-family: 'Public Sans';
}

QLineEdit:hover, QTextEdit:hover, QPlainTextEdit:hover, QSpinBox:hover, QDoubleSpinBox:hover, QDateEdit:hover, QTimeEdit:hover, QDateTimeEdit:hover, QComboBox:hover {
    background-color: #FFFFFF;
    color: #333333;
}

QLabel, QComboBox, QCheckBox {
    font-size: 12px;
    color: #333333;
    font-family: 'Public Sans';
}

/* Chat window */
QTextEdit#chat_display {
    background-color: #E0F2F7;
    font-family: 'Public Sans';
}

/* Sent message */
QTextEdit#chat_display QLabel[text^="You"] {
    background-color: #FF8500;
    color: #FFFFFF;
    padding: 5px;
    border-radius: 5px;
}

/* Received message */
QTextEdit#chat_display QLabel[text^="Watts"] {
    background-color: #59C8E3;
    color: #FFFFFF;
    padding: 5px;
    border-radius: 5px;
}

/* Selected row in QTreeWidget */
QTreeWidget::item:selected {
    background-color: #59C8E3;
    color: #333333;
}

/* Settlement Period Label */
QLabel#enroll_period_label, QLabel#settlement_period_label, QLabel#pay_label {
    font-size: 20px;
    font-weight: bold;
    color: #020244;
    font-family: 'Public Sans';
}

/* Section labels */
QLabel[text="Event Performance Calculation"], QLabel[text="Report Generation"], QLabel[text="Data Transfers (External)"], QLabel[text="Data Downloads (Internal)"],
QLabel[text="Data Retrieval"], QLabel[text="Submit Data"], QLabel[text="Customer Outreach"], QLabel[text="Review Data"], 
QLabel[text="Issuances"], QLabel[text="Sales Outreach"], QLabel[text="Statements"], QLabel[text="Payment Files"], QLabel[text="Finance Reports"], 
QLabel[text="Data"]{
    font-size: 16px;
    font-weight: bold;
    font-family: 'Public Sans';
    color: #363C49;
    margin-top: 20px;
    margin-bottom: 5px;
}

/* Navigation Bar */
QWidget#search_layout {
    background-color: #FFFFFF;
    border-radius: 10px;
}

/* Data Viewer */
QWidget#data_viewer_widget {
    background-color: #FFFFFF;
    border-radius: 10px;
}

/* Web Browser */
QWebEngineView {
    border-radius: 10px;
}
"""