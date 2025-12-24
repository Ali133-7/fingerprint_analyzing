# styles.py

# Define a professional color palette
class ColorPalette:
    Primary = "#4CAF50"   # Green
    Secondary = "#2196F3" # Blue
    Accent = "#FFC107"    # Amber
    Background = "#F5F5F5" # Light Gray
    Surface = "#FFFFFF"   # White
    TextPrimary = "#212121" # Dark Gray
    TextSecondary = "#757575" # Medium Gray
    Success = "#4CAF50"   # Green
    Warning = "#FF9800"   # Orange
    Danger = "#F44336"    # Red
    Info = "#03A9F4"      # Light Blue

# Define common QSS (Qt Style Sheet) snippets
class StyleSheets:
    # Base button style
    ButtonBase = """
        QPushButton {
            background-color: %s;
            color: %s;
            border-radius: 8px;
            padding: 10px 20px;
            font-size: 14px;
            font-weight: bold;
            border: none;
        }
        QPushButton:hover {
            background-color: %s;
        }
        QPushButton:pressed {
            background-color: %s;
        }
    """ % (ColorPalette.Primary, ColorPalette.Surface, "#43A047", "#388E3C") # Darker shades for hover/pressed

    # Primary button style (e.g., Generate Report, Import & Analyze)
    PrimaryButton = """
        QPushButton {
            background-color: %s;
            color: %s;
            border-radius: 8px;
            padding: 10px 20px;
            font-size: 14px;
            font-weight: bold;
            border: none;
        }
        QPushButton:hover {
            background-color: #43A047; /* Darker green */
        }
        QPushButton:pressed {
            background-color: #388E3C; /* Even darker green */
        }
    """ % (ColorPalette.Primary, ColorPalette.Surface)

    # Secondary button style (e.g., Choose File)
    SecondaryButton = """
        QPushButton {
            background-color: %s;
            color: %s;
            border-radius: 8px;
            padding: 8px 15px;
            font-size: 13px;
            font-weight: normal;
            border: 1px solid #BBBBBB;
        }
        QPushButton:hover {
            background-color: #E0E0E0; /* Lighter gray */
        }
        QPushButton:pressed {
            background-color: #BDBDBD; /* Darker gray */
        }
    """ % (ColorPalette.Surface, ColorPalette.TextPrimary)

    # Export button style (e.g., Export Report)
    ExportButton = """
        QPushButton {
            background-color: %s;
            color: %s;
            border-radius: 8px;
            padding: 10px 20px;
            font-size: 14px;
            font-weight: bold;
            border: none;
        }
        QPushButton:hover {
            background-color: #2693F0; /* Darker blue */
        }
        QPushButton:pressed {
            background-color: #1A7BBF; /* Even darker blue */
        }
    """ % (ColorPalette.Secondary, ColorPalette.Surface)

    # Title Label Style
    TitleLabel = """
        QLabel {
            font-size: 24px;
            font-weight: bold;
            color: %s;
            padding: 10px;
        }
    """ % ColorPalette.TextPrimary

    # GroupBox Title Style
    GroupBoxTitle = """
        QGroupBox::title {
            color: %s;
            subcontrol-origin: margin;
            subcontrol-position: top left;
            padding: 0 3px;
            font-weight: bold;
        }
        QGroupBox {
            border: 1px solid %s;
            border-radius: 5px;
            margin-top: 1ex;
        }
    """ % (ColorPalette.TextPrimary, ColorPalette.TextSecondary)

    # Table Header Style
    TableHeader = """
        QHeaderView::section {
            background-color: %s;
            color: %s;
            padding: 5px;
            border: 1px solid %s;
            font-weight: bold;
        }
    """ % (ColorPalette.Primary, ColorPalette.Surface, "#E0E0E0")

    # Table View Style
    TableView = """
        QTableWidget {
            gridline-color: #E0E0E0;
            alternate-background-color: %s;
            background-color: %s;
            border: 1px solid #CCCCCC;
            selection-background-color: #C0DFFD; /* Light blue selection */
        }
        QTableWidget::item {
            padding: 5px;
        }
    """ % (ColorPalette.Background, ColorPalette.Surface)

    # Tab Widget Style
    TabWidget = """
        QTabWidget::pane { /* The tab widget frame */
            border: 1px solid %s;
            background: %s;
        }
        QTabWidget::tab-bar {
            left: 5px; /* move to the right by 5px */
        }
        QTabBar::tab {
            background: %s;
            border: 1px solid %s;
            border-bottom-color: %s; /* same as pane color */
            border-top-left-radius: 4px;
            border-top-right-radius: 4px;
            min-width: 8ex;
            padding: 5px 10px;
            margin-right: 1px;
            color: %s;
        }
        QTabBar::tab:selected {
            background: %s;
            border-color: %s;
            border-bottom-color: %s; /* same as pane color */
            color: %s;
            font-weight: bold;
        }
        QTabBar::tab:hover {
            background: %s;
        }
    """ % (ColorPalette.TextSecondary, ColorPalette.Background, # Pane and background
           ColorPalette.Surface, ColorPalette.TextSecondary, ColorPalette.Background, ColorPalette.TextPrimary, # Tab base
           ColorPalette.Surface, ColorPalette.TextSecondary, ColorPalette.Surface, ColorPalette.TextPrimary, # Tab selected
           ColorPalette.Background # Tab hover
          )

