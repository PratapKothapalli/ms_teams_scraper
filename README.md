# Microsoft Teams Chat Scraper

## Overview

The Microsoft Teams Chat Scraper is a powerful Python tool for automated extraction of chat histories from Microsoft Teams. The tool uses Selenium with the Microsoft Edge WebDriver to log into Teams, navigate through the chat list, and extract messages. It handles the challenges of virtual DOM loading in Teams and provides robust features for message extraction, deduplication, and optional image/attachment downloads.

The project also includes a web-based visualization component that allows you to browse and view the extracted chats in a user-friendly interface. This makes it easy to review conversations, search for specific content, and view images directly in your browser.

### Key Features

- **Automated Chat Extraction**: Navigates through the Teams interface and extracts chat histories
- **Virtual Scrolling Support**: Handles Teams' dynamic loading of messages through intelligent scrolling strategies
- **Image and Attachment Extraction**: Downloads images and captures information about attachments from messages
- **Chat Selection**: Allows selection of specific chats for extraction
- **Search Functionality**: Searches for specific chats using search terms
- **Automatic Edge Driver Management**: Detects and manages the Microsoft Edge WebDriver
- **Data Export**: Saves extracted data in JSON and CSV formats
- **Web-based Visualization**: Browse and view extracted chats in a user-friendly web interface

## Requirements

- Python 3.6 or higher
- Microsoft Edge browser
- Internet connection
- Microsoft Teams account with SSO login

### Dependencies

The project requires the following Python packages:
- selenium (for web scraping)
- requests (for HTTP requests)
- msedgedriver (optional, will be automatically attempted to install)
- flask (for the visualization web server)
- markupsafe (for safe HTML rendering in the visualization)

## Installation

1. Ensure Python 3.6+ is installed
2. Install the required packages:

```bash
pip install -r requirements.txt
```

3. Ensure Microsoft Edge is installed

## Usage

### Basic Usage

```bash
python teams_chat_scraper.py
```

### Command Line Arguments

The script supports the following command line arguments:

| Argument | Description |
|----------|-------------|
| `--output-dir DIRECTORY` | Output directory for extracted data (default: "teams_complete_export") |
| `--headless` | Run browser in headless mode (without GUI) |
| `--no-images` | Don't download images |
| `--auto-select-all` | Automatically select all chats (no user prompt) |

### Examples

Extract all chats with default settings:
```bash
python teams_chat_scraper.py
```

Extract chats in headless mode:
```bash
python teams_chat_scraper.py --headless
```

Extract chats without downloading images:
```bash
python teams_chat_scraper.py --no-images
```

Change output directory:
```bash
python teams_chat_scraper.py --output-dir my_teams_export
```

Automatically select all chats:
```bash
python teams_chat_scraper.py --auto-select-all
```

## How It Works

### Login Process

The script opens a Microsoft Edge browser and navigates to the Teams web interface. It then waits for SSO login by the user. After successful login, it navigates to the chat area.

### Chat Selection

After loading the chat list, the script offers two options:
1. **Direct Selection**: Choose chats from the displayed list
2. **Search**: Search for specific chats using search terms

For direct selection, you can:
- Select individual chats (e.g., "1,3,5")
- Select ranges (e.g., "5-7")
- Combined selection (e.g., "1,3,5-7")
- Select all chats ("all")

### Message Extraction

For each selected chat:
1. The script clicks on the chat to open it
2. It scrolls up to load older messages
3. It extracts messages, including:
   - Message text
   - Author
   - Timestamp
   - Images (optional)
   - Attachment information

### Deduplication

The script uses a hash-based mechanism to detect and remove duplicate messages that may arise from virtual scrolling.

### Image Download

When enabled, the script downloads images from messages and saves them in the "images" subdirectory. It supports various image sources:
- HTTP/HTTPS URLs
- Blob URLs
- Base64-encoded images

## Output Formats

### For Each Chat

- **JSON file**: `[chat_name]_[timestamp].json` - Contains complete message data
- **CSV file**: `[chat_name]_[timestamp].csv` - Contains basic message information

### Summary Files

- **Combined JSON file**: `teams_export_[timestamp].json` - All extracted messages
- **Combined CSV file**: `teams_export_[timestamp].csv` - All messages in CSV format
- **Image summary**: `image_summary_[timestamp].json` - Statistics about downloaded images

### Image Directory

Downloaded images are stored in the `images` subdirectory, with filenames in the format `ChatName_ImageHash.ext` to easily associate images with their respective chats.

## Chat Visualization

The project includes a web-based visualization tool for the extracted chat data, making it easy to browse and view the conversations in a user-friendly interface.

### Key Features

- **Web Interface**: Browser-based visualization of chat histories
- **Chat List Sidebar**: Easy navigation between different chats
- **Real-time Filtering**: Filter chats by name using the search box
- **Responsive Design**: Clean, modern interface with corporate design elements
- **Link Detection**: Automatic conversion of URLs to clickable links
- **Complete Message Display**: Shows all messages from the chat history without filtering


### Usage

To start the visualization server:

```bash
python visualize_chat.py
```

Then open a web browser and navigate to:
```
http://localhost:5000
```

### Interface Components

- **Sidebar**: Lists all available chats with real-time filtering capability
- **Main Content Area**: Displays the selected chat's messages with author, timestamp, content, and images
- **Message Display**: Shows messages in chronological order with clear author attribution
- **Image Rendering**: Displays images inline with the messages

### Technical Implementation

The visualization tool is built with:
- **Flask**: Lightweight web framework for serving the application
- **HTML/CSS**: Responsive design with a corporate color scheme
- **JavaScript**: Real-time filtering and keyboard shortcuts
- **Jinja2 Templates**: Dynamic content rendering

## Main Components

### TeamsCollector Class

The main class that controls the entire extraction process:

- **Initialization**: Configures browser and settings
- **Edge Driver Management**: Detects and manages the Microsoft Edge WebDriver
- **Navigation**: Controls navigation through the Teams interface
- **Chat Processing**: Extracts and processes chat messages
- **Image Processing**: Downloads and saves images from messages
- **Data Export**: Saves extracted data in various formats

### Scrolling Strategies

The script uses multiple strategies to load older messages:
- Scrolling up
- Clicking on "Show more" buttons
- Waiting for content to load

### Message Deduplication

Uses a hash-based approach to identify duplicate messages:
- Creates a hash from author, timestamp, and message content
- Tracks already seen messages to avoid duplicates

## Known Limitations

- **Screenshots**: Embedded screenshots may not be extracted (known bug)
- **Long Chats**: For very long chats, the scraping process may stop for unknown reasons
- **Teams Updates**: Changes to the Teams interface may affect the selectors
- **Authentication**: Requires manual SSO login (no automatic login)
- **Attachments**: Captures information about attachments but does not download files

## Troubleshooting

### Edge Driver Issues

If you encounter issues with the Edge driver:
1. Ensure Microsoft Edge is installed
2. Try installing msedgedriver manually: `pip install msedgedriver`
3. Check if the Edge version is compatible with the driver version

### Login Issues

If you encounter login issues:
1. Ensure you have valid Teams credentials
2. Try logging into Teams manually before running the script
3. Disable headless mode to observe the login process

### Extraction Issues

If no messages are extracted:
1. Ensure you have access to the selected chats
2. Try disabling headless mode
3. Increase wait times between scrolling operations

### Visualization Issues

If you encounter issues with the visualization tool:

1. **Images not displaying**:
   - Ensure images were downloaded during scraping (don't use `--no-images` flag)
   - Check that the images exist in the `images` directory
   - Verify that image paths in the JSON files are correct

2. **Flask server won't start**:
   - Ensure Flask is installed: `pip install flask`
   - Check if another application is using port 5000
   - Try running with a different port: modify the line `app.run(host='0.0.0.0', port=5000, debug=True)` in `visualize_chat.py`

3. **Empty chat list**:
   - Verify that chat JSON files exist in the output directory
   - Check that the files follow the expected format
   - Ensure the output directory path is correct in `visualize_chat.py`

## License

[Insert license information here, if available]
