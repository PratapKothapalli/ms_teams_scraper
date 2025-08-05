import os
import json
import glob
import re
from markupsafe import Markup
from flask import Flask, render_template_string, send_from_directory, request, redirect, url_for


app = Flask(__name__)

BASE_DIR = os.path.dirname(__file__)
EXPORT_DIR = os.path.join(BASE_DIR, 'teams_complete_export')
IMAGES_DIR = os.path.join(EXPORT_DIR, 'images')

# No helper functions needed as we're showing all messages

# Function to load all available chat files
def get_all_chats():
    chat_files = glob.glob(os.path.join(EXPORT_DIR, '*.json'))
    chats = []
    
    for file_path in chat_files:
        # Ignore image_summary files
        if 'image_summary' in file_path:
            continue
            
        filename = os.path.basename(file_path)
        chat_name = os.path.splitext(filename)[0]
        chats.append({
            'id': filename,
            'name': chat_name
        })
    
    # Sort by name
    chats.sort(key=lambda x: x['name'])
    return chats

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Teams Chat Visualization</title>
    <style>
        :root {
            /* Aumovio-inspired color scheme */
            --primary-color: #ff6b00;         /* Orange as main color */
            --primary-dark: #e05a00;          /* Darker orange for hover */
            --secondary-color: #333333;       /* Dark gray for text */
            --background-light: #f7f7f7;      /* Light gray for background */
            --background-white: #ffffff;      /* White for cards */
            --accent-light: #ffe0cc;          /* Light orange for accents */
            --text-light: #ffffff;            /* White for text on dark background */
            --text-dark: #333333;             /* Dark gray for text on light background */
            --border-color: #e0e0e0;          /* Light gray for borders */
            --hover-color: #fff0e6;           /* Very light orange for hover effects */
            --active-color: #ffcca3;          /* Medium orange for active elements */
        }
        
        body { 
            font-family: Arial, sans-serif; 
            background: var(--background-light); 
            margin: 0; 
            padding: 0;
            display: flex;
            height: 100vh;
            color: var(--text-dark);
        }
        
        .sidebar {
            width: 300px;
            background: var(--secondary-color);
            color: var(--text-light);
            overflow-y: auto;
            padding: 10px;
            box-sizing: border-box;
            display: flex;
            flex-direction: column;
            box-shadow: 2px 0 5px rgba(0,0,0,0.1);
        }
        
        .sidebar h2 {
            margin-top: 0;
            padding: 10px;
            border-bottom: 1px solid rgba(255,255,255,0.2);
            color: var(--text-light);
        }
        
        .filter-container {
            padding: 10px;
            margin-bottom: 10px;
            background: rgba(255,255,255,0.1);
            border-radius: 4px;
        }
        
        .filter-input {
            width: 100%;
            padding: 8px;
            border: none;
            border-radius: 4px;
            background: rgba(255,255,255,0.2);
            color: var(--text-light);
            box-sizing: border-box;
        }
        
        .filter-input::placeholder {
            color: rgba(255,255,255,0.7);
        }
        
        .chat-list-container {
            flex: 1;
            overflow-y: auto;
        }
        
        .chat-list {
            list-style: none;
            padding: 0;
            margin: 0;
        }
        
        .chat-list li {
            padding: 10px 15px;
            border-radius: 4px;
            margin-bottom: 5px;
            cursor: pointer;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            transition: background-color 0.2s ease;
        }
        
        .chat-list li:hover {
            background: rgba(255,255,255,0.1);
        }
        
        .chat-list li.active {
            background: var(--primary-color);
            font-weight: bold;
        }
        
        .chat-list a {
            color: var(--text-light);
            text-decoration: none;
            display: block;
        }
        
        .main-content {
            flex: 1;
            overflow-y: auto;
            padding: 20px;
            box-sizing: border-box;
            background: var(--background-light);
        }
        
        .chat-container { 
            max-width: 900px; 
            margin: 0 auto; 
            background: var(--background-white); 
            border-radius: 8px; 
            box-shadow: 0 2px 8px rgba(0,0,0,0.1); 
            padding: 24px; 
        }
        
        .chat-container h2 {
            color: var(--primary-color);
            margin-top: 0;
            padding-bottom: 15px;
            border-bottom: 1px solid var(--border-color);
        }
        
        .message { 
            margin-bottom: 24px; 
            padding-bottom: 12px; 
            border-bottom: 1px solid var(--border-color); 
        }
        
        .author { 
            font-weight: bold; 
            color: var(--primary-color); 
        }
        
        .timestamp { 
            color: #888; 
            font-size: 0.9em; 
            margin-left: 8px; 
        }
        
        .content { 
            margin: 8px 0; 
            white-space: pre-line; 
        }
        
        .images { 
            margin: 8px 0; 
        }
        
        .images img { 
            max-height: 200px; 
            margin-right: 6px; 
            vertical-align: middle; 
            border-radius: 4px;
            border: 1px solid var(--border-color);
        }
        
        .attachments { 
            margin: 8px 0; 
        }
        
        .attachment { 
            display: inline-block; 
            background: var(--accent-light); 
            color: var(--primary-dark);
            padding: 4px 8px; 
            border-radius: 4px; 
            margin-right: 6px; 
            font-size: 0.95em; 
        }
        
        .welcome-message {
            text-align: center;
            padding: 50px 20px;
            color: var(--text-dark);
        }
        
        .welcome-message h2 {
            margin-bottom: 20px;
            color: var(--primary-color);
        }
        
        .hidden {
            display: none;
        }
        
        /* Logo und Header */
        .app-header {
            display: flex;
            align-items: center;
            padding: 10px;
            background: var(--secondary-color);
            border-bottom: 2px solid var(--primary-color);
        }
        
        .app-logo {
            height: 30px;
            margin-right: 10px;
        }
        
        .app-title {
            color: var(--text-light);
            margin: 0;
            font-size: 1.2em;
        }
    </style>
</head>
<body>
    <div class="sidebar">
        <div class="app-header">
            <h2 class="app-title">Chat Visualization</h2>
        </div>
        <div class="filter-container">
            <input type="text" class="filter-input" id="chatFilter" placeholder="Filter chat names..." autocomplete="off">
        </div>
        <div class="chat-list-container">
            <ul class="chat-list" id="chatList">
                {% for chat in all_chats %}
                    <li {% if current_chat_id == chat.id %}class="active"{% endif %} data-chat-name="{{ chat.name.lower() }}">
                        <a href="{{ url_for('show_specific_chat', chat_id=chat.id) }}">{{ chat.name }}</a>
                    </li>
                {% endfor %}
            </ul>
        </div>
    </div>
    
    <div class="main-content">
        {% if messages %}
            <div class="chat-container">
                <h2>{{ chat_name }}</h2>
                {% for msg in messages %}
                <div class="message">
                    <span class="author">{{ msg.author }}</span>
                    <span class="timestamp">{{ msg.timestamp }}</span>
                    <div class="content">{{ msg.content }}</div>
                    {% if msg.images %}
                    <div class="images">
                        {% for img in msg.images %}
                            {% if img.src.startswith('http') %}
                                <img src="{{ img.src }}" alt="{{ img.alt }}" title="{{ img.title }}" width="{{ img.width }}" height="{{ img.height }}">
                            {% elif img.local_path %}
                                <img src="/images/{{ img.local_path }}" alt="{{ img.alt }}" title="{{ img.title }}" width="{{ img.width }}" height="{{ img.height }}">
                            {% endif %}
                        {% endfor %}
                    </div>
                    {% endif %}
                    {% if msg.attachments %}
                    <div class="attachments">
                        {% for att in msg.attachments %}
                            <span class="attachment">{{ att }}</span>
                        {% endfor %}
                    </div>
                    {% endif %}
                </div>
                {% endfor %}
            </div>
        {% else %}
            <div class="welcome-message">
                <h2>Teams Chat Visualization</h2>
                <p>Please select a chat from the list on the left to view messages.</p>
            </div>
        {% endif %}
    </div>
    
    <script>
        document.addEventListener('DOMContentLoaded', function() {
            const filterInput = document.getElementById('chatFilter');
            const chatItems = document.querySelectorAll('#chatList li');
            
            filterInput.addEventListener('input', function() {
                const filterText = this.value.toLowerCase().trim();
                
                chatItems.forEach(function(item) {
                    const chatName = item.getAttribute('data-chat-name');
                    if (chatName.includes(filterText)) {
                        item.classList.remove('hidden');
                    } else {
                        item.classList.add('hidden');
                    }
                });
            });
            
            // Set focus on the search field when Ctrl+F is pressed
            document.addEventListener('keydown', function(e) {
                if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
                    e.preventDefault();
                    filterInput.focus();
                }
            });
        });
    </script>
</body>
</html>
'''

# Route for local images
@app.route('/images/<filename>')
def serve_image(filename):
    return send_from_directory(IMAGES_DIR, filename)

# Main page - shows the chat list and a welcome screen
@app.route('/')
def index():
    all_chats = get_all_chats()
    return render_template_string(HTML_TEMPLATE, all_chats=all_chats, messages=None, chat_name=None, current_chat_id=None)

# Shows a specific chat
@app.route('/chat/<chat_id>')
def show_specific_chat(chat_id):
    chat_file = os.path.join(EXPORT_DIR, chat_id)
    all_chats = get_all_chats()
    
    try:
        with open(chat_file, encoding='utf-8') as f:
            chat_data = json.load(f)
    except Exception as e:
        return f"Error loading chat file: {e}", 500
    
    if not chat_data:
        return "No messages found.", 404
    
    chat_name = chat_data[0].get('chat_name', os.path.splitext(chat_id)[0]) if chat_data else os.path.splitext(chat_id)[0]
    
    # Process all messages
    url_pattern = re.compile(r'(https?://[\w\-\.\?&=/#%]+)')
    for msg in chat_data:
        # Adjust image paths
        for img in msg.get('images', []):
            if img['local_path']:
                img['local_path'] = os.path.basename(img['local_path'])
        # Show links in content as HTML links
        if msg.get('content'):
            def repl(m):
                url = m.group(1)
                return f'<a href="{url}" target="_blank">{url}</a>'
            msg['content'] = Markup(url_pattern.sub(repl, msg['content']))
    
    return render_template_string(HTML_TEMPLATE, all_chats=all_chats, messages=chat_data, chat_name=chat_name, current_chat_id=chat_id)

if __name__ == '__main__':
    # Only bind to localhost for security
    app.run(host='127.0.0.1', port=5000, debug=False)
