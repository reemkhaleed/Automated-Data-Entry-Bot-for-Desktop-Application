# Notepad Automator

This project automates data entry into **Notepad** using Python and the `pywinauto` library.  
It fetches posts from the [JSONPlaceholder API](https://jsonplaceholder.typicode.com/) and saves them as `.txt` files.

## 📌 Features
- Launches Notepad automatically.
- Fetches the first **10 posts** from the JSONPlaceholder API.
- Types each post (title + body) into Notepad.
- Saves each post in a `tjm-project` folder on the Desktop.
- Handles errors (fallbacks to saving directly as a `.txt` file if automation fails).
- Can be packaged into a **standalone executable (.exe)** for Windows.

## 📂 Output Example
On your Desktop, a folder named `tjm-project` will be created containing:


## ⚙️ Setup Instructions

### 1. Clone the repository
```bash
git clone https://github.com/reemkhaleed/Automated-Data-Entry-Bot-for-Desktop-Application.git
cd Automated-Data-Entry-Bot-for-Desktop-Application
```

