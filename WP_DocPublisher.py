import mammoth
import tkinter as tk
import customtkinter
from tkinter import filedialog, StringVar, Toplevel
from docx import Document
import re
import os
import traceback

# WordPress Imports
from wordpress_xmlrpc import Client, WordPressPost, WordPressPage
from wordpress_xmlrpc.methods import posts
from wordpress_xmlrpc.compat import xmlrpc_client
from wordpress_xmlrpc.exceptions import InvalidCredentialsError
# Import specific network/protocol errors
from socket import gaierror
from xmlrpc.client import ProtocolError
import ssl

# --- Global Variables ---
app = None
main = None
file_path_var = None
post_title_var = None
content_type_var = None
result_label = None
wp_url = None
wp_username = None
wp_password = None
url_entry = None
user_entry = None
password_entry = None
error_label = None
login_button = None

# --- Helper Function ---
def sanitize_style_name(name):
    """Converts DOCX style name to a CSS class name."""
    if not name: return "style-unnamed"
    s = name.lower()
    s = re.sub(r'[^a-z0-9\s-]', '', s)
    s = re.sub(r'\s+', '-', s)
    s = s.strip('-')
    return f"style-{s}" if s else "style-paragraph"

# --- Core Functionality ---

def select_file():
    """Opens file dialog to select DOCX."""
    if file_path_var:
        openpath = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=[("Word Documents", "*.docx")]
        )
        if openpath:
            file_path_var.delete(0, tk.END)
            file_path_var.insert(0, openpath)

def publish_content():
    """Converts DOCX and publishes to WordPress."""

    if not all([file_path_var, post_title_var, content_type_var,
                wp_url, wp_username, wp_password, result_label]):
        print("Error: Core UI/credential variables missing.")
        if result_label: result_label.configure(text="Error: Setup incomplete.", text_color="red")
        return

    docx_file = file_path_var.get().strip()
    title = post_title_var.get().strip()
    content_type_str = content_type_var.get()

    if not docx_file or not title or not content_type_str:
        result_label.configure(text="Error: File, Title, and Type required.", text_color="red")
        return

    try:
        result_label.configure(text="Processing...", text_color="orange")
        app.update_idletasks()

        # --- DOCX Style Reading -> Mammoth Map -> HTML Conversion ---
        print(f"Reading styles from: {docx_file}")
        doc = Document(docx_file)
        mammoth_style_map_parts = []
        for style in doc.styles:
            if style.type == 1: # Paragraph styles only
                style_class = sanitize_style_name(style.name)
                mammoth_style_map_parts.append(f"p[style-name='{style.name}'] => p.{style_class}:fresh")
        mammoth_style_map = "\n".join(mammoth_style_map_parts)

        print("Converting DOCX to HTML with Mammoth...")
        result = mammoth.convert_to_html(docx_file, style_map=mammoth_style_map)
        html = result.value
        if result.messages: print("Mammoth Warnings/Messages:", result.messages)
        if not html or not html.strip(): raise ValueError("Mammoth conversion resulted in empty HTML.")
        print("DOCX to HTML conversion successful.")

        # --- WordPress Client Setup ---
        client = Client(wp_url, wp_username, wp_password)

        # --- Prepare WordPress Post/Page Object ---
        content_object = WordPressPost() if content_type_str == "POST" else WordPressPage()
        content_object.post_type = content_type_str.lower()
        content_object.post_status = 'publish'
        content_object.title = title
        content_object.content = html

        # --- Publish to WordPress ---
        print(f"Publishing '{title}' as {content_type_str}...")
        content_id = client.call(posts.NewPost(content_object))
        print(f"Publish successful! New Content ID: {content_id}")
        result_label.configure(text=f'"{title}"\npublished as {content_type_str}!', text_color="green")

    # --- Error Handling ---
    except FileNotFoundError:
        result_label.configure(text=f"Error: File not found\n'{docx_file}'", text_color="red")
    except InvalidCredentialsError:
        result_label.configure(text='Error: Invalid WordPress credentials.', text_color="red")
    except ImportError as ie:
         result_label.configure(text=f"Error: Missing library\n{ie.name}", text_color="red")
    except ValueError as ve:
        result_label.configure(text=f"Error:\n{ve}", text_color="red")
    except (OSError, gaierror, ProtocolError, ssl.SSLError) as network_err:
         result_label.configure(text=f"Network Error:\n{type(network_err).__name__}", text_color="red")
         print(f"Network error during publishing: {network_err}")
    except Exception as e:
        print(f"An unexpected error occurred during publishing: {type(e).__name__} - {e}")
        traceback.print_exc()
        result_label.configure(text=f"Error: {type(e).__name__}\nCheck console log.", text_color="red")

# --- Function to close the entire application ---
def quit_application():
    """Destroys the root window, terminating the application."""
    if app:
        app.destroy()

# --- UI Window Definitions ---
def create_main_window():
    """Creates the main content publishing UI window."""
    global main, file_path_var, post_title_var, content_type_var, result_label

    if main is not None and main.winfo_exists():
        main.lift()
        return

    customtkinter.set_appearance_mode("System")
    customtkinter.set_default_color_theme("blue")

    main = Toplevel(app) # Create as Toplevel, child of 'app'
    main.geometry("420x450")
    main.title('WordPress Publisher')
    main.resizable(False, False)

    main.protocol("WM_DELETE_WINDOW", quit_application) # Handle closing window

    # Center window relative to login window
    main.update_idletasks()
    try:
        login_x, login_y = app.winfo_x(), app.winfo_y()
        login_w, login_h = app.winfo_width(), app.winfo_height()
        main_w, main_h = main.winfo_width(), main.winfo_height()
        center_x = login_x + (login_w // 2) - (main_w // 2)
        center_y = login_y + (login_h // 2) - (main_h // 2)
        main.geometry(f"+{center_x}+{center_y}")
    except Exception: pass


    # --- Content Frame ---
    frame = customtkinter.CTkFrame(master=main, width=360, height=410, corner_radius=15)
    frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

    title_label = customtkinter.CTkLabel(master=frame, text="Publish DOCX to WordPress", font=('Century Gothic', 20, 'bold'))
    title_label.place(relx=0.5, y=35, anchor=tk.CENTER)

    # File Selection
    file_frame = customtkinter.CTkFrame(master=frame, fg_color="transparent", width=300)
    file_frame.place(x=30, y=80)
    file_path_var = customtkinter.CTkEntry(master=file_frame, width=200, placeholder_text='Path to .docx file')
    file_path_var.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
    select_file_button = customtkinter.CTkButton(master=file_frame, width=80, text="Browse...", command=select_file)
    select_file_button.pack(side=tk.LEFT)

    # Post Title
    title_label_inner = customtkinter.CTkLabel(master=frame, text="Content Title:", font=('Century Gothic', 12))
    title_label_inner.place(x=30, y=125)
    post_title_var = customtkinter.CTkEntry(master=frame, width=300, placeholder_text='Enter title for WordPress')
    post_title_var.place(x=30, y=150)

    # Content Type
    type_label = customtkinter.CTkLabel(master=frame, text="Content Type:", font=('Century Gothic', 12))
    type_label.place(x=30, y=190)
    content_type_var = StringVar(value="POST")
    post_radio = customtkinter.CTkRadioButton(master=frame, text="Post", variable=content_type_var, value="POST")
    post_radio.place(x=50, y=215)
    page_radio = customtkinter.CTkRadioButton(master=frame, text="Page", variable=content_type_var, value="PAGE")
    page_radio.place(x=150, y=215)

    # --- Publish Button ---
    publish_button = customtkinter.CTkButton(master=frame, width=300, text="Publish to WordPress", command=publish_content, corner_radius=6, font=('Century Gothic', 14, 'bold'))
    publish_button.place(x=30, y=270)

    # --- Result Label ---
    result_label = customtkinter.CTkLabel(master=frame, text="", font=('Century Gothic', 12), wraplength=320, justify=tk.CENTER)
    result_label.place(relx=0.5, y=330, anchor=tk.CENTER)


def create_login_window():
    """Creates the initial login window."""
    global app, url_entry, user_entry, password_entry, error_label, login_button

    customtkinter.set_appearance_mode("System")
    customtkinter.set_default_color_theme("blue")


    app = customtkinter.CTk()
    app.geometry("400x510")
    app.title('WordPress Login')
    app.resizable(False, False)

    app.protocol("WM_DELETE_WINDOW", quit_application) # Ensure closing login window quits app

    # Center window
    app.update_idletasks()
    try:
        screen_width = app.winfo_screenwidth()
        screen_height = app.winfo_screenheight()
        window_width = app.winfo_width()
        window_height = app.winfo_height()
        center_x = (screen_width // 2) - (window_width // 2)
        center_y = (screen_height // 2) - (window_height // 2)
        app.geometry(f"+{center_x}+{center_y}")
    except Exception: pass

    # Login Frame
    frame = customtkinter.CTkFrame(master=app, width=320, height=450, corner_radius=15)
    frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

    logintitle = customtkinter.CTkLabel(master=frame, text="Login to WordPress", font=('Century Gothic', 20, 'bold'))
    logintitle.place(relx=0.5, y=45, anchor=tk.CENTER)

    # URL Input
    url_label = customtkinter.CTkLabel(master=frame, text="Site URL (e.g., mysite.local or https://mysite.com):", font=('Century Gothic', 10))
    url_label.place(x=30, y=90)
    url_entry = customtkinter.CTkEntry(master=frame, width=260, placeholder_text='Enter site address')
    url_entry.place(x=30, y=115)

    # Username Input
    user_label = customtkinter.CTkLabel(master=frame, text="Username:", font=('Century Gothic', 12))
    user_label.place(x=30, y=175)
    user_entry = customtkinter.CTkEntry(master=frame, width=260, placeholder_text='Enter WordPress username')
    user_entry.place(x=30, y=200)

    # Password Input
    password_label = customtkinter.CTkLabel(master=frame, text="Application Password (Recommended):", font=('Century Gothic', 12))
    password_label.place(x=30, y=245)
    password_entry = customtkinter.CTkEntry(master=frame, width=260, placeholder_text='Enter Application Password', show='*')
    password_entry.place(x=30, y=270)

    # Error Label
    error_label = customtkinter.CTkLabel(master=frame, text="", font=('Century Gothic', 12), text_color="red", wraplength=280)
    error_label.place(relx=0.5, y=410, anchor=tk.CENTER)

    def submit_login_event(event=None): # Wrapper for binding
        """Wrapper function to call submit_login, compatible with binds."""
        submit_login()

    def submit_login():
        """Validates login, connects, and opens main window."""
        global wp_url, wp_username, wp_password

        error_label.configure(text="")
        temp_url_base = url_entry.get().strip()
        temp_user = user_entry.get().strip()
        temp_pass = password_entry.get().strip()

        if not temp_url_base or not temp_user or not temp_pass:
            error_label.configure(text="Error: All fields are required.")
            return

        # Construct potential URLs
        temp_url_cleaned = temp_url_base.lower().removeprefix("http://").removeprefix("https://").rstrip('/')
        if not temp_url_cleaned:
            error_label.configure(text="Error: Invalid URL entered.")
            return
        possible_urls = [ f"https://{temp_url_cleaned}/xmlrpc.php", f"http://{temp_url_cleaned}/xmlrpc.php" ]

        login_button.configure(state='disabled', text='Validating...')
        app.update_idletasks()

        login_successful = False
        last_error = None

        # Try connecting with possible URLs
        for i, url_to_try in enumerate(possible_urls):
            print(f"Attempting connection to: {url_to_try}")
            try:
                # Test connection needs a client instance
                client_test = Client(url_to_try, temp_user, temp_pass)
                from wordpress_xmlrpc.methods import users
                client_test.call(users.GetProfile()) # Test auth

                print("Login validation successful.")
                wp_url = url_to_try
                wp_username = temp_user
                wp_password = temp_pass
                login_successful = True
                break # Exit loop on success
            except InvalidCredentialsError as e:
                last_error = e; break
            except (OSError, gaierror, ProtocolError, ssl.SSLError) as e:
                print(f"Connection attempt failed for {url_to_try}: {type(e).__name__} - {e}")
                last_error = e
            except Exception as e:
                print(f"Unexpected login error with {url_to_try}: {type(e).__name__}")
                last_error = e; traceback.print_exc(); break

        # Post-Loop Actions
        if login_successful:
            # --- Open Main Window ---
            print("Login successful, opening publisher window.")
            app.withdraw()
            create_main_window()

        else: # Login failed
            if isinstance(last_error, InvalidCredentialsError):
                error_label.configure(text='Login Failed: Invalid username or Application Password.\n(Use an Application Password, not your WP Admin login password)')
            elif isinstance(last_error, (OSError, gaierror, ProtocolError, ssl.SSLError)):
                error_label.configure(text=f'Connection Error:\nCould not reach site or XML-RPC disabled.')
                print(f"Detailed connection error: {last_error}")
            elif last_error:
                 error_label.configure(text=f'Login Error:\n{type(last_error).__name__}. Check details.')
            else:
                error_label.configure(text='Login failed. Check details and ensure XML-RPC is enabled.')
            # Re-enable button only if the window still exists
            if login_button and login_button.winfo_exists():
                login_button.configure(state='normal', text='Log in')


    # Bind Enter key to password field
    password_entry.bind("<Return>", submit_login_event)

    # Login Button
    login_button = customtkinter.CTkButton(master=frame, width=260, text='Log in', command=submit_login_event, corner_radius=6, font=('Century Gothic', 14, 'bold'))
    login_button.place(relx=0.5, y=365, anchor=tk.CENTER)

    app.mainloop() # Start Tkinter event loop

# --- Script Entry Point ---
if __name__ == "__main__":
    # Startup library check
    missing = []
    try: import mammoth
    except ImportError: missing.append("mammoth")
    # try: import pytz # Removed
    # except ImportError: missing.append("pytz")
    try: import customtkinter
    except ImportError: missing.append("customtkinter")
    try: from PIL import Image, ImageTk
    except ImportError: missing.append("Pillow")
    try: import docx # Check python-docx
    except ImportError: missing.append("python-docx")
    try: import wordpress_xmlrpc
    except ImportError: missing.append("python-wordpress-xmlrpc")

    if missing:
         print("[ERROR] Required libraries are missing:")
         for lib in missing:
             print(f" - {lib}")
         print("\nPlease install them (e.g., using 'pip install <library_name>') and restart.")
         try:
             root = tk.Tk()
             root.withdraw()
             tk.messagebox.showerror("Missing Libraries", "Required libraries missing:\n- " + "\n- ".join(missing) + "\nPlease install them and restart.")
             root.destroy()
         except Exception:
             pass
         exit(1)
    else:
        print("Required libraries seem to be installed.")

    print("\n--- WordPress Login Recommendation ---")
    print("It is strongly recommended to use Application Passwords for this tool.")
    print("You can generate one in your WordPress User > Profile > Application Passwords section.")
    print("Do NOT use your main WordPress admin password here.\n")


    create_login_window() # Launch the application