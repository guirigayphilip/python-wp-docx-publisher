# Python WordPress DOCX Publisher

A Python GUI application built with CustomTkinter that allows users to publish `.docx` documents directly to a WordPress site as either Posts or Pages. It uses `mammoth` to convert documents to HTML, attempting to preserve semantic structure by mapping Word paragraph styles to corresponding CSS classes in the output.

**Features**

*   **Graphical User Interface:** Easy-to-use interface built with CustomTkinter.
*   **DOCX to HTML Conversion:** Uses the `mammoth` library for robust conversion.
*   **WordPress Style Mapping:** Converts Word paragraph styles (e.g., "Heading 1") into CSS classes (e.g., `style-heading-1`) in the HTML output. **Requires corresponding CSS in your WordPress theme to render correctly.**
*   **Publish as Post or Page:** Choose whether the content becomes a WordPress Post or Page.
*   **Flexible URL Input:** Enter your site address (e.g., `mysite.local` or `https://mysite.com`) and the script will attempt connection via HTTPS/HTTP and automatically append `/xmlrpc.php`.
*   **Sample Document:** Includes a `sample/a.docx` file for quick testing.

**Requirements**

*   **Python:** 3.8 or higher recommended (Tested with 3.x). Make sure Python is added to your system's PATH during installation.
*   **WordPress Site:** A self-hosted WordPress site or a local instance (like LocalWP) with XML-RPC enabled (usually enabled by default in Settings > Writing).
*   **Required Python Libraries:** Should be listed in `requirements.txt`. Key libraries include:
    *   `customtkinter`
    *   `mammoth`
    *   `python-wordpress-xmlrpc`
    *   `python-docx`
    *   `Pillow`

**Installation**

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/guirigayphilip/python-wp-docx-publisher.git

    cd python-wp-docx-publisher
    ```

2.  **Prepare `requirements.txt` (if not included):**
    Create a file named `requirements.txt` in the project directory and list the required libraries:
    ```
    customtkinter
    mammoth
    python-wordpress-xmlrpc
    python-docx
    Pillow
    ```

3.  **Install the required libraries:**
    Open your terminal or command prompt in the project directory and run:
    ```bash
    pip install -r requirements.txt
    ```
    *   **Note:** This installs the libraries globally or into your user-specific Python packages. Using a Virtual Environment is recommended (see note below).

---

**Optional but Recommended: Using a Virtual Environment**

A virtual environment creates an isolated space for your project's dependencies.

1.  **Create:** `python -m venv venv`
2.  **Activate:**
    *   Windows: `.\venv\Scripts\activate`
    *   macOS/Linux: `source venv/bin/activate`
3.  **Install (while activated):** `pip install -r requirements.txt`
4.  **Deactivate:** `deactivate`

---

**Usage**

1.  **Run the script:**
    Make sure you are in the project directory in your terminal (and activate the virtual environment if you created one).
    ```bash
    python WP_DocPublisher.py
    ```

2.  **Login Window:**
    *   **Site URL:** Enter your WordPress site's address (e.g., `myblog.com` or `http://mylocalwp.local`).
    *   **Username:** Enter your WordPress username.
    *   **Password:** Enter your **Application Password**.
        *   **Important Security Note:** **Do NOT use your main WordPress admin password.** Generate an **Application Password** specifically for this tool in your WordPress Admin area (Users > Profile > Application Passwords). This is much more secure.
    *   Click "Log in".

3.  **Publisher Window (after successful login):**
    *   **Browse:** Click to select a `.docx` file.
    *   **Content Title:** Enter the desired title for your WordPress Post or Page.
    *   **Content Type:** Select either "Post" or "Page".
    *   **Publish:** Click the "Publish to WordPress" button.

4.  **Check Results:**
    *   The application window will display a success or error message.
    *   Check your terminal for detailed logs or error messages.
    *   Verify the content on your WordPress site. WordPress will handle the display of the publication date based on its own settings and your theme.

**Important Note: CSS Styling**

This script converts the _structure_ of your DOCX and maps Word **paragraph styles** to **CSS classes** in the format `.style-your-style-name` (e.g., `.style-heading-1`, `.style-custom-quote`).

For these styles to appear correctly on your WordPress site, **you MUST add corresponding CSS rules to your WordPress theme.**

*   Go to `Appearance` -> `Customize` -> `Additional CSS` in your WordPress admin dashboard.
*   Add rules matching the generated classes. Example:

    ```css

    /* Example for a Heading 1 style in Word */
    .style-heading-1 {
        font-size: 2.2em; /* Adjust as needed */
        font-weight: bold;
        margin-top: 1.5em;
        margin-bottom: 0.5em;
        /* add text-align, color, etc. */
    }

    /* Example for a Body Text or Normal style */
    .style-body-text, /* Add common/expected names */
    .style-normal,
    .style-paragraph { /* Fallback */
        line-height: 1.6;
        margin-bottom: 1em;
    }
    ```

_Without adding this CSS in WordPress, the published content will mostly use your theme's default styles._

**License**

This project is licensed under the MIT License - see the `LICENSE` file for details.