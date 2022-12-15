from main import start_browser, read_keywords_and_urls, search_keywords, save_results, close_browser
import tkinter as tk
from tkinter import filedialog


class GoogleSearch:
    def __init__(self):
        # Create a new tkinter window with a dark blue background
        self.window = tk.Tk()
        self.window.geometry("400x250")
        self.window.title("Google Search")

        # Add a label and a text field for the file containing the keywords and URLs
        tk.Label(self.window, text="Excel file with links:").grid(row=0, column=0, padx=10, pady=30)
        self.keywords_urls_file = tk.Entry(self.window)
        self.keywords_urls_file.grid(row=0, column=1)

        # Add a button to open the file browser for the keywords and URLs file
        def open_keywords_urls_file_browser():
            # Use filedialog.askopenfilename to create a file browser and allow the user to select a file
            file_name = filedialog.askopenfilename()
            # Set the text of the keywords_urls_file text field to the selected file
            self.keywords_urls_file.insert(0, file_name)

        tk.Button(self.window, text="Browse", command=open_keywords_urls_file_browser).grid(row=0, column=2)

        # Add a label and a text field for the file to save the results
        tk.Label(self.window, text="Results Filename:").grid(row=1, column=0)
        self.results_file = tk.Entry(self.window)
        self.results_file.grid(row=1, column=1)


        # Add a button to start the search
        tk.Button(self.window, text="Run script", padx=25, pady=5, command=self.start).grid(row=2, column=0,
                                                                                              columnspan=3, pady=20)

    def start(self):
        # Start the browser
        browser = start_browser()

        # Get the keywords and URLs file and results file from the text fields
        keywords_urls_file_name = self.keywords_urls_file.get()
        results_file_name = self.results_file.get()

        # Read the keywords and URLs
        keywords, urls, pages = read_keywords_and_urls(keywords_urls_file_name)

        # Search for the keywords and get the results
        results = search_keywords(browser, keywords, urls, pages)

        # Save the results
        try:
            save_results(results, results_file_name)
        except:
            close_browser(browser)

        # Close the browser
        close_browser(browser)


# Create a new GoogleSearch instance and start the keyword search
search = GoogleSearch()

# Call the mainloop method on the window instance variable to display the tkinter window
search.window.mainloop()
