import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pyautogui
import keyboard
import pytesseract
import time
from PIL import Image, ImageTk, ImageGrab
import os
from ebooklib import epub
import threading

# Set Tesseract path for Windows
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


class RegionSelector:
    def __init__(self):
        self.root = tk.Tk()
        self.root.attributes('-alpha', 0.3)
        self.root.attributes('-fullscreen', True)
        self.root.attributes('-topmost', True)

        # Variables to store coordinates
        self.start_x = None
        self.start_y = None
        self.current_x = None
        self.current_y = None

        # Create canvas
        self.canvas = tk.Canvas(self.root, highlightthickness=0)
        self.canvas.pack(fill='both', expand=True)

        # Bind mouse events
        self.canvas.bind('<Button-1>', self.on_press)
        self.canvas.bind('<B1-Motion>', self.on_drag)
        self.canvas.bind('<ButtonRelease-1>', self.on_release)

        self.rectangle = None
        self.region = None

        # Add instruction label
        self.label = tk.Label(self.root, text="Click and drag to select region",
                              bg='white', fg='black', font=('Arial', 14))
        self.label.place(relx=0.5, rely=0.1, anchor='center')

    def on_press(self, event):
        self.start_x = event.x
        self.start_y = event.y

    def on_drag(self, event):
        self.current_x = event.x
        self.current_y = event.y

        # Clear previous rectangle
        if self.rectangle:
            self.canvas.delete(self.rectangle)

        # Draw new rectangle
        self.rectangle = self.canvas.create_rectangle(
            self.start_x, self.start_y,
            self.current_x, self.current_y,
            outline='red', width=2
        )

    def on_release(self, event):
        if (self.start_x is None or self.current_x is None or
                self.start_y is None or self.current_y is None):
            return None

        # Get the final coordinates
        x1 = min(self.start_x, self.current_x)
        y1 = min(self.start_y, self.current_y)
        x2 = max(self.start_x, self.current_x)
        y2 = max(self.start_y, self.current_y)

        # Ensure minimum size
        if x2 - x1 < 10 or y2 - y1 < 10:
            return None

        self.region = (x1, y1, x2, y2)  # Store absolute coordinates
        self.root.quit()  # Use quit() instead of destroy()

    def get_region(self):
        self.root.mainloop()
        region = self.region
        self.root.destroy()  # Destroy after getting region
        if region is None:
            return None
        x1, y1, x2, y2 = region
        return (x1, y1, x2-x1, y2-y1)  # Return as (x, y, width, height)


class ScreenshotOCRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Screenshot OCR Automation")

        # State variables
        self.is_running = False
        self.is_paused = False
        self.screenshot_region = None
        self.next_button_pos = None
        self.pages_to_process = 0
        self.current_page = 0
        self.start_time = 0
        self.save_location = os.path.join(os.path.expanduser(
            '~'), 'document.epub')  # Default location

        # Main frame
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Setup UI elements
        self.setup_ui()

    def setup_ui(self):
        # Pages input
        ttk.Label(self.main_frame, text="Number of pages:").grid(
            row=0, column=0, pady=5)
        self.pages_entry = ttk.Entry(self.main_frame)
        self.pages_entry.grid(row=0, column=1, pady=5)
        self.pages_entry.insert(0, "1")

        # Delay input
        ttk.Label(self.main_frame, text="Delay between pages (seconds):").grid(
            row=1, column=0, pady=5)
        self.delay_entry = ttk.Entry(self.main_frame)
        self.delay_entry.grid(row=1, column=1, pady=5)
        self.delay_entry.insert(0, "1")

        # Save location frame
        save_frame = ttk.Frame(self.main_frame)
        save_frame.grid(row=2, column=0, columnspan=2, pady=5)

        self.save_location_label = ttk.Label(save_frame,
                                             text=f"Save to: {self.save_location}",
                                             wraplength=300)
        self.save_location_label.pack(pady=5)

        save_button = ttk.Button(save_frame,
                                 text="Choose Save Location",
                                 command=self.choose_save_location)
        save_button.pack(pady=5)

        # Region selection button
        self.region_btn = ttk.Button(self.main_frame, text="Select Screenshot Region",
                                     command=self.select_region)
        self.region_btn.grid(row=2, column=0, columnspan=2, pady=5)

        # Selection status
        self.region_status = ttk.Label(
            self.main_frame, text="Region: Not Selected")
        self.region_status.grid(row=3, column=0, columnspan=2, pady=5)

        # Create style for centered buttons
        style = ttk.Style()
        style.configure('Centered.TButton', justify='center', width=25)

        # Next button selection
        self.next_btn = ttk.Button(self.main_frame, text="Select 'Next' Button Location\n(Press Space when ready)",
                                   command=self.select_next_button, style='Centered.TButton')
        self.next_btn.grid(row=4, column=0, columnspan=2, pady=5)

        # Next button status
        self.next_status = ttk.Label(
            self.main_frame, text="Next Button: Not Selected")
        self.next_status.grid(row=5, column=0, columnspan=2, pady=5)

        # Start/Stop button
        self.start_btn = ttk.Button(self.main_frame, text="Start",
                                    command=self.toggle_processing)
        self.start_btn.grid(row=6, column=0, columnspan=2, pady=5)

        # Pause button
        self.pause_btn = ttk.Button(self.main_frame, text="Pause",
                                    command=self.toggle_pause, state=tk.DISABLED)
        self.pause_btn.grid(row=7, column=0, columnspan=2, pady=5)

        # Progress bar
        self.progress = ttk.Progressbar(
            self.main_frame, length=200, mode='determinate')
        self.progress.grid(row=8, column=0, columnspan=2, pady=5)

        # Status label
        self.status_label = ttk.Label(self.main_frame, text="Ready")
        self.status_label.grid(row=9, column=0, columnspan=2, pady=5)

        # Adjust the row numbers for all subsequent elements
        self.region_btn.grid(row=3, column=0, columnspan=2, pady=5)
        self.region_status.grid(row=4, column=0, columnspan=2, pady=5)
        self.next_btn.grid(row=5, column=0, columnspan=2, pady=5)
        self.next_status.grid(row=6, column=0, columnspan=2, pady=5)
        self.start_btn.grid(row=7, column=0, columnspan=2, pady=5)
        self.pause_btn.grid(row=8, column=0, columnspan=2, pady=5)
        self.progress.grid(row=9, column=0, columnspan=2, pady=5)
        self.status_label.grid(row=10, column=0, columnspan=2, pady=5)

    def select_region(self):
        self.root.iconify()
        time.sleep(0.5)

        try:
            selector = RegionSelector()
            region = selector.get_region()
            if region:
                self.screenshot_region = region
                self.region_status.config(
                    text=f"Region: Selected ({region[2]}x{region[3]})")
                messagebox.showinfo("Success", "Screenshot region selected!")
            else:
                messagebox.showerror(
                    "Error", "No region selected. Please try again.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to select region: {str(e)}")
        finally:
            self.root.deiconify()

    def select_next_button(self):
        self.root.iconify()
        time.sleep(0.5)
        self.status_label.config(
            text="Move mouse to 'Next' button and press Space...")

        def on_space():
            self.next_button_pos = pyautogui.position()
            keyboard.unhook_all()
            self.root.deiconify()
            self.next_status.config(
                text=f"Next Button: Selected ({self.next_button_pos[0]}, {self.next_button_pos[1]})")
            messagebox.showinfo("Success", "'Next' button location saved!")
            self.status_label.config(text="Ready")

        keyboard.on_press_key('space', lambda _: on_space())

    def toggle_processing(self):
        if not self.is_running:
            if not self.validate_inputs():
                return

            self.is_running = True
            self.start_time = time.time()  # Add this line to track start time
            self.start_btn.config(text="Stop")
            self.pause_btn.config(state=tk.NORMAL)
            self.processing_thread = threading.Thread(
                target=self.process_pages)
            self.processing_thread.start()
        else:
            self.is_running = False
            self.start_btn.config(text="Start")
            self.pause_btn.config(state=tk.DISABLED)

    def toggle_pause(self):
        self.is_paused = not self.is_paused
        self.pause_btn.config(text="Resume" if self.is_paused else "Pause")

    def validate_inputs(self):
        if not self.screenshot_region:
            messagebox.showerror(
                "Error", "Please select screenshot region first")
            return False
        if not self.next_button_pos:
            messagebox.showerror(
                "Error", "Please select 'Next' button location first")
            return False
        try:
            self.pages_to_process = int(self.pages_entry.get())
            self.delay = float(self.delay_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numbers")
            return False
        return True

    def process_pages(self):
        book = epub.EpubBook()
        book.set_identifier('id123456')
        book.set_title('Automated Document Scan')
        book.set_language('en')

        # Add CSS style
        style = '''
            @namespace epub "http://www.idpf.org/2007/ops";
            body {
                font-family: Arial, sans-serif;
                line-height: 1.5;
                padding: 1em;
            }
            p {
                margin: 1em 0;
                text-align: justify;
            }
            h1 {
                text-align: center;
                color: #333;
                margin: 1em 0;
            }
            .paragraph {
                margin-bottom: 1em;
            }
            .page-break {
                page-break-after: always;
            }
        '''
        nav_css = epub.EpubItem(uid="style_nav",
                                file_name="style/nav.css",
                                media_type="text/css",
                                content=style)
        book.add_item(nav_css)

        chapters = []
        self.current_page = 0

        while self.current_page < self.pages_to_process and self.is_running:
            while self.is_paused:
                time.sleep(0.1)

            # Take screenshot
            x, y, width, height = self.screenshot_region
            screenshot = ImageGrab.grab(bbox=(x, y, x+width, y+height))

            # Perform OCR with additional parameters
            custom_config = r'--oem 3 --psm 6'

            # Get text and detailed data about text blocks
            ocr_data = pytesseract.image_to_data(
                screenshot, output_type=pytesseract.Output.DICT, config=custom_config)

            # Process OCR data to maintain formatting
            formatted_content = self.process_ocr_data(ocr_data)

            # Create chapter with formatted content
            chapter = epub.EpubHtml(title=f'Page {self.current_page + 1}',
                                    file_name=f'page_{self.current_page + 1}.xhtml',
                                    lang='en')

            chapter_content = f'''
                <h1>Page {self.current_page + 1}</h1>
                {formatted_content}
                <div class="page-break"></div>
            '''

            chapter.content = chapter_content
            chapter.add_item(nav_css)
            chapters.append(chapter)

            # Update progress
            self.current_page += 1
            self.progress['value'] = (
                self.current_page / self.pages_to_process) * 100
            self.status_label.config(
                text=f"Processing page {self.current_page}/{self.pages_to_process}")

            # Click next if not last page
            if self.current_page < self.pages_to_process:
                pyautogui.click(self.next_button_pos)
                time.sleep(self.delay)

        if self.is_running:  # Only save if not stopped manually
            # Add chapters to book
            for chapter in chapters:
                book.add_item(chapter)

            # Create table of contents
            book.toc = [(epub.Section('Document'), chapters)]

            # Add default NCX and Nav file
            book.add_item(epub.EpubNcx())
            book.add_item(epub.EpubNav())

            # Create spine
            book.spine = ['nav'] + chapters

            # Save epub with proper formatting
            epub.write_epub(self.save_location, book, {})

            # Show completion window
            completion_window = tk.Toplevel(self.root)
            completion_window.title("Scanning Complete")
            completion_window.geometry("500x650")
            completion_window.lift()
            completion_window.attributes('-topmost', True)

            # Calculate some statistics
            total_time = round(time.time() - self.start_time, 1)
            avg_time_per_page = round(total_time / self.pages_to_process, 1)

            completion_text = f"""Scanning Complete! 🎉

Summary:
• Pages processed: {self.pages_to_process}
• Total time: {total_time} seconds
• Average time per page: {avg_time_per_page} seconds
• File saved as: document.epub

The EPUB file has been created with:
• Full text content
• Formatted paragraphs
• Proper page breaks
• Consistent styling

You can find your document at:
{os.path.abspath('document.epub')}
            """

            # Add completion message
            msg_label = ttk.Label(
                completion_window, text=completion_text, justify=tk.LEFT, wraplength=380)
            msg_label.pack(pady=20, padx=20)

            # Add button to open file location
            def open_file_location():
                os.startfile(os.path.dirname(os.path.abspath('document.epub')))

            open_button = ttk.Button(completion_window, text="Open File Location",
                                     command=open_file_location)
            open_button.pack(pady=10)

            # Add close button
            close_button = ttk.Button(completion_window, text="Close",
                                      command=completion_window.destroy)
            close_button.pack(pady=5)

            # Update main window status
            self.status_label.config(text="Scanning complete! Check the summary window.",
                                     foreground="green")
            self.progress['value'] = 100

            # Enable/disable buttons appropriately
            self.start_btn.config(text="Start New Scan")
            self.pause_btn.config(state=tk.DISABLED)

            # Clear the progress bar gradually
            def fade_progress():
                current = self.progress['value']
                if current > 0:
                    self.progress['value'] = current - 1
                    self.root.after(50, fade_progress)

            # Start fading after 3 seconds
            self.root.after(3000, fade_progress)

        self.is_running = False
        self.start_btn.config(text="Start")
        self.pause_btn.config(state=tk.DISABLED)

    def open_file_location(self):
        os.startfile(os.path.dirname(self.save_location))

    def process_ocr_data(self, ocr_data):
        """Process OCR data to maintain formatting and structure."""
        formatted_html = []
        current_block = []
        last_block_num = -1

        # Zip all the OCR data together
        for i in range(len(ocr_data['text'])):
            # Skip empty text
            if not ocr_data['text'][i].strip():
                continue

            block_num = ocr_data['block_num'][i]
            text = ocr_data['text'][i]
            conf = ocr_data['conf'][i]

            # New block detected
            if block_num != last_block_num:
                if current_block:
                    # Join the previous block and add it to formatted_html
                    formatted_html.append(
                        f'<p class="paragraph">{" ".join(current_block)}</p>'
                    )
                    current_block = []

            current_block.append(text)
            last_block_num = block_num

        # Add the last block if it exists
        if current_block:
            formatted_html.append(
                f'<p class="paragraph">{" ".join(current_block)}</p>'
            )

        return '\n'.join(formatted_html)

    def choose_save_location(self):
        initial_dir = os.path.dirname(self.save_location)
        filename = filedialog.asksaveasfilename(
            initialdir=initial_dir,
            initialfile='document.epub',
            defaultextension='.epub',
            filetypes=[('EPUB files', '*.epub'), ('All files', '*.*')]
        )
        if filename:
            self.save_location = filename
            self.save_location_label.config(
                text=f"Save to: {self.save_location}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ScreenshotOCRApp(root)
    root.mainloop()
