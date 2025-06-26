import cv2
import pytesseract
import pandas as pd
from difflib import SequenceMatcher
import re
from pdf2image import convert_from_path
import os
import numpy as np
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
from tkinter.font import Font
import platform
import time
from datetime import datetime

class SiemensRCCComparator:
    def __init__(self, root):
        self.root = root
        self.setup_window()
        self.cancel_requested = False
        self.create_styles()
        self.build_ui()
        self.configure_paths()
        self.setup_bindings()
        
    def setup_window(self):
        """Configure the main application window"""
        self.root.title("SIEMENS RCC Comparator")
        self.root.geometry("1000x750")
        self.root.minsize(900, 700)
        self.root.configure(bg="#f5f5f5")  # Light gray background
        
        # Set window icon
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass
        
    def create_styles(self):
        """Create modern, clean styles"""
        self.style = ttk.Style()
        
        # Color palette
        self.colors = {
            'primary': "#0066b3",  # Siemens blue
            'secondary': "#e6f2ff",  # Light blue
            'accent': '#00b0f0',  # Bright blue
            'dark': "#333333",  # Dark gray
            'light': "#ffffff",  # White
            'success': '#4CAF50',  # Green
            'warning': '#FF9800',  # Orange
            'error': '#F44336',  # Red
            'background': "#f5f5f5"  # Light gray
        }
        
        # Configure theme
        self.style.theme_use('clam')
        
        # Base styles
        self.style.configure('.', 
                           background=self.colors['background'], 
                           foreground=self.colors['dark'],
                           font=('Segoe UI', 10))
        
        # Custom widget styles
        self.style.configure('Title.TLabel', 
                           font=('Segoe UI', 20, 'bold'), 
                           foreground=self.colors['primary'])
        self.style.configure('Subtitle.TLabel', 
                           font=('Segoe UI', 10), 
                           foreground=self.colors['dark'])
        
        # Buttons
        self.style.configure('TButton', 
                           font=('Segoe UI', 10), 
                           padding=8,
                           relief="flat",
                           borderwidth=0)
        self.style.map('TButton',
                      background=[('active', self.colors['secondary'])],
                      foreground=[('active', self.colors['dark'])])
        
        self.style.configure('Primary.TButton', 
                           background=self.colors['primary'], 
                           foreground="white",
                           font=('Segoe UI', 10, 'bold'))
        self.style.map('Primary.TButton',
                      background=[('active', self.colors['accent'])])
        
        # Frames
        self.style.configure('TLabelframe', 
                           background=self.colors['light'], 
                           relief="flat",
                           borderwidth=0)
        self.style.configure('TLabelframe.Label', 
                           background=self.colors['light'],
                           font=('Segoe UI', 11, 'bold'),
                           foreground=self.colors['primary'])
        
        # Entries
        self.style.configure('TEntry', 
                           fieldbackground="white",
                           padding=8,
                           relief="solid",
                           borderwidth=1,
                           bordercolor="#ddd")
        
        # Progress bar
        self.style.configure('Horizontal.TProgressbar',
                           thickness=20,
                           troughcolor="#e0e0e0",
                           background=self.colors['primary'],
                           lightcolor=self.colors['primary'],
                           darkcolor=self.colors['primary'])
        
        # Status label
        self.style.configure('Status.TLabel',
                           font=('Segoe UI', 9),
                           foreground="#666666")
        
    def build_ui(self):
        """Construct the user interface with modern design"""
        # Main container with padding
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Header section
        self.create_header(main_container)
        
        # File selection section
        self.create_file_selection(main_container)
        
        # Settings section
        self.create_settings_panel(main_container)
        
        # Progress section
        self.create_progress_section(main_container)
        
        # Action buttons
        self.create_action_buttons(main_container)
        
        # Results panel
        self.create_results_panel(main_container)
        
        # Status bar
        self.create_status_bar()
        
    def create_header(self, parent):
        """Create the application header"""
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Title frame
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(side=tk.LEFT)
        
        self.title_label = ttk.Label(title_frame, 
                                   text="SIEMENS RCC Comparator", 
                                   style='Title.TLabel')
        self.title_label.pack(anchor=tk.W)
        
        self.subtitle_label = ttk.Label(title_frame, 
                                      text="Document Comparison Tool", 
                                      style='Subtitle.TLabel')
        self.subtitle_label.pack(anchor=tk.W)
        
    def create_file_selection(self, parent):
        """Create file selection controls with modern design"""
        file_frame = ttk.LabelFrame(parent, text=" Document Selection ")
        file_frame.pack(fill=tk.X, pady=(0, 15), ipadx=5, ipady=5)
        
        # Original document
        orig_frame = ttk.Frame(file_frame)
        orig_frame.pack(fill=tk.X, pady=8, padx=10)
        
        ttk.Label(orig_frame, text="Original Document:").pack(side=tk.LEFT, padx=5)
        
        self.orig_entry = ttk.Entry(orig_frame)
        self.orig_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        orig_button = ttk.Button(orig_frame, 
                               text="Browse", 
                               command=lambda: self.browse_file(self.orig_entry))
        orig_button.pack(side=tk.LEFT, padx=5)
        
        # Modified document
        mod_frame = ttk.Frame(file_frame)
        mod_frame.pack(fill=tk.X, pady=8, padx=10)
        
        ttk.Label(mod_frame, text="Modified Document:").pack(side=tk.LEFT, padx=5)
        
        self.mod_entry = ttk.Entry(mod_frame)
        self.mod_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        mod_button = ttk.Button(mod_frame, 
                              text="Browse", 
                              command=lambda: self.browse_file(self.mod_entry))
        mod_button.pack(side=tk.LEFT, padx=5)
        
    def create_settings_panel(self, parent):
        """Create comparison settings panel"""
        settings_frame = ttk.LabelFrame(parent, text=" Comparison Settings ")
        settings_frame.pack(fill=tk.X, pady=(0, 15), ipadx=5, ipady=5)
        
        # Output location
        output_frame = ttk.Frame(settings_frame)
        output_frame.pack(fill=tk.X, pady=8, padx=10)
        
        ttk.Label(output_frame, text="Output Folder:").pack(side=tk.LEFT, padx=5)
        
        self.output_entry = ttk.Entry(output_frame)
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.output_entry.insert(0, "comparison_results")
        
        output_button = ttk.Button(output_frame, 
                                 text="Browse", 
                                 command=lambda: self.browse_folder(self.output_entry))
        output_button.pack(side=tk.LEFT, padx=5)
        
        # Comparison options
        options_frame = ttk.Frame(settings_frame)
        options_frame.pack(fill=tk.X, pady=8, padx=10)
        
        ttk.Label(options_frame, text="Comparison Sensitivity:").pack(side=tk.LEFT, padx=5)
        
        self.sensitivity = tk.DoubleVar(value=0.75)
        sensitivity_slider = ttk.Scale(options_frame, 
                                     from_=0.5, 
                                     to=1.0, 
                                     variable=self.sensitivity, 
                                     orient=tk.HORIZONTAL)
        sensitivity_slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        self.sensitivity_label = ttk.Label(options_frame, 
                                         text="75%",
                                         font=('Segoe UI', 9),
                                         foreground=self.colors['primary'])
        self.sensitivity_label.pack(side=tk.LEFT, padx=5)
        
        # Bind slider to update label
        self.sensitivity.trace_add("write", lambda *_: self.update_sensitivity_label())
        
    def create_progress_section(self, parent):
        """Create progress tracking section"""
        progress_frame = ttk.LabelFrame(parent, text=" Comparison Progress ")
        progress_frame.pack(fill=tk.X, pady=(0, 15), ipadx=5, ipady=5)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            orient=tk.HORIZONTAL,
            length=400,
            mode='determinate',
            style='Horizontal.TProgressbar'
        )
        self.progress_bar.pack(fill=tk.X, padx=10, pady=(8, 12))
        
        # Progress info frame
        progress_info_frame = ttk.Frame(progress_frame)
        progress_info_frame.pack(fill=tk.X, padx=10, pady=(0, 5))
        
        # Status text
        self.status_label = ttk.Label(progress_info_frame, 
                                    style='Status.TLabel',
                                    anchor=tk.W)
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Time estimation
        self.time_label = ttk.Label(progress_info_frame, 
                                  style='Status.TLabel',
                                  anchor=tk.E)
        self.time_label.pack(side=tk.RIGHT)
        
    def create_action_buttons(self, parent):
        """Create action buttons with modern design"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Left-aligned buttons
        left_button_frame = ttk.Frame(button_frame)
        left_button_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Compare button (primary action)
        self.compare_button = ttk.Button(
            left_button_frame,
            text="Compare Documents",
            style='Primary.TButton',
            command=self.start_comparison
        )
        self.compare_button.pack(side=tk.LEFT, padx=5)
        
        # Open results button
        self.open_button = ttk.Button(
            left_button_frame,
            text="Open Results",
            command=self.open_results,
            state=tk.DISABLED
        )
        self.open_button.pack(side=tk.LEFT, padx=5)
        
        # Right-aligned buttons
        right_button_frame = ttk.Frame(button_frame)
        right_button_frame.pack(side=tk.RIGHT)
        
        # Cancel button
        self.cancel_button = ttk.Button(
            right_button_frame,
            text="Cancel",
            command=self.cancel_operation,
            state=tk.DISABLED
        )
        self.cancel_button.pack(side=tk.RIGHT, padx=5)
        
    def create_results_panel(self, parent):
        """Create results display panel with modern design"""
        results_frame = ttk.LabelFrame(parent, text=" Comparison Results ")
        results_frame.pack(fill=tk.BOTH, expand=True, ipadx=5, ipady=5)
        
        # Results text area with scrollbar
        self.results_text = scrolledtext.ScrolledText(
            results_frame,
            wrap=tk.WORD,
            font=('Segoe UI', 10),
            bg="white",
            padx=10,
            pady=10,
            relief="solid",
            borderwidth=1,
            highlightthickness=0
        )
        self.results_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Configure tags for colored text
        self.results_text.tag_config("success", foreground=self.colors['success'])
        self.results_text.tag_config("error", foreground=self.colors['error'])
        self.results_text.tag_config("warning", foreground=self.colors['warning'])
        self.results_text.tag_config("info", foreground=self.colors['primary'])
        
    def create_status_bar(self):
        """Create status bar at bottom of window"""
        self.status_bar = ttk.Frame(self.root, relief=tk.SUNKEN)
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)
        
        self.status_message = ttk.Label(
            self.status_bar,
            text="Ready to compare documents",
            style='Status.TLabel'
        )
        self.status_message.pack(side=tk.LEFT, padx=10)
        
        # Version and copyright
        version_frame = ttk.Frame(self.status_bar)
        version_frame.pack(side=tk.RIGHT, padx=10)
        
        ttk.Label(version_frame, 
                 text="v1.0", 
                 style='Status.TLabel').pack(side=tk.LEFT, padx=5)
        
        ttk.Label(version_frame, 
                 text="© 2023 Siemens", 
                 style='Status.TLabel').pack(side=tk.LEFT)
        
    # [Rest of the methods remain exactly the same as in your original code]
    # All the functional methods below this point would be identical to your original implementation
    # I'm omitting them here for brevity, but they should be included in your actual code

    def configure_paths(self):
        """Configure paths for Tesseract and Poppler"""
        pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        self.POPPLER_PATH = r"C:\Users\Vidyadhar\Downloads\Release-24.08.0-0\poppler-24.08.0\Library\bin"
        
    def setup_bindings(self):
        """Set up keyboard and UI bindings"""
        self.root.bind("<F1>", lambda e: self.show_help())
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
    def browse_file(self, entry_widget):
        """Open file dialog and update entry widget"""
        file_path = filedialog.askopenfilename(
            title="Select PDF Document",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        if file_path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, file_path)
            
    def browse_folder(self, entry_widget):
        """Open folder dialog and update entry widget"""
        folder_path = filedialog.askdirectory(
            title="Select Output Folder"
        )
        if folder_path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, folder_path)
            
    def update_sensitivity_label(self):
        """Update the sensitivity percentage label"""
        sensitivity = int(self.sensitivity.get() * 100)
        self.sensitivity_label.config(text=f"{sensitivity}%")
        
    def start_comparison(self):
        """Start the document comparison process"""
        original = self.orig_entry.get()
        modified = self.mod_entry.get()
        output_dir = self.output_entry.get()
        
        # Validate inputs
        if not original or not modified:
            messagebox.showerror("Input Error", "Please select both documents to compare")
            return
            
        if not os.path.exists(original) or not os.path.exists(modified):
            messagebox.showerror("File Error", "One or more selected files don't exist")
            return
            
        if not output_dir:
            output_dir = "comparison_results"
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, output_dir)
            
        # Create output directory if needed
        os.makedirs(output_dir, exist_ok=True)
        
        # Disable UI during operation
        self.toggle_ui_state(False)
        self.results_text.delete(1.0, tk.END)
        self.progress_bar["value"] = 0
        self.update_status("Preparing to compare documents...")
        
        # Store start time for time estimation
        self.start_time = time.time()
        self.last_update_time = self.start_time
        
        # Run comparison in separate thread
        self.comparison_thread = threading.Thread(
            target=self.compare_documents,
            args=(original, modified, output_dir),
            daemon=True
        )
        self.comparison_thread.start()
        
        # Start progress monitoring
        self.monitor_progress()
        
    def compare_documents(self, original_path, modified_path, output_dir):
        """Main document comparison function"""
        try:
            # Step 1: Convert PDFs to images
            self.update_progress(5, "Converting documents to images...")
            
            original_pages = convert_from_path(
                original_path, 
                dpi=300, 
                poppler_path=self.POPPLER_PATH
            )
            modified_pages = convert_from_path(
                modified_path, 
                dpi=300, 
                poppler_path=self.POPPLER_PATH
            )
            
            total_pages = min(len(original_pages), len(modified_pages))
            if total_pages == 0:
                raise ValueError("No pages found in one or both documents")
                
            self.log_result(f"Processing {total_pages} page comparisons...", "info")
            
            # Step 2: Process each page
            result_images = []
            sensitivity = self.sensitivity.get()
            
            for i, (orig_page, mod_page) in enumerate(zip(original_pages, modified_pages)):
                if self.cancel_requested:
                    self.log_result("Comparison cancelled by user", "warning")
                    break
                    
                progress = 10 + (i * 80 / total_pages)
                self.update_progress(
                    progress,
                    f"Comparing page {i+1}/{total_pages}..."
                )
                
                # Extract text data
                orig_data, _ = self.extract_text_data(orig_page)
                mod_data, mod_img = self.extract_text_data(mod_page)
                
                # Find differences
                orig_words = orig_data['text'].tolist()
                mod_words = mod_data['text'].tolist()
                
                changed_indices = [
                    idx for idx, mod_word in enumerate(mod_words)
                    if not any(self.is_text_similar(mod_word, orig_word, sensitivity) 
                             for orig_word in orig_words)
                ]
                
                # Highlight changes if any
                if changed_indices:
                    highlighted_img = self.highlight_changes(mod_img, mod_data, changed_indices)
                    output_path = os.path.join(output_dir, f"page_{i+1}_diff.png")
                    cv2.imwrite(output_path, highlighted_img)
                    result_images.append(output_path)
                    
                    self.log_result(
                        f"Page {i+1}: Found {len(changed_indices)} differences → {output_path}",
                        "success"
                    )
                else:
                    self.log_result(f"Page {i+1}: No differences found", "info")
                
            # Step 3: Create final PDF if we have results
            if result_images and not self.cancel_requested:
                self.update_progress(95, "Generating final report...")
                pdf_path = os.path.join(output_dir, "comparison_report.pdf")
                self.create_pdf_report(result_images, pdf_path)
                self.log_result(f"Final report generated: {pdf_path}", "success")
                self.result_pdf_path = pdf_path
                
            # Final update
            if not self.cancel_requested:
                self.update_progress(100, "Comparison complete!")
                self.log_result("Document comparison finished successfully", "success")
            else:
                self.update_progress(0, "Comparison cancelled")
                
        except Exception as e:
            self.update_progress(0, f"Error: {str(e)}")
            self.log_result(f"Error during comparison: {str(e)}", "error")
            messagebox.showerror("Comparison Error", f"An error occurred:\n{str(e)}")
            
        finally:
            # Re-enable UI
            self.toggle_ui_state(True)
            self.cancel_requested = False
            
    def highlight_changes(self, image, text_data, change_indices):
        """Highlight changed text regions on the image"""
        overlay = image.copy()
        padding = 4
        
        for idx in change_indices:
            if idx < len(text_data):
                row = text_data.loc[idx]
                x, y, w, h = int(row['left']), int(row['top']), int(row['width']), int(row['height'])
                x1 = max(x - padding, 0)
                y1 = max(y - padding, 0)
                x2 = x + w + padding
                y2 = y + h + padding
                cv2.rectangle(overlay, (x1, y1), (x2, y2), (0, 0, 255), -1)  # Red highlight
        
        # Blend overlay with original image
        return cv2.addWeighted(overlay, 0.3, image, 0.7, 0)
        
    def create_pdf_report(self, image_paths, output_path):
        """Combine highlighted images into a PDF report"""
        images = [Image.open(img).convert("RGB") for img in image_paths]
        images[0].save(output_path, save_all=True, append_images=images[1:])
        
    def monitor_progress(self):
        """Monitor and update progress estimation"""
        if not self.comparison_thread.is_alive() or self.cancel_requested:
            return
            
        # Calculate time remaining
        current_progress = self.progress_bar["value"]
        if current_progress > 5:  # Wait until we have some progress
            elapsed = time.time() - self.start_time
            estimated_total = elapsed / (current_progress / 100)
            remaining = max(0, estimated_total - elapsed)
            
            # Format time string
            if remaining > 60:
                mins = int(remaining // 60)
                secs = int(remaining % 60)
                time_str = f"Estimated time remaining: {mins}m {secs}s"
            else:
                time_str = f"Estimated time remaining: {int(remaining)}s"
                
            self.time_label.config(text=time_str)
            
        # Check again in 1 second
        self.root.after(1000, self.monitor_progress)
        
    def cancel_operation(self):
        """Cancel the current comparison operation"""
        self.cancel_requested = True
        self.cancel_button.config(state=tk.DISABLED)
        self.update_status("Cancelling... Please wait")
        
    def open_results(self):
        """Open the results folder"""
        if hasattr(self, 'result_pdf_path') and os.path.exists(self.result_pdf_path):
            if platform.system() == "Windows":
                os.startfile(os.path.dirname(self.result_pdf_path))
            elif platform.system() == "Darwin":
                os.system(f'open "{os.path.dirname(self.result_pdf_path)}"')
            else:
                os.system(f'xdg-open "{os.path.dirname(self.result_pdf_path)}"')
        else:
            output_dir = self.output_entry.get()
            if os.path.exists(output_dir):
                if platform.system() == "Windows":
                    os.startfile(output_dir)
                elif platform.system() == "Darwin":
                    os.system(f'open "{output_dir}"')
                else:
                    os.system(f'xdg-open "{output_dir}"')
            else:
                messagebox.showwarning(
                    "No Results", 
                    "No comparison results available yet"
                )
                
    def toggle_ui_state(self, enabled):
        """Enable or disable UI controls"""
        state = tk.NORMAL if enabled else tk.DISABLED
        self.compare_button.config(state=state)
        self.orig_entry.config(state=state)
        self.mod_entry.config(state=state)
        self.output_entry.config(state=state)
        self.open_button.config(state=tk.NORMAL if hasattr(self, 'result_pdf_path') else tk.DISABLED)
        self.cancel_button.config(state=tk.DISABLED if enabled else tk.NORMAL)
        
    def update_progress(self, value, message):
        """Update progress bar and status message"""
        self.progress_bar["value"] = value
        self.update_status(message)
        
    def update_status(self, message):
        """Update status message"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_label.config(text=f"[{timestamp}] {message}")
        self.status_message.config(text=message)
        self.root.update()
        
    def log_result(self, message, tag=None):
        """Add a message to the results log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.results_text.insert(tk.END, f"[{timestamp}] {message}\n", tag)
        self.results_text.see(tk.END)
        self.root.update()
        
    def show_help(self):
        """Show help information in a styled dialog"""
        help_window = tk.Toplevel(self.root)
        help_window.title("SIEMENS RCC Comparator - Help")
        help_window.geometry("600x400")
        help_window.resizable(False, False)
        
        # Style the help window
        help_window.configure(bg=self.colors['background'])
        
        # Header
        header_frame = ttk.Frame(help_window)
        header_frame.pack(fill=tk.X, padx=20, pady=15)
        
        ttk.Label(header_frame, 
                 text="Help Guide", 
                 font=('Segoe UI', 16, 'bold'),
                 foreground=self.colors['primary']).pack()
        
        # Content
        content_frame = ttk.Frame(help_window)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))
        
        help_text = scrolledtext.ScrolledText(
            content_frame,
            wrap=tk.WORD,
            font=('Segoe UI', 10),
            padx=15,
            pady=15,
            bg="white"
        )
        help_text.pack(fill=tk.BOTH, expand=True)
        
        # Insert help content
        help_content = """SIEMENS RCC Comparator - Help Guide

1. SELECTING DOCUMENTS
   - Click "Browse" to select original and modified PDFs
   - Both documents must be PDF files with readable text

2. CONFIGURING COMPARISON
   - Output folder: Where results will be saved
   - Sensitivity: Adjust how strict the comparison should be
     - Higher values = only major changes
     - Lower values = more subtle differences

3. RUNNING COMPARISON
   - Click "Compare Documents" to start
   - Progress bar shows current status
   - Cancel button stops the operation

4. VIEWING RESULTS
   - Results show differences found
   - "Open Results" opens the output folder
   - Each page with changes is saved as an image
   - Complete report is generated as PDF

TIPS:
- Use high-quality PDFs for best results
- Standard fonts improve OCR accuracy
- Larger documents take more time to process

For support contact: automation.support@siemens.com
"""
        help_text.insert(tk.END, help_content)
        help_text.config(state=tk.DISABLED)
        
        # Close button
        close_button = ttk.Button(
            help_window,
            text="Close",
            command=help_window.destroy
        )
        close_button.pack(pady=(0, 15))
        
    def on_close(self):
        """Handle window close event"""
        if hasattr(self, 'comparison_thread') and self.comparison_thread.is_alive():
            if messagebox.askokcancel("Quit", "Comparison in progress. Are you sure you want to quit?"):
                self.cancel_requested = True
                self.root.destroy()
        else:
            self.root.destroy()
            
    @staticmethod
    def extract_text_data(pil_image):
        """Extract text data from image using OCR"""
        img_cv = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
        config = r'--oem 3 --psm 6'
        data = pytesseract.image_to_data(
            img_cv, 
            output_type=pytesseract.Output.DATAFRAME, 
            config=config
        )
        return data[data.text.notnull()].reset_index(drop=True), img_cv
        
    @staticmethod
    def is_text_similar(a, b, threshold=0.75):
        """Compare two text strings with given similarity threshold"""
        def normalize(text):
            text = text.strip().lower()
            text = re.sub(r'[^a-z0-9]', '', text)
            text = text.replace('0', 'o').replace('1', 'i').replace('5', 's').replace('8', 'b')
            return text
            
        norm_a = normalize(a)
        norm_b = normalize(b)
        if norm_a == norm_b:
            return True
        return SequenceMatcher(None, norm_a, norm_b).ratio() > threshold


if __name__ == "__main__":
    root = tk.Tk()
    app = SiemensRCCComparator(root)
    root.mainloop()