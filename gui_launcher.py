"""
Exemption Applicant Comparison Tool - GUI Launcher
===================================================
For School District / Special District Parcel Tax Exemptions

Simple GUI wrapper for the Exemption Applicant Comparison script.
Place this file in the same directory as exemption_applicant_comparison.py

Author: Brandon Van Leuven
Version: 1.0
Date: January 2026
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import os
import sys
from pathlib import Path
from datetime import datetime


class ExemptionComparisonGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Exemption Applicant Comparison Tool v1.0")
        self.root.geometry("850x780")
        self.root.resizable(False, False)
        
        # Variables
        self.input_file = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.similarity_threshold = tk.DoubleVar(value=0.85)
        self.prior_period_label = tk.StringVar(value="Prior Period")
        self.current_period_label = tk.StringVar(value="Current Period")
        
        # Build UI
        self.create_widgets()
        
    def create_widgets(self):
        """Create all GUI widgets."""
        # Title bar
        title_frame = tk.Frame(self.root, bg='#1a5276', height=100)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(
            title_frame, 
            text="Exemption Applicant Comparison", 
            font=('Arial', 20, 'bold'), 
            bg='#1a5276', 
            fg='white'
        )
        title_label.pack(pady=(15, 2))
        
        subtitle_label = tk.Label(
            title_frame, 
            text="Compare Owner Names Between Periods to Validate Exemptions", 
            font=('Arial', 10), 
            bg='#1a5276', 
            fg='#d5d8dc'
        )
        subtitle_label.pack(pady=(0, 2))
        
        version_label = tk.Label(
            title_frame, 
            text="Version 1.0", 
            font=('Arial', 9), 
            bg='#1a5276', 
            fg='#aeb6bf'
        )
        version_label.pack()
        
        # Main content frame with padding
        content_frame = tk.Frame(self.root, padx=40, pady=20, bg='#ecf0f1')
        content_frame.pack(fill='both', expand=True)
        
        # ===== Input file section =====
        input_label = tk.Label(
            content_frame, 
            text="Input Excel File (Exemption Data):", 
            font=('Arial', 10, 'bold'),
            bg='#ecf0f1'
        )
        input_label.grid(row=0, column=0, sticky='w', pady=(0, 5))
        
        input_frame = tk.Frame(content_frame, bg='#ecf0f1')
        input_frame.grid(row=1, column=0, sticky='ew', pady=(0, 15))
        
        input_entry = tk.Entry(
            input_frame, 
            textvariable=self.input_file, 
            width=72, 
            font=('Arial', 9),
            state='readonly',
            readonlybackground='white'
        )
        input_entry.pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        input_btn = tk.Button(
            input_frame, 
            text="Browse...", 
            command=self.browse_input, 
            width=12,
            bg='#3498db',
            fg='white',
            font=('Arial', 9, 'bold'),
            relief='flat',
            cursor='hand2'
        )
        input_btn.pack(side='right')
        
        # ===== Output directory section =====
        output_label = tk.Label(
            content_frame, 
            text="Output Directory:", 
            font=('Arial', 10, 'bold'),
            bg='#ecf0f1'
        )
        output_label.grid(row=2, column=0, sticky='w', pady=(0, 5))
        
        output_frame = tk.Frame(content_frame, bg='#ecf0f1')
        output_frame.grid(row=3, column=0, sticky='ew', pady=(0, 15))
        
        output_entry = tk.Entry(
            output_frame, 
            textvariable=self.output_dir, 
            width=72, 
            font=('Arial', 9),
            state='readonly',
            readonlybackground='white'
        )
        output_entry.pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        output_btn = tk.Button(
            output_frame, 
            text="Browse...", 
            command=self.browse_output, 
            width=12,
            bg='#3498db',
            fg='white',
            font=('Arial', 9, 'bold'),
            relief='flat',
            cursor='hand2'
        )
        output_btn.pack(side='right')
        
        # Separator
        separator1 = ttk.Separator(content_frame, orient='horizontal')
        separator1.grid(row=4, column=0, sticky='ew', pady=(5, 15))
        
        # ===== Configuration Options Frame =====
        config_frame = tk.LabelFrame(
            content_frame,
            text="  Configuration Options  ",
            font=('Arial', 9, 'bold'),
            bg='#ecf0f1',
            fg='#2c3e50',
            relief='groove',
            bd=2
        )
        config_frame.grid(row=5, column=0, sticky='ew', pady=(0, 15))
        
        # Period Labels
        period_label = tk.Label(
            config_frame,
            text="Period Labels (for output headers):",
            font=('Arial', 9),
            bg='#ecf0f1'
        )
        period_label.grid(row=0, column=0, columnspan=4, sticky='w', padx=10, pady=(10, 5))
        
        # Prior Period
        prior_frame = tk.Frame(config_frame, bg='#ecf0f1')
        prior_frame.grid(row=1, column=0, columnspan=2, sticky='w', padx=20, pady=(0, 5))
        
        prior_lbl = tk.Label(prior_frame, text="Prior Period:", font=('Arial', 9), bg='#ecf0f1', width=12, anchor='e')
        prior_lbl.pack(side='left')
        
        prior_entry = tk.Entry(prior_frame, textvariable=self.prior_period_label, width=20, font=('Arial', 9))
        prior_entry.pack(side='left', padx=(5, 0))
        
        # Current Period
        current_frame = tk.Frame(config_frame, bg='#ecf0f1')
        current_frame.grid(row=1, column=2, columnspan=2, sticky='w', padx=20, pady=(0, 5))
        
        current_lbl = tk.Label(current_frame, text="Current Period:", font=('Arial', 9), bg='#ecf0f1', width=12, anchor='e')
        current_lbl.pack(side='left')
        
        current_entry = tk.Entry(current_frame, textvariable=self.current_period_label, width=20, font=('Arial', 9))
        current_entry.pack(side='left', padx=(5, 0))
        
        # Hint for period labels
        hint_label = tk.Label(
            config_frame,
            text="(e.g., '202502' and '202507' for February to July comparison)",
            font=('Arial', 8, 'italic'),
            bg='#ecf0f1',
            fg='#7f8c8d'
        )
        hint_label.grid(row=2, column=0, columnspan=4, sticky='w', padx=30, pady=(0, 10))
        
        # Separator inside config frame
        sep_inner = ttk.Separator(config_frame, orient='horizontal')
        sep_inner.grid(row=3, column=0, columnspan=4, sticky='ew', padx=10, pady=5)
        
        # Similarity Threshold
        threshold_label = tk.Label(
            config_frame,
            text="Name Similarity Threshold:",
            font=('Arial', 9),
            bg='#ecf0f1'
        )
        threshold_label.grid(row=4, column=0, sticky='w', padx=10, pady=(5, 5))
        
        # Threshold slider frame
        slider_frame = tk.Frame(config_frame, bg='#ecf0f1')
        slider_frame.grid(row=5, column=0, columnspan=4, sticky='ew', padx=20, pady=(0, 10))
        
        # Current value label
        self.threshold_value_label = tk.Label(
            slider_frame,
            text=f"{self.similarity_threshold.get():.0%}",
            font=('Arial', 11, 'bold'),
            bg='#ecf0f1',
            fg='#1a5276',
            width=6
        )
        self.threshold_value_label.pack(side='left')
        
        # Slider
        threshold_slider = tk.Scale(
            slider_frame,
            from_=0.5,
            to=1.0,
            resolution=0.05,
            orient='horizontal',
            variable=self.similarity_threshold,
            command=self.update_threshold_label,
            length=350,
            bg='#ecf0f1',
            highlightthickness=0,
            troughcolor='#bdc3c7',
            activebackground='#3498db',
            showvalue=False
        )
        threshold_slider.pack(side='left', padx=(10, 10))
        
        # Threshold hint
        threshold_hint = tk.Label(
            slider_frame,
            text="(0.85 = Conservative)",
            font=('Arial', 8, 'italic'),
            bg='#ecf0f1',
            fg='#7f8c8d'
        )
        threshold_hint.pack(side='left')
        
        # Threshold explanation
        threshold_explain = tk.Label(
            config_frame,
            text="Higher = stricter matching (fewer false keeps). Lower = more lenient (fewer false removes).",
            font=('Arial', 8),
            bg='#ecf0f1',
            fg='#7f8c8d'
        )
        threshold_explain.grid(row=6, column=0, columnspan=4, sticky='w', padx=30, pady=(0, 10))
        
        # Separator
        separator2 = ttk.Separator(content_frame, orient='horizontal')
        separator2.grid(row=6, column=0, sticky='ew', pady=(5, 15))
        
        # ===== Process Button =====
        self.process_btn = tk.Button(
            content_frame,
            text="▶  PROCESS FILE",
            command=self.process_file,
            width=25,
            height=2,
            bg='#27ae60',
            fg='white',
            font=('Arial', 12, 'bold'),
            relief='flat',
            cursor='hand2'
        )
        self.process_btn.grid(row=7, column=0, pady=(0, 15))
        
        # Progress bar
        self.progress = ttk.Progressbar(
            content_frame,
            mode='indeterminate',
            length=500
        )
        self.progress.grid(row=8, column=0, pady=(0, 15))
        
        # ===== Log Output =====
        log_label = tk.Label(
            content_frame,
            text="Processing Log:",
            font=('Arial', 9, 'bold'),
            bg='#ecf0f1'
        )
        log_label.grid(row=9, column=0, sticky='w', pady=(0, 5))
        
        self.log_text = scrolledtext.ScrolledText(
            content_frame,
            width=95,
            height=12,
            font=('Consolas', 9),
            bg='#2c3e50',
            fg='#ecf0f1',
            insertbackground='white'
        )
        self.log_text.grid(row=10, column=0, sticky='ew', pady=(0, 10))
        
        # Initial log message
        self.log("Exemption Applicant Comparison Tool ready.")
        self.log("Select an input file and output directory to begin.")
        self.log("")
        self.log("This tool compares owner names between periods to determine")
        self.log("if exemptions should be kept or removed based on ownership changes.")
    
    def update_threshold_label(self, value):
        """Update the threshold display label when slider moves."""
        self.threshold_value_label.config(text=f"{float(value):.0%}")
    
    def browse_input(self):
        """Browse for input Excel file."""
        filename = filedialog.askopenfilename(
            title="Select Exemption Data File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ],
            initialdir=os.getcwd()
        )
        if filename:
            self.input_file.set(filename)
            self.log(f"✓ Input file selected: {os.path.basename(filename)}")
            
            # Auto-set output directory to same as input
            if not self.output_dir.get():
                self.output_dir.set(os.path.dirname(filename))
                self.log(f"✓ Output directory set to: {os.path.dirname(filename)}")
    
    def browse_output(self):
        """Browse for output directory."""
        directory = filedialog.askdirectory(
            title="Select Output Directory",
            initialdir=self.output_dir.get() or os.getcwd()
        )
        if directory:
            self.output_dir.set(directory)
            self.log(f"✓ Output directory selected: {directory}")
    
    def log(self, message):
        """Add message to log with timestamp."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert('end', f"[{timestamp}] {message}\n")
        self.log_text.see('end')
        self.root.update_idletasks()
    
    def process_file(self):
        """Process the selected file using the comparison tool."""
        # Validate inputs
        if not self.input_file.get():
            messagebox.showerror("Input Required", "Please select an input Excel file.")
            return
        
        if not self.output_dir.get():
            messagebox.showerror("Output Required", "Please select an output directory.")
            return
        
        if not os.path.exists(self.input_file.get()):
            messagebox.showerror("File Not Found", "The selected input file does not exist.")
            return
        
        if not os.path.exists(self.output_dir.get()):
            messagebox.showerror("Directory Not Found", "The selected output directory does not exist.")
            return
        
        # Clear log
        self.log_text.delete('1.0', 'end')
        
        # Disable process button
        self.process_btn.config(state='disabled', text='PROCESSING...', bg='#95a5a6')
        
        # Start progress bar
        self.progress.start(10)
        
        try:
            self.log("=" * 70)
            self.log("EXEMPTION APPLICANT COMPARISON - Processing Started")
            self.log("=" * 70)
            self.log("")
            
            # Import the main processing module
            try:
                # Add current directory to path
                current_dir = os.path.dirname(os.path.abspath(__file__))
                if current_dir not in sys.path:
                    sys.path.insert(0, current_dir)
                
                # Import the processing module
                import exemption_applicant_comparison as eac
                
                self.log("✓ Loaded exemption comparison module")
                
                # Report library status
                if eac.HAVE_JELLYFISH:
                    self.log("  → Using jellyfish for Jaro-Winkler similarity")
                else:
                    self.log("  → Using difflib fallback (install jellyfish for better results)")
                
                if eac.HAVE_RAPIDFUZZ:
                    self.log("  → Using rapidfuzz for token-based matching")
                else:
                    self.log("  → Using basic token matching (install rapidfuzz for better results)")
                
                # Generate output path
                input_basename = os.path.basename(self.input_file.get())
                base = os.path.splitext(input_basename)[0]
                timestamp = datetime.now().strftime("%Y-%m-%d-%H%M%S")
                output_filename = f"{base}_Comparison_Results_{timestamp}.xlsx"
                output_path = os.path.join(self.output_dir.get(), output_filename)
                
                self.log("")
                self.log(f"→ Input file: {input_basename}")
                self.log(f"→ Output file: {output_filename}")
                self.log(f"→ Similarity threshold: {self.similarity_threshold.get():.0%}")
                self.log(f"→ Prior period label: {self.prior_period_label.get()}")
                self.log(f"→ Current period label: {self.current_period_label.get()}")
                self.log("")
                
                # Initialize processor
                processor = eac.ExemptionApplicantComparison(
                    input_file=self.input_file.get(),
                    output_file=output_path,
                    similarity_threshold=self.similarity_threshold.get(),
                    prior_period_label=self.prior_period_label.get(),
                    current_period_label=self.current_period_label.get()
                )
                
                # Load data
                self.log("→ Loading data...")
                processor.load_data()
                self.log(f"  ✓ Loaded {len(processor.data):,} records")
                
                # Check for empty file
                if len(processor.data) == 0:
                    self.log("")
                    self.log("⚠ WARNING: Input file contains no data rows (only headers)")
                    self.log("  Please check your input file has data to process.")
                    messagebox.showwarning(
                        "No Data",
                        "The input file contains no data rows (only headers).\n\n"
                        "Please verify your input file contains data to process."
                    )
                    return
                
                # Show mapped columns
                self.log("  ✓ Column mapping:")
                for std_name, orig_col in processor.column_map.items():
                    if orig_col:
                        self.log(f"      {std_name} → {orig_col}")
                
                # Process comparison
                self.log("")
                self.log("→ Processing name comparisons...")
                processor.process_comparison()
                
                # Get results summary
                results = processor.results
                keep_count = (results['recommend_keep'] == 'Keep').sum()
                remove_count = (results['recommend_keep'] == 'Remove').sum()
                # Handle empty dataframe case
                if len(results) > 0:
                    review_count = results['confidence_level'].astype(str).str.contains('REVIEW', na=False).sum()
                    heirs_count = results['heirs_flag'].sum()
                else:
                    review_count = 0
                    heirs_count = 0
                
                self.log(f"  ✓ Recommend Keep: {keep_count:,}")
                self.log(f"  ✓ Recommend Remove: {remove_count:,}")
                self.log(f"  ✓ Needs Review: {review_count:,}")
                self.log(f"  ✓ Heirs/Estate flagged: {heirs_count:,}")
                
                # Save results
                self.log("")
                self.log("→ Saving results...")
                self.log("  → Creating 'All Records' tab...")
                self.log("  → Creating 'Recommend Keep' tab...")
                self.log("  → Creating 'Recommend Remove' tab...")
                if review_count > 0:
                    self.log("  → Creating 'Needs Review' tab...")
                if heirs_count > 0:
                    self.log("  → Creating 'Heirs-Estate Flagged' tab...")
                self.log("  → Creating 'Summary' tab...")
                
                processor.save_results()
                
                self.log("")
                self.log("=" * 70)
                self.log("✓✓✓ SUCCESS! Output file created. ✓✓✓")
                self.log("=" * 70)
                self.log("")
                self.log(f"Output location: {output_path}")
                
                # Show success message
                result = messagebox.askyesno(
                    "Success!", 
                    f"File processed successfully!\n\n"
                    f"Results:\n"
                    f"  • Total Records: {len(results):,}\n"
                    f"  • Recommend Keep: {keep_count:,}\n"
                    f"  • Recommend Remove: {remove_count:,}\n"
                    f"  • Needs Review: {review_count:,}\n"
                    f"  • Heirs/Estate Flagged: {heirs_count:,}\n\n"
                    f"Output saved to:\n{output_path}\n\n"
                    f"Would you like to open the output folder?",
                    icon='info'
                )
                
                if result:
                    # Open output directory
                    if sys.platform == 'win32':
                        os.startfile(self.output_dir.get())
                    elif sys.platform == 'darwin':
                        os.system(f'open "{self.output_dir.get()}"')
                    else:
                        os.system(f'xdg-open "{self.output_dir.get()}"')
            
            except ImportError as e:
                self.log(f"✗ ERROR: Could not import exemption_applicant_comparison module")
                self.log(f"  Make sure 'exemption_applicant_comparison.py' is in the same directory")
                self.log(f"  Error details: {str(e)}")
                messagebox.showerror(
                    "Module Not Found",
                    "Could not find 'exemption_applicant_comparison.py' in the same directory.\n\n"
                    "Please make sure both files are in the same folder."
                )
                raise
        
        except Exception as e:
            self.log("")
            self.log(f"✗✗✗ ERROR: {str(e)}")
            import traceback
            self.log("")
            self.log("Full error traceback:")
            self.log(traceback.format_exc())
            messagebox.showerror(
                "Processing Error", 
                f"An error occurred during processing:\n\n{str(e)}\n\n"
                f"Check the log for details."
            )
        
        finally:
            # Stop progress bar and re-enable button
            self.progress.stop()
            self.process_btn.config(state='normal', text='▶  PROCESS FILE', bg='#27ae60')


def main():
    """Main entry point for the GUI."""
    root = tk.Tk()
    
    # Set window icon if available
    try:
        root.iconbitmap('icon.ico')
    except:
        pass
    
    app = ExemptionComparisonGUI(root)
    
    # Center window on screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()


if __name__ == "__main__":
    main()
