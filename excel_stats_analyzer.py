"""
Excel Stats Analyzer
A desktop GUI application to analyze Excel files and calculate statistics for each column.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from pathlib import Path
import threading
from typing import Dict, List, Optional
import os


class ExcelStatsAnalyzer:
    """Main application class for Excel Stats Analyzer."""
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Excel Stats Analyzer")
        self.root.geometry("700x650")
        self.root.minsize(600, 550)
        
        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self._configure_styles()
        
        # Application state
        self.input_file_path: Optional[str] = None
        self.sheet_names: List[str] = []
        self.sheet_checkboxes: Dict[str, tk.BooleanVar] = {}
        self.is_processing = False
        self.output_mode_var: Optional[tk.StringVar] = None  # 'separate' or 'consolidated'
        
        # Build UI
        self._create_widgets()
        
    def _configure_styles(self):
        """Configure custom styles for the application."""
        # Main colors
        bg_color = "#1e1e2e"
        fg_color = "#cdd6f4"
        accent_color = "#89b4fa"
        success_color = "#a6e3a1"
        button_bg = "#313244"
        
        self.root.configure(bg=bg_color)
        
        self.style.configure("TFrame", background=bg_color)
        self.style.configure("TLabel", background=bg_color, foreground=fg_color, font=("Helvetica", 11))
        self.style.configure("Title.TLabel", font=("Helvetica", 18, "bold"), foreground=accent_color)
        self.style.configure("Subtitle.TLabel", font=("Helvetica", 10), foreground="#a6adc8")
        self.style.configure("Status.TLabel", font=("Helvetica", 10), foreground=success_color)
        
        self.style.configure("TButton", 
                           font=("Helvetica", 11),
                           padding=(15, 8))
        self.style.map("TButton",
                      background=[("active", accent_color), ("!active", button_bg)],
                      foreground=[("active", bg_color), ("!active", fg_color)])
        
        self.style.configure("Accent.TButton",
                           font=("Helvetica", 11, "bold"),
                           padding=(20, 10))
        
        self.style.configure("TCheckbutton", 
                           background=bg_color, 
                           foreground=fg_color,
                           font=("Helvetica", 11))
        self.style.map("TCheckbutton",
                      background=[("active", bg_color)],
                      foreground=[("active", accent_color)])
        
        self.style.configure("TRadiobutton", 
                           background=bg_color, 
                           foreground=fg_color,
                           font=("Helvetica", 11))
        self.style.map("TRadiobutton",
                      background=[("active", bg_color)],
                      foreground=[("active", accent_color)])
        
        self.style.configure("Horizontal.TProgressbar",
                           background=accent_color,
                           troughcolor=button_bg,
                           thickness=8)
        
    def _create_widgets(self):
        """Create all UI widgets."""
        # Main container
        main_frame = ttk.Frame(self.root, padding=30)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title Section
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 25))
        
        ttk.Label(title_frame, text="📊 Excel Stats Analyzer", style="Title.TLabel").pack(anchor=tk.W)
        ttk.Label(title_frame, 
                 text="Calculate statistics for each column in your Excel sheets",
                 style="Subtitle.TLabel").pack(anchor=tk.W, pady=(5, 0))
        
        # File Selection Section
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(file_frame, text="Step 1: Select Excel File").pack(anchor=tk.W, pady=(0, 10))
        
        file_input_frame = ttk.Frame(file_frame)
        file_input_frame.pack(fill=tk.X)
        
        self.file_path_var = tk.StringVar()
        self.file_entry = ttk.Entry(file_input_frame, textvariable=self.file_path_var, state="readonly", font=("Helvetica", 10))
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        self.browse_button = ttk.Button(file_input_frame, text="Browse...", command=self._browse_file)
        self.browse_button.pack(side=tk.RIGHT)
        
        # Sheet Selection Section
        sheet_frame = ttk.Frame(main_frame)
        sheet_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        sheet_header = ttk.Frame(sheet_frame)
        sheet_header.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(sheet_header, text="Step 2: Select Sheets to Analyze").pack(side=tk.LEFT)
        
        self.select_all_var = tk.BooleanVar()
        self.select_all_checkbox = ttk.Checkbutton(
            sheet_header, 
            text="Select All", 
            variable=self.select_all_var,
            command=self._toggle_select_all
        )
        self.select_all_checkbox.pack(side=tk.RIGHT)
        
        # Sheets list container with scrollbar
        sheets_container = ttk.Frame(sheet_frame)
        sheets_container.pack(fill=tk.BOTH, expand=True)
        
        # Create canvas for scrolling
        self.sheets_canvas = tk.Canvas(sheets_container, bg="#1e1e2e", highlightthickness=0)
        scrollbar = ttk.Scrollbar(sheets_container, orient=tk.VERTICAL, command=self.sheets_canvas.yview)
        
        self.sheets_inner_frame = ttk.Frame(self.sheets_canvas)
        
        self.sheets_canvas.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.sheets_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.sheets_window = self.sheets_canvas.create_window((0, 0), window=self.sheets_inner_frame, anchor=tk.NW)
        
        # Configure canvas scrolling
        self.sheets_inner_frame.bind("<Configure>", self._on_frame_configure)
        self.sheets_canvas.bind("<Configure>", self._on_canvas_configure)
        
        # Placeholder label
        self.placeholder_label = ttk.Label(
            self.sheets_inner_frame, 
            text="No file selected. Please browse and select an Excel file.",
            style="Subtitle.TLabel"
        )
        self.placeholder_label.pack(pady=20)
        
        # Output Format Section
        output_format_frame = ttk.Frame(main_frame)
        output_format_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(output_format_frame, text="Step 3: Output Format").pack(anchor=tk.W, pady=(0, 10))
        
        self.output_mode_var = tk.StringVar(value="separate")
        
        radio_frame = ttk.Frame(output_format_frame)
        radio_frame.pack(fill=tk.X)
        
        ttk.Radiobutton(
            radio_frame,
            text="Separate Sheets (each sheet's stats in its own tab)",
            variable=self.output_mode_var,
            value="separate"
        ).pack(anchor=tk.W, padx=10, pady=2)
        
        ttk.Radiobutton(
            radio_frame,
            text="Consolidated Sheet (all stats in a single tab)",
            variable=self.output_mode_var,
            value="consolidated"
        ).pack(anchor=tk.W, padx=10, pady=2)
        
        # Progress Section
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            variable=self.progress_var,
            maximum=100,
            mode='determinate',
            style="Horizontal.TProgressbar"
        )
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))
        
        self.status_var = tk.StringVar(value="Ready")
        self.status_label = ttk.Label(progress_frame, textvariable=self.status_var, style="Status.TLabel")
        self.status_label.pack(anchor=tk.W)
        
        # Action Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        self.analyze_button = ttk.Button(
            button_frame, 
            text="🔍 Analyze & Generate Report",
            style="Accent.TButton",
            command=self._start_analysis
        )
        self.analyze_button.pack(side=tk.RIGHT)
        
        self.clear_button = ttk.Button(
            button_frame,
            text="Clear",
            command=self._clear_selection
        )
        self.clear_button.pack(side=tk.RIGHT, padx=(0, 10))
        
    def _on_frame_configure(self, event):
        """Update scroll region when frame size changes."""
        self.sheets_canvas.configure(scrollregion=self.sheets_canvas.bbox("all"))
        
    def _on_canvas_configure(self, event):
        """Update inner frame width when canvas size changes."""
        self.sheets_canvas.itemconfig(self.sheets_window, width=event.width)
        
    def _browse_file(self):
        """Open file dialog to select Excel file."""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[
                ("Excel Files", "*.xlsx *.xls *.xlsm"),
                ("All Files", "*.*")
            ]
        )
        
        if file_path:
            self._load_excel_file(file_path)
            
    def _load_excel_file(self, file_path: str):
        """Load Excel file and extract sheet names."""
        try:
            self.status_var.set("Loading file...")
            self.root.update()
            
            # Read Excel file to get sheet names
            excel_file = pd.ExcelFile(file_path)
            self.sheet_names = excel_file.sheet_names
            excel_file.close()
            
            self.input_file_path = file_path
            self.file_path_var.set(file_path)
            
            # Update sheet selection UI
            self._populate_sheets()
            
            self.status_var.set(f"Loaded: {len(self.sheet_names)} sheet(s) found")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel file:\n{str(e)}")
            self.status_var.set("Error loading file")
            
    def _populate_sheets(self):
        """Populate the sheets selection area with checkboxes."""
        # Clear existing checkboxes
        for widget in self.sheets_inner_frame.winfo_children():
            widget.destroy()
            
        self.sheet_checkboxes.clear()
        self.select_all_var.set(False)
        
        if not self.sheet_names:
            self.placeholder_label = ttk.Label(
                self.sheets_inner_frame,
                text="No sheets found in the file.",
                style="Subtitle.TLabel"
            )
            self.placeholder_label.pack(pady=20)
            return
            
        # Create checkbox for each sheet
        for i, sheet_name in enumerate(self.sheet_names):
            var = tk.BooleanVar()
            self.sheet_checkboxes[sheet_name] = var
            
            checkbox = ttk.Checkbutton(
                self.sheets_inner_frame,
                text=f"📄 {sheet_name}",
                variable=var,
                command=self._update_select_all_state
            )
            checkbox.pack(anchor=tk.W, pady=3, padx=10)
            
    def _toggle_select_all(self):
        """Toggle all sheet checkboxes based on Select All state."""
        select_all = self.select_all_var.get()
        for var in self.sheet_checkboxes.values():
            var.set(select_all)
            
    def _update_select_all_state(self):
        """Update Select All checkbox based on individual selections."""
        if not self.sheet_checkboxes:
            return
            
        all_selected = all(var.get() for var in self.sheet_checkboxes.values())
        self.select_all_var.set(all_selected)
        
    def _get_selected_sheets(self) -> List[str]:
        """Get list of selected sheet names."""
        return [name for name, var in self.sheet_checkboxes.items() if var.get()]
        
    def _clear_selection(self):
        """Clear file and sheet selection."""
        self.input_file_path = None
        self.file_path_var.set("")
        self.sheet_names = []
        self.sheet_checkboxes.clear()
        self.select_all_var.set(False)
        self.output_mode_var.set("separate")
        self.progress_var.set(0)
        self.status_var.set("Ready")
        
        # Clear sheets area
        for widget in self.sheets_inner_frame.winfo_children():
            widget.destroy()
            
        self.placeholder_label = ttk.Label(
            self.sheets_inner_frame,
            text="No file selected. Please browse and select an Excel file.",
            style="Subtitle.TLabel"
        )
        self.placeholder_label.pack(pady=20)
        
    def _start_analysis(self):
        """Start the analysis process in a separate thread."""
        if self.is_processing:
            return
            
        if not self.input_file_path:
            messagebox.showwarning("Warning", "Please select an Excel file first.")
            return
            
        selected_sheets = self._get_selected_sheets()
        if not selected_sheets:
            messagebox.showwarning("Warning", "Please select at least one sheet to analyze.")
            return
            
        # Disable buttons during processing
        self.is_processing = True
        self.analyze_button.configure(state=tk.DISABLED)
        self.browse_button.configure(state=tk.DISABLED)
        self.clear_button.configure(state=tk.DISABLED)
        
        # Run analysis in separate thread
        thread = threading.Thread(target=self._run_analysis, args=(selected_sheets,))
        thread.daemon = True
        thread.start()
        
    def _run_analysis(self, selected_sheets: List[str]):
        """Run the analysis on selected sheets."""
        try:
            total_sheets = len(selected_sheets)
            results: Dict[str, pd.DataFrame] = {}
            
            for i, sheet_name in enumerate(selected_sheets):
                # Update progress
                progress = (i / total_sheets) * 100
                self._update_ui(progress, f"Processing: {sheet_name}...")
                
                # Read sheet data
                df = pd.read_excel(
                    self.input_file_path,
                    sheet_name=sheet_name,
                    header=0,  # First row is header
                    dtype=str  # Read all as strings to preserve data
                )
                
                # Calculate statistics for this sheet
                stats_df = self._calculate_stats(df, sheet_name)
                results[sheet_name] = stats_df
                
            # Generate output file
            self._update_ui(90, "Generating output file...")
            output_mode = self.output_mode_var.get()
            output_path = self._generate_output(results, output_mode)
            
            self._update_ui(100, f"Complete! Output saved to: {os.path.basename(output_path)}")
            
            # Show success message
            self.root.after(0, lambda: messagebox.showinfo(
                "Success",
                f"Analysis complete!\n\nOutput file saved to:\n{output_path}"
            ))
            
        except Exception as e:
            self._update_ui(0, f"Error: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("Error", f"Analysis failed:\n{str(e)}"))
            
        finally:
            # Re-enable buttons
            self.root.after(0, self._enable_buttons)
            
    def _calculate_stats(self, df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
        """Calculate statistics for each column in the dataframe."""
        stats_data = []
        
        for column in df.columns:
            col_data = df[column]
            
            # Total transactions (all rows, including blanks)
            total_transactions = len(col_data)
            
            # Count of availability (non-blank values)
            # Consider NaN, None, empty string, whitespace-only as blank
            non_blank_mask = col_data.notna() & (col_data.astype(str).str.strip() != '')
            count_availability = non_blank_mask.sum()
            
            # Percentage availability
            if total_transactions > 0:
                pct_availability = (count_availability / total_transactions) * 100
            else:
                pct_availability = 0.0
                
            # Unique values (excluding blanks)
            non_blank_values = col_data[non_blank_mask]
            unique_values = non_blank_values.nunique()
            
            stats_data.append({
                'Header Name': column,
                'Total Number of Transactions': total_transactions,
                'Count of Availability': count_availability,
                '% Availability': round(pct_availability, 2),
                'No of Unique Values': unique_values
            })
            
        return pd.DataFrame(stats_data)
        
    def _generate_output(self, results: Dict[str, pd.DataFrame], output_mode: str) -> str:
        """Generate output Excel file with results."""
        # Create output filename
        input_path = Path(self.input_file_path)
        base_filename = f"{input_path.stem}_stats"
        output_path = input_path.parent / f"{base_filename}.xlsx"
        
        # Handle existing file - append number if file exists
        counter = 1
        while output_path.exists():
            output_path = input_path.parent / f"{base_filename}_{counter}.xlsx"
            counter += 1
        
        # Write results to Excel
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            if output_mode == "consolidated":
                # Combine all results into a single sheet with Sheet Name column
                consolidated_data = []
                for sheet_name, stats_df in results.items():
                    # Add Sheet Name column at the front
                    stats_with_sheet = stats_df.copy()
                    stats_with_sheet.insert(0, 'Sheet Name', sheet_name)
                    consolidated_data.append(stats_with_sheet)
                
                # Concatenate all dataframes
                consolidated_df = pd.concat(consolidated_data, ignore_index=True)
                consolidated_df.to_excel(writer, sheet_name="Consolidated Stats", index=False)
                
                # Auto-adjust column widths
                worksheet = writer.sheets["Consolidated Stats"]
                for idx, column in enumerate(consolidated_df.columns):
                    max_length = max(
                        consolidated_df[column].astype(str).apply(len).max(),
                        len(column)
                    )
                    # Add a little extra space
                    worksheet.column_dimensions[chr(65 + idx)].width = min(max_length + 2, 50)
            else:
                # Separate sheets mode (original behavior)
                for sheet_name, stats_df in results.items():
                    # Truncate sheet name if too long (Excel limit is 31 chars)
                    safe_sheet_name = sheet_name[:28] + "..." if len(sheet_name) > 31 else sheet_name
                    stats_df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                    
                    # Auto-adjust column widths
                    worksheet = writer.sheets[safe_sheet_name]
                    for idx, column in enumerate(stats_df.columns):
                        max_length = max(
                            stats_df[column].astype(str).apply(len).max(),
                            len(column)
                        )
                        # Add a little extra space
                        worksheet.column_dimensions[chr(65 + idx)].width = min(max_length + 2, 50)
                    
        return str(output_path)
        
    def _update_ui(self, progress: float, status: str):
        """Update UI elements from background thread."""
        self.root.after(0, lambda: self.progress_var.set(progress))
        self.root.after(0, lambda: self.status_var.set(status))
        
    def _enable_buttons(self):
        """Re-enable buttons after processing."""
        self.is_processing = False
        self.analyze_button.configure(state=tk.NORMAL)
        self.browse_button.configure(state=tk.NORMAL)
        self.clear_button.configure(state=tk.NORMAL)


def main():
    """Main entry point."""
    root = tk.Tk()
    app = ExcelStatsAnalyzer(root)
    root.mainloop()


if __name__ == "__main__":
    main()

