#!/usr/bin/env python3
"""
Marc Mentat - Batch/Bash Generator
Vibracoustic - European FEA Department
VERSION 1.0.1

Maintains the original Windows tool logic and adds Target OS option
to generate .bat (Windows) or .sh (Linux) scripts. In Linux mode the
executable field is fixed and the Cancel button continues to obey the safe Mentat flow.

"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import glob
import sys
import subprocess

# Official paths (Windows)
OFFICIAL_NOMINAL_STRAIN_PATH = (
    "X:\\VC-Marc_Post\\Marc_Subroutine\\nominal_strain_R2020"
)
OFFICIAL_NOMINAL_STRAIN_DISPLAY = "Marc's Official nominal strain R2020"

# Linux constants
LINUX_PROGRAM_SAVED = "/home/10_MT_tools/emSP/emSPr01a_R2020.marc"
LINUX_EXECUTABLE_LABEL = "Marc executable: marc2020"
TITLE_BG_WINDOWS = "#d4542a"
TITLE_BG_LINUX = "#4CAF50"

# Color palette for different folders (10 distinct colors)
FOLDER_COLORS = [
    "#1E88E5",  # Blue
    "#43A047",  # Green
    "#E53935",  # Red
    "#8E24AA",  # Purple
    "#FB8C00",  # Orange
    "#00ACC1",  # Cyan
    "#D81B60",  # Pink
    "#5E35B1",  # Deep Purple
    "#00897B",  # Teal
    "#FFB300",  # Amber
]

# Guideline PDF path
GUIDELINE_PDF_PATH = r"\\frafil002\VC_FEA\VC-Marc_Post\Marc_Tools_Guideline\Marc_Batch_Generator_User_Guide_v1.html"


def open_guideline_pdf():
    """Open the guideline PDF file"""
    try:
        if os.path.exists(GUIDELINE_PDF_PATH):
            os.startfile(GUIDELINE_PDF_PATH)
        else:
            messagebox.showwarning("File Not Found", 
                f"Guideline PDF not found at:\n{GUIDELINE_PDF_PATH}")
    except Exception as e:
        messagebox.showerror("Error", f"Could not open guideline: {e}")


def get_user_temp_directory():
    """Get the user-specific temporary directory for Marc Python files"""
    try:
        local_appdata = os.environ.get('LOCALAPPDATA')
        if local_appdata:
            user_temp_dir = os.path.join(local_appdata, 'MARC_PYTHON_TEMP')
        else:
            username = os.environ.get('USERNAME', 'DefaultUser')
            user_temp_dir = f"C:\\Users\\{username}\\AppData\\Local\\MARC_PYTHON_TEMP"
        os.makedirs(user_temp_dir, exist_ok=True)
        return user_temp_dir
    except Exception:
        fallback_dir = os.path.join(os.getcwd(), 'MARC_PYTHON_TEMP')
        os.makedirs(fallback_dir, exist_ok=True)
        return fallback_dir


def create_default_history_file(history_file_path):
    """Create default history file with predefined paths if it doesn't exist."""
    try:
        if not os.path.exists(history_file_path):
            default_run_marc = (
                "C:/Program Files/MSC.Software/Marc/2020.1.0/marc2020.1/tools/run_marc.bat"
            )
            default_nominal_strain = OFFICIAL_NOMINAL_STRAIN_PATH
            with open(history_file_path, 'w') as f:
                f.write("Windows\n")
                f.write(default_run_marc + "\n")
                f.write(default_nominal_strain + "\n")
            return True
        return False
    except Exception:
        return False


def get_script_directory():
    """Get script directory with user AppData integration"""
    try:
        return get_user_temp_directory()
    except Exception:
        try:
            if '__file__' in globals() and __file__:
                return os.path.dirname(os.path.abspath(__file__))
        except Exception:
            pass
        try:
            if len(sys.argv) > 0 and sys.argv[0]:
                return os.path.dirname(os.path.abspath(sys.argv[0]))
        except Exception:
            pass
        try:
            return get_user_temp_directory()
        except Exception:
            pass
        return os.getcwd()


def is_mentat_context():
    """Check if running in Mentat context"""
    try:
        import py_mentat  # noqa: F401
        return True
    except ImportError:
        return False


def setup_tkinter_for_mentat():
    """Setup Tkinter to work properly in Mentat context"""
    try:
        if os.name == 'nt':
            os.environ['TK_SILENCE_DEPRECATION'] = '1'
        if hasattr(tk, 'call'):
            try:
                tk.call('tk', 'scaling', 1.0)
            except Exception:
                pass
    except Exception:
        pass


class MarcMentatGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("European FEA Department - Vibracoustic")
        self.root.geometry("900x750")

        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.after_idle(lambda: self.root.attributes('-topmost', False))

        # State variables
        self.dat_folder = tk.StringVar()
        self.run_marc_path = tk.StringVar()
        self.nominal_strain_path = tk.StringVar()
        self.cpu_count = tk.StringVar(value="1")
        self.target_os = tk.StringVar(value="Windows")  # Windows or Linux

        self.restart_counter = 1
        self.loadcase_counter = 1
        self.dat_files = []
        # Modified: selected_files now stores (filename, file_type, number, folder_path)
        self.selected_files = []
        
        # Folder color mapping: folder_path -> color
        self.folder_colors = {}
        self.next_color_index = 0

        self.script_dir = get_script_directory()
        self.history_file = os.path.join(self.script_dir, "History_Address_Files.txt")
        create_default_history_file(self.history_file)

        self.create_widgets()
        self.load_history()
        self.apply_target_os(style_only=True)

        self.root.update()
        self.root.deiconify()

    def get_folder_color(self, folder_path):
        """Get or assign a color for a folder path"""
        if folder_path not in self.folder_colors:
            color = FOLDER_COLORS[self.next_color_index % len(FOLDER_COLORS)]
            self.folder_colors[folder_path] = color
            self.next_color_index += 1
        return self.folder_colors[folder_path]

    def create_widgets(self):
        # Create main canvas with scrollbar
        self.canvas = tk.Canvas(self.root)
        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas, padding="10")
        
        # Configure scrollable frame
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        # Create window inside canvas
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        # Configure canvas
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Bind canvas resize to adjust inner frame width
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        
        # Enable mouse wheel scrolling
        self.canvas.bind_all("<MouseWheel>", self.on_mousewheel)
        self.canvas.bind_all("<Button-4>", self.on_mousewheel)  # Linux scroll up
        self.canvas.bind_all("<Button-5>", self.on_mousewheel)  # Linux scroll down
        
        # Pack scrollbar and canvas
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        
        # Use scrollable_frame as main_frame
        main_frame = self.scrollable_frame
        main_frame.columnconfigure(1, weight=1)

        # Title bar
        self.title_frame = tk.Frame(main_frame, bg=TITLE_BG_WINDOWS,
                                    relief='raised', bd=3)
        self.title_frame.grid(row=0, column=0, columnspan=3,
                              sticky=(tk.W, tk.E), pady=(0, 15))
        self.title_frame.columnconfigure(0, weight=1)
        
        # Title content frame for layout (Guideline button + centered title)
        title_side_padding = 95
        self.title_content = tk.Frame(self.title_frame, bg=TITLE_BG_WINDOWS)
        self.title_content.grid(row=0, column=0, sticky=(tk.W, tk.E),
                                padx=(title_side_padding, title_side_padding), pady=5)
        self.title_content.columnconfigure(0, weight=1)
        
        # Keep the button pinned to the title bar's top-right corner.
        self.guideline_btn = tk.Button(self.title_frame, text="Guideline", command=open_guideline_pdf,
                                       font=('Arial', 8, 'bold'), bg='#FFEB3B', fg='#333333',
                                       relief='raised', bd=2, cursor='hand2', padx=8, pady=2,
                                       activebackground='#FFC107', activeforeground='#333333')
        self.guideline_btn.place(relx=1.0, x=-8, y=8, anchor='ne')
        
        # Centered title frame
        title_center = tk.Frame(self.title_content, bg=TITLE_BG_WINDOWS)
        title_center.grid(row=0, column=0, sticky=(tk.W, tk.E))
        self.title_center = title_center
        
        current_user = os.environ.get('USERNAME', 'DefaultUser')
        base_title = "Marc Mentat - Batch Generator"
        self.title_label_1 = tk.Label(title_center, text=base_title,
                                      font=('Arial', 18, 'bold'),
                                      fg='white', bg=TITLE_BG_WINDOWS, pady=5)
        self.title_label_1.pack(anchor='center')
        self.title_label_2 = tk.Label(title_center,
                                      text=f"User: {current_user}",
                                      font=('Arial', 12),
                                      fg='white', bg=TITLE_BG_WINDOWS, pady=2)
        self.title_label_2.pack(anchor='center')
        self.title_label_3 = tk.Label(title_center,
                                      text="Vibracoustic - European FEA Department",
                                      font=('Arial', 12),
                                      fg='white', bg=TITLE_BG_WINDOWS, pady=5)
        self.title_label_3.pack(anchor='center')

        # .dat folder
        ttk.Label(main_frame, text="Folder with .dat files:").grid(row=1, column=0,
                                                                   sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.dat_folder, width=50)\
            .grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_dat_folder)\
            .grid(row=1, column=2, padx=5)

        # File lists
        lists_frame = ttk.Frame(main_frame)
        lists_frame.grid(row=2, column=0, columnspan=3,
                         sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        lists_frame.columnconfigure(0, weight=1)
        lists_frame.columnconfigure(2, weight=1)
        lists_frame.rowconfigure(1, weight=1)
        ttk.Label(lists_frame, text="Available .dat files:").grid(row=0, column=0, sticky=tk.W)

        dat_frame = ttk.Frame(lists_frame)
        dat_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        dat_frame.columnconfigure(0, weight=1)
        dat_frame.rowconfigure(0, weight=1)
        self.dat_listbox = tk.Listbox(dat_frame, height=12)
        self.dat_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        dat_scrollbar = ttk.Scrollbar(dat_frame, orient="vertical",
                                      command=self.dat_listbox.yview)
        dat_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.dat_listbox.configure(yscrollcommand=dat_scrollbar.set)

        buttons_frame = ttk.Frame(lists_frame)
        buttons_frame.grid(row=1, column=1, padx=10, pady=20)
        tk.Button(buttons_frame, text="Restart >>", command=self.add_restart,
                  font=('Arial', 10, 'bold'), bg='#d4542a', fg='white',
                  width=10, relief='raised', bd=3).grid(row=0, column=0, pady=8)
        tk.Button(buttons_frame, text="Loadcase >>", command=self.add_loadcase,
                  font=('Arial', 10, 'bold'), bg='#2196F3', fg='white',
                  width=10, relief='raised', bd=3).grid(row=1, column=0, pady=8)
        tk.Button(buttons_frame, text="<< Remove", command=self.remove_selected,
                  font=('Arial', 10, 'bold'), bg='#f44336', fg='white',
                  width=10, relief='raised', bd=3).grid(row=2, column=0, pady=8)

        ttk.Label(lists_frame, text="Selected files:").grid(row=0, column=2, sticky=tk.W)

        selected_frame = ttk.Frame(lists_frame)
        selected_frame.grid(row=1, column=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
        selected_frame.columnconfigure(0, weight=1)
        selected_frame.rowconfigure(0, weight=1)
        self.selected_listbox = tk.Listbox(selected_frame, height=12)
        self.selected_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        selected_scrollbar = ttk.Scrollbar(selected_frame, orient="vertical",
                                           command=self.selected_listbox.yview)
        selected_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.selected_listbox.configure(yscrollcommand=selected_scrollbar.set)
        
        # Color legend frame
        self.legend_frame = ttk.LabelFrame(lists_frame, text="Folder Legend", padding="5")
        self.legend_frame.grid(row=2, column=2, sticky=(tk.W, tk.E), pady=(5, 0), padx=(5, 0))

        # Target OS (centered)
        os_row = ttk.Frame(main_frame)
        os_row.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 8))
        os_row.columnconfigure(0, weight=1)
        os_row.columnconfigure(1, weight=0)
        os_row.columnconfigure(2, weight=1)
        self.os_frame = ttk.LabelFrame(os_row, text="Target OS", padding="10")
        self.os_frame.grid(row=0, column=1, sticky="")
        ttk.Radiobutton(self.os_frame, text="Windows (.bat)",
                        variable=self.target_os, value="Windows",
                        command=self.on_os_toggle).grid(row=0, column=0, padx=5, sticky=tk.W)
        ttk.Radiobutton(self.os_frame, text="Linux (.sh)",
                        variable=self.target_os, value="Linux",
                        command=self.on_os_toggle).grid(row=0, column=1, padx=5, sticky=tk.W)

        # Paths frame
        paths_frame = ttk.LabelFrame(main_frame, text="File Paths", padding="10")
        paths_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        paths_frame.columnconfigure(1, weight=1)

        self.run_marc_label = ttk.Label(paths_frame, text="run_marc.bat:")
        self.run_marc_label.grid(row=0, column=0, sticky=tk.W, pady=5)
        self.run_marc_entry = ttk.Entry(paths_frame, textvariable=self.run_marc_path, width=50)
        self.run_marc_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        self.run_marc_button = ttk.Button(paths_frame, text="Browse", command=self.browse_run_marc)
        self.run_marc_button.grid(row=0, column=2, padx=5)

        self.marc_hint_label = ttk.Label(paths_frame, text=LINUX_EXECUTABLE_LABEL)

        self.nominal_label = ttk.Label(paths_frame, text="nominal_strain_R2020.f:")
        self.nominal_label.grid(row=1, column=0, sticky=tk.W, pady=5)
        self.nominal_entry = ttk.Entry(paths_frame, textvariable=self.nominal_strain_path, width=50)
        self.nominal_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5)
        self.nominal_button = ttk.Button(paths_frame, text="Browse", command=self.browse_nominal_strain)
        self.nominal_button.grid(row=1, column=2, padx=5)

        # CPUS
        cpu_frame = ttk.Frame(main_frame)
        cpu_frame.grid(row=5, column=0, columnspan=3, pady=10)
        ttk.Label(cpu_frame, text="CPUS:").grid(row=0, column=0, padx=5)
        cpu_combo = ttk.Combobox(cpu_frame, textvariable=self.cpu_count,
                                 values=[str(i) for i in range(1, 16)],
                                 state="readonly", width=10)
        cpu_combo.grid(row=0, column=1, padx=5)
        cpu_combo.set("1")

        # Main buttons
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=6, column=0, columnspan=3, pady=20)
        self.export_button = tk.Button(action_frame, text="Export batch file",
                                       command=self.export_batch,
                                       font=('Arial', 10, 'bold'), bg='#4CAF50',
                                       fg='white', width=15, relief='raised', bd=3)
        self.export_button.grid(row=0, column=0, padx=15)
        tk.Button(action_frame, text="Cancel", command=self.cancel,
                  font=('Arial', 10), bg='#757575', fg='white',
                  width=12, relief='raised', bd=3).grid(row=0, column=1, padx=15)

        # Credits
        author_frame = tk.Frame(main_frame, bg='#f0f0f0')
        author_frame.grid(row=7, column=0, columnspan=3, pady=(15, 5))
        tk.Label(author_frame, text="Author: Leandro Barbosa",
                 font=('Arial', 9), bg='#f0f0f0', fg='#666666',
                 justify=tk.CENTER).pack()

        main_frame.rowconfigure(2, weight=1)
    
    def on_canvas_configure(self, event):
        """Adjust the inner frame width when canvas is resized"""
        self.canvas.itemconfig(self.canvas_window, width=event.width)
    
    def on_mousewheel(self, event):
        """Handle mouse wheel scrolling"""
        if event.num == 4:  # Linux scroll up
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:  # Linux scroll down
            self.canvas.yview_scroll(1, "units")
        else:  # Windows
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # ---------------- Original handlers ---------------- #

    def browse_dat_folder(self):
        try:
            self.root.attributes('-topmost', True)
            folder = filedialog.askdirectory(title="Select folder with .dat files",
                                             parent=self.root)
            self.root.attributes('-topmost', False)
            if folder:
                self.dat_folder.set(folder)
                self.load_dat_files()
        except Exception as e:
            messagebox.showerror("Error", f"Error selecting folder: {e}")

    def browse_run_marc(self):
        try:
            self.root.attributes('-topmost', True)
            file_path = filedialog.askopenfilename(
                title="Select run_marc.bat",
                filetypes=[("Batch files", "*.bat"), ("All files", "*.*")],
                parent=self.root
            )
            self.root.attributes('-topmost', False)
            if file_path:
                self.run_marc_path.set(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Error selecting run_marc.bat: {e}")

    def get_nominal_strain_display_path(self):
        current_path = self.nominal_strain_path.get()
        if current_path == OFFICIAL_NOMINAL_STRAIN_PATH:
            return OFFICIAL_NOMINAL_STRAIN_DISPLAY
        return current_path

    def update_nominal_strain_display(self):
        self.nominal_strain_path.set(self.get_nominal_strain_display_path())

    def browse_nominal_strain(self):
        try:
            self.root.attributes('-topmost', True)
            file_path = filedialog.askopenfilename(
                title="Select nominal_strain_R2020.f",
                filetypes=[("Fortran files", "*.f"), ("All files", "*.*")],
                parent=self.root
            )
            self.root.attributes('-topmost', False)
            if file_path:
                path_without_extension = os.path.splitext(file_path)[0]
                self.nominal_strain_path.set(path_without_extension)
                self.update_nominal_strain_display()
                if not os.path.exists(os.path.dirname(file_path)):
                    messagebox.showerror("Error", "Selected directory does not exist!")
        except Exception as e:
            messagebox.showerror("Error", f"Error selecting nominal_strain_R2020.f: {e}")

    def load_dat_files(self):
        folder = self.dat_folder.get()
        if not folder:
            return
        try:
            self.dat_files = []
            self.dat_listbox.delete(0, tk.END)
            for file_path in glob.glob(os.path.join(folder, "*.dat")):
                filename = os.path.basename(file_path)
                self.dat_files.append(filename)
                self.dat_listbox.insert(tk.END, filename)
        except Exception as e:
            messagebox.showerror("Error", f"Error loading .dat files: {e}")

    def update_legend(self):
        """Update the folder color legend"""
        # Clear existing legend items
        for widget in self.legend_frame.winfo_children():
            widget.destroy()
        
        # Get unique folders from selected files
        folders_in_use = set()
        for item in self.selected_files:
            if len(item) >= 4:
                folders_in_use.add(item[3])
        
        # Create legend entries for folders in use
        for i, folder in enumerate(sorted(folders_in_use)):
            color = self.folder_colors.get(folder, "#000000")
            # Get just the folder name for display
            folder_name = os.path.basename(folder) if folder else "Unknown"
            if len(folder_name) > 30:
                folder_name = "..." + folder_name[-27:]
            
            legend_item = tk.Frame(self.legend_frame)
            legend_item.pack(anchor=tk.W, pady=1)
            
            color_box = tk.Label(legend_item, text="■", fg=color, font=('Arial', 12))
            color_box.pack(side=tk.LEFT, padx=(0, 5))
            
            folder_label = tk.Label(legend_item, text=folder_name, font=('Arial', 8))
            folder_label.pack(side=tk.LEFT)

    def renumber_and_refresh_display(self):
        restart_counter = 1
        loadcase_counter = 1
        self.selected_listbox.delete(0, tk.END)

        for i, item in enumerate(self.selected_files):
            # Handle both old format (filename, file_type, number) and new format (filename, file_type, number, folder_path)
            if len(item) == 3:
                filename, file_type, _old_number = item
                folder_path = self.dat_folder.get()  # Use current folder as fallback
            else:
                filename, file_type, _old_number, folder_path = item
            
            if file_type == "Restart":
                new_number = restart_counter
                restart_counter += 1
            else:
                new_number = loadcase_counter
                loadcase_counter += 1
            self.selected_files[i] = (filename, file_type, new_number, folder_path)

        for i, (filename, file_type, number, folder_path) in enumerate(self.selected_files):
            color = self.get_folder_color(folder_path)
            
            if file_type == "Restart":
                display_text = f"{filename} - R{number:02d}"
            else:
                is_last_child = True
                for j in range(i + 1, len(self.selected_files)):
                    if self.selected_files[j][1] == "Restart":
                        break
                    elif self.selected_files[j][1] == "Loadcase":
                        is_last_child = False
                        break
                symbol = "   +-- " if is_last_child else "   |-- "
                display_text = f"{symbol}{filename} - L{number:02d}"
            
            self.selected_listbox.insert(tk.END, display_text)
            self.selected_listbox.itemconfig(tk.END, fg=color)

        self.restart_counter = restart_counter
        self.loadcase_counter = loadcase_counter
        
        # Update the legend
        self.update_legend()

    def add_restart(self):
        selection = self.dat_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a .dat file first!")
            return
        filename = self.dat_files[selection[0]]
        folder_path = self.dat_folder.get()
        # Store filename without extension and the folder path
        self.selected_files.append((os.path.splitext(filename)[0], "Restart", self.restart_counter, folder_path))
        self.restart_counter += 1
        self.renumber_and_refresh_display()

    def add_loadcase(self):
        selection = self.dat_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a .dat file first!")
            return
        filename = self.dat_files[selection[0]]
        folder_path = self.dat_folder.get()
        # Store filename without extension and the folder path
        self.selected_files.append((os.path.splitext(filename)[0], "Loadcase", self.loadcase_counter, folder_path))
        self.loadcase_counter += 1
        self.renumber_and_refresh_display()

    def remove_selected(self):
        selection = self.selected_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select an item from 'Selected files' list first!")
            return
        self.selected_files.pop(selection[0])
        self.renumber_and_refresh_display()

    def cancel(self):
        """
        FIXED: Cancel method adapted to prevent Mentat freezing.
        
        In Mentat context:
        - Uses withdraw() to hide the window
        - Uses destroy() to clean up resources
        - Does NOT use quit() which interferes with Mentat's loop
        
        Outside Mentat context:
        - Uses quit() to exit mainloop
        - Uses destroy() to close the window
        """
        try:
            # Unbind mousewheel events to prevent errors after destroy
            try:
                self.canvas.unbind_all("<MouseWheel>")
                self.canvas.unbind_all("<Button-4>")
                self.canvas.unbind_all("<Button-5>")
            except Exception:
                pass
            
            if is_mentat_context():
                # In Mentat: just hide and destroy the window without calling quit()
                self.root.withdraw()
                self.root.update_idletasks()
                self.root.destroy()
            else:
                # Outside Mentat: normal behavior
                self.root.quit()
                self.root.destroy()
        except Exception as e:
            # Safety fallback
            try:
                self.root.destroy()
            except Exception:
                pass

    def get_unique_folders(self):
        """Get set of unique folders from selected files"""
        folders = set()
        for item in self.selected_files:
            if len(item) >= 4:
                folders.add(item[3])
        return folders

    def export_batch(self):
        if not self.validate_inputs():
            return
        try:
            self.save_history()
            
            # Check if files are from multiple folders
            unique_folders = self.get_unique_folders()
            multiple_folders = len(unique_folders) > 1
            
            if self.target_os.get() == "Windows":
                # Windows mode: offers automatic execution
                content = self.generate_batch_content()
                
                if multiple_folders:
                    # Ask user where to save when multiple folders are involved
                    self.root.attributes('-topmost', True)
                    path = filedialog.asksaveasfilename(
                        title="Save Batch File As",
                        defaultextension=".bat",
                        filetypes=[("Batch files", "*.bat"), ("All files", "*.*")],
                        initialfile="Marc_Mentat_Start_01.bat",
                        parent=self.root
                    )
                    self.root.attributes('-topmost', False)
                    
                    if not path:  # User cancelled
                        return
                    
                    folder = os.path.dirname(path)
                    filename = os.path.basename(path)
                else:
                    # Use folder of first selected file for output
                    folder = self.selected_files[0][3] if self.selected_files else self.dat_folder.get()
                    filename = self.get_next_batch_filename(folder)
                    path = os.path.join(folder, filename)
                
                with open(path, 'w') as f:
                    f.write(content)
                
                if messagebox.askyesno(
                    "Batch File Created",
                    f"Batch file created successfully!\n\n"
                    f"File: {filename}\nLocation: {folder}\n\n"
                    f"Do you want to execute the batch file now?"
                ):
                    self.execute_batch_file(path)
            else:
                # Linux mode: only creates file, no automatic execution option
                content = self.generate_bash_content()
                
                if multiple_folders:
                    # Ask user where to save when multiple folders are involved
                    self.root.attributes('-topmost', True)
                    path = filedialog.asksaveasfilename(
                        title="Save Shell Script As",
                        defaultextension=".sh",
                        filetypes=[("Shell scripts", "*.sh"), ("All files", "*.*")],
                        initialfile="Marc_Mentat_Start_01.sh",
                        parent=self.root
                    )
                    self.root.attributes('-topmost', False)
                    
                    if not path:  # User cancelled
                        return
                    
                    folder = os.path.dirname(path)
                    filename = os.path.basename(path)
                else:
                    # Use folder of first selected file for output
                    folder = self.selected_files[0][3] if self.selected_files else self.dat_folder.get()
                    filename = self.get_next_bash_filename(folder)
                    path = os.path.join(folder, filename)
                
                with open(path, 'w', newline='\n') as f:
                    f.write(content)
                
                messagebox.showinfo(
                    "Shell Script Created",
                    f"Shell script created successfully!\n\n"
                    f"File: {filename}\n"
                    f"Location: {folder}\n\n"
                    f"Transfer this file to your Linux system and execute it there.\n"
                    f"Remember to set execution permission: chmod +x {filename}"
                )
        except Exception as e:
            messagebox.showerror("Error", f"Error saving script: {str(e)}")

    def execute_batch_file(self, batch_path):
        try:
            batch_dir = os.path.dirname(batch_path)
            batch_filename = os.path.basename(batch_path)
            if os.name == 'nt':
                try:
                    if hasattr(subprocess, 'CREATE_NEW_CONSOLE'):
                        subprocess.Popen([batch_filename], cwd=batch_dir,
                                         shell=True, creationflags=subprocess.CREATE_NEW_CONSOLE)
                    else:
                        subprocess.Popen([batch_filename], cwd=batch_dir, shell=True)
                except Exception:
                    subprocess.Popen(f'cmd.exe /c "{batch_filename}"',
                                     cwd=batch_dir, shell=True)
            else:
                subprocess.Popen([batch_path], cwd=batch_dir, shell=True)
        except Exception as e:
            messagebox.showerror("Error", f"Error executing batch file: {str(e)}")

    def get_next_batch_filename(self, folder):
        base = "Marc_Mentat_Start"
        counter = 1
        while True:
            filename = f"{base}_{counter:02d}.bat"
            if not os.path.exists(os.path.join(folder, filename)):
                return filename
            counter += 1
            if counter > 999:
                return f"{base}_{counter}.bat"

    def get_next_bash_filename(self, folder):
        base = "Marc_Mentat_Start"
        counter = 1
        while True:
            filename = f"{base}_{counter:02d}.sh"
            if not os.path.exists(os.path.join(folder, filename)):
                return filename
            counter += 1
            if counter > 999:
                return f"{base}_{counter}.sh"

    def validate_inputs(self):
        if not self.selected_files:
            messagebox.showerror("Error", "Select at least one .dat file!")
            return False
        if self.target_os.get() == "Windows":
            if not self.run_marc_path.get():
                messagebox.showerror("Error", "Select the run_marc.bat file!")
                return False
            if not self.nominal_strain_path.get():
                messagebox.showerror("Error", "Select the nominal_strain_R2020.f file!")
                return False
        else:
            if not self.nominal_strain_path.get():
                self.nominal_strain_path.set(LINUX_PROGRAM_SAVED)
        return True

    def generate_batch_content(self):
        """Standard Windows (.bat) content with full paths."""
        content = [
            "REM Calculations starting now...",
            "",
            "SET MSC_AUTHQUE=36000",
            ""
        ]
        run_marc_path = self.run_marc_path.get()
        nominal_strain_path = self.nominal_strain_path.get()
        if nominal_strain_path == OFFICIAL_NOMINAL_STRAIN_DISPLAY:
            nominal_strain_path = OFFICIAL_NOMINAL_STRAIN_PATH
        run_marc_short_path = self.get_short_path(run_marc_path)
        # Convert to backslashes for Windows
        run_marc_short_path = run_marc_short_path.replace('/', '\\')
        nominal_strain_path = nominal_strain_path.replace('/', '\\')
        content.append(f'SET PATH_RUN_MARC={run_marc_short_path}')
        content.append(f'SET NOMINAL_STRAIN={nominal_strain_path}')
        content.append("")
        
        # Generate variables with full paths (folder + filename without extension)
        # Also generate working directory variables for each file
        for filename, file_type, number, folder_path in self.selected_files:
            # Create full path: folder_path + filename (without extension)
            full_path = os.path.join(folder_path, filename)
            # Convert to backslashes for Windows
            full_path = full_path.replace('/', '\\')
            folder_path_win = folder_path.replace('/', '\\')
            if file_type == "Restart":
                content.append(f"SET Restart_{number:02d}={full_path}")
                content.append(f"SET WD_Restart_{number:02d}={folder_path_win}")
            else:
                content.append(f"SET Loadcase_{number:02d}={full_path}")
                content.append(f"SET WD_Loadcase_{number:02d}={folder_path_win}")
        
        content.append("")
        content.append("REM MATRIX SOLVER TYPE: MUMPS Parallel Direct")
        content.append("")
        content.append("")
        cpu_count = self.cpu_count.get()
        last_restart = None
        for filename, file_type, number, folder_path in self.selected_files:
            if file_type == "Restart":
                content.append(f"REM Changing to working directory for Restart_{number:02d}")
                content.append(f"cd /d %WD_Restart_{number:02d}%")
                content.append(
                    f'CALL %PATH_RUN_MARC% -jid %Restart_{number:02d}% '
                    f'-pr %NOMINAL_STRAIN% -nsolver {cpu_count} -b no -v no'
                )
                content.append("")
                last_restart = f"Restart_{number:02d}"
            else:
                content.append(f"REM Changing to working directory for Loadcase_{number:02d}")
                content.append(f"cd /d %WD_Loadcase_{number:02d}%")
                if last_restart:
                    content.append(
                        f'CALL %PATH_RUN_MARC% -jid %Loadcase_{number:02d}% '
                        f'-rid %{last_restart}% -pr %NOMINAL_STRAIN% '
                        f'-nsolver {cpu_count} -b no -v no'
                    )
                else:
                    content.append(
                        f'CALL %PATH_RUN_MARC% -jid %Loadcase_{number:02d}% '
                        f'-pr %NOMINAL_STRAIN% -nsolver {cpu_count} -b no -v no'
                    )
                content.append("")
        content.append("")
        content.append("msg * Calculations finished!")
        content.append("")
        content.append("")
        content.append("exit")
        content.append("")
        return "\n".join(content)

    def generate_bash_content(self):
        """Content for .sh file (Linux) with full paths."""
        lines = [
            "#!/bin/bash",
            "",
            "# License wait time (minutes)",
            "export MSC_AUTHQUE=36000",
            "",
            "# Program-saved (.marc) file",
            f'export PROGRAM_SAVED="{LINUX_PROGRAM_SAVED}"',
            ""
        ]
        
        # Generate variables with full paths
        # Also generate working directory variables for each file
        for filename, file_type, number, folder_path in self.selected_files:
            # Create full path: folder_path + filename (without extension)
            full_path = os.path.join(folder_path, filename)
            # Convert Windows path to Linux path if needed
            full_path_linux = full_path.replace('\\', '/')
            folder_path_linux = folder_path.replace('\\', '/')
            if file_type == "Restart":
                lines.append(f'Restart_{number:02d}="{full_path_linux}"')
                lines.append(f'WD_Restart_{number:02d}="{folder_path_linux}"')
            else:
                lines.append(f'Loadcase_{number:02d}="{full_path_linux}"')
                lines.append(f'WD_Loadcase_{number:02d}="{folder_path_linux}"')
        
        lines.append("")
        lines.append("# Threads")
        lines.append(f'export NTS={self.cpu_count.get()}')
        lines.append("")
        marc_cmd = "marc2020"
        last_restart_var = None
        for filename, file_type, number, folder_path in self.selected_files:
            if file_type == "Restart":
                varname = f"Restart_{number:02d}"
                wd_varname = f"WD_Restart_{number:02d}"
                lines.append(f'# Changing to working directory for {varname}')
                lines.append(f'cd "${{{wd_varname}}}"')
                lines.append(f'# Run restart-producing job: ${varname}')
                lines.append(
                    f'{marc_cmd} -jid "${{{varname}}}" '
                    f'-pr "$PROGRAM_SAVED" -nts "$NTS" -b no -v no'
                )
                lines.append("")
                last_restart_var = varname
            else:
                varname = f"Loadcase_{number:02d}"
                wd_varname = f"WD_Loadcase_{number:02d}"
                lines.append(f'# Changing to working directory for {varname}')
                lines.append(f'cd "${{{wd_varname}}}"')
                if last_restart_var:
                    lines.append(f'# Run continuation from restart: ${varname}')
                    lines.append(
                        f'{marc_cmd} -jid "${{{varname}}}" -rid "${{{last_restart_var}}}" '
                        f'-pr "$PROGRAM_SAVED" -nts "$NTS" -b no -v no'
                    )
                else:
                    lines.append(f'# Run loadcase without restart: ${varname}')
                    lines.append(
                        f'{marc_cmd} -jid "${{{varname}}}" '
                        f'-pr "$PROGRAM_SAVED" -nts "$NTS" -b no -v no'
                    )
                lines.append("")
        lines.append('echo "Calculations finished!"')
        lines.append("")
        return "\n".join(lines)

    def get_short_path(self, path):
        try:
            if os.name == 'nt':
                try:
                    import ctypes
                    from ctypes import wintypes
                    _GetShortPathNameW = ctypes.windll.kernel32.GetShortPathNameW
                    _GetShortPathNameW.argtypes = [wintypes.LPCWSTR,
                                                   wintypes.LPWSTR, wintypes.DWORD]
                    _GetShortPathNameW.restype = wintypes.DWORD
                    output_buf_size = 260
                    output_buf = ctypes.create_unicode_buffer(output_buf_size)
                    ret = _GetShortPathNameW(path, output_buf, output_buf_size)
                    if ret and ret < output_buf_size:
                        return output_buf.value
                except Exception:
                    pass
            parts = path.replace('/', '\\').split('\\')
            short_parts = []
            for part in parts:
                if not part:
                    short_parts.append(part)
                    continue
                if ' ' not in part:
                    short_parts.append(part)
                    continue
                if '.' in part:
                    name, ext = part.rsplit('.', 1)
                    clean_name = name.replace(' ', '').replace('-', '').replace('_', '')
                    if len(clean_name) > 6:
                        short_part = clean_name[:6].upper() + '~1.' + ext.upper()
                    else:
                        short_part = clean_name.upper() + '.' + ext.upper()
                else:
                    clean_name = part.replace(' ', '').replace('-', '').replace('_', '')
                    if len(clean_name) > 8:
                        short_part = clean_name[:6].upper() + '~1'
                    else:
                        short_part = clean_name.upper()
                short_parts.append(short_part)
            return '\\'.join(short_parts)
        except Exception:
            return path

    def load_history(self):
        try:
            if os.path.exists(self.history_file):
                with open(self.history_file, 'r') as f:
                    lines = [ln.strip() for ln in f.readlines()]
                    if len(lines) >= 3 and lines[0] in ("Windows", "Linux"):
                        self.target_os.set(lines[0])
                        self.run_marc_path.set(lines[1])
                        if self.target_os.get() == "Windows":
                            self.nominal_strain_path.set(OFFICIAL_NOMINAL_STRAIN_PATH)
                            self.update_nominal_strain_display()
                        else:
                            self.nominal_strain_path.set(LINUX_PROGRAM_SAVED)
                    elif len(lines) >= 2:
                        self.target_os.set("Windows")
                        self.run_marc_path.set(lines[0])
                        self.nominal_strain_path.set(OFFICIAL_NOMINAL_STRAIN_PATH)
                        self.update_nominal_strain_display()
        except Exception:
            self.target_os.set("Windows")
            self.nominal_strain_path.set(OFFICIAL_NOMINAL_STRAIN_PATH)
            self.update_nominal_strain_display()

        self.on_os_toggle()

    def save_history(self):
        try:
            os.makedirs(os.path.dirname(self.history_file), exist_ok=True)
            run_marc_path = self.run_marc_path.get()
            nominal_strain_path = self.nominal_strain_path.get()
            if self.target_os.get() == "Windows" and nominal_strain_path == OFFICIAL_NOMINAL_STRAIN_DISPLAY:
                nominal_strain_path = OFFICIAL_NOMINAL_STRAIN_PATH
            if self.target_os.get() == "Linux":
                nominal_strain_path = LINUX_PROGRAM_SAVED
            with open(self.history_file, 'w') as f:
                f.write(self.target_os.get() + "\n")
                f.write(run_marc_path + "\n")
                f.write(nominal_strain_path + "\n")
        except Exception:
            pass

    def on_os_toggle(self):
        self.apply_target_os(style_only=False)

    def apply_target_os(self, style_only=False):
        os_sel = self.target_os.get()
        if os_sel == "Windows":
            color = TITLE_BG_WINDOWS
            self.title_label_1.config(text="Marc Mentat - Batch Generator")
            self.export_button.config(text="Export batch file")
            if not style_only:
                self.run_marc_label.grid(row=0, column=0, sticky=tk.W, pady=5)
                self.run_marc_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
                self.run_marc_button.grid(row=0, column=2, padx=5)
                self.marc_hint_label.grid_forget()
                self.nominal_label.config(text="nominal_strain_R2020.f:")
                self.nominal_entry.config(state='normal')
                self.nominal_button.config(state='normal')
                if self.nominal_strain_path.get() == LINUX_PROGRAM_SAVED:
                    self.nominal_strain_path.set(OFFICIAL_NOMINAL_STRAIN_PATH)
                    self.update_nominal_strain_display()
        else:
            color = TITLE_BG_LINUX
            self.title_label_1.config(text="Marc Mentat - Bash Generator")
            self.export_button.config(text="Export bash file")
            if not style_only:
                self.run_marc_label.grid_remove()
                self.run_marc_entry.grid_remove()
                self.run_marc_button.grid_remove()
                self.marc_hint_label.grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=5)
                self.nominal_label.config(text="Nominal Strain:")
                self.nominal_entry.config(state='readonly')
                self.nominal_button.config(state='disabled')
                self.nominal_strain_path.set(LINUX_PROGRAM_SAVED)

        self.title_frame.config(bg=color)
        self.title_content.config(bg=color)
        self.title_center.config(bg=color)
        self.title_label_1.config(bg=color)
        self.title_label_2.config(bg=color)
        self.title_label_3.config(bg=color)


def main():
    setup_tkinter_for_mentat()
    root = tk.Tk()
    if is_mentat_context():
        root.withdraw()
        root.update_idletasks()
    app = MarcMentatGUI(root)
    if is_mentat_context():
        root.deiconify()
        root.lift()
        root.focus_force()
    root.mainloop()


if __name__ == "__main__":
    main()
