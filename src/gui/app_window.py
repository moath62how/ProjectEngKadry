"""
Main application window for the Engineer Syndicate Lookup tool
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import time
from pathlib import Path
import sys
import os

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from scraper import get_engineer_syndicate_safe
from excel_handler import read_national_ids_from_excel, write_results_to_excel


class AppWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("البحث في نقابة المهندسين")
        self.root.geometry("600x500")
        self.root.resizable(False, False)
        
        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        # Event to request stopping long-running processing
        self._stop_event = threading.Event()
        self._current_results = None
        self._current_output_path = None
        self._processing = False
        
        self.setup_ui()
        
    def setup_ui(self):
        """Set up the user interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        # Configure three columns for RTL layout (rightmost is labels)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)
        main_frame.grid_columnconfigure(2, weight=1)
        
        # Title
        title_label = ttk.Label(
            main_frame, 
            text="🔍 البحث في نقابة المهندسين", 
            font=('Segoe UI', 16, 'bold'),
            anchor='e'
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # === Single ID Lookup Section ===
        section1_label = ttk.Label(
            main_frame, 
            text="بحث رقم قومي فردي", 
            font=('Segoe UI', 12, 'bold'),
            anchor='e'
        )
        section1_label.grid(row=1, column=0, columnspan=3, sticky=tk.E, pady=(10, 10))
        
        # National ID input
        id_label = ttk.Label(main_frame, text="الرقم القومي:", font=('Segoe UI', 10))
        id_label.grid(row=2, column=2, sticky=tk.E, pady=5)
        
        self.entry_id = ttk.Entry(main_frame, width=25, font=('Segoe UI', 10), justify='right')
        self.entry_id.grid(row=2, column=1, pady=5, padx=(10, 5), sticky=tk.W)
        self.entry_id.bind('<Return>', lambda e: self.lookup_single_id())
        
        # Lookup button
        self.btn_lookup = ttk.Button(
            main_frame, 
            text="بحث", 
            command=self.lookup_single_id
        )
        self.btn_lookup.grid(row=2, column=0, pady=5, padx=(5, 0))
        
        # Result display
        result_label = ttk.Label(main_frame, text="النتيجة:", font=('Segoe UI', 10))
        result_label.grid(row=3, column=0, columnspan=3, sticky=tk.E, pady=(10, 5))
        
        self.text_result = tk.Text(
            main_frame, 
            height=6, 
            width=60, 
            font=('Segoe UI', 9),
            wrap=tk.WORD,
            state='disabled'
        )
        self.text_result.grid(row=4, column=0, columnspan=3, pady=5)
        # Right-align text content for RTL
        self.text_result.tag_configure('right', justify='right')
        
        # Separator
        separator = ttk.Separator(main_frame, orient='horizontal')
        separator.grid(row=5, column=0, columnspan=3, sticky='ew', pady=20)
        
        # === Excel Batch Processing Section ===
        section2_label = ttk.Label(
            main_frame, 
            text="معالجة دفعة من ملف إكسل", 
            font=('Segoe UI', 12, 'bold'),
            anchor='e'
        )
        section2_label.grid(row=6, column=0, columnspan=3, sticky=tk.E, pady=(10, 10))
        
        # File selection
        file_label = ttk.Label(main_frame, text="ملف إكسل:", font=('Segoe UI', 10))
        file_label.grid(row=7, column=2, sticky=tk.E, pady=5)
        
        self.entry_file = ttk.Entry(main_frame, width=35, font=('Segoe UI', 9), state='readonly', justify='right')
        self.entry_file.grid(row=7, column=1, pady=5, padx=(10, 5), sticky=tk.W)
        
        btn_browse = ttk.Button(
            main_frame, 
            text="استعراض...", 
            command=self.browse_file
        )
        btn_browse.grid(row=7, column=0, pady=5, padx=(5, 0))
        
        # Process and Stop buttons
        self.btn_process = ttk.Button(
            main_frame,
            text="معالجة ملف إكسل",
            command=self.process_excel,
            state='disabled'
        )
        self.btn_process.grid(row=8, column=0, columnspan=2, pady=15)

        self.btn_stop = ttk.Button(
            main_frame,
            text="إيقاف",
            command=self.request_stop,
            state='disabled'
        )
        self.btn_stop.grid(row=8, column=2, pady=15)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(
            main_frame,
            orient='horizontal',
            length=400,
            mode='determinate'
        )
        self.progress_bar.grid(row=9, column=0, columnspan=3, pady=(5,2))

        # Progress label (detailed text)
        self.progress_label = ttk.Label(
            main_frame, 
            text="", 
            font=('Segoe UI', 9),
            foreground='blue'
        )
        self.progress_label.grid(row=10, column=0, columnspan=3, pady=2)
        
        # Status bar at bottom
        self.status_label = ttk.Label(
            main_frame, 
            text="جاهز", 
            font=('Segoe UI', 9),
            relief=tk.SUNKEN,
            anchor='e'
        )
        self.status_label.grid(row=11, column=0, columnspan=3, sticky='ew', pady=(10, 0))
        
    def lookup_single_id(self):
        """Look up a single national ID"""
        national_id = self.entry_id.get().strip()
        
        if not national_id:
            messagebox.showwarning("مطلوب إدخال", "يرجى إدخال الرقم القومي")
            return
        
        # Disable button and show loading
        self.btn_lookup.config(state='disabled')
        self.status_label.config(text="جار البحث...")
        self.update_result("جار البحث...", clear=True)
        
        # Run in thread to prevent UI freeze
        def lookup_thread():
            result = get_engineer_syndicate_safe(national_id)
            self.root.after(0, self.display_single_result, result)
        
        thread = threading.Thread(target=lookup_thread, daemon=True)
        thread.start()
    
    def display_single_result(self, result):
        """Display the result of a single ID lookup"""
        self.btn_lookup.config(state='normal')
        
        if result['success']:
            output = f"✓ تم بنجاح\n\n"
            output += f"الرقم القومي: {result['national_id']}\n"
            output += f"النقابة: {result['syndicate']}\n"
            self.status_label.config(text="تم البحث بنجاح")
        else:
            output = f"✗ فشل\n\n"
            output += f"الرقم القومي: {result['national_id']}\n"
            output += f"الخطأ: {result.get('error', 'غير معروف')}\n"
            self.status_label.config(text="فشل البحث")
        
        self.update_result(output, clear=True)
    
    def browse_file(self):
        """Open file browser to select Excel file"""
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        
        if filename:
            self.entry_file.config(state='normal')
            self.entry_file.delete(0, tk.END)
            self.entry_file.insert(0, filename)
            self.entry_file.config(state='readonly')
            self.btn_process.config(state='normal')
            self.status_label.config(text=f"تم تحديد الملف: {Path(filename).name}")
    
    def process_excel(self):
        """Process all national IDs from the Excel file"""
        file_path = self.entry_file.get()
        
        if not file_path:
            messagebox.showwarning("No File", "Please select an Excel file first")
            return
        
        # Ask for output location
        output_path = filedialog.asksaveasfilename(
            title="Save Results As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="syndicate_results.xlsx"
        )
        
        if not output_path:
            return
        
        # Disable button and start processing
        self.btn_process.config(state='disabled')
        self.btn_stop.config(state='normal')
        self._stop_event.clear()
        self._processing = True
        self.progress_label.config(text="جار المعالجة...")
        self.status_label.config(text="جار قراءة ملف الإكسل...")
        
        def process_thread():
            try:
                # Read IDs from Excel
                national_ids = read_national_ids_from_excel(file_path)
                total = len(national_ids)

                # If no IDs found, surface a helpful error to the user
                if total == 0:
                    self.root.after(0, self.show_process_error, "لم يتم العثور على أرقام قومية في الملف. تحقق من أسماء الأعمدة.")
                    return

                self.root.after(0, lambda: self.status_label.config(
                    text=f"تم العثور على {total} رقم قومي. جارٍ المعالجة..."
                ))

                # Process each ID (handle per-ID errors so one failure doesn't stop the batch)
                results = []
                # store partial state so Stop can save
                self._current_results = results
                self._current_output_path = output_path
                self._processing = True
                start_time = time.time()
                for i, national_id in enumerate(national_ids, 1):
                    # Check for stop request and break early if requested
                    if self._stop_event.is_set():
                        print("Stop requested, breaking processing loop")
                        break
                    try:
                        print(f"Processing {i}/{total}: {national_id}")
                        result = get_engineer_syndicate_safe(national_id)
                    except Exception as item_exc:
                        # Record failure for this ID and continue
                        result = {
                            'success': False,
                            'national_id': national_id,
                            'syndicate': None,
                            'error': str(item_exc)
                        }
                        print(f"Error processing {national_id}: {item_exc}")

                    results.append(result)
                    # Update progress bar and labels
                    # Update progress value
                    try:
                        self.root.after(0, lambda v=i: self.progress_bar.config(value=v))
                        self.root.after(0, lambda m=total: self.progress_bar.config(maximum=m))
                    except Exception:
                        pass

                    # Estimate remaining time
                    elapsed = time.time() - start_time
                    avg_per = elapsed / i if i else 0
                    remaining = max(0, int(avg_per * (total - i)))
                    mins, secs = divmod(remaining, 60)
                    hours, mins = divmod(mins, 60)
                    eta = f"{hours:d}h {mins:d}m {secs:d}s" if hours else f"{mins:d}m {secs:d}s"

                    progress_text = f"جار المعالجة {i}/{total} — متوقع: {eta}"
                    self.root.after(0, lambda pt=progress_text: self.progress_label.config(text=pt))

                # Write results
                try:
                    write_results_to_excel(results, output_path)
                except Exception as write_exc:
                    print(f"Error writing results: {write_exc}")
                    self.root.after(0, self.show_process_error, str(write_exc))
                    return

                # Show success
                self.root.after(0, self.show_process_complete, output_path, results)
                
            except Exception as e:
                print(f"Batch processing error: {e}")
                self.root.after(0, self.show_process_error, str(e))
            finally:
                # Ensure the process button is re-enabled and stop button disabled in all cases
                def finish_buttons():
                    try:
                        self.btn_process.config(state='normal')
                    except Exception:
                        pass
                    try:
                        self.btn_stop.config(state='disabled')
                    except Exception:
                        pass
                    self._processing = False

                self.root.after(0, finish_buttons)
        
        thread = threading.Thread(target=process_thread, daemon=True)
        thread.start()
    
    def show_process_complete(self, output_path, results):
        """Show completion message"""
        self.btn_process.config(state='normal')
        self.progress_label.config(text="")
        try:
            self.progress_bar.config(value=0)
            self.progress_bar.config(maximum=100)
        except Exception:
            pass
        # Clear current state
        try:
            self._current_results = None
            self._current_output_path = None
        except Exception:
            pass
        
        success_count = sum(1 for r in results if r['success'])
        total_count = len(results)
        
        messagebox.showinfo(
            "اكتمل",
            f"اكتملت المعالجة!\n\n"
            f"الناجحة: {success_count}/{total_count}\n"
            f"تم حفظ النتائج في:\n{output_path}"
        )
        self.status_label.config(text=f"اكتملت: {success_count}/{total_count} ناجحة")
    
    def show_process_error(self, error_msg):
        """Show error message"""
        self.btn_process.config(state='normal')
        self.progress_label.config(text="")
        messagebox.showerror("خطأ", f"حدث خطأ:\n\n{error_msg}")
        self.status_label.config(text="حدث خطأ")
        # Clear current state
        try:
            self._current_results = None
            self._current_output_path = None
        except Exception:
            pass

    def request_stop(self):
        """Request the worker thread to stop and save current partial results."""
        if not self._processing:
            return
        self._stop_event.set()
        self.status_label.config(text="تم طلب الإيقاف — جار حفظ النتائج الجزئية...")

        # If we have partial results, save them immediately
        try:
            if self._current_results and self._current_output_path:
                # create a filename indicating partial
                out_path = self._current_output_path
                p = Path(out_path)
                partial_name = p.with_name(p.stem + "_partial" + p.suffix)
                write_results_to_excel(self._current_results, str(partial_name))
                self.status_label.config(text=f"تم حفظ النتائج الجزئية: {partial_name.name}")
                messagebox.showinfo("مؤقت", f"تم حفظ النتائج الجزئية في:\n{partial_name}")
        except Exception as e:
            print(f"Error saving partial results: {e}")
            messagebox.showerror("خطأ حفظ جزئي", f"خطأ أثناء حفظ النتائج الجزئية:\n{e}")
    
    def update_result(self, text, clear=False):
        """Update the result text box"""
        self.text_result.config(state='normal')
        if clear:
            self.text_result.delete(1.0, tk.END)
        # Insert text with right justification for RTL
        self.text_result.insert(tk.END, text, 'right')
        self.text_result.config(state='disabled')