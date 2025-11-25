"""
Meeting Scheduler Application
Helps find the best meeting time based on team member availability
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
from datetime import datetime, timedelta
import re
from typing import List, Dict, Tuple, Optional
import os


class MeetingScheduler:
    def __init__(self):
        self.availability_data = None
        self.parsed_schedules = {}
        
    def parse_time_slot(self, time_str: str) -> List[Tuple[int, int]]:
        """
        Parse time slot string like '1-2PM', '3-6PM', '9AM-12PM'
        Returns list of (start_hour, end_hour) tuples in 24-hour format
        """
        if not isinstance(time_str, str):
            return []
        
        time_str = time_str.strip().upper()
        
        # Check if unavailable
        unavailable_keywords = ['NA', 'N/A', 'ON LEAVE', 'LEAVE', 'UNAVAILABLE', 'OFF']
        if any(keyword in time_str for keyword in unavailable_keywords):
            return []
        
        # Pattern for time ranges like "1-2PM", "9AM-12PM", "1PM-3PM"
        slots = []
        
        # Handle multiple slots separated by comma or semicolon
        time_parts = re.split('[,;]', time_str)
        
        for part in time_parts:
            part = part.strip()
            if not part:
                continue
                
            # Pattern: "9AM-12PM" or "9-12PM" or "1-2PM"
            match = re.search(r'(\d+)\s*(AM|PM)?\s*[-–]\s*(\d+)\s*(AM|PM)', part)
            
            if match:
                start_time = int(match.group(1))
                start_period = match.group(2) if match.group(2) else None
                end_time = int(match.group(3))
                end_period = match.group(4)
                
                # Convert to 24-hour format
                if end_period == 'PM' and end_time != 12:
                    end_time += 12
                elif end_period == 'AM' and end_time == 12:
                    end_time = 0
                    
                if start_period:
                    if start_period == 'PM' and start_time != 12:
                        start_time += 12
                    elif start_period == 'AM' and start_time == 12:
                        start_time = 0
                else:
                    # If no AM/PM for start, infer from end period
                    if end_period == 'PM' and start_time < 12:
                        start_time += 12  # Assume same period as end
                    # If end is AM, keep start as is (already in correct range)
                        
                slots.append((start_time, end_time))
        
        return slots
    
    def load_availability_file(self, filepath: str) -> bool:
        """Load availability data from CSV or Excel file"""
        try:
            if filepath.endswith('.csv'):
                self.availability_data = pd.read_csv(filepath)
            elif filepath.endswith(('.xlsx', '.xls')):
                self.availability_data = pd.read_excel(filepath)
            else:
                return False
            
            return True
        except Exception as e:
            print(f"Error loading file: {e}")
            return False
    
    def parse_availability(self) -> Dict:
        """
        Parse the availability data
        Expected format: First column = Member names, Other columns = Dates with availability
        """
        if self.availability_data is None:
            return {}
        
        df = self.availability_data
        member_column = df.columns[0]
        date_columns = df.columns[1:]
        
        schedules = {}
        
        for _, row in df.iterrows():
            member_name = str(row[member_column]).strip()
            if not member_name or member_name.lower() in ['nan', 'none', '']:
                continue
                
            member_schedule = {}
            
            for date_col in date_columns:
                availability_str = str(row[date_col])
                time_slots = self.parse_time_slot(availability_str)
                
                if time_slots:
                    member_schedule[date_col] = time_slots
                else:
                    member_schedule[date_col] = []
            
            schedules[member_name] = member_schedule
        
        self.parsed_schedules = schedules
        return schedules
    
    def find_best_meeting_time(self, duration_hours: float) -> Dict:
        """
        Find the best meeting time where all members are available
        Returns dict with 'perfect_slots' and 'best_alternative_slots'
        """
        if not self.parsed_schedules:
            return {'perfect_slots': [], 'best_alternative_slots': []}
        
        all_dates = set()
        for member_schedule in self.parsed_schedules.values():
            all_dates.update(member_schedule.keys())
        
        perfect_slots = []
        alternative_slots = []
        
        for date in sorted(all_dates):
            # Get all members' availability for this date
            members_available = {}
            
            for member, schedule in self.parsed_schedules.items():
                if date in schedule and schedule[date]:
                    members_available[member] = schedule[date]
            
            total_members = len(self.parsed_schedules)
            available_count = len(members_available)
            
            if available_count == 0:
                continue
            
            # Find overlapping time slots
            if available_count == total_members:
                # All members available - find perfect overlap
                common_slots = self._find_common_slots(members_available, duration_hours)
                for slot in common_slots:
                    perfect_slots.append({
                        'date': date,
                        'start_time': slot[0],
                        'end_time': slot[1],
                        'members_available': list(members_available.keys()),
                        'members_unavailable': []
                    })
            
            # Find best alternative (most members available)
            if available_count >= 2:
                common_slots = self._find_common_slots(members_available, duration_hours)
                unavailable_members = [m for m in self.parsed_schedules.keys() 
                                     if m not in members_available]
                
                for slot in common_slots:
                    alternative_slots.append({
                        'date': date,
                        'start_time': slot[0],
                        'end_time': slot[1],
                        'members_available': list(members_available.keys()),
                        'members_unavailable': unavailable_members,
                        'coverage': available_count / total_members
                    })
        
        # Sort alternative slots by coverage (most members first)
        alternative_slots.sort(key=lambda x: x['coverage'], reverse=True)
        
        return {
            'perfect_slots': perfect_slots,
            'best_alternative_slots': alternative_slots[:10]  # Top 10 alternatives
        }
    
    def _find_common_slots(self, members_available: Dict, duration_hours: float) -> List[Tuple[int, int]]:
        """Find common time slots across all provided members"""
        if not members_available:
            return []
        
        # Create list of all time slots for each member
        member_slots = list(members_available.values())
        
        if not member_slots:
            return []
        
        # Find intersection of all slots
        common_intervals = self._intersect_all_slots(member_slots, duration_hours)
        
        return common_intervals
    
    def _intersect_all_slots(self, all_member_slots: List[List[Tuple[int, int]]], 
                            duration_hours: float) -> List[Tuple[int, int]]:
        """Find intersection of time slots across all members"""
        if not all_member_slots:
            return []
        
        # Start with first member's slots
        current_intersection = all_member_slots[0]
        
        # Intersect with each subsequent member
        for member_slots in all_member_slots[1:]:
            new_intersection = []
            
            for slot1 in current_intersection:
                for slot2 in member_slots:
                    # Find overlap between slot1 and slot2
                    overlap_start = max(slot1[0], slot2[0])
                    overlap_end = min(slot1[1], slot2[1])
                    
                    if overlap_end > overlap_start:
                        # There's an overlap
                        new_intersection.append((overlap_start, overlap_end))
            
            current_intersection = new_intersection
            
            if not current_intersection:
                break
        
        # Filter slots that meet the duration requirement
        valid_slots = []
        for slot in current_intersection:
            slot_duration = slot[1] - slot[0]
            if slot_duration >= duration_hours:
                valid_slots.append(slot)
        
        return valid_slots
    
    def format_time(self, hour: int) -> str:
        """Convert 24-hour format to 12-hour format with AM/PM"""
        if hour == 0:
            return "12AM"
        elif hour < 12:
            return f"{hour}AM"
        elif hour == 12:
            return "12PM"
        else:
            return f"{hour - 12}PM"


class MeetingSchedulerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Meeting Scheduler")
        self.root.geometry("900x700")
        
        self.scheduler = MeetingScheduler()
        self.filepath = None
        
        self.setup_gui()
    
    def setup_gui(self):
        """Setup the GUI components"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="Meeting Scheduler", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=10)
        
        # File selection
        file_frame = ttk.LabelFrame(main_frame, text="1. Select Availability File", 
                                    padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.file_label = ttk.Label(file_frame, text="No file selected", 
                                    foreground="gray")
        self.file_label.grid(row=0, column=0, padx=5)
        
        browse_btn = ttk.Button(file_frame, text="Browse File", 
                               command=self.browse_file)
        browse_btn.grid(row=0, column=1, padx=5)
        
        load_btn = ttk.Button(file_frame, text="Load & Parse", 
                             command=self.load_and_parse)
        load_btn.grid(row=0, column=2, padx=5)
        
        # Meeting duration
        duration_frame = ttk.LabelFrame(main_frame, text="2. Select Meeting Duration", 
                                       padding="10")
        duration_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(duration_frame, text="Duration (hours):").grid(row=0, column=0, padx=5)
        
        self.duration_var = tk.StringVar(value="1.0")
        duration_spinbox = ttk.Spinbox(duration_frame, from_=0.5, to=8.0, 
                                       increment=0.5, textvariable=self.duration_var,
                                       width=10)
        duration_spinbox.grid(row=0, column=1, padx=5)
        
        find_btn = ttk.Button(duration_frame, text="Find Best Time", 
                             command=self.find_best_time)
        find_btn.grid(row=0, column=2, padx=5)
        
        # Results area
        results_frame = ttk.LabelFrame(main_frame, text="3. Results", padding="10")
        results_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), 
                          pady=5)
        
        # Create text widget with scrollbar
        self.results_text = scrolledtext.ScrolledText(results_frame, width=100, height=30,
                                                      wrap=tk.WORD, font=('Courier', 10))
        self.results_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Instructions button
        help_btn = ttk.Button(main_frame, text="Show Instructions", 
                             command=self.show_instructions)
        help_btn.grid(row=5, column=0, columnspan=3, pady=5)
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
    
    def browse_file(self):
        """Open file dialog to select availability file"""
        self.filepath = filedialog.askopenfilename(
            title="Select Availability File",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx *.xls"), 
                      ("All files", "*.*")]
        )
        
        if self.filepath:
            filename = os.path.basename(self.filepath)
            self.file_label.config(text=filename, foreground="black")
    
    def load_and_parse(self):
        """Load and parse the availability file"""
        if not self.filepath:
            messagebox.showerror("Error", "Please select a file first!")
            return
        
        success = self.scheduler.load_availability_file(self.filepath)
        
        if not success:
            messagebox.showerror("Error", "Failed to load file. Please check the format.")
            return
        
        schedules = self.scheduler.parse_availability()
        
        if not schedules:
            messagebox.showerror("Error", "No valid availability data found!")
            return
        
        # Display parsed data
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, "=== PARSED AVAILABILITY DATA ===\n\n")
        
        for member, schedule in schedules.items():
            self.results_text.insert(tk.END, f"📋 {member}:\n")
            for date, slots in schedule.items():
                if slots:
                    slot_strs = [f"{self.scheduler.format_time(s[0])}-{self.scheduler.format_time(s[1])}" 
                               for s in slots]
                    self.results_text.insert(tk.END, f"  {date}: {', '.join(slot_strs)}\n")
                else:
                    self.results_text.insert(tk.END, f"  {date}: Unavailable\n")
            self.results_text.insert(tk.END, "\n")
        
        messagebox.showinfo("Success", f"Successfully parsed data for {len(schedules)} members!")
    
    def find_best_time(self):
        """Find the best meeting time"""
        if not self.scheduler.parsed_schedules:
            messagebox.showerror("Error", "Please load and parse a file first!")
            return
        
        try:
            duration = float(self.duration_var.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid duration value!")
            return
        
        results = self.scheduler.find_best_meeting_time(duration)
        
        # Display results
        self.results_text.delete(1.0, tk.END)
        
        self.results_text.insert(tk.END, "=" * 80 + "\n")
        self.results_text.insert(tk.END, f"MEETING SCHEDULER RESULTS (Duration: {duration} hour(s))\n")
        self.results_text.insert(tk.END, "=" * 80 + "\n\n")
        
        # Perfect slots (everyone available)
        if results['perfect_slots']:
            self.results_text.insert(tk.END, "✅ PERFECT TIME SLOTS (All members available):\n")
            self.results_text.insert(tk.END, "-" * 80 + "\n\n")
            
            for i, slot in enumerate(results['perfect_slots'], 1):
                self.results_text.insert(tk.END, f"Option {i}:\n")
                self.results_text.insert(tk.END, f"  Date: {slot['date']}\n")
                self.results_text.insert(tk.END, 
                    f"  Time: {self.scheduler.format_time(slot['start_time'])} - "
                    f"{self.scheduler.format_time(slot['end_time'])}\n")
                self.results_text.insert(tk.END, 
                    f"  Available: {', '.join(slot['members_available'])}\n")
                self.results_text.insert(tk.END, "\n")
        else:
            self.results_text.insert(tk.END, "⚠️  NO PERFECT TIME SLOT FOUND\n")
            self.results_text.insert(tk.END, 
                "There is no time when ALL members are available.\n\n")
        
        # Alternative slots
        if results['best_alternative_slots']:
            self.results_text.insert(tk.END, "📊 BEST ALTERNATIVE TIME SLOTS:\n")
            self.results_text.insert(tk.END, 
                "(Times when MOST members are available)\n")
            self.results_text.insert(tk.END, "-" * 80 + "\n\n")
            
            for i, slot in enumerate(results['best_alternative_slots'], 1):
                coverage_pct = slot['coverage'] * 100
                self.results_text.insert(tk.END, 
                    f"Alternative {i} ({coverage_pct:.0f}% coverage):\n")
                self.results_text.insert(tk.END, f"  Date: {slot['date']}\n")
                self.results_text.insert(tk.END, 
                    f"  Time: {self.scheduler.format_time(slot['start_time'])} - "
                    f"{self.scheduler.format_time(slot['end_time'])}\n")
                self.results_text.insert(tk.END, 
                    f"  Available ({len(slot['members_available'])}): "
                    f"{', '.join(slot['members_available'])}\n")
                
                if slot['members_unavailable']:
                    self.results_text.insert(tk.END, 
                        f"  Unavailable ({len(slot['members_unavailable'])}): "
                        f"{', '.join(slot['members_unavailable'])}\n")
                
                self.results_text.insert(tk.END, "\n")
        elif not results['perfect_slots']:
            self.results_text.insert(tk.END, 
                "❌ No suitable meeting times found with the current availability.\n")
        
        self.results_text.insert(tk.END, "=" * 80 + "\n")

    
    def show_instructions(self):
        """Show instructions dialog"""
        instructions = """
MEETING SCHEDULER - INSTRUCTIONS

1. PREPARE YOUR AVAILABILITY FILE:
   Create a CSV or Excel file with the following format:
   
   | Member Name | Date 1      | Date 2      | Date 3      |
   |-------------|-------------|-------------|-------------|
   | John        | 1-2PM       | 9AM-12PM    | NA          |
   | Sarah       | 2-4PM       | 10AM-1PM    | 3-5PM       |
   | Mike        | 1-5PM       | on leave    | 2-6PM       |

2. TIME FORMAT:
   - Use formats like: 1-2PM, 9AM-12PM, 3-6PM
   - For unavailability use: NA, n/a, on leave, unavailable

3. USE THE APP:
   - Browse and select your availability file
   - Click "Load & Parse" to process the data
   - Select meeting duration (in hours)
   - Click "Find Best Time" to see results

4. RESULTS:
   - Perfect slots: Times when ALL members are available
   - Alternative slots: Times when MOST members are available
   - If no perfect slot exists, alternatives are ranked by coverage

        """
        
        msg = tk.Toplevel(self.root)
        msg.title("Instructions")
        msg.geometry("600x500")
        
        text = scrolledtext.ScrolledText(msg, wrap=tk.WORD, font=('Arial', 10))
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        text.insert(tk.END, instructions)
        text.config(state=tk.DISABLED)
        
        close_btn = ttk.Button(msg, text="Close", command=msg.destroy)
        close_btn.pack(pady=5)


def main():
    root = tk.Tk()
    app = MeetingSchedulerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()