"""
Meeting Scheduler - Web App Version (Streamlit)
Converts the desktop GUI app to a web-based application
"""

import streamlit as st
import pandas as pd
import re
from typing import List, Dict, Tuple
from io import StringIO, BytesIO

# Set page config
st.set_page_config(
    page_title="Meeting Scheduler",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS to make it look like the desktop app
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .section-header {
        font-size: 1.5rem;
        font-weight: bold;
        color: #2c3e50;
        margin-top: 2rem;
        margin-bottom: 1rem;
        padding: 0.5rem;
        background-color: #f0f2f6;
        border-radius: 5px;
    }
    .success-box {
        padding: 1rem;
        background-color: #d4edda;
        border-left: 5px solid #28a745;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .warning-box {
        padding: 1rem;
        background-color: #fff3cd;
        border-left: 5px solid #ffc107;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .info-box {
        padding: 1rem;
        background-color: #d1ecf1;
        border-left: 5px solid #17a2b8;
        border-radius: 5px;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)


class MeetingScheduler:
    def __init__(self):
        self.availability_data = None
        self.parsed_schedules = {}
    
    def parse_time_slot(self, time_str: str) -> List[Tuple[int, int]]:
        """Parse time slot string and return list of (start_hour, end_hour) tuples"""
        if not isinstance(time_str, str):
            return []
        
        time_str = time_str.strip().upper()
        
        # Check if unavailable
        unavailable_keywords = ['NA', 'N/A', 'ON LEAVE', 'LEAVE', 'UNAVAILABLE', 'OFF']
        if any(keyword in time_str for keyword in unavailable_keywords):
            return []
        
        slots = []
        time_parts = re.split('[,;]', time_str)
        
        for part in time_parts:
            part = part.strip()
            if not part:
                continue
                
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
                    if end_period == 'PM' and start_time < 12:
                        start_time += 12
                        
                slots.append((start_time, end_time))
        
        return slots
    
    def load_availability_data(self, df: pd.DataFrame) -> bool:
        """Load availability data from DataFrame"""
        try:
            self.availability_data = df
            return True
        except Exception as e:
            st.error(f"Error loading data: {e}")
            return False
    
    def parse_availability(self) -> Dict:
        """Parse the availability data"""
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
                member_schedule[date_col] = time_slots
            
            schedules[member_name] = member_schedule
        
        self.parsed_schedules = schedules
        return schedules
    
    def find_best_meeting_time(self, duration_hours: float) -> Dict:
        """Find the best meeting time"""
        if not self.parsed_schedules:
            return {'perfect_slots': [], 'best_alternative_slots': []}
        
        all_dates = set()
        for member_schedule in self.parsed_schedules.values():
            all_dates.update(member_schedule.keys())
        
        perfect_slots = []
        alternative_slots = []
        
        for date in sorted(all_dates):
            members_available = {}
            
            for member, schedule in self.parsed_schedules.items():
                if date in schedule and schedule[date]:
                    members_available[member] = schedule[date]
            
            total_members = len(self.parsed_schedules)
            available_count = len(members_available)
            
            if available_count == 0:
                continue
            
            if available_count == total_members:
                common_slots = self._find_common_slots(members_available, duration_hours)
                for slot in common_slots:
                    perfect_slots.append({
                        'date': date,
                        'start_time': slot[0],
                        'end_time': slot[1],
                        'members_available': list(members_available.keys()),
                        'members_unavailable': []
                    })
            
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
        
        alternative_slots.sort(key=lambda x: x['coverage'], reverse=True)
        
        return {
            'perfect_slots': perfect_slots,
            'best_alternative_slots': alternative_slots[:10]
        }
    
    def _find_common_slots(self, members_available: Dict, duration_hours: float) -> List[Tuple[int, int]]:
        """Find common time slots across all provided members"""
        if not members_available:
            return []
        
        member_slots = list(members_available.values())
        
        if not member_slots:
            return []
        
        current_intersection = member_slots[0]
        
        for member_slots in member_slots[1:]:
            new_intersection = []
            
            for slot1 in current_intersection:
                for slot2 in member_slots:
                    overlap_start = max(slot1[0], slot2[0])
                    overlap_end = min(slot1[1], slot2[1])
                    
                    if overlap_end > overlap_start:
                        new_intersection.append((overlap_start, overlap_end))
            
            current_intersection = new_intersection
            
            if not current_intersection:
                break
        
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


# Initialize session state
if 'scheduler' not in st.session_state:
    st.session_state.scheduler = MeetingScheduler()
if 'parsed' not in st.session_state:
    st.session_state.parsed = False
if 'results' not in st.session_state:
    st.session_state.results = None


# Main App
def main():
    # Header
    st.markdown('<div class="main-header">📅 Meeting Scheduler</div>', unsafe_allow_html=True)
    
    # Sidebar - Instructions
    with st.sidebar:
        st.markdown("### 📖 Instructions")
        st.markdown("""
        **How to use:**
        1. Upload your availability file (CSV or Excel)
        2. Review the parsed data
        3. Select meeting duration
        4. Find the best meeting time!
        
        **File Format:**
        - First column: Member names
        - Other columns: Dates with availability
        - Time format: 9AM-12PM, 1-3PM, etc.
        - Unavailable: NA, on leave
        
        **Example:**
        ```
        Name    | Mon      | Tue
        Alice   | 9AM-12PM | 1-3PM
        Bob     | NA       | 2-5PM
        ```
        """)
        
        st.markdown("---")
        st.markdown("### 📊 About")
        st.markdown("""
        This tool helps find optimal meeting times by analyzing team availability.
        
        **Features:**
        - ✅ Perfect match finding
        - 📊 Best alternatives
        - 📈 Coverage analysis
        - 📤 Export results
        """)
    
    # Section 1: File Upload
    st.markdown('<div class="section-header">1️⃣ Upload Availability File</div>', unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader(
        "Choose a CSV or Excel file",
        type=['csv', 'xlsx', 'xls'],
        help="Upload a file with member names in first column and dates in other columns"
    )
    
    # Sample file download
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📥 Download Sample CSV"):
            sample_data = """Member Name,Monday Dec 9,Tuesday Dec 10,Wednesday Dec 11
Alice,9AM-12PM,1-3PM,NA
Bob,10AM-2PM,2-5PM,on leave
Carol,9AM-5PM,NA,10AM-3PM"""
            st.download_button(
                label="Click to Download",
                data=sample_data,
                file_name="sample_availability.csv",
                mime="text/csv"
            )
    
    # Load and Parse
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            st.success(f"✅ File loaded: {uploaded_file.name}")
            
            if st.button("🔄 Parse Availability Data", type="primary"):
                with st.spinner("Parsing availability..."):
                    success = st.session_state.scheduler.load_availability_data(df)
                    if success:
                        schedules = st.session_state.scheduler.parse_availability()
                        st.session_state.parsed = True
                        
                        st.markdown('<div class="success-box">✅ Successfully parsed data for {} team members!</div>'.format(len(schedules)), unsafe_allow_html=True)
                        
                        # Section 2: Display Parsed Data
                        st.markdown('<div class="section-header">2️⃣ Parsed Availability</div>', unsafe_allow_html=True)
                        
                        for member, schedule in schedules.items():
                            with st.expander(f"📋 {member}", expanded=False):
                                for date, slots in schedule.items():
                                    if slots:
                                        slot_strs = [f"{st.session_state.scheduler.format_time(s[0])}-{st.session_state.scheduler.format_time(s[1])}" for s in slots]
                                        st.write(f"**{date}:** {', '.join(slot_strs)}")
                                    else:
                                        st.write(f"**{date}:** ❌ Unavailable")
        
        except Exception as e:
            st.error(f"❌ Error reading file: {e}")
    
    # Section 3: Find Meeting Time
    if st.session_state.parsed:
        st.markdown('<div class="section-header">3️⃣ Find Best Meeting Time</div>', unsafe_allow_html=True)
        
        col1, col2 = st.columns([1, 3])
        
        with col1:
            duration = st.selectbox(
                "Meeting Duration (hours)",
                options=[0.5, 1.0, 1.5, 2.0, 2.5, 3.0],
                index=1,
                help="Select how long your meeting will be"
            )
        
        with col2:
            if st.button("🔍 Find Best Time", type="primary"):
                with st.spinner("Analyzing availability..."):
                    results = st.session_state.scheduler.find_best_meeting_time(duration)
                    st.session_state.results = results
        
        # Section 4: Display Results
        if st.session_state.results:
            st.markdown('<div class="section-header">4️⃣ Results</div>', unsafe_allow_html=True)
            
            results = st.session_state.results
            
            # Perfect Slots
            if results['perfect_slots']:
                st.markdown('<div class="success-box"><strong>✅ PERFECT TIME SLOTS</strong> (All members available)</div>', unsafe_allow_html=True)
                
                for i, slot in enumerate(results['perfect_slots'], 1):
                    with st.expander(f"Option {i}: {slot['date']} at {st.session_state.scheduler.format_time(slot['start_time'])}-{st.session_state.scheduler.format_time(slot['end_time'])}", expanded=(i==1)):
                        st.write(f"**Date:** {slot['date']}")
                        st.write(f"**Time:** {st.session_state.scheduler.format_time(slot['start_time'])} - {st.session_state.scheduler.format_time(slot['end_time'])}")
                        st.write(f"**Available:** {', '.join(slot['members_available'])}")
                        st.success(f"✅ All {len(slot['members_available'])} members can attend!")
            else:
                st.markdown('<div class="warning-box"><strong>⚠️ NO PERFECT TIME SLOT FOUND</strong><br>There is no time when ALL members are available.</div>', unsafe_allow_html=True)
            
            # Alternative Slots
            if results['best_alternative_slots']:
                st.markdown('<div class="info-box"><strong>📊 BEST ALTERNATIVE TIME SLOTS</strong><br>Times when MOST members are available</div>', unsafe_allow_html=True)
                
                for i, slot in enumerate(results['best_alternative_slots'], 1):
                    coverage_pct = slot['coverage'] * 100
                    
                    with st.expander(f"Alternative {i}: {slot['date']} at {st.session_state.scheduler.format_time(slot['start_time'])}-{st.session_state.scheduler.format_time(slot['end_time'])} ({coverage_pct:.0f}% coverage)", expanded=(i==1)):
                        st.write(f"**Date:** {slot['date']}")
                        st.write(f"**Time:** {st.session_state.scheduler.format_time(slot['start_time'])} - {st.session_state.scheduler.format_time(slot['end_time'])}")
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            st.write(f"**✅ Available ({len(slot['members_available'])}):**")
                            st.write(", ".join(slot['members_available']))
                        
                        if slot['members_unavailable']:
                            with col2:
                                st.write(f"**❌ Unavailable ({len(slot['members_unavailable'])}):**")
                                st.write(", ".join(slot['members_unavailable']))
                        
                        # Progress bar for coverage
                        st.progress(slot['coverage'])
            elif not results['perfect_slots']:
                st.error("❌ No suitable meeting times found with the current availability.")
            
            # Export Results
            st.markdown("---")
            results_text = generate_results_text(results, st.session_state.scheduler, duration)
            st.download_button(
                label="📥 Download Results (TXT)",
                data=results_text,
                file_name=f"meeting_results_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.txt",
                mime="text/plain"
            )


def generate_results_text(results, scheduler, duration):
    """Generate text file content for results"""
    lines = []
    lines.append("=" * 80)
    lines.append(f"MEETING SCHEDULER RESULTS (Duration: {duration} hour(s))")
    lines.append("=" * 80)
    lines.append("")
    
    if results['perfect_slots']:
        lines.append("✅ PERFECT TIME SLOTS (All members available):")
        lines.append("-" * 80)
        lines.append("")
        
        for i, slot in enumerate(results['perfect_slots'], 1):
            lines.append(f"Option {i}:")
            lines.append(f"  Date: {slot['date']}")
            lines.append(f"  Time: {scheduler.format_time(slot['start_time'])} - {scheduler.format_time(slot['end_time'])}")
            lines.append(f"  Available: {', '.join(slot['members_available'])}")
            lines.append("")
    else:
        lines.append("⚠️  NO PERFECT TIME SLOT FOUND")
        lines.append("There is no time when ALL members are available.")
        lines.append("")
    
    if results['best_alternative_slots']:
        lines.append("📊 BEST ALTERNATIVE TIME SLOTS:")
        lines.append("(Times when MOST members are available)")
        lines.append("-" * 80)
        lines.append("")
        
        for i, slot in enumerate(results['best_alternative_slots'], 1):
            coverage_pct = slot['coverage'] * 100
            lines.append(f"Alternative {i} ({coverage_pct:.0f}% coverage):")
            lines.append(f"  Date: {slot['date']}")
            lines.append(f"  Time: {scheduler.format_time(slot['start_time'])} - {scheduler.format_time(slot['end_time'])}")
            lines.append(f"  Available ({len(slot['members_available'])}): {', '.join(slot['members_available'])}")
            
            if slot['members_unavailable']:
                lines.append(f"  Unavailable ({len(slot['members_unavailable'])}): {', '.join(slot['members_unavailable'])}")
            
            lines.append("")
    
    lines.append("=" * 80)
    
    return "\n".join(lines)


if __name__ == "__main__":
    main()