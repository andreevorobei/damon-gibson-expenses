#!/usr/bin/env python3
"""
Expense Reconciliation Web App
Web application for reconciling CapitalOne bank statements with Jobber expenses
Created using Streamlit
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io
from pathlib import Path

# Page configuration
st.set_page_config(
    page_title="Expense Reconciliation Tool",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

class ExpenseReconciler:
    def __init__(self, tolerance_amount=2.0, tolerance_days=2, check_person=True):
        self.tolerance_amount = tolerance_amount
        self.tolerance_days = tolerance_days
        self.check_person = check_person
        
        # Mapping card numbers to people names
        self.card_to_person = {
            '9265': 'Aaron Davidson',
            '4298': 'Alex Masuda', 
            '1725': 'Jericho Taylor-Daves',
            '3253': 'Jerry Morales',
            '2984': 'Antonio'
        }
        
    def load_capitalone_data(self, uploaded_file):
        """Loads data from CapitalOne file"""
        try:
            df = pd.read_excel(uploaded_file)
            
            # Looking for data columns
            date_col = None
            amount_col = None
            desc_col = None
            card_col = None
            
            for col in df.columns:
                col_lower = str(col).lower()
                if 'transaction' in col_lower and 'date' in col_lower:
                    date_col = col
                elif 'debit' in col_lower or ('amount' in col_lower and 'debit' not in col_lower):
                    amount_col = col
                elif 'description' in col_lower:
                    desc_col = col
                elif 'card' in col_lower and 'no' in col_lower:
                    card_col = col
            
            if not date_col or not amount_col:
                st.error("❌ Required columns not found in CapitalOne file. Expected columns with transaction date and debit amount.")
                return None
                
            if not card_col:
                st.warning("⚠️ 'Card No.' column not found in CapitalOne file. Person verification will be skipped.")
            
            # Clean data
            df = df.dropna(subset=[date_col, amount_col])
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
            df[amount_col] = pd.to_numeric(df[amount_col], errors='coerce')
            df = df.dropna(subset=[date_col, amount_col])
            
            # Filter only debit transactions
            df = df[df[amount_col] > 0]
            
            # Create result DataFrame
            result_data = {
                'date': df[date_col],
                'amount': df[amount_col],
                'description': df[desc_col] if desc_col else 'N/A',
                'source': 'CapitalOne'
            }
            
            # Add person information if card number exists
            if card_col:
                def get_person_name(card_no):
                    if pd.isna(card_no):
                        return 'Unknown'
                    card_str = str(int(float(card_no))) if isinstance(card_no, (int, float)) else str(card_no)
                    return self.card_to_person.get(card_str, f'Unknown Card {card_str}')
                
                result_data['person'] = df[card_col].apply(get_person_name)
                result_data['card_no'] = df[card_col].astype(str)
            else:
                result_data['person'] = 'N/A'
                result_data['card_no'] = 'N/A'
            
            result_df = pd.DataFrame(result_data)
            return result_df
            
        except Exception as e:
            st.error(f"❌ Error loading CapitalOne file: {e}")
            return None
    
    def load_jobber_data(self, uploaded_file):
        """Loads data from Jobber file"""
        try:
            df = pd.read_excel(uploaded_file)
            
            # Looking for data columns
            date_col = None
            amount_col = None
            item_col = None
            entered_by_col = None
            
            for col in df.columns:
                col_lower = str(col).lower()
                if 'date' in col_lower and 'transaction' not in col_lower:
                    date_col = col
                elif 'total' in col_lower and '$' in col_lower:
                    amount_col = col
                elif 'item' in col_lower and 'name' in col_lower:
                    item_col = col
                elif 'entered' in col_lower and 'by' in col_lower:
                    entered_by_col = col
            
            if not date_col or not amount_col:
                st.error("❌ Required columns not found in Jobber file. Expected columns Date and Total $.")
                return None
                
            if not entered_by_col:
                st.warning("⚠️ 'Entered by' column not found in Jobber file. Person verification will be skipped.")
            
            # Clean data
            df = df.dropna(subset=[date_col, amount_col])
            
            # Process dates in different formats
            if df[date_col].dtype == 'object':
                # Try different date formats
                def parse_date(date_str):
                    if pd.isna(date_str):
                        return pd.NaT
                    date_str = str(date_str).strip()
                    for fmt in ['%b %d, %Y', '%m/%d/%Y', '%Y-%m-%d', '%d.%m.%Y']:
                        try:
                            return datetime.strptime(date_str, fmt)
                        except:
                            continue
                    return pd.to_datetime(date_str, errors='coerce')
                
                df[date_col] = df[date_col].apply(parse_date)
            else:
                df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
            
            df[amount_col] = pd.to_numeric(df[amount_col], errors='coerce')
            df = df.dropna(subset=[date_col, amount_col])
            
            # Filter only positive amounts
            df = df[df[amount_col] > 0]
            
            # Create result DataFrame
            result_data = {
                'date': df[date_col],
                'amount': df[amount_col],
                'description': df[item_col] if item_col else 'N/A',
                'source': 'Jobber'
            }
            
            # Add person information
            if entered_by_col:
                result_data['person'] = df[entered_by_col].fillna('Unknown')
            else:
                result_data['person'] = 'N/A'
            
            result_df = pd.DataFrame(result_data)
            return result_df
            
        except Exception as e:
            st.error(f"❌ Error loading Jobber file: {e}")
            return None
    
    def find_matches(self, capitalone_df, jobber_df):
        """Finds matches between CapitalOne and Jobber data"""
        matches = []
        
        for cap_idx, cap_row in capitalone_df.iterrows():
            for job_idx, job_row in jobber_df.iterrows():
                # Calculate differences in date and amount
                date_diff = abs((cap_row['date'] - job_row['date']).days)
                amount_diff = abs(cap_row['amount'] - job_row['amount'])
                
                # Check tolerances for date and amount
                if date_diff <= self.tolerance_days and amount_diff <= self.tolerance_amount:
                    
                    # Check person correspondence (if enabled)
                    person_match = True
                    person_match_quality = 1.0
                    
                    if self.check_person and cap_row['person'] != 'N/A' and job_row['person'] != 'N/A':
                        cap_person = str(cap_row['person']).strip()
                        job_person = str(job_row['person']).strip()
                        
                        # Exact name match
                        if cap_person.lower() == job_person.lower():
                            person_match = True
                            person_match_quality = 1.0
                        # Check partial match (e.g., "Aaron Davidson" vs "Aaron")
                        elif cap_person.lower() in job_person.lower() or job_person.lower() in cap_person.lower():
                            person_match = True
                            person_match_quality = 0.8
                        else:
                            person_match = False
                            person_match_quality = 0.0
                    
                    # If all criteria are met, add match
                    if person_match or not self.check_person:
                        match_quality = self._calculate_match_quality(date_diff, amount_diff, person_match_quality)
                        
                        matches.append({
                            'cap_idx': cap_idx,
                            'job_idx': job_idx,
                            'cap_date': cap_row['date'],
                            'job_date': job_row['date'],
                            'cap_amount': cap_row['amount'],
                            'job_amount': job_row['amount'],
                            'cap_desc': cap_row['description'],
                            'job_desc': job_row['description'],
                            'cap_person': cap_row['person'],
                            'job_person': job_row['person'],
                            'cap_card_no': cap_row.get('card_no', 'N/A'),
                            'date_diff': date_diff,
                            'amount_diff': amount_diff,
                            'person_match': person_match,
                            'person_match_quality': person_match_quality,
                            'match_quality': match_quality
                        })
        
        # Sort by match quality
        matches.sort(key=lambda x: x['match_quality'], reverse=True)
        
        # Remove duplicates
        used_cap_idx = set()
        used_job_idx = set()
        final_matches = []
        
        for match in matches:
            if match['cap_idx'] not in used_cap_idx and match['job_idx'] not in used_job_idx:
                final_matches.append(match)
                used_cap_idx.add(match['cap_idx'])
                used_job_idx.add(match['job_idx'])
        
        return final_matches, used_cap_idx, used_job_idx
    
    def _calculate_match_quality(self, date_diff, amount_diff, person_match_quality=1.0):
        """Calculates match quality considering date, amount and person"""
        date_score = max(0, (self.tolerance_days - date_diff) / self.tolerance_days)
        amount_score = max(0, (self.tolerance_amount - amount_diff) / self.tolerance_amount)
        
        # If person checking is enabled, weight it 40%
        if self.check_person:
            return (date_score * 0.2 + amount_score * 0.4 + person_match_quality * 0.4)
        else:
            # Standard formula without person checking
            return (date_score * 0.3 + amount_score * 0.7)
    
    def generate_report_df(self, capitalone_df, jobber_df, matches, used_cap_idx, used_job_idx):
        """Generates report DataFrame"""
        report_data = []
        
        # Add matches
        for match in matches:
            person_status = "✅ Match" if match['person_match'] else "⚠️ Different Person"
            person_note = f"Cap: {match['cap_person']} | Job: {match['job_person']}"
            
            report_data.append({
                'Status': 'MATCH',
                'CapitalOne_Date': match['cap_date'].strftime('%Y-%m-%d'),
                'CapitalOne_Amount': match['cap_amount'],
                'CapitalOne_Person': match['cap_person'],
                'CapitalOne_Card': match['cap_card_no'],
                'CapitalOne_Description': match['cap_desc'],
                'Jobber_Date': match['job_date'].strftime('%Y-%m-%d'),
                'Jobber_Amount': match['job_amount'],
                'Jobber_Person': match['job_person'],
                'Jobber_Description': match['job_desc'],
                'Date_Diff_Days': match['date_diff'],
                'Amount_Diff_$': round(match['amount_diff'], 2),
                'Person_Check': person_status,
                'Match_Quality_%': round(match['match_quality'] * 100, 1),
                'Notes': f'Found matching transaction. {person_note}'
            })
        
        # CapitalOne transactions without matches
        for idx, row in capitalone_df.iterrows():
            if idx not in used_cap_idx:
                report_data.append({
                    'Status': 'MISSING_IN_JOBBER',
                    'CapitalOne_Date': row['date'].strftime('%Y-%m-%d'),
                    'CapitalOne_Amount': row['amount'],
                    'CapitalOne_Person': row['person'],
                    'CapitalOne_Card': row.get('card_no', 'N/A'),
                    'CapitalOne_Description': row['description'],
                    'Jobber_Date': '',
                    'Jobber_Amount': '',
                    'Jobber_Person': '',
                    'Jobber_Description': '',
                    'Date_Diff_Days': '',
                    'Amount_Diff_$': '',
                    'Person_Check': '',
                    'Match_Quality_%': '',
                    'Notes': f'Transaction by {row["person"]} not found in Jobber - may need to be added'
                })
        
        # Jobber expenses without matches
        for idx, row in jobber_df.iterrows():
            if idx not in used_job_idx:
                report_data.append({
                    'Status': 'MISSING_IN_CAPITALONE',
                    'CapitalOne_Date': '',
                    'CapitalOne_Amount': '',
                    'CapitalOne_Person': '',
                    'CapitalOne_Card': '',
                    'CapitalOne_Description': '',
                    'Jobber_Date': row['date'].strftime('%Y-%m-%d'),
                    'Jobber_Amount': row['amount'],
                    'Jobber_Person': row['person'],
                    'Jobber_Description': row['description'],
                    'Date_Diff_Days': '',
                    'Amount_Diff_$': '',
                    'Person_Check': '',
                    'Match_Quality_%': '',
                    'Notes': f'Expense by {row["person"]} not found in CapitalOne - verify entry'
                })
        
        return pd.DataFrame(report_data)

def main():
    # Application header
    st.title("💰 Expense Reconciliation Tool")
    st.markdown("### Automatic reconciliation of CapitalOne bank statements with Jobber expenses")
    
    # Sidebar with settings
    with st.sidebar:
        st.header("⚙️ Settings")
        
        tolerance_amount = st.slider(
            "💵 Allowable amount difference ($)",
            min_value=0.1,
            max_value=10.0,
            value=2.0,
            step=0.1,
            help="Transactions with amount difference up to this value are considered matching"
        )
        
        tolerance_days = st.slider(
            "📅 Allowable date difference (days)",
            min_value=0,
            max_value=7,
            value=2,
            step=1,
            help="Transactions with date difference up to this value are considered matching"
        )
        
        check_person = st.checkbox(
            "👤 Check person correspondence",
            value=True,
            help="If enabled, transactions will also be matched by person (CapitalOne card number vs Entered by in Jobber)"
        )
        
        if check_person:
            st.info("""
            **Person mapping:**
            - 9265 → Aaron Davidson
            - 4298 → Alex Masuda  
            - 1725 → Jericho Taylor-Daves
            - 3253 → Jerry Morales
            - 2984 → Antonio
            """)
        
        st.markdown("---")
        st.markdown("### 📋 Instructions")
        st.markdown("""
        1. **Upload** CapitalOne file (bank statement)
        2. **Upload** Jobber file (expenses)  
        3. **Adjust** tolerances if needed
        4. **Click** "Run reconciliation"
        5. **Download** Excel report
        """)
    
    # Main area
    col1, col2 = st.columns(2)
    
    with col1:
        st.header("🏦 Upload CapitalOne File")
        capitalone_file = st.file_uploader(
            "Select CapitalOne bank statement file (.xlsx)",
            type=['xlsx'],
            key='capitalone',
            help="File should contain columns with transaction date, debit amount and description"
        )
        
        if capitalone_file:
            st.success(f"✅ Loaded: {capitalone_file.name}")
    
    with col2:
        st.header("💼 Upload Jobber File")
        jobber_file = st.file_uploader(
            "Select Jobber expenses file (.xlsx)",
            type=['xlsx'],
            key='jobber',
            help="File should contain Date, Total $ and Item name columns"
        )
        
        if jobber_file:
            st.success(f"✅ Loaded: {jobber_file.name}")
    
    # Run reconciliation button
    if capitalone_file and jobber_file:
        st.markdown("---")
        
        if st.button("🚀 Run Reconciliation", type="primary", use_container_width=True):
            with st.spinner("🔄 Processing files and finding matches..."):
                # Create reconciler instance
                reconciler = ExpenseReconciler(tolerance_amount, tolerance_days, check_person)
                
                # Load data
                capitalone_df = reconciler.load_capitalone_data(capitalone_file)
                jobber_df = reconciler.load_jobber_data(jobber_file)
                
                if capitalone_df is not None and jobber_df is not None:
                    # Perform reconciliation
                    matches, used_cap_idx, used_job_idx = reconciler.find_matches(capitalone_df, jobber_df)
                    
                    # Generate report
                    report_df = reconciler.generate_report_df(capitalone_df, jobber_df, matches, used_cap_idx, used_job_idx)
                    
                    # Statistics
                    matched_count = len(matches)
                    cap_only_count = len(capitalone_df) - matched_count
                    job_only_count = len(jobber_df) - matched_count
                    match_percentage = (matched_count / max(len(capitalone_df), len(jobber_df))) * 100
                    
                    st.success("✅ Reconciliation completed successfully!")
                    
                    # Display metrics
                    st.markdown("---")
                    st.header("📊 Reconciliation Results")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.metric(
                            label="✅ Matches",
                            value=matched_count,
                            help="Found matches between CapitalOne and Jobber"
                        )
                    
                    with col2:
                        st.metric(
                            label="🟡 Only in CapitalOne",
                            value=cap_only_count,
                            help="Transactions that exist in bank but are missing in Jobber"
                        )
                    
                    with col3:
                        st.metric(
                            label="🔴 Only in Jobber",
                            value=job_only_count,
                            help="Expenses that exist in Jobber but are missing in bank"
                        )
                    
                    with col4:
                        st.metric(
                            label="📈 Match Percentage",
                            value=f"{match_percentage:.1f}%",
                            help="Percentage of found matches from total transactions"
                        )
                    
                    # Display results table
                    st.markdown("---")
                    st.header("📋 Detailed Report")
                    
                    # Table color coding
                    def style_dataframe(df):
                        def color_status(val):
                            if val == 'MATCH':
                                return 'background-color: #006E6E; color: white'  # Dark teal
                            elif val == 'MISSING_IN_JOBBER':
                                return 'background-color: #002A93; color: white'  # Dark blue
                            elif val == 'MISSING_IN_CAPITALONE':
                                return 'background-color: #933B00; color: white'  # Dark orange
                            return ''
                        
                        def color_person_check(val):
                            if val == '✅ Match':
                                return 'background-color: #90EE90; color: #2E7D32'  # Green
                            elif val == '⚠️ Different Person':
                                return 'background-color: #FFF9C4; color: #F57C00'  # Yellow
                            return ''
                        
                        styled = df.style.applymap(color_status, subset=['Status'])
                        if 'Person_Check' in df.columns:
                            styled = styled.applymap(color_person_check, subset=['Person_Check'])
                        return styled
                    
                    styled_df = style_dataframe(report_df)
                    st.dataframe(styled_df, use_container_width=True, height=600)
                    
                    # Legend
                    st.markdown("### 🎨 Color Coding:")
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.markdown("""
                        <div style='background-color: #006E6E; color: white; padding: 10px; border-radius: 5px; text-align: center;'>
                        <strong>🟢 MATCH</strong><br>
                        Found corresponding transactions
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown("""
                        <div style='background-color: #002A93; color: white; padding: 10px; border-radius: 5px; text-align: center;'>
                        <strong>🟡 NOT IN JOBBER</strong><br>
                        Transaction only in bank
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col3:
                        st.markdown("""
                        <div style='background-color: #933B00; color: white; padding: 10px; border-radius: 5px; text-align: center;'>
                        <strong>🔴 NOT IN CAPITALONE</strong><br>
                        Expense only in Jobber
                        </div>
                        """, unsafe_allow_html=True)
                    
                    # Download button
                    st.markdown("---")
                    
                    # Create Excel file for download
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        report_df.to_excel(writer, sheet_name='Reconciliation Report', index=False)
                        
                        # Add statistics on separate sheet
                        stats_df = pd.DataFrame({
                            'Metric': ['Total CapitalOne transactions', 'Total Jobber expenses', 'Found matches', 
                                       'Only in CapitalOne', 'Only in Jobber', 'Match percentage'],
                            'Value': [len(capitalone_df), len(jobber_df), matched_count, 
                                      cap_only_count, job_only_count, f"{match_percentage:.1f}%"]
                        })
                        stats_df.to_excel(writer, sheet_name='Statistics', index=False)
                    
                    excel_data = output.getvalue()
                    
                    st.download_button(
                        label="📥 Download Excel Report",
                        data=excel_data,
                        file_name=f"reconciliation_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )
                    
                    # Action recommendations
                    st.markdown("---")
                    st.header("💡 Action Recommendations")
                    
                    if cap_only_count > 0:
                        st.warning(f"🟡 **{cap_only_count} transactions found only in CapitalOne.** "
                                  "Recommend adding these expenses to Jobber.")
                    
                    if job_only_count > 0:
                        st.error(f"🔴 **{job_only_count} expenses found only in Jobber.** "
                                "Recommend verifying entry accuracy (amounts, dates).")
                    
                    if matched_count > 0:
                        st.success(f"✅ **{matched_count} matches found.** These transactions are correctly reconciled.")
                        
                        # Check person mismatches among matches
                        if check_person:
                            person_mismatches = sum(1 for match in matches if not match['person_match'])
                            if person_mismatches > 0:
                                st.warning(f"⚠️ **{person_mismatches} matches have different people.** "
                                          "Verify that transactions were performed with correct cards.")
                            else:
                                st.info("👤 **All matches have correct people.** Excellent work!")
                    
                    # Save results to session_state for reuse
                    st.session_state.last_report = report_df
                    st.session_state.last_stats = {
                        'matched': matched_count,
                        'cap_only': cap_only_count,
                        'job_only': job_only_count,
                        'percentage': match_percentage
                    }
    
    else:
        st.info("📁 Upload both files to start reconciliation")
        
        # Show file format examples
        st.markdown("---")
        st.header("📋 File Requirements")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("🏦 CapitalOne file should contain:")
            st.markdown("""
            **Required columns:**
            - **Transaction Date** - transaction date
            - **Debit** or **Amount** - transaction amount  
            
            **Recommended columns:**
            - **Card No.** - card number (for person verification)
            - **Description** - transaction description
            
            *Supported format: .xlsx*
            """)
        
        with col2:
            st.subheader("💼 Jobber file should contain:")
            st.markdown("""
            **Required columns:**
            - **Date** - expense date
            - **Total $** - expense amount
            
            **Recommended columns:**
            - **Entered by** - who entered the record (for person verification)
            - **Item name** - item name
            
            *Supported format: .xlsx*
            """)

    # Footer
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: #666;'>
    💰 Expense Reconciliation Tool | Created for automating expense reconciliation
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
