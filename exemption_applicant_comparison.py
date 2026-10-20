"""
Exemption Applicant Comparison Tool
====================================
For School District / Special District Parcel Tax Exemptions

Compares owner names between periods to determine if exemptions should
be maintained or removed based on ownership changes.

Logic:
- Compare Name (prior period) vs CurrentPeriodName (current period)
- Check if ApplicantName tokens appear in CurrentPeriodName
- Detect "heirs/estate" flags indicating deceased owners
- Generate recommend_keep decision

Author: Brandon Van Leuven
Version: 1.0
Date: January 2026
"""

import re
import pandas as pd
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Try to load similarity libraries
try:
    from rapidfuzz.fuzz import token_set_ratio
    HAVE_RAPIDFUZZ = True
except ImportError:
    token_set_ratio = None
    HAVE_RAPIDFUZZ = False

try:
    import jellyfish
    def jw_sim(a, b):
        return jellyfish.jaro_winkler_similarity(a, b)
    HAVE_JELLYFISH = True
except ImportError:
    from difflib import SequenceMatcher
    def jw_sim(a, b):
        return SequenceMatcher(None, a, b).ratio()
    HAVE_JELLYFISH = False

# Default similarity threshold
DEFAULT_SIMILARITY_THRESHOLD = 0.85

# Words to ignore when comparing names (noise words)
NOISE_WORDS = {
    "TRUST", "TRUSTEE", "TTEE", "TR", "TRS", "TRE", "ET", "AL", "ETAL",
    "JR", "SR", "II", "III", "IV", "V",
    "REV", "REVOCABLE", "LIVING", "FAMILY", "SURVIVOR", "SURVIVORS",
    "EST", "ESTATE", "HEIRS", "OF", "HEIR",
    "THE", "AND", "OR", "A", "AN",
    "LLC", "INC", "CO", "CORP", "LP", "LLP", "PLC", "COMPANY",
    "DECEASED", "DEC", "DECD"
}

# Words to protect from being stripped (meaningful for identification)
PROTECT_WORDS = {
    "SCHOOL", "DISTRICT", "FOUNDATION", "UNIVERSITY", "CHURCH", 
    "CITY", "COUNTY", "STATE", "FEDERAL"
}

# Patterns indicating deceased/estate (should recommend removal)
HEIRS_PATTERNS = [
    r'\bHEIRS\s+OF\b',
    r'\bEST\s+OF\b',
    r'\bESTATE\s+OF\b',
    r'\bEST\s*$',
    r'\bDECEASED\b',
    r'\bDEC\'?D\b',
    r'\bDECD\b'
]


class ExemptionApplicantComparison:
    """
    Process exemption files to compare owner names between periods
    and recommend whether exemptions should be kept or removed.
    """
    
    def __init__(self, input_file, output_file=None, similarity_threshold=None, 
                 prior_period_label=None, current_period_label=None):
        """
        Initialize the comparison processor.
        
        Args:
            input_file: Path to input Excel file
            output_file: Path for output file (auto-generated if None)
            similarity_threshold: Threshold for name matching (0.0-1.0)
            prior_period_label: Label for prior period (e.g., "202502")
            current_period_label: Label for current period (e.g., "202507")
        """
        self.input_file = input_file
        self.output_file = output_file or self._generate_output_filename()
        self.similarity_threshold = similarity_threshold or DEFAULT_SIMILARITY_THRESHOLD
        self.prior_period_label = prior_period_label or "Prior Period"
        self.current_period_label = current_period_label or "Current Period"
        self.data = None
        self.column_map = {}
        self.results = None
        
    def _generate_output_filename(self):
        """Generate output filename with timestamp."""
        import os
        base = os.path.splitext(os.path.basename(self.input_file))[0]
        timestamp = datetime.now().strftime("%Y-%m-%d-%H%M%S")
        return f"{base}_Comparison_Results_{timestamp}.xlsx"
    
    def _map_columns(self, df):
        """
        Map flexible column names to standardized internal names.
        Handles various naming conventions from different districts.
        """
        # Column variations for auto-detection
        column_variations = {
            'parcel': ['PARCEL', 'APN', 'PARCEL_NUM', 'PARCEL_NUMBER', 'PARCELNUMBER'],
            'status_check': ['STATUSCHECK', 'STATUS_CHECK', 'STATUS', 'MATCHSTATUS'],
            'status_out': ['STATUSOUT', 'STATUS_OUT'],
            'prior_period_apn': ['PRIORPERIODAPN', 'PRIOR_PERIOD_APN', 'LASTAPN', 'PRIOR_APN'],
            'prior_name': ['NAME', 'NAME1', 'OWNERNAME', 'OWNER_NAME', 'LASTYEARNAME', 
                          'LAST_YEAR_NAME', 'PRIORNAME', 'PRIOR_NAME'],
            'prior_name2': ['NAME2', 'OWNERNAME2', 'OWNER_NAME2'],
            'current_name': ['CURRENTPERIODNAME', 'CURRENT_PERIOD_NAME', 'CURRENTNAME', 
                            'CURRENT_NAME', 'CURRENTYEARNAME', 'CURRENT_YEAR_NAME'],
            'current_name2': ['CURRENTPERIODNAME2', 'CURRENT_PERIOD_NAME2', 'CURRENTNAME2'],
            'exemption_type': ['EXEMPTIONTYPE', 'EXEMPTION_TYPE', 'EXTYPE', 'TYPE'],
            'last_exemption_type': ['LASTYEAREXEMPTIONTYPE', 'LAST_YEAR_EXEMPTION_TYPE', 
                                   'PRIOREXEMPTIONTYPE', 'PRIOR_EXEMPTION_TYPE'],
            'applicant_phone': ['APPLICANTPHONENUMBER', 'APPLICANT_PHONE', 'PHONE'],
            'applicant_name': ['APPLICANTNAME', 'APPLICANT_NAME', 'APPLICANT'],
            'applicant_email': ['APPLICANTEMAIL', 'APPLICANT_EMAIL', 'EMAIL'],
            'date_entered': ['DATEENTERED', 'DATE_ENTERED', 'ENTRYDATE', 'ENTRY_DATE'],
            'docs_verified': ['DOCUMENTSVERIFIED', 'DOCUMENTS_VERIFIED', 'VERIFIED'],
            'app_format': ['APPLICATIONFORMAT', 'APPLICATION_FORMAT', 'FORMAT'],
            'username': ['USERNAME', 'USER_NAME', 'USER', 'ANALYST'],
            'note': ['NOTE', 'NOTES', 'COMMENTS', 'COMMENT'],
            'situs_address': ['SITUSADDRESS', 'SITUS_ADDRESS', 'SITUS', 'PROPERTY_ADDRESS'],
            'mail_address': ['MAILADDRESS', 'MAIL_ADDRESS', 'MAILING_ADDRESS', 'MAIL']
        }
        
        # Create uppercase column name mapping
        df_cols_upper = {col.upper().replace(' ', ''): col for col in df.columns}
        
        for standard_name, possible_names in column_variations.items():
            found = False
            for possible_name in possible_names:
                # Check exact match (case-insensitive, no spaces)
                possible_clean = possible_name.replace(' ', '').replace('_', '')
                for col_upper, col_orig in df_cols_upper.items():
                    col_clean = col_upper.replace('_', '')
                    if col_clean == possible_clean or col_upper == possible_name:
                        self.column_map[standard_name] = col_orig
                        found = True
                        break
                if found:
                    break
            
            if not found:
                # Column not found - will handle missing columns gracefully
                self.column_map[standard_name] = None
        
        return df
    
    @staticmethod
    def normalize_name_upper(s):
        """
        Normalize a name string for comparison.
        - Converts to uppercase
        - Replaces & with AND
        - Removes special characters (keeps letters, digits, spaces, commas)
        - Collapses multiple spaces
        """
        if pd.isna(s) or not str(s).strip():
            return ""
        s = str(s).upper().replace("&", " AND ")
        s = re.sub(r"[^A-Z0-9 ,]+", " ", s)
        s = re.sub(r"\s+", " ", s).strip()
        return s
    
    @staticmethod
    def strip_noise_tokens(s):
        """Remove noise words from a name string."""
        toks = [t for t in re.split(r"[ ,]+", s) if t]
        out = []
        for t in toks:
            if t in PROTECT_WORDS:
                out.append(t)
            elif t not in NOISE_WORDS:
                out.append(t)
        return " ".join(out)
    
    @staticmethod
    def canonical_token_string(s):
        """
        Create a canonical representation of name tokens.
        - Removes noise words
        - Removes single-character tokens
        - Sorts alphabetically
        - Deduplicates
        """
        toks = [t for t in re.split(r"[ ,]+", s) if t]
        toks = [t for t in toks if t in PROTECT_WORDS or t not in NOISE_WORDS]
        toks = [t for t in toks if len(t) > 1]
        return " ".join(sorted(set(toks)))
    
    def best_name_similarity(self, a, b):
        """
        Calculate the best similarity score between two names.
        Uses both Jaro-Winkler and token set ratio, returns the maximum.
        """
        a0 = self.normalize_name_upper(a)
        b0 = self.normalize_name_upper(b)
        a1 = self.strip_noise_tokens(a0)
        b1 = self.strip_noise_tokens(b0)
        a_can = self.canonical_token_string(a1)
        b_can = self.canonical_token_string(b1)
        
        # Handle empty strings
        if not a1 or not b1:
            return 0.0
        
        # Jaro-Winkler similarity
        jw = jw_sim(a1, b1)
        
        # Token set ratio (order-insensitive)
        if HAVE_RAPIDFUZZ and token_set_ratio is not None:
            ts = token_set_ratio(a_can, b_can) / 100.0
        else:
            # Fallback: Jaccard-like similarity
            aset = set(a_can.split()) if a_can else set()
            bset = set(b_can.split()) if b_can else set()
            if aset | bset:
                ts = len(aset & bset) / len(aset | bset)
            else:
                ts = 0.0
        
        return max(jw, ts)
    
    def applicant_tokens_in_current(self, applicant, current):
        """
        Check if the applicant's name tokens appear in the current owner name.
        Returns True if all significant applicant tokens are found in current.
        """
        a = self.canonical_token_string(self.strip_noise_tokens(self.normalize_name_upper(applicant)))
        c = self.canonical_token_string(self.strip_noise_tokens(self.normalize_name_upper(current)))
        
        if not a:
            return False
        
        atoks = set(a.split())
        ctoks = set(c.split())
        
        # Check if applicant tokens are subset of current tokens
        return atoks.issubset(ctoks)
    
    def has_heirs_flag(self, name):
        """
        Check if the name contains indicators of deceased owner / estate.
        """
        s = self.normalize_name_upper(name)
        for pattern in HEIRS_PATTERNS:
            if re.search(pattern, s):
                return True
        return False
    
    def check_name_token_overlap(self, prior, current):
        """
        Check if there's significant name token overlap between prior and current.
        Returns tuple of (has_overlap, overlap_ratio, matching_tokens, is_surname_match).
        """
        p_norm = self.normalize_name_upper(prior)
        c_norm = self.normalize_name_upper(current)
        p_clean = self.strip_noise_tokens(p_norm)
        c_clean = self.strip_noise_tokens(c_norm)
        
        # Get token sets (excluding single chars)
        p_toks = set(t for t in p_clean.split() if len(t) > 1)
        c_toks = set(t for t in c_clean.split() if len(t) > 1)
        
        if not p_toks or not c_toks:
            return (False, 0.0, set(), False)
        
        # Find matching tokens
        matching = p_toks & c_toks
        
        # Calculate overlap ratio based on current tokens found in prior
        # (handles case where current has fewer names due to spouse death)
        overlap_ratio = len(matching) / len(p_toks) if p_toks else 0.0
        
        # Check for surname match (first token is often surname in "LASTNAME FIRSTNAME" format)
        # Also check if any multi-syllable token matches (likely surname)
        p_list = list(p_toks)
        c_list = list(c_toks)
        
        # Consider tokens with 4+ chars as potential surnames
        p_surnames = [t for t in p_toks if len(t) >= 4]
        c_surnames = [t for t in c_toks if len(t) >= 4]
        surname_match = bool(set(p_surnames) & set(c_surnames))
        
        return (len(matching) > 0, overlap_ratio, matching, surname_match)
    
    def decide_recommend_keep(self, row):
        """
        Determine the recommendation for keeping or removing the exemption.
        
        Logic:
        1. If heirs/estate flag present → Remove (deceased owner)
        2. If applicant name found in current owner → Keep (applicant still owns)
        3. If applicant NOT in current, but there's strong name overlap between
           prior and current → Keep (likely surviving spouse scenario)
        4. If surname matches and there's meaningful token overlap → Keep
        5. If at least 50% of prior tokens persist → Keep
        6. Fallback: Use owner similarity match
        """
        # Always remove if heirs/estate language present in current name
        if row.get('heirs_flag', False):
            return "Remove"
        
        # Get key metrics
        has_overlap = row.get('name_token_overlap', False)
        overlap_ratio = row.get('name_overlap_ratio', 0.0)
        surname_match = row.get('surname_match', False)
        sim_score = row.get('sim_score', 0)
        applicant_in_current = row.get('applicant_in_current', False)
        
        # Check if Applicant Name exists
        applicant_norm = row.get('applicant_name_normalized', '')
        applicant_clean = self.canonical_token_string(
            self.strip_noise_tokens(applicant_norm if applicant_norm else '')
        )
        
        if applicant_clean:
            # Applicant name exists
            if applicant_in_current:
                # Best case: applicant is still in current owner
                return "Keep"
            else:
                # Applicant not found in current - but check if this is a surviving spouse scenario
                # If there's strong overlap between prior and current names, the exemption
                # likely transferred to a surviving spouse who was on the original title
                if surname_match and has_overlap and overlap_ratio >= 0.5:
                    return "Keep"  # Likely surviving spouse
                if has_overlap and overlap_ratio >= 0.66:  # 2/3 of prior names still present
                    return "Keep"
                if sim_score >= 0.85:  # High similarity even without applicant
                    return "Keep"
                # Applicant truly gone and no strong continuity
                return "Remove"
        
        # No applicant name - use name matching logic
        # If surname matches and we have at least one other token matching, Keep
        if surname_match and has_overlap:
            if overlap_ratio >= 0.3 or sim_score >= 0.60:
                return "Keep"
        
        # If there's significant overlap (at least 50% of prior tokens found)
        if has_overlap and overlap_ratio >= 0.5:
            return "Keep"
        
        # If similarity is high enough based on threshold
        if row.get('owner_match', False):
            return "Keep"
        
        # Last resort: if similarity is borderline (0.60-0.85) and there's any token overlap
        if has_overlap and 0.60 <= sim_score < self.similarity_threshold:
            return "Keep"
        
        return "Remove"
    
    def get_confidence_level(self, row):
        """
        Assign a confidence level to the recommendation.
        """
        sim_score = row.get('sim_score', 0)
        heirs_flag = row.get('heirs_flag', False)
        applicant_in_current = row.get('applicant_in_current', False)
        owner_match = row.get('owner_match', False)
        recommend = row.get('recommend_keep', '')
        
        # High confidence removal cases
        if heirs_flag:
            return "HIGH"
        
        # Perfect or near-perfect match
        if sim_score >= 0.95:
            return "HIGH"
        
        # Clear mismatch (low similarity, no applicant match)
        if sim_score < 0.5 and not applicant_in_current:
            return "HIGH"
        
        # Good match
        if owner_match and sim_score >= 0.85:
            return "MEDIUM"
        
        # Borderline cases need review
        if 0.7 <= sim_score < 0.85:
            return "LOW - REVIEW"
        
        return "MEDIUM"
    
    def load_data(self):
        """Load and prepare the input data."""
        print(f"Loading data from: {self.input_file}")
        
        # Read Excel file
        self.data = pd.read_excel(self.input_file)
        print(f"  Loaded {len(self.data):,} records")
        print(f"  Columns found: {list(self.data.columns)}")
        
        # Map columns
        self._map_columns(self.data)
        
        # Report column mapping
        print("\n  Column mapping:")
        for std_name, orig_col in self.column_map.items():
            if orig_col:
                print(f"    {std_name} → {orig_col}")
        
        return self.data
    
    def process_comparison(self):
        """
        Process the comparison between prior and current period names.
        """
        if self.data is None:
            raise ValueError("Data not loaded. Call load_data() first.")
        
        print("\nProcessing name comparisons...")
        
        # Get column names (use original column names from mapping)
        parcel_col = self.column_map.get('parcel')
        status_check_col = self.column_map.get('status_check')
        prior_name_col = self.column_map.get('prior_name')
        current_name_col = self.column_map.get('current_name')
        applicant_name_col = self.column_map.get('applicant_name')
        
        # Create results dataframe
        results = self.data.copy()
        
        # Extract relevant columns with safe access
        def safe_get(row, col):
            if col and col in row.index:
                return row[col]
            return ""
        
        # Normalize names for output
        results['prior_name_normalized'] = results.apply(
            lambda r: self.normalize_name_upper(safe_get(r, prior_name_col)), axis=1
        )
        results['current_name_normalized'] = results.apply(
            lambda r: self.normalize_name_upper(safe_get(r, current_name_col)), axis=1
        )
        results['applicant_name_normalized'] = results.apply(
            lambda r: self.normalize_name_upper(safe_get(r, applicant_name_col)), axis=1
        )
        
        # Calculate similarity scores
        print("  Calculating similarity scores...")
        results['sim_score'] = results.apply(
            lambda r: self.best_name_similarity(
                safe_get(r, prior_name_col),
                safe_get(r, current_name_col)
            ), axis=1
        )
        
        # Owner match based on threshold
        results['owner_match'] = results['sim_score'] >= self.similarity_threshold
        
        # Check if applicant name tokens appear in current name
        print("  Checking applicant name containment...")
        results['applicant_in_current'] = results.apply(
            lambda r: self.applicant_tokens_in_current(
                safe_get(r, applicant_name_col),
                safe_get(r, current_name_col)
            ), axis=1
        )
        
        # Check for heirs/estate flags in current name
        print("  Checking for heirs/estate flags...")
        results['heirs_flag'] = results.apply(
            lambda r: self.has_heirs_flag(safe_get(r, current_name_col)), axis=1
        )
        
        # Calculate name token overlap (for partial name persistence)
        print("  Calculating name token overlap...")
        def calc_overlap(row):
            return self.check_name_token_overlap(
                safe_get(row, prior_name_col),
                safe_get(row, current_name_col)
            )
        
        overlap_results = results.apply(calc_overlap, axis=1)
        results['name_token_overlap'] = overlap_results.apply(lambda x: x[0])
        results['name_overlap_ratio'] = overlap_results.apply(lambda x: x[1])
        results['matching_tokens'] = overlap_results.apply(lambda x: ', '.join(x[2]) if x[2] else '')
        results['surname_match'] = overlap_results.apply(lambda x: x[3])
        
        # Generate recommendations
        print("  Generating recommendations...")
        results['recommend_keep'] = results.apply(self.decide_recommend_keep, axis=1)
        
        # Add confidence level
        results['confidence_level'] = results.apply(self.get_confidence_level, axis=1)
        
        # Add reason for recommendation
        def get_reason(row):
            if row['heirs_flag']:
                return "Estate/Heirs language detected - likely deceased"
            
            # Get key metrics
            has_overlap = row.get('name_token_overlap', False)
            overlap_ratio = row.get('name_overlap_ratio', 0.0)
            surname_match = row.get('surname_match', False)
            sim_score = row.get('sim_score', 0)
            matching = row.get('matching_tokens', '')
            applicant_in_current = row.get('applicant_in_current', False)
            
            if row['applicant_name_normalized']:
                if applicant_in_current:
                    return "Applicant name found in current owner"
                elif row['recommend_keep'] == 'Keep':
                    return f"Surviving spouse likely ({matching}) - applicant may be deceased"
                else:
                    return "Applicant name NOT found in current owner, no family continuity"
            
            if row['recommend_keep'] == 'Keep':
                if surname_match and has_overlap:
                    return f"Surname match + tokens persist ({matching}) - likely same family"
                if has_overlap and overlap_ratio >= 0.5:
                    return f"Name tokens persist ({matching}) - {overlap_ratio:.0%} overlap"
                if row['owner_match']:
                    return f"Name similarity {sim_score:.1%} >= {self.similarity_threshold:.0%} threshold"
                if has_overlap:
                    return f"Token overlap ({matching}) + {sim_score:.1%} similarity"
            
            return f"Name similarity {sim_score:.1%} - no significant name continuity"
        
        results['reason'] = results.apply(get_reason, axis=1)
        
        self.results = results
        
        # Summary statistics
        keep_count = (results['recommend_keep'] == 'Keep').sum()
        remove_count = (results['recommend_keep'] == 'Remove').sum()
        heirs_count = results['heirs_flag'].sum()
        
        print(f"\n  Summary:")
        print(f"    Total records: {len(results):,}")
        print(f"    Recommend Keep: {keep_count:,}")
        print(f"    Recommend Remove: {remove_count:,}")
        print(f"    Heirs/Estate flagged: {heirs_count:,}")
        
        return results
    
    def save_results(self):
        """
        Save results to Excel with formatting and multiple tabs.
        """
        if self.results is None:
            raise ValueError("No results to save. Call process_comparison() first.")
        
        print(f"\nSaving results to: {self.output_file}")
        
        # Prepare output columns
        parcel_col = self.column_map.get('parcel')
        status_check_col = self.column_map.get('status_check')
        prior_name_col = self.column_map.get('prior_name')
        current_name_col = self.column_map.get('current_name')
        applicant_name_col = self.column_map.get('applicant_name')
        situs_col = self.column_map.get('situs_address')
        
        # Build output column order
        output_cols = []
        col_labels = {}
        
        # Key identification columns
        if parcel_col:
            output_cols.append(parcel_col)
            col_labels[parcel_col] = 'Parcel'
        
        if status_check_col:
            output_cols.append(status_check_col)
            col_labels[status_check_col] = 'Status Check'
        
        # Name columns - side by side for easy comparison
        if prior_name_col:
            output_cols.append(prior_name_col)
            col_labels[prior_name_col] = f'{self.prior_period_label} Name'
        
        if current_name_col:
            output_cols.append(current_name_col)
            col_labels[current_name_col] = f'{self.current_period_label} Name'
        
        if applicant_name_col:
            output_cols.append(applicant_name_col)
            col_labels[applicant_name_col] = 'Applicant Name'
        
        # Analysis columns
        output_cols.extend([
            'sim_score', 'owner_match', 'matching_tokens', 'applicant_in_current', 
            'heirs_flag', 'recommend_keep', 'confidence_level', 'reason'
        ])
        col_labels.update({
            'sim_score': 'Similarity Score',
            'owner_match': 'Name Match',
            'matching_tokens': 'Matching Tokens',
            'applicant_in_current': 'Applicant in Current',
            'heirs_flag': 'Heirs/Estate Flag',
            'recommend_keep': 'Recommendation',
            'confidence_level': 'Confidence',
            'reason': 'Reason'
        })
        
        # Add situs if available
        if situs_col:
            output_cols.append(situs_col)
            col_labels[situs_col] = 'Situs Address'
        
        # Filter to existing columns
        output_cols = [c for c in output_cols if c in self.results.columns]
        
        # Create output dataframe
        output_df = self.results[output_cols].copy()
        output_df = output_df.rename(columns=col_labels)
        
        # Create subsets
        keep_df = output_df[self.results['recommend_keep'] == 'Keep'].copy()
        remove_df = output_df[self.results['recommend_keep'] == 'Remove'].copy()
        review_df = output_df[self.results['confidence_level'].str.contains('REVIEW', na=False)].copy()
        heirs_df = output_df[self.results['heirs_flag'] == True].copy()
        
        # Write to Excel with formatting
        with pd.ExcelWriter(self.output_file, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Define formats
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#2c3e50',
                'font_color': 'white',
                'border': 1,
                'text_wrap': True,
                'valign': 'vcenter',
                'align': 'center'
            })
            
            keep_format = workbook.add_format({
                'bg_color': '#d4edda',
                'font_color': '#155724'
            })
            
            remove_format = workbook.add_format({
                'bg_color': '#f8d7da',
                'font_color': '#721c24'
            })
            
            review_format = workbook.add_format({
                'bg_color': '#fff3cd',
                'font_color': '#856404'
            })
            
            heirs_format = workbook.add_format({
                'bg_color': '#f5c6cb',
                'font_color': '#721c24',
                'bold': True
            })
            
            high_sim_format = workbook.add_format({
                'bg_color': '#c3e6cb',
                'num_format': '0.0%'
            })
            
            low_sim_format = workbook.add_format({
                'bg_color': '#f5c6cb',
                'num_format': '0.0%'
            })
            
            medium_sim_format = workbook.add_format({
                'bg_color': '#ffeeba',
                'num_format': '0.0%'
            })
            
            percent_format = workbook.add_format({'num_format': '0.0%'})
            
            def write_sheet(df, sheet_name, highlight_col=None):
                """Write a dataframe to a sheet with formatting."""
                df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1, header=False)
                worksheet = writer.sheets[sheet_name]
                
                # Write headers
                for col_idx, col_name in enumerate(df.columns):
                    worksheet.write(0, col_idx, col_name, header_format)
                
                # Set column widths
                col_widths = {
                    'Parcel': 18,
                    'Status Check': 12,
                    f'{self.prior_period_label} Name': 35,
                    f'{self.current_period_label} Name': 35,
                    'Applicant Name': 30,
                    'Similarity Score': 14,
                    'Name Match': 12,
                    'Applicant in Current': 16,
                    'Heirs/Estate Flag': 14,
                    'Recommendation': 14,
                    'Confidence': 14,
                    'Reason': 45,
                    'Situs Address': 35
                }
                
                for col_idx, col_name in enumerate(df.columns):
                    width = col_widths.get(col_name, 15)
                    worksheet.set_column(col_idx, col_idx, width)
                
                # Apply conditional formatting for Recommendation column
                if 'Recommendation' in df.columns:
                    rec_col_idx = list(df.columns).index('Recommendation')
                    worksheet.conditional_format(
                        1, rec_col_idx, len(df) + 1, rec_col_idx,
                        {'type': 'text', 'criteria': 'containing', 'value': 'Keep',
                         'format': keep_format}
                    )
                    worksheet.conditional_format(
                        1, rec_col_idx, len(df) + 1, rec_col_idx,
                        {'type': 'text', 'criteria': 'containing', 'value': 'Remove',
                         'format': remove_format}
                    )
                
                # Apply conditional formatting for Similarity Score column
                if 'Similarity Score' in df.columns:
                    sim_col_idx = list(df.columns).index('Similarity Score')
                    worksheet.conditional_format(
                        1, sim_col_idx, len(df) + 1, sim_col_idx,
                        {'type': 'cell', 'criteria': '>=', 'value': 0.85,
                         'format': high_sim_format}
                    )
                    worksheet.conditional_format(
                        1, sim_col_idx, len(df) + 1, sim_col_idx,
                        {'type': 'cell', 'criteria': '<', 'value': 0.7,
                         'format': low_sim_format}
                    )
                    worksheet.conditional_format(
                        1, sim_col_idx, len(df) + 1, sim_col_idx,
                        {'type': 'cell', 'criteria': 'between', 'minimum': 0.7, 'maximum': 0.849,
                         'format': medium_sim_format}
                    )
                
                # Apply conditional formatting for Heirs/Estate Flag column
                if 'Heirs/Estate Flag' in df.columns:
                    heirs_col_idx = list(df.columns).index('Heirs/Estate Flag')
                    worksheet.conditional_format(
                        1, heirs_col_idx, len(df) + 1, heirs_col_idx,
                        {'type': 'text', 'criteria': 'containing', 'value': 'True',
                         'format': heirs_format}
                    )
                
                # Apply conditional formatting for Confidence column
                if 'Confidence' in df.columns:
                    conf_col_idx = list(df.columns).index('Confidence')
                    worksheet.conditional_format(
                        1, conf_col_idx, len(df) + 1, conf_col_idx,
                        {'type': 'text', 'criteria': 'containing', 'value': 'REVIEW',
                         'format': review_format}
                    )
                
                # Freeze top row
                worksheet.freeze_panes(1, 0)
                
                # Add autofilter
                if len(df) > 0:
                    worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
            
            # Write all sheets
            print("  Creating 'All Records' tab...")
            write_sheet(output_df, 'All Records')
            
            print("  Creating 'Recommend Keep' tab...")
            write_sheet(keep_df, 'Recommend Keep')
            
            print("  Creating 'Recommend Remove' tab...")
            write_sheet(remove_df, 'Recommend Remove')
            
            if len(review_df) > 0:
                print("  Creating 'Needs Review' tab...")
                write_sheet(review_df, 'Needs Review')
            
            if len(heirs_df) > 0:
                print("  Creating 'Heirs-Estate Flagged' tab...")
                write_sheet(heirs_df, 'Heirs-Estate Flagged')
            
            # Create Summary tab
            print("  Creating 'Summary' tab...")
            summary_data = {
                'Metric': [
                    'Total Records',
                    'Recommend Keep',
                    'Recommend Remove',
                    'Needs Review (Low Confidence)',
                    'Heirs/Estate Flagged',
                    '',
                    'Similarity Threshold Used',
                    'Prior Period Label',
                    'Current Period Label',
                    '',
                    'Processing Date',
                    'Input File'
                ],
                'Value': [
                    len(output_df),
                    len(keep_df),
                    len(remove_df),
                    len(review_df),
                    len(heirs_df),
                    '',
                    f'{self.similarity_threshold:.0%}',
                    self.prior_period_label,
                    self.current_period_label,
                    '',
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    self.input_file
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False, startrow=1, header=False)
            
            summary_ws = writer.sheets['Summary']
            summary_ws.write(0, 0, 'Metric', header_format)
            summary_ws.write(0, 1, 'Value', header_format)
            summary_ws.set_column(0, 0, 35)
            summary_ws.set_column(1, 1, 50)
            
            # Highlight key metrics
            summary_ws.write(1, 1, len(output_df), workbook.add_format({'bold': True, 'num_format': '#,##0'}))
            summary_ws.write(2, 1, len(keep_df), workbook.add_format({'bold': True, 'bg_color': '#d4edda', 'num_format': '#,##0'}))
            summary_ws.write(3, 1, len(remove_df), workbook.add_format({'bold': True, 'bg_color': '#f8d7da', 'num_format': '#,##0'}))
            summary_ws.write(4, 1, len(review_df), workbook.add_format({'bold': True, 'bg_color': '#fff3cd', 'num_format': '#,##0'}))
            summary_ws.write(5, 1, len(heirs_df), workbook.add_format({'bold': True, 'bg_color': '#f5c6cb', 'num_format': '#,##0'}))
        
        print(f"\n  ✓ Output saved to: {self.output_file}")
        
        return self.output_file


def main(input_file, output_file=None, similarity_threshold=None,
         prior_period_label=None, current_period_label=None):
    """
    Main function to run the exemption applicant comparison.
    """
    processor = ExemptionApplicantComparison(
        input_file=input_file,
        output_file=output_file,
        similarity_threshold=similarity_threshold,
        prior_period_label=prior_period_label,
        current_period_label=current_period_label
    )
    
    print("=" * 70)
    print("EXEMPTION APPLICANT COMPARISON TOOL")
    print("Ownership-change review for parcel tax exemptions")
    print("=" * 70)
    print()
    
    # Check for optional libraries
    if HAVE_JELLYFISH:
        print("✓ Using jellyfish for Jaro-Winkler similarity")
    else:
        print("→ Using difflib fallback for similarity (install jellyfish for better results)")
    
    if HAVE_RAPIDFUZZ:
        print("✓ Using rapidfuzz for token-based matching")
    else:
        print("→ Using basic token matching (install rapidfuzz for better results)")
    
    print()
    
    processor.load_data()
    processor.process_comparison()
    output_path = processor.save_results()
    
    print()
    print("=" * 70)
    print("✓✓✓ PROCESSING COMPLETE ✓✓✓")
    print("=" * 70)
    
    return output_path, processor


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python exemption_applicant_comparison.py <input_file> [output_file] [threshold]")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    threshold = float(sys.argv[3]) if len(sys.argv) > 3 else None
    
    main(input_file, output_file, threshold)
